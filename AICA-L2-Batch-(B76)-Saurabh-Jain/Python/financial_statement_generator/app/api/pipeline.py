"""Run extract → classify → map, then wait for review before Excel generation."""

from __future__ import annotations

import logging
import threading
from pathlib import Path
from typing import Any

from app.api.jobs import (
    CLASSIFIED_FILENAME,
    MAPPED_FILENAME,
    OUTPUT_FILENAME,
    REVIEW_FILENAME,
    TRIAL_BALANCE_CLASSIFICATION_FILENAME,
    JobRecord,
    JobStore,
    public_issue,
    sanitize_issue,
)
from app.classification.trial_balance_classifier import TrialBalanceClassifier
from app.config import settings
from app.excel.workbook_generator import WorkbookGenerator
from app.extraction.exceptions import PDFExtractionError, TrialBalanceExtractionError
from app.extraction.pdf_extractor import PDFExtractor
from app.extraction.trial_balance_reader import TrialBalanceReader, is_trial_balance_upload
from app.mapping.schedule_iii_mapper import ScheduleIIIMapper
from app.trial_balance.generator import (
    TrialBalanceStatementGenerator,
    classification_to_legacy_mapped,
)
from app.trial_balance.parser import TrialBalanceParser
from app.trial_balance.classifier import TrialBalanceAccountClassifier
from app.review.review_service import (
    ReviewIncompleteError,
    apply_review_updates,
    approved_mapped,
    build_review,
)
from app.validation.financial_validator import FinancialValidator

logger = logging.getLogger(__name__)


def run_pipeline(filename: str, content: bytes, *, store: JobStore | None = None) -> JobRecord:
    """Store the PDF and run extraction and mapping, pausing for human review."""
    job_store = store or JobStore()
    job = job_store.create(filename)
    _execute(job, job_store, filename, content)
    return job


def start_pipeline(filename: str, content: bytes, *, store: JobStore | None = None) -> JobRecord:
    """Create a job and process it in a background thread so clients can poll status."""
    job_store = store or JobStore()
    job = job_store.create(filename)
    job.status = "processing"
    job.stage = "uploaded"
    job_store.save(job)
    thread = threading.Thread(
        target=_execute,
        args=(job, job_store, filename, content),
        daemon=True,
        name=f"fs-job-{job.job_id[:8]}",
    )
    thread.start()
    logger.info("Job %s started in background for %s", job.job_id, filename)
    return job


def save_review(job_id: str, payload: dict | None, *, store: JobStore | None = None) -> tuple[JobRecord, dict]:
    """Apply reviewer edits and persist the review table."""
    job_store = store or JobStore()
    job = job_store.get(job_id)
    if job is None:
        raise KeyError(job_id)
    review = job_store.load_artifact(job_id, REVIEW_FILENAME)
    if review is None:
        mapped = job_store.load_artifact(job_id, MAPPED_FILENAME)
        classified = job_store.load_artifact(job_id, CLASSIFIED_FILENAME) or {}
        if mapped is None:
            raise ValueError("Review is not available for this job.")
        review = build_review(mapped, classified)
    has_updates = bool(payload) and (payload.get("items") or payload.get("financial_year"))
    if has_updates and job.status != "review_required":
        raise ValueError("Review can only be updated while the job is waiting for review.")
    updated = apply_review_updates(review, payload if job.status == "review_required" else None)
    if job.status == "review_required":
        job_store.save_artifact(job_id, REVIEW_FILENAME, updated)
        job.review_summary = updated.get("summary")
        job.financial_year = updated.get("financial_year")
        job_store.save(job)
        logger.info(
            "Job %s review saved approved=%s needs_review=%s",
            job_id,
            (updated.get("summary") or {}).get("approved"),
            (updated.get("summary") or {}).get("needs_review"),
        )
    return job, updated


def approve_job(job_id: str, *, store: JobStore | None = None) -> JobRecord:
    """Generate Excel from approved mappings only, then validate."""
    job_store = store or JobStore()
    job = job_store.get(job_id)
    if job is None:
        raise KeyError(job_id)
    if job.status == "completed":
        return job
    if job.status == "failed":
        return job
    if job.status != "review_required":
        raise ValueError("This job is not waiting for review.")
    review = job_store.load_artifact(job_id, REVIEW_FILENAME)
    mapped = job_store.load_artifact(job_id, MAPPED_FILENAME)
    classified = job_store.load_artifact(job_id, CLASSIFIED_FILENAME) or {}
    if review is None or mapped is None:
        raise ValueError("Review data was not found for this job.")
    approved = approved_mapped(mapped, review)
    job_dir = job_store.job_dir(job.job_id)
    try:
        _finalize_workbook(job, job_store, job_dir, approved, classified)
    except ReviewIncompleteError:
        raise
    except Exception:
        logger.exception("Job %s failed during approved generation", job.job_id)
        _fail(job, job_store, "processing_error", "Document processing failed.")
    return job


def _execute(job: JobRecord, store: JobStore, filename: str, content: bytes) -> None:
    job_dir = store.job_dir(job.job_id)
    try:
        _process(job, store, job_dir, filename, content)
    except PDFExtractionError as exc:
        logger.warning("Job %s failed during extraction: %s", job.job_id, exc.message)
        _fail(job, store, exc.error_code, exc.message)
    except TrialBalanceExtractionError as exc:
        logger.warning("Job %s failed during trial balance read: %s", job.job_id, exc.message)
        _fail(job, store, exc.error_code, exc.message)
    except Exception:
        logger.exception("Job %s failed with an unexpected error", job.job_id)
        _fail(job, store, "processing_error", "Document processing failed.")


def _process(job: JobRecord, store: JobStore, job_dir: Path, filename: str, content: bytes) -> None:
    job.status = "processing"
    job.stage = "uploaded"
    store.save(job)
    logger.info("Job %s stored upload %s (%s bytes)", job.job_id, filename, len(content))

    if is_trial_balance_upload(filename):
        classified_model, mapped, review = _process_trial_balance(job, store, job_dir, filename, content)
    else:
        classified_model, mapped, review = _process_pdf(job, store, job_dir, filename, content)

    store.save_artifact(job.job_id, MAPPED_FILENAME, mapped)
    store.save_artifact(job.job_id, CLASSIFIED_FILENAME, classified_model.model_dump_classified())
    store.save_artifact(job.job_id, REVIEW_FILENAME, review)
    job.review_summary = review.get("summary")
    job.financial_year = review.get("financial_year")
    job.stage = "review"
    job.status = "review_required"
    store.save(job)
    logger.info(
        "Job %s waiting for review items=%s needs_review=%s",
        job.job_id,
        (review.get("summary") or {}).get("total"),
        (review.get("summary") or {}).get("needs_review"),
    )


def _process_pdf(
    job: JobRecord,
    store: JobStore,
    job_dir: Path,
    filename: str,
    content: bytes,
) -> tuple[Any, dict, dict]:
    from app.classification.document_classifier import DocumentClassifier

    job.stage = "extracting"
    store.save(job)
    extractor = PDFExtractor(upload_dir=job_dir)
    extracted = extractor.extract_upload(filename, content)
    stored_path = extracted.get("document", {}).get("stored_path")
    job.stored_pdf = str(stored_path) if stored_path else None
    logger.info("Job %s extracted %s page(s)", job.job_id, extracted.get("document", {}).get("pages"))

    job.stage = "classifying"
    store.save(job)
    classified_model = DocumentClassifier().classify(extracted)
    logger.info("Job %s classified %s section(s)", job.job_id, len(classified_model.sections))

    job.stage = "mapping"
    store.save(job)
    mapped = ScheduleIIIMapper().map_classified(classified_model)
    for warning in mapped.get("warnings") or []:
        issue = sanitize_issue(warning)
        if issue:
            job.warnings.append(issue)
    logger.info("Job %s mapped %s placement(s)", job.job_id, len(mapped.get("placements") or []))
    review = build_review(mapped, classified_model.model_dump_classified())
    return classified_model, mapped, review


def _process_trial_balance(
    job: JobRecord,
    store: JobStore,
    job_dir: Path,
    filename: str,
    content: bytes,
) -> tuple[Any, dict, dict]:
    job.stage = "extracting"
    store.save(job)
    parser = TrialBalanceParser(upload_dir=job_dir)
    parsed = parser.read_upload(filename, content)
    document = TrialBalanceReader(upload_dir=job_dir).read_path(Path(parsed.stored_path))
    job.stored_pdf = parsed.stored_path
    logger.info("Job %s read %s leaf trial balance account(s)", job.job_id, len(document.rows))

    job.stage = "classifying"
    store.save(job)
    classification = TrialBalanceAccountClassifier().classify(parsed)
    classified_model = TrialBalanceClassifier().classify(document)
    logger.info(
        "Job %s classified %s balance sheet and %s P&L line(s)",
        job.job_id,
        len(classified_model.balance_sheet.line_items),
        len(classified_model.profit_and_loss.line_items),
    )

    job.stage = "mapping"
    store.save(job)
    mapped = classification_to_legacy_mapped(classification)
    store.save_artifact(job.job_id, TRIAL_BALANCE_CLASSIFICATION_FILENAME, {
        "company_name": classification.company_name,
        "period_label": classification.period_label,
        "financial_year": classification.financial_year,
        "line_items": classification.line_items,
        "totals": classification.totals.to_dict(),
        "accounts": [item.to_dict() for item in classification.mapped],
    })
    for item in classification.review_accounts():
        issue = sanitize_issue(
            f"Review required: {item.account.account_name} — {item.reason or 'unusual balance'}"
        )
        if issue:
            job.warnings.append(issue)
    for item in classification.unmapped_accounts():
        issue = sanitize_issue(f"Unmapped account: {item.account.account_name}")
        if issue:
            job.warnings.append(issue)
    for warning in parsed.warnings:
        issue = sanitize_issue(warning)
        if issue:
            job.warnings.append(issue)
    logger.info("Job %s mapped %s placement(s)", job.job_id, len(mapped.get("placements") or []))
    review = build_review(mapped, classified_model.model_dump_classified())
    return classified_model, mapped, review


def _finalize_workbook(
    job: JobRecord,
    store: JobStore,
    job_dir: Path,
    mapped: dict,
    classified: dict,
) -> None:
    job.status = "processing"
    job.stage = "generating"
    store.save(job)
    if mapped.get("generation_mode") == "trial_balance" and job.stored_pdf:
        generator = TrialBalanceStatementGenerator(output_dir=job_dir)
        result = generator.generate_from_path(
            job.stored_pdf,
            output_filename=OUTPUT_FILENAME,
        )
        generated_path = result.path
        try:
            job.output_file = str(generated_path.resolve().relative_to(Path(settings.output_dir).resolve()))
        except ValueError:
            job.output_file = str(generated_path)
        job.workbook_validation = {
            "ok": all(c.status != "FAIL" for c in result.validation.checks),
            "checks": result.validation.to_dict(),
        }
        job.validation = {
            "status": "PASS" if job.workbook_validation["ok"] else "WARNING",
            "checks": [c.name for c in result.validation.checks],
            "warnings": [c.detail for c in result.validation.checks if c.status == "WARNING"],
            "errors": [c.detail for c in result.validation.checks if c.status == "FAIL"],
        }
        job.validation_status = job.validation["status"]
        job.stage = "completed"
        job.status = "completed"
        store.save(job)
        logger.info(
            "Job %s completed TB-only output=%s written=%s",
            job.job_id,
            job.output_file,
            result.written_count,
        )
        return

    generated = WorkbookGenerator(output_dir=job_dir).generate_detailed(mapped)
    try:
        job.output_file = str(generated.path.resolve().relative_to(Path(settings.output_dir).resolve()))
    except ValueError:
        job.output_file = str(generated.path)
    if generated.validation.warnings:
        for warning in generated.validation.warnings:
            issue = sanitize_issue(warning)
            if issue:
                job.warnings.append(issue)
    if not generated.validation.ok:
        for error in generated.validation.errors:
            issue = sanitize_issue({"code": "workbook_validation", "message": error})
            if issue:
                job.errors.append(issue)
        logger.warning("Job %s workbook structure validation failed", job.job_id)

    job.stage = "validating"
    store.save(job)
    financial = FinancialValidator().validate(mapped, classified)
    job.validation = {
        "status": financial.get("status"),
        "checks": financial.get("checks") or [],
        "warnings": financial.get("warnings") or [],
        "errors": financial.get("errors") or [],
    }
    job.workbook_validation = generated.validation.to_dict()
    job.validation_status = str(financial.get("status") or "WARNING")
    if not generated.validation.ok:
        job.validation_status = "ERROR"
    for item in financial.get("warnings") or []:
        issue = sanitize_issue(item)
        if issue:
            job.warnings.append(issue)

    job.stage = "completed"
    job.status = "completed"
    store.save(job)
    logger.info(
        "Job %s completed output=%s validation=%s approved_writes=%s",
        job.job_id,
        job.output_file,
        job.validation_status,
        len(mapped.get("placements") or []),
    )


def _fail(job: JobRecord, store: JobStore, code: str, message: str) -> None:
    job.status = "failed"
    job.errors.append(public_issue(code, message))
    store.save(job)
    logger.info("Job %s marked failed at stage %s: %s", job.job_id, job.stage, message)
