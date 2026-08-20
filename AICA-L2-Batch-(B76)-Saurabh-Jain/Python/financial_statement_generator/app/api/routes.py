"""HTTP routes for the Financial Statement Generator API."""

from __future__ import annotations

import logging
from pathlib import Path

from fastapi import APIRouter, Body, File, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse

from app.api.jobs import JobStore
from app.api.pipeline import approve_job, run_pipeline, save_review, start_pipeline
from app.classification.document_classifier import DocumentClassifier
from app.config import settings
from app.extraction.pdf_extractor import PDFExtractor
from app.mapping.schedule_iii_mapper import ScheduleIIIMapper
from app.review.review_service import ReviewIncompleteError
from app.validation.financial_validator import FinancialValidator

logger = logging.getLogger(__name__)

router = APIRouter()


@router.get("/health")
def health() -> dict[str, str]:
    """Liveness check used to confirm the API process is running."""
    return {
        "status": "ok",
        "service": settings.app_name,
        "version": settings.app_version,
    }


@router.post("/upload")
async def upload_statement(file: UploadFile = File(...)) -> dict:
    """Store a trial balance (.xlsx) or PDF and run the generation pipeline."""
    filename = file.filename or "upload.pdf"
    content = await file.read()
    logger.info("Received upload %s (%s bytes)", filename, len(content))
    job = start_pipeline(filename, content)
    return job.upload_response()


@router.post("/generate")
async def generate_workbook(
    file: UploadFile | None = File(default=None),
    job_id: str | None = Query(default=None),
) -> dict:
    """Return a generated workbook job. Accepts a PDF or an existing job_id."""
    if job_id:
        job = _require_job(job_id)
        logger.info("Generate requested for existing job %s", job_id)
        return job.generate_response()
    if file is None:
        raise HTTPException(
            status_code=400,
            detail={"error": "missing_input", "message": "Upload a trial balance (.xlsx) or PDF, or provide job_id."},
        )
    filename = file.filename or "upload.pdf"
    content = await file.read()
    logger.info("Generate requested for upload %s (%s bytes)", filename, len(content))
    job = run_pipeline(filename, content)
    return job.generate_response()


@router.get("/status/{job_id}")
def job_status(job_id: str) -> dict:
    """Return processing status for a job."""
    job = _require_job(job_id)
    logger.info("Status requested for job %s (%s)", job_id, job.status)
    return job.status_response()


@router.get("/download/{job_id}")
def download_workbook(job_id: str) -> FileResponse:
    """Download the generated Excel workbook for a completed job."""
    job = _require_job(job_id)
    if job.status != "completed" or not job.output_file:
        raise HTTPException(
            status_code=409,
            detail={
                "error": "not_ready",
                "message": "The workbook is not available for this job.",
                "job_id": job_id,
                "status": job.status,
            },
        )
    path = JobStore().output_path(job_id)
    if not path.exists() and job.output_file:
        candidate = Path(job.output_file)
        path = candidate if candidate.is_absolute() else settings.output_dir / candidate
    if not path.exists():
        logger.error("Job %s marked completed but output file is missing", job_id)
        raise HTTPException(
            status_code=404,
            detail={"error": "not_found", "message": "Generated workbook was not found."},
        )
    logger.info("Download requested for job %s", job_id)
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Financial_Statements_Generated.xlsx",
    )


@router.get("/validation/{job_id}")
def job_validation(job_id: str) -> dict:
    """Return the financial validation report for a job."""
    job = _require_job(job_id)
    logger.info("Validation requested for job %s (%s)", job_id, job.validation_status)
    return job.validation_response()


@router.post("/review/{job_id}")
def review_job(job_id: str, payload: dict | None = Body(default=None)) -> dict:
    """Return or update the human-review table for a mapped job."""
    _require_job(job_id)
    logger.info("Review requested for job %s", job_id)
    try:
        job, review = save_review(job_id, payload if isinstance(payload, dict) else None)
    except KeyError:
        raise HTTPException(status_code=404, detail={"error": "not_found", "message": "Job was not found."})
    except ValueError as exc:
        raise HTTPException(status_code=409, detail={"error": "review_unavailable", "message": str(exc)})
    return job.review_response(review)


@router.post("/approve/{job_id}")
def approve_reviewed_job(job_id: str) -> dict:
    """Write only approved mappings into Excel, then validate."""
    _require_job(job_id)
    logger.info("Approve requested for job %s", job_id)
    try:
        job = approve_job(job_id)
    except KeyError:
        raise HTTPException(status_code=404, detail={"error": "not_found", "message": "Job was not found."})
    except ReviewIncompleteError as exc:
        raise HTTPException(
            status_code=409,
            detail={
                "error": "review_incomplete",
                "message": exc.message,
                "needs_review": exc.needs_review,
            },
        )
    except ValueError as exc:
        raise HTTPException(status_code=409, detail={"error": "not_ready", "message": str(exc)})
    return job.generate_response()


@router.post("/extract")
async def extract_pdf(file: UploadFile = File(...)) -> dict:
    """Upload a financial statement PDF and return structured extraction data."""
    filename = file.filename or "upload.pdf"
    content = await file.read()
    extractor = PDFExtractor()
    return extractor.extract_upload(filename, content)


@router.post("/classify")
async def classify_pdf(file: UploadFile = File(...)) -> dict:
    """Upload a PDF, extract pages, then classify sourced financial data."""
    filename = file.filename or "upload.pdf"
    content = await file.read()
    extracted = PDFExtractor().extract_upload(filename, content)
    classified = DocumentClassifier().classify(extracted)
    return {
        "document": extracted.get("document"),
        "classified": classified.model_dump_classified(),
    }


@router.post("/map")
async def map_pdf(file: UploadFile = File(...)) -> dict:
    """Upload a PDF, classify it, and map values onto the Excel template."""
    filename = file.filename or "upload.pdf"
    content = await file.read()
    extracted = PDFExtractor().extract_upload(filename, content)
    classified = DocumentClassifier().classify(extracted)
    mapped = ScheduleIIIMapper().map_classified(classified)
    return {
        "document": extracted.get("document"),
        "report": mapped.get("report"),
        "placements": mapped.get("placements"),
        "unmapped_sources": mapped.get("unmapped_sources"),
        "warnings": mapped.get("warnings"),
    }


@router.post("/validate")
async def validate_pdf(file: UploadFile = File(...)) -> dict:
    """Upload a PDF and return financial validation without changing values."""
    filename = file.filename or "upload.pdf"
    content = await file.read()
    extracted = PDFExtractor().extract_upload(filename, content)
    classified = DocumentClassifier().classify(extracted)
    mapped = ScheduleIIIMapper().map_classified(classified)
    validation = FinancialValidator().validate(mapped, classified)
    return {
        "document": extracted.get("document"),
        "validation": validation,
    }


def _require_job(job_id: str):
    job = JobStore().get(job_id)
    if job is None:
        raise HTTPException(
            status_code=404,
            detail={"error": "not_found", "message": "Job was not found."},
        )
    return job
