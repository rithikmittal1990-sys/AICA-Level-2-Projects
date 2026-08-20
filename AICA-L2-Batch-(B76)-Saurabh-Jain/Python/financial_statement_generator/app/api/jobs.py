"""Persist pipeline jobs without exposing internals to API clients."""

from __future__ import annotations

import json
import logging
import time
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from app.config import settings

logger = logging.getLogger(__name__)

JOB_FILENAME = "job.json"
OUTPUT_FILENAME = "Financial_Statements_Generated.xlsx"
MAPPED_FILENAME = "mapped.json"
CLASSIFIED_FILENAME = "classified.json"
REVIEW_FILENAME = "review.json"
TRIAL_BALANCE_CLASSIFICATION_FILENAME = "trial_balance_classification.json"


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _read_json(path: Path) -> dict[str, Any] | None:
    """Read a job record, retrying briefly if a write is still in flight."""
    last_error: Exception | None = None
    for _ in range(5):
        try:
            text = path.read_text(encoding="utf-8")
            if not text.strip():
                last_error = json.JSONDecodeError("empty", text, 0)
            else:
                payload = json.loads(text)
                return payload if isinstance(payload, dict) else None
        except (OSError, json.JSONDecodeError) as exc:
            last_error = exc
        time.sleep(0.02)
    logger.warning("Failed to read job record %s: %s", path, last_error)
    return None


def public_issue(code: str, message: str) -> dict[str, str]:
    """User-facing error/warning object. Never includes a traceback."""
    return {"code": str(code), "message": str(message)}


def sanitize_issue(item: Any) -> dict[str, str] | None:
    if item is None:
        return None
    if isinstance(item, str):
        text = item.strip()
        return public_issue("warning", text) if text else None
    if isinstance(item, dict):
        message = item.get("message") or item.get("detail") or item.get("error")
        code = item.get("code") or item.get("id") or "processing_error"
        if not message:
            return None
        return public_issue(str(code), str(message))
    return public_issue("processing_error", str(item))


@dataclass
class JobRecord:
    job_id: str
    status: str = "pending"
    stage: str = "created"
    original_filename: str = "upload.pdf"
    stored_pdf: str | None = None
    output_file: str | None = None
    validation_status: str | None = None
    validation: dict[str, Any] | None = None
    workbook_validation: dict[str, Any] | None = None
    errors: list[dict[str, str]] = field(default_factory=list)
    warnings: list[dict[str, str]] = field(default_factory=list)
    review_summary: dict[str, Any] | None = None
    financial_year: dict[str, Any] | None = None
    created_at: str = field(default_factory=_utc_now)
    updated_at: str = field(default_factory=_utc_now)

    def to_dict(self) -> dict[str, Any]:
        return {
            "job_id": self.job_id,
            "status": self.status,
            "stage": self.stage,
            "original_filename": self.original_filename,
            "stored_pdf": self.stored_pdf,
            "output_file": self.output_file,
            "validation_status": self.validation_status,
            "validation": self.validation,
            "workbook_validation": self.workbook_validation,
            "errors": list(self.errors),
            "warnings": list(self.warnings),
            "review_summary": self.review_summary,
            "financial_year": self.financial_year,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
        }

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> JobRecord:
        return cls(
            job_id=str(payload.get("job_id") or ""),
            status=str(payload.get("status") or "pending"),
            stage=str(payload.get("stage") or "created"),
            original_filename=str(payload.get("original_filename") or "upload.pdf"),
            stored_pdf=payload.get("stored_pdf"),
            output_file=payload.get("output_file"),
            validation_status=payload.get("validation_status"),
            validation=payload.get("validation") if isinstance(payload.get("validation"), dict) else None,
            workbook_validation=(
                payload.get("workbook_validation")
                if isinstance(payload.get("workbook_validation"), dict)
                else None
            ),
            errors=[item for item in (payload.get("errors") or []) if isinstance(item, dict)],
            warnings=[item for item in (payload.get("warnings") or []) if isinstance(item, dict)],
            review_summary=payload.get("review_summary") if isinstance(payload.get("review_summary"), dict) else None,
            financial_year=payload.get("financial_year") if isinstance(payload.get("financial_year"), dict) else None,
            created_at=str(payload.get("created_at") or _utc_now()),
            updated_at=str(payload.get("updated_at") or _utc_now()),
        )

    def upload_response(self) -> dict[str, Any]:
        payload: dict[str, Any] = {"job_id": self.job_id, "status": self.status}
        if self.status == "failed":
            payload["errors"] = list(self.errors)
            payload["warnings"] = list(self.warnings)
        return payload

    def generate_response(self) -> dict[str, Any]:
        if self.status == "failed":
            return {
                "job_id": self.job_id,
                "status": "failed",
                "errors": list(self.errors),
                "warnings": list(self.warnings),
            }
        payload = {
            "job_id": self.job_id,
            "status": self.status,
            "output_file": self.output_file,
            "validation_status": self.validation_status,
        }
        if self.status == "review_required":
            payload["review_summary"] = self.review_summary
        return payload

    def status_response(self) -> dict[str, Any]:
        return {
            "job_id": self.job_id,
            "status": self.status,
            "stage": self.stage,
            "output_file": self.output_file,
            "validation_status": self.validation_status,
            "review_summary": self.review_summary,
            "financial_year": self.financial_year,
            "errors": list(self.errors),
            "warnings": list(self.warnings),
        }

    def review_response(self, review: dict[str, Any] | None) -> dict[str, Any]:
        payload = review or {"items": [], "summary": {}, "financial_year": self.financial_year}
        return {
            "job_id": self.job_id,
            "status": self.status,
            "stage": self.stage,
            "threshold": payload.get("threshold"),
            "financial_year": payload.get("financial_year") or self.financial_year,
            "summary": payload.get("summary") or self.review_summary,
            "items": payload.get("items") or [],
            "unmapped_sources": payload.get("unmapped_sources") or [],
        }

    def validation_response(self) -> dict[str, Any]:
        return {
            "job_id": self.job_id,
            "status": self.status,
            "validation_status": self.validation_status,
            "validation": self.validation
            or {"status": self.validation_status, "checks": [], "warnings": [], "errors": []},
        }


class JobStore:
    """JSON job records under ``output/jobs/{job_id}/``."""

    def __init__(self, root: Path | None = None) -> None:
        self.root = Path(root) if root else settings.jobs_dir
        self.root.mkdir(parents=True, exist_ok=True)

    def create(self, original_filename: str) -> JobRecord:
        job_id = uuid.uuid4().hex
        job = JobRecord(job_id=job_id, original_filename=original_filename, status="pending")
        self.job_dir(job_id).mkdir(parents=True, exist_ok=True)
        self.save(job)
        logger.info("Created job %s for %s", job_id, original_filename)
        return job

    def job_dir(self, job_id: str) -> Path:
        return self.root / job_id

    def get(self, job_id: str) -> JobRecord | None:
        path = self.job_dir(job_id) / JOB_FILENAME
        if not path.exists():
            return None
        payload = _read_json(path)
        if not isinstance(payload, dict):
            return None
        return JobRecord.from_dict(payload)

    def save(self, job: JobRecord) -> None:
        job.updated_at = _utc_now()
        directory = self.job_dir(job.job_id)
        directory.mkdir(parents=True, exist_ok=True)
        path = directory / JOB_FILENAME
        temporary = directory / f".{JOB_FILENAME}.tmp"
        temporary.write_text(json.dumps(job.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8")
        for _ in range(10):
            try:
                temporary.replace(path)
                break
            except PermissionError:
                time.sleep(0.05)

    def output_path(self, job_id: str) -> Path:
        return self.job_dir(job_id) / OUTPUT_FILENAME

    def save_artifact(self, job_id: str, filename: str, payload: dict[str, Any]) -> None:
        directory = self.job_dir(job_id)
        directory.mkdir(parents=True, exist_ok=True)
        path = directory / filename
        temporary = directory / f".{filename}.tmp"
        temporary.write_text(json.dumps(payload, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
        for _ in range(10):
            try:
                temporary.replace(path)
                break
            except PermissionError:
                time.sleep(0.05)

    def load_artifact(self, job_id: str, filename: str) -> dict[str, Any] | None:
        path = self.job_dir(job_id) / filename
        if not path.exists():
            return None
        payload = _read_json(path)
        return payload if isinstance(payload, dict) else None
