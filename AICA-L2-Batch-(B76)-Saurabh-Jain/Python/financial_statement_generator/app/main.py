"""FastAPI application entry point."""

from __future__ import annotations

import logging
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.exceptions import HTTPException as StarletteHTTPException

from app.api.routes import router
from app.config import settings
from app.extraction.exceptions import PDFExtractionError

STATIC_DIR = Path(__file__).resolve().parent / "static"

logger = logging.getLogger("financial_statement_generator")

logging.basicConfig(
    level=logging.DEBUG if settings.debug else logging.INFO,
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
)

app = FastAPI(
    title=settings.app_name,
    description=(
        "Converts uploaded financial statement documents into Excel workbooks "
        "following ICAI Division I – Non-Ind AS Schedule III requirements."
    ),
    version=settings.app_version,
)

app.include_router(router)


@app.get("/", include_in_schema=False)
def frontend() -> FileResponse:
    """Serve the generator user interface."""
    return FileResponse(
        STATIC_DIR / "index.html",
        headers={"Cache-Control": "no-cache, no-store, must-revalidate"},
    )


@app.get("/api")
def api_info() -> dict[str, str]:
    """Machine-readable service information."""
    return {
        "name": settings.app_name,
        "version": settings.app_version,
        "status": "running",
        "docs": "/docs",
        "health": "/health",
        "upload": "/upload",
        "generate": "/generate",
        "status_job": "/status/{job_id}",
        "download": "/download/{job_id}",
        "validation": "/validation/{job_id}",
        "review": "/review/{job_id}",
        "approve": "/approve/{job_id}",
    }


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.exception_handler(PDFExtractionError)
async def pdf_extraction_error_handler(_request: Request, exc: PDFExtractionError) -> JSONResponse:
    logger.warning("Upload extraction error: %s", exc.message)
    return JSONResponse(status_code=exc.status_code, content=exc.to_dict())


@app.exception_handler(StarletteHTTPException)
async def http_exception_handler(_request: Request, exc: StarletteHTTPException) -> JSONResponse:
    detail = exc.detail
    if isinstance(detail, dict):
        payload = {key: value for key, value in detail.items() if key not in {"traceback", "exception", "stack"}}
    else:
        payload = {"error": "http_error", "message": str(detail)}
    return JSONResponse(status_code=exc.status_code, content=payload)


@app.exception_handler(RequestValidationError)
async def request_validation_handler(_request: Request, exc: RequestValidationError) -> JSONResponse:
    logger.info("Request validation failed: %s", exc.errors())
    return JSONResponse(
        status_code=422,
        content={"error": "invalid_request", "message": "The request was not valid."},
    )


@app.exception_handler(Exception)
async def unhandled_error_handler(_request: Request, exc: Exception) -> JSONResponse:
    logger.exception("Unhandled API error: %s", type(exc).__name__)
    return JSONResponse(
        status_code=500,
        content={"error": "internal_error", "message": "An unexpected error occurred."},
    )


