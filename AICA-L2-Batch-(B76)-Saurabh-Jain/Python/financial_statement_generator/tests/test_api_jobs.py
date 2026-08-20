"""Tests for job-based FastAPI upload/generate/status/download/validation endpoints."""

from __future__ import annotations

import time
from io import BytesIO

from fastapi.testclient import TestClient

from app.config import settings
from app.main import app

TEXT_PDF = b"""%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>
endobj
4 0 obj
<< /Length 168 >>
stream
BT
/F1 12 Tf
72 720 Td
(BALANCE SHEET as at 31-03-2024) Tj
0 -20 Td
(Share Capital 1,00,000) Tj
0 -20 Td
(Revenue from operations 250000.50) Tj
ET
endstream
endobj
5 0 obj
<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>
endobj
xref
0 6
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000266 00000 n 
0000000485 00000 n 
trailer
<< /Size 6 /Root 1 0 R >>
startxref
563
%%EOF
"""


def _client(tmp_path, monkeypatch: object) -> TestClient:
    monkeypatch.setattr(settings, "output_dir", tmp_path)
    monkeypatch.setattr(settings, "input_dir", tmp_path / "input")
    settings.ensure_directories()
    return TestClient(app)


def _wait_for_job(client: TestClient, job_id: str, timeout: float = 60.0) -> dict:
    deadline = time.monotonic() + timeout
    last = None
    while time.monotonic() < deadline:
        response = client.get(f"/status/{job_id}")
        if response.status_code == 200:
            last = response.json()
            if last.get("status") in {"completed", "failed", "review_required"}:
                return last
        time.sleep(0.1)
    raise AssertionError(f"job {job_id} did not finish: {last}")


def test_upload_runs_pipeline_and_returns_job_id(tmp_path, monkeypatch) -> None:
    client = _client(tmp_path, monkeypatch)
    response = client.post("/upload", files={"file": ("statement.pdf", BytesIO(TEXT_PDF), "application/pdf")})
    assert response.status_code == 200, response.text
    body = response.json()
    assert "job_id" in body
    assert body["status"] in {"processing", "review_required", "completed", "failed"}
    finished = _wait_for_job(client, body["job_id"])
    assert finished["status"] in {"review_required", "completed", "failed"}
    assert "traceback" not in response.text.lower()
    assert "Traceback" not in response.text


def test_generate_status_download_and_validation(tmp_path, monkeypatch) -> None:
    client = _client(tmp_path, monkeypatch)
    uploaded = client.post("/upload", files={"file": ("statement.pdf", BytesIO(TEXT_PDF), "application/pdf")})
    assert uploaded.status_code == 200, uploaded.text
    job_id = uploaded.json()["job_id"]
    finished = _wait_for_job(client, job_id)
    assert finished["status"] == "review_required"

    reviewed = client.post(f"/review/{job_id}", json={})
    assert reviewed.status_code == 200, reviewed.text
    items = [{"item_id": item["item_id"], "status": "approved"} for item in reviewed.json().get("items") or []]
    if items:
        saved = client.post(f"/review/{job_id}", json={"items": items})
        assert saved.status_code == 200, saved.text
    approved = client.post(f"/approve/{job_id}")
    assert approved.status_code == 200, approved.text
    assert approved.json()["status"] == "completed"

    generated = client.post("/generate", params={"job_id": job_id})
    assert generated.status_code == 200, generated.text
    payload = generated.json()
    assert payload == {
        "job_id": job_id,
        "status": "completed",
        "output_file": payload["output_file"],
        "validation_status": payload["validation_status"],
    }
    assert payload["output_file"]
    assert payload["validation_status"] in {"PASS", "WARNING", "ERROR"}
    assert "traceback" not in generated.text.lower()

    status = client.get(f"/status/{job_id}")
    assert status.status_code == 200
    assert status.json()["job_id"] == job_id
    assert status.json()["status"] == "completed"

    validation = client.get(f"/validation/{job_id}")
    assert validation.status_code == 200
    report = validation.json()
    assert report["job_id"] == job_id
    assert "validation" in report
    assert report["validation"]["status"] in {"PASS", "WARNING", "ERROR"}
    assert "checks" in report["validation"]

    download = client.get(f"/download/{job_id}")
    assert download.status_code == 200
    assert download.headers["content-type"].startswith("application/vnd.openxmlformats")
    assert download.content[:2] == b"PK"


def test_generate_with_file_returns_job_payload(tmp_path, monkeypatch) -> None:
    client = _client(tmp_path, monkeypatch)
    response = client.post("/generate", files={"file": ("statement.pdf", BytesIO(TEXT_PDF), "application/pdf")})
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["status"] == "review_required"
    assert body["job_id"]
    assert body.get("output_file") in {None, ""}


def test_invalid_upload_returns_failed_job_without_stack_trace(tmp_path, monkeypatch) -> None:
    client = _client(tmp_path, monkeypatch)
    response = client.post("/upload", files={"file": ("notes.txt", BytesIO(b"hello world"), "text/plain")})
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["job_id"]
    assert body["status"] in {"processing", "failed"}
    finished = _wait_for_job(client, body["job_id"])
    assert finished["status"] == "failed"
    assert finished["errors"]
    assert all("message" in item for item in finished["errors"])
    assert "traceback" not in response.text.lower()
    assert "Traceback" not in response.text

    status = client.get(f"/status/{body['job_id']}")
    assert status.json()["status"] == "failed"
    assert "traceback" not in status.text.lower()
    download = client.get(f"/download/{body['job_id']}")
    assert download.status_code == 409
    assert "traceback" not in download.text.lower()


def test_generate_without_file_or_job_id_is_bad_request(tmp_path, monkeypatch) -> None:
    client = _client(tmp_path, monkeypatch)
    response = client.post("/generate")
    assert response.status_code == 400, response.text
    body = response.json()
    assert body["error"] == "missing_input"
    assert "traceback" not in response.text.lower()


def test_generate_invalid_file_returns_failed_payload(tmp_path, monkeypatch) -> None:
    client = _client(tmp_path, monkeypatch)
    response = client.post("/generate", files={"file": ("notes.txt", BytesIO(b"hello world"), "text/plain")})
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["status"] == "failed"
    assert body["job_id"]
    assert isinstance(body["errors"], list)
    assert isinstance(body["warnings"], list)
    assert body["errors"]
    assert "output_file" not in body
    assert "traceback" not in response.text.lower()


def test_unknown_job_is_not_found_without_internals(tmp_path, monkeypatch) -> None:
    client = _client(tmp_path, monkeypatch)
    response = client.get("/status/does-not-exist")
    assert response.status_code == 404
    body = response.json()
    assert body["error"] == "not_found"
    assert "traceback" not in response.text.lower()
    assert "src/" not in response.text
