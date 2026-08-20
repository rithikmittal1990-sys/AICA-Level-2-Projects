"""Tests for the generator web UI served by FastAPI."""

from fastapi.testclient import TestClient

from app.main import app


def test_homepage_serves_generator_ui() -> None:
    client = TestClient(app)
    response = client.get("/")
    assert response.status_code == 200
    assert "text/html" in response.headers["content-type"]
    body = response.text
    assert "ICAI Schedule III Financial Statement Generator" in body
    assert "Upload Trial Balance or Financial Statement" in body
    assert "Choose a trial balance (.xlsx) or PDF" in body
    assert "Generate Excel" in body
    assert "Uploading" in body
    assert "Reading source file" in body
    assert "Identifying financial statements" in body
    assert "Mapping to Schedule III" in body
    assert "Review mappings" in body
    assert "Generating Excel" in body
    assert "Source Field" in body
    assert "Extracted Value" in body
    assert "Human review" in body
    assert "Validating workbook" in body
    assert "Validation summary" in body
    assert "Download Excel" in body
    assert "Errors and warnings" in body


def test_static_assets_are_served() -> None:
    client = TestClient(app)
    css = client.get("/static/styles.css")
    js = client.get("/static/app.js")
    assert css.status_code == 200
    assert js.status_code == 200
    assert "/upload" in js.text
    assert "/review/" in js.text
    assert "/approve/" in js.text
    assert "/download/" in js.text
