"""Flask route tests using the test client."""

from __future__ import annotations

import io

import pytest
from docx import Document

from app import create_app


@pytest.fixture
def client(test_config, service, monkeypatch):
    """Build a Flask app, inject the in-memory translation service, return a test client."""
    # Bypass the real GCS/Gemini wiring by patching the factory's builders to no-ops.
    monkeypatch.setattr("app._build_storage", lambda cfg: None)
    monkeypatch.setattr("app._build_provider", lambda cfg: None)
    app = create_app(test_config)
    app.config["TRANSLATION_SERVICE"] = service
    app.config["TESTING"] = True
    with app.test_client() as c:
        yield c


def _docx_bytes(paragraphs: list[str]) -> bytes:
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def test_healthz(client):
    resp = client.get("/healthz")
    assert resp.status_code == 200
    assert resp.get_json() == {"status": "ok"}


def test_api_health_reports_service_status(client):
    resp = client.get("/api/v1/health")
    body = resp.get_json()
    assert resp.status_code == 200
    assert body["status"] == "ok"
    assert ".docx" in body["supported_extensions"]


def test_api_languages(client):
    resp = client.get("/api/v1/languages")
    assert resp.status_code == 200
    assert "English" in resp.get_json()["languages"]


def test_api_translation_requires_file(client):
    resp = client.post("/api/v1/translations", data={"target_language": "Spanish"})
    assert resp.status_code == 400


def test_api_translation_rejects_unknown_extension(client):
    data = {
        "file": (io.BytesIO(b"hello"), "doc.pdf"),
        "target_language": "Spanish",
    }
    resp = client.post("/api/v1/translations", data=data, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert "unsupported" in resp.get_json()["error"].lower()


def test_api_translation_happy_path(client):
    data = {
        "file": (io.BytesIO(_docx_bytes(["hi"])), "doc.docx"),
        "target_language": "Spanish",
    }
    resp = client.post("/api/v1/translations", data=data, content_type="multipart/form-data")
    assert resp.status_code == 201
    body = resp.get_json()
    assert body["download_filename"] == "translated_doc.docx"
    assert body["download_url"].startswith("https://signed.example/")
