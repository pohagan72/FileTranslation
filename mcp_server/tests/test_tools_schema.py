"""Schema-level tests — exercise Pydantic validation without needing the SDK."""

from __future__ import annotations

import base64

import pytest
from pydantic import ValidationError

from filetranslation_mcp.tools import (
    Base64Source,
    DocumentSourceAdapter,
    GcsKeySource,
    HttpUrlSource,
    LocalPathSource,
    TranslateInput,
)


def test_path_source_round_trips():
    payload = TranslateInput.model_validate(
        {
            "source": {"kind": "path", "path": "/tmp/x.docx"},
            "target_language": "Spanish",
        }
    )
    assert isinstance(payload.source, LocalPathSource)
    assert payload.source.path == "/tmp/x.docx"


def test_url_source_round_trips():
    payload = TranslateInput.model_validate(
        {
            "source": {"kind": "url", "url": "https://example.com/x.docx"},
            "target_language": "German",
        }
    )
    assert isinstance(payload.source, HttpUrlSource)


def test_gcs_source_round_trips():
    payload = TranslateInput.model_validate(
        {
            "source": {
                "kind": "gcs_key",
                "gcs_key": "uploads/abc/x.docx",
                "original_filename": "x.docx",
            },
            "target_language": "French",
        }
    )
    assert isinstance(payload.source, GcsKeySource)


def test_base64_source_under_cap_ok():
    data = base64.b64encode(b"hi").decode()
    payload = TranslateInput.model_validate(
        {
            "source": {"kind": "base64", "data": data, "filename": "x.docx"},
            "target_language": "Spanish",
        }
    )
    assert isinstance(payload.source, Base64Source)


def test_base64_source_over_cap_rejected():
    # ~2 MiB of base64 — comfortably over the 1 MiB cap.
    big = "A" * (2 * 1_398_101)
    with pytest.raises(ValidationError):
        TranslateInput.model_validate(
            {
                "source": {"kind": "base64", "data": big, "filename": "x.docx"},
                "target_language": "Spanish",
            }
        )


def test_target_language_required():
    with pytest.raises(ValidationError):
        TranslateInput.model_validate(
            {"source": {"kind": "path", "path": "/tmp/x.docx"}, "target_language": ""}
        )


def test_document_source_adapter_picks_variant_by_kind():
    src = DocumentSourceAdapter.validate_python({"kind": "url", "url": "https://x.com/a.docx"})
    assert isinstance(src, HttpUrlSource)


def test_document_source_adapter_rejects_unknown_kind():
    with pytest.raises(ValidationError):
        DocumentSourceAdapter.validate_python({"kind": "ftp", "url": "ftp://x.com/a.docx"})


def test_document_source_adapter_enforces_base64_cap():
    big = "A" * (2 * 1_398_101)
    with pytest.raises(ValidationError):
        DocumentSourceAdapter.validate_python(
            {"kind": "base64", "data": big, "filename": "x.docx"}
        )
