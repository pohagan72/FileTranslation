"""JSON API for machine clients.

Endpoints:
    GET  /api/v1/health        — liveness probe (always open)
    GET  /api/v1/languages     — supported target languages (always open)
    POST /api/v1/translations  — multipart upload, returns signed-URL JSON
                                 requires X-API-Key header if API_KEY is set
"""

from __future__ import annotations

import hmac
import logging
import os

from flask import Blueprint, current_app, jsonify, request
from werkzeug.utils import secure_filename

from app.core import TranslationService, UnsupportedFileType

logger = logging.getLogger(__name__)

bp = Blueprint("api", __name__, url_prefix="/api/v1")


def _require_api_key():
    """Return None if the request is authorised, else a Flask 401 response.

    If the server has no API_KEY configured, the endpoint is open — this
    keeps local dev frictionless. Compare with hmac.compare_digest to avoid
    timing-side-channel attacks on key comparison.
    """
    expected = current_app.config["APP_CONFIG"].api_key
    if not expected:
        return None
    supplied = request.headers.get("X-API-Key", "")
    if not hmac.compare_digest(supplied, expected):
        return jsonify({"error": "unauthorized"}), 401
    return None


@bp.get("/health")
def health():
    config = current_app.config["APP_CONFIG"]
    service = current_app.config.get("TRANSLATION_SERVICE")
    return jsonify(
        {
            "status": "ok" if service is not None else "degraded",
            "gemini_configured": config.gemini_configured,
            "gcs_configured": config.gcs_configured,
            "supported_extensions": list(TranslationService.supported_extensions()),
        }
    )


@bp.get("/languages")
def languages():
    config = current_app.config["APP_CONFIG"]
    return jsonify({"languages": config.supported_languages})


@bp.post("/translations")
def create_translation():
    auth_error = _require_api_key()
    if auth_error is not None:
        return auth_error

    service: TranslationService | None = current_app.config.get("TRANSLATION_SERVICE")
    if service is None:
        return jsonify({"error": "service unavailable"}), 503

    file = request.files.get("file")
    target_language = request.form.get("target_language")

    if not file or not file.filename:
        return jsonify({"error": "file is required"}), 400
    if not target_language:
        return jsonify({"error": "target_language is required"}), 400

    safe_filename = secure_filename(file.filename) or "upload"
    extension = os.path.splitext(safe_filename)[1].lower()
    if extension not in service.supported_extensions():
        return jsonify({"error": f"unsupported file type: {extension}"}), 400

    try:
        result = service.translate_document(
            file_stream=file.stream,
            original_filename=safe_filename,
            extension=extension,
            target_language=target_language,
        )
    except UnsupportedFileType as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception:
        logger.exception("translation job failed")
        return jsonify({"error": "translation failed"}), 500

    return (
        jsonify(
            {
                "job_id": result.job_id,
                "download_url": result.download_url,
                "download_filename": result.download_filename,
                "detected_language": result.detected_language,
            }
        ),
        201,
    )
