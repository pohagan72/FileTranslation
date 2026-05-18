"""HTML form for browser-based translation requests."""

from __future__ import annotations

import logging
import os

from flask import Blueprint, current_app, flash, get_flashed_messages, render_template, request
from werkzeug.utils import secure_filename

from app.core import TranslationService, UnsupportedFileType

logger = logging.getLogger(__name__)

bp = Blueprint("web", __name__)


@bp.route("/", methods=["GET", "POST"])
def index():
    config = current_app.config["APP_CONFIG"]
    service: TranslationService | None = current_app.config.get("TRANSLATION_SERVICE")

    services_ready = service is not None
    context = {
        "languages": config.supported_languages,
        "services_ready": services_ready,
        "gemini_configured": config.gemini_configured,
        "gcs_configured": config.gcs_configured,
        "download_url": None,
    }

    if request.method == "GET":
        # Consume any leftover flashes so a refresh doesn't replay them.
        get_flashed_messages()
        return render_template("index.html", **context)

    if service is None:
        flash("Service is not fully configured. Check server logs.")
        return render_template("index.html", **context), 503

    file = request.files.get("file")
    target_language = request.form.get("target_language")

    if not file or not file.filename:
        flash("No file selected.")
        return render_template("index.html", **context), 400
    if not target_language:
        flash("No target language selected.")
        return render_template("index.html", **context), 400

    safe_filename = secure_filename(file.filename) or "upload"
    extension = os.path.splitext(safe_filename)[1].lower()
    if extension not in service.supported_extensions():
        flash(f"Unsupported file type: {extension or 'unknown'}.")
        return render_template("index.html", **context), 400

    try:
        result = service.translate_document(
            file_stream=file.stream,
            original_filename=safe_filename,
            extension=extension,
            target_language=target_language,
        )
    except UnsupportedFileType:
        flash("Unsupported file type.")
        return render_template("index.html", **context), 400
    except Exception:
        logger.exception("translation job failed")
        flash("Translation failed. Please try again.")
        return render_template("index.html", **context), 500

    flash("Translation completed. Use the link below to download.")
    context["download_url"] = result.download_url
    return render_template("index.html", **context)
