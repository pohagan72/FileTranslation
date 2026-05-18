"""Flask application factory.

The factory wires together: config, logging, the storage backend, the
translation provider, and the translation service. Each piece is constructed
behind try/except so the app boots into a degraded-but-running state if any
optional dependency is misconfigured. Boot failures surface via /healthz
and the HTML form's disabled controls.
"""

from __future__ import annotations

import logging
import uuid
from datetime import timedelta

from flask import Flask, g, jsonify, request

from app.config import Config, load_config
from app.core import (
    GeminiProvider,
    GCSBackend,
    StorageError,
    TranslationProvider,
    TranslationService,
)
from app.logging_setup import configure_logging

logger = logging.getLogger(__name__)


def create_app(config: Config | None = None) -> Flask:
    """Build and configure the Flask application.

    Tests can pass a pre-built `config` to override env-driven settings.
    """
    configure_logging()

    cfg = config or load_config()
    app = Flask(__name__, template_folder="../templates")
    app.secret_key = cfg.secret_key
    app.config["APP_CONFIG"] = cfg
    app.config["MAX_CONTENT_LENGTH"] = cfg.max_content_length

    if cfg.secret_key and len(cfg.secret_key) < 32:
        logger.warning("SECRET_KEY is short — set a 32+ char random string in production")

    storage_backend = _build_storage(cfg)
    provider = _build_provider(cfg)

    if storage_backend is not None and provider is not None:
        app.config["TRANSLATION_SERVICE"] = TranslationService(
            storage=storage_backend,
            provider=provider,
            translation_threads=cfg.translation_threads,
            signed_url_expiry=timedelta(minutes=cfg.signed_url_expiry_minutes),
        )
        logger.info("translation service ready")
    else:
        app.config["TRANSLATION_SERVICE"] = None
        logger.warning(
            "translation service NOT ready (storage=%s, provider=%s)",
            storage_backend is not None,
            provider is not None,
        )

    from app.api import bp as api_bp
    from app.web import bp as web_bp

    app.register_blueprint(api_bp)
    app.register_blueprint(web_bp)

    @app.before_request
    def _assign_request_id() -> None:
        # Use the incoming X-Request-Id if a caller supplied one (handy for
        # tracing through a CDN / load balancer); otherwise mint a fresh one.
        g.request_id = request.headers.get("X-Request-Id") or uuid.uuid4().hex

    @app.after_request
    def _emit_request_id(response):
        if hasattr(g, "request_id"):
            response.headers["X-Request-Id"] = g.request_id
        return response

    @app.get("/healthz")
    def healthz():
        return jsonify({"status": "ok"})

    @app.errorhandler(413)
    def too_large(_):
        return (
            jsonify(
                {
                    "error": "file too large",
                    "max_bytes": cfg.max_content_length,
                }
            ),
            413,
        )

    return app


def _build_storage(cfg: Config) -> GCSBackend | None:
    project = cfg.google_cloud_project
    bucket = cfg.gcs_bucket_name
    if not project or not bucket:
        logger.warning("GCS not configured — set GCS_BUCKET_NAME and GOOGLE_CLOUD_PROJECT")
        return None
    try:
        return GCSBackend(project=project, bucket_name=bucket)
    except StorageError as exc:
        logger.error("failed to initialise GCS backend: %s", exc)
        return None
    except Exception as exc:  # noqa: BLE001 — boot must not crash on cloud auth errors
        logger.exception("unexpected error initialising GCS backend: %s", exc)
        return None


def _build_provider(cfg: Config) -> TranslationProvider | None:
    api_key = cfg.google_api_key
    model = cfg.gemini_model
    if not api_key or not model:
        logger.warning("Gemini not configured — set GOOGLE_API_KEY and GEMINI_MODEL")
        return None
    try:
        return GeminiProvider(api_key=api_key, model_name=model)
    except Exception as exc:  # noqa: BLE001
        logger.exception("failed to initialise Gemini provider: %s", exc)
        return None
