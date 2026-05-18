"""Flask application factory.

The factory wires together: config, logging, the storage backend, the
translation provider, and the translation service. Each piece is constructed
behind try/except so the app boots into a degraded-but-running state if any
optional dependency is misconfigured. Boot failures surface via /healthz
and the HTML form's disabled controls.
"""

from __future__ import annotations

import logging
from datetime import timedelta

from flask import Flask, jsonify

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


def _build_storage(cfg: Config):
    if not cfg.gcs_configured:
        logger.warning("GCS not configured — set GCS_BUCKET_NAME and GOOGLE_CLOUD_PROJECT")
        return None
    try:
        return GCSBackend(
            project=cfg.google_cloud_project,
            bucket_name=cfg.gcs_bucket_name,
        )
    except StorageError as exc:
        logger.error("failed to initialise GCS backend: %s", exc)
        return None
    except Exception as exc:  # noqa: BLE001 — boot must not crash on cloud auth errors
        logger.exception("unexpected error initialising GCS backend: %s", exc)
        return None


def _build_provider(cfg: Config) -> TranslationProvider | None:
    if not cfg.gemini_configured:
        logger.warning("Gemini not configured — set GOOGLE_API_KEY and GEMINI_MODEL")
        return None
    try:
        return GeminiProvider(api_key=cfg.google_api_key, model_name=cfg.gemini_model)
    except Exception as exc:  # noqa: BLE001
        logger.exception("failed to initialise Gemini provider: %s", exc)
        return None
