"""Application configuration loaded from environment variables."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import List

from dotenv import load_dotenv

load_dotenv()


class ConfigError(RuntimeError):
    """Raised when required configuration is missing or invalid."""


def _env(name: str, default: str | None = None) -> str | None:
    value = os.getenv(name)
    return value if value not in (None, "") else default


def _env_int(name: str, default: int) -> int:
    raw = _env(name)
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError as exc:
        raise ConfigError(f"{name} must be an integer, got {raw!r}") from exc


def _env_bool(name: str, default: bool) -> bool:
    raw = _env(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "t", "yes", "y"}


@dataclass(frozen=True)
class Config:
    secret_key: str
    google_api_key: str | None
    gemini_model: str
    gcs_bucket_name: str | None
    google_cloud_project: str | None
    port: int
    debug: bool
    max_content_length: int
    translation_threads: int
    signed_url_expiry_minutes: int
    supported_languages: List[str] = field(
        default_factory=lambda: [
            "English",
            "Spanish",
            "French",
            "German",
            "Chinese",
            "Japanese",
        ]
    )

    @property
    def gemini_configured(self) -> bool:
        return bool(self.google_api_key and self.gemini_model)

    @property
    def gcs_configured(self) -> bool:
        return bool(self.gcs_bucket_name and self.google_cloud_project)


def load_config() -> Config:
    """Build a Config from the current environment.

    Secrets are not required at import time so the app can still start in a
    degraded state and surface configuration errors through the UI/API.
    """
    secret_key = _env("SECRET_KEY")
    if not secret_key:
        # Fall back to an ephemeral key so dev still works; logged by app factory.
        secret_key = os.urandom(24).hex()

    return Config(
        secret_key=secret_key,
        google_api_key=_env("GOOGLE_API_KEY"),
        gemini_model=_env("GEMINI_MODEL", "gemini-1.5-flash") or "gemini-1.5-flash",
        gcs_bucket_name=_env("GCS_BUCKET_NAME"),
        google_cloud_project=_env("GOOGLE_CLOUD_PROJECT"),
        port=_env_int("PORT", 8080),
        debug=_env_bool("FLASK_DEBUG", False),
        max_content_length=_env_int("MAX_CONTENT_LENGTH_BYTES", 25 * 1024 * 1024),
        translation_threads=_env_int("TRANSLATION_THREADS", 8),
        signed_url_expiry_minutes=_env_int("SIGNED_URL_EXPIRY_MINUTES", 15),
    )
