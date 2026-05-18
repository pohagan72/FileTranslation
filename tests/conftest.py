"""Shared pytest fixtures and in-memory test doubles."""

from __future__ import annotations

import io
from datetime import timedelta
from typing import IO, Dict, List, Optional

import pytest

from app.config import Config
from app.core import (
    StorageBackend,
    TranslationProvider,
    TranslationService,
)


class InMemoryStorage(StorageBackend):
    """StorageBackend stub backed by a dict — no GCS required."""

    def __init__(self) -> None:
        self.objects: Dict[str, bytes] = {}
        self.deleted: List[str] = []

    def upload(self, key: str, stream: IO[bytes], content_type: Optional[str] = None) -> None:
        stream.seek(0)
        self.objects[key] = stream.read()

    def download(self, key: str) -> io.BytesIO:
        return io.BytesIO(self.objects[key])

    def exists(self, key: str) -> bool:
        return key in self.objects

    def delete(self, key: str) -> None:
        self.deleted.append(key)
        self.objects.pop(key, None)

    def signed_url(
        self,
        key: str,
        expiry: timedelta,
        download_filename: Optional[str] = None,
    ) -> str:
        return f"https://signed.example/{key}?name={download_filename or ''}"


class UppercaseProvider(TranslationProvider):
    """Deterministic provider that 'translates' by uppercasing."""

    def __init__(self) -> None:
        self.calls: List[str] = []

    def translate(self, text: str, target_language: str, source_language=None) -> str:
        self.calls.append(text)
        return text.upper()


@pytest.fixture
def storage() -> InMemoryStorage:
    return InMemoryStorage()


@pytest.fixture
def provider() -> UppercaseProvider:
    return UppercaseProvider()


@pytest.fixture
def service(storage: InMemoryStorage, provider: UppercaseProvider) -> TranslationService:
    # Single-threaded keeps test ordering deterministic without thread races.
    return TranslationService(storage=storage, provider=provider, translation_threads=1)


@pytest.fixture
def test_config() -> Config:
    return Config(
        secret_key="x" * 32,
        google_api_key=None,
        gemini_model="gemini-test",
        gcs_bucket_name=None,
        google_cloud_project=None,
        port=8080,
        debug=False,
        max_content_length=1024 * 1024,
        translation_threads=1,
        signed_url_expiry_minutes=5,
        api_key=None,
    )
