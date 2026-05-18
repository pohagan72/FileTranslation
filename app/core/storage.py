"""Object storage abstraction.

Defines a minimal storage interface so the rest of the app does not depend
directly on Google Cloud Storage. A `GCSBackend` implementation is provided.
"""

from __future__ import annotations

import io
import logging
from abc import ABC, abstractmethod
from datetime import timedelta
from typing import IO, Optional

logger = logging.getLogger(__name__)


class StorageError(RuntimeError):
    """Raised when a storage operation fails."""


class ObjectNotFound(StorageError):
    """Raised when a requested object does not exist."""


class StorageBackend(ABC):
    """Abstract object-storage interface used by the translation service."""

    @abstractmethod
    def upload(self, key: str, stream: IO[bytes], content_type: Optional[str] = None) -> None: ...

    @abstractmethod
    def download(self, key: str) -> io.BytesIO: ...

    @abstractmethod
    def exists(self, key: str) -> bool: ...

    @abstractmethod
    def delete(self, key: str) -> None: ...

    @abstractmethod
    def signed_url(
        self, key: str, expiry: timedelta, download_filename: Optional[str] = None
    ) -> str: ...


class GCSBackend(StorageBackend):
    """Google Cloud Storage implementation of `StorageBackend`."""

    def __init__(self, project: str, bucket_name: str):
        # Imported lazily so unit tests can stub the module without GCP credentials.
        from google.cloud import storage  # type: ignore[attr-defined]
        from google.cloud.exceptions import NotFound

        self._NotFound = NotFound
        self._client = storage.Client(project=project)
        self._bucket = self._client.bucket(bucket_name)
        # Validate bucket exists / is accessible. Surface as StorageError so
        # the app factory can mark the backend unavailable without crashing.
        try:
            self._bucket.reload()
        except NotFound as exc:
            raise StorageError(f"GCS bucket '{bucket_name}' not found") from exc

    def upload(self, key: str, stream: IO[bytes], content_type: Optional[str] = None) -> None:
        blob = self._bucket.blob(key)
        try:
            blob.upload_from_file(stream, content_type=content_type, rewind=True)
        except Exception as exc:
            raise StorageError(f"upload to {key} failed: {exc}") from exc

    def download(self, key: str) -> io.BytesIO:
        blob = self._bucket.blob(key)
        buffer = io.BytesIO()
        try:
            blob.download_to_file(buffer)
        except self._NotFound as exc:
            raise ObjectNotFound(key) from exc
        except Exception as exc:
            raise StorageError(f"download of {key} failed: {exc}") from exc
        buffer.seek(0)
        return buffer

    def exists(self, key: str) -> bool:
        return self._bucket.blob(key).exists()

    def delete(self, key: str) -> None:
        try:
            self._bucket.blob(key).delete()
        except self._NotFound:
            return  # Idempotent
        except Exception as exc:
            raise StorageError(f"delete of {key} failed: {exc}") from exc

    def signed_url(
        self,
        key: str,
        expiry: timedelta,
        download_filename: Optional[str] = None,
    ) -> str:
        blob = self._bucket.blob(key)
        disposition = f'attachment; filename="{download_filename}"' if download_filename else None
        return blob.generate_signed_url(
            version="v4",
            expiration=expiry,
            method="GET",
            response_disposition=disposition,
        )
