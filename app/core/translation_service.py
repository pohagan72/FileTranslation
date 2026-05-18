"""Orchestrates a single translation job end-to-end.

The service depends only on `StorageBackend`, `TranslationProvider`, and the
reader registry — no Flask, no GCS specifics, no Gemini specifics. That seam
makes the workflow trivially unit-testable with fakes.

Concurrency: segments are translated in parallel using a `ThreadPoolExecutor`.
Gemini calls are I/O-bound (HTTPS round-trip), so threads give a real
throughput win without needing async.
"""

from __future__ import annotations

import logging
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import timedelta
from typing import IO, List, Optional

from .providers import TranslationError, TranslationProvider
from .readers import get_handler, supported_extensions
from .storage import StorageBackend

logger = logging.getLogger(__name__)


class UnsupportedFileType(ValueError):
    """Raised when the uploaded file extension is not supported."""


@dataclass
class TranslationResult:
    job_id: str
    download_url: str
    download_filename: str


class TranslationService:
    """Coordinates upload -> extract -> translate -> upload -> signed URL."""

    def __init__(
        self,
        storage: StorageBackend,
        provider: TranslationProvider,
        translation_threads: int = 8,
        signed_url_expiry: timedelta = timedelta(minutes=15),
        max_retries: int = 2,
    ):
        self._storage = storage
        self._provider = provider
        self._threads = max(1, translation_threads)
        self._expiry = signed_url_expiry
        self._max_retries = max_retries

    @staticmethod
    def supported_extensions() -> tuple[str, ...]:
        return supported_extensions()

    def translate_document(
        self,
        file_stream: IO[bytes],
        original_filename: str,
        extension: str,
        target_language: str,
    ) -> TranslationResult:
        try:
            handler = get_handler(extension)
        except KeyError as exc:
            raise UnsupportedFileType(str(exc)) from exc

        job_id = uuid.uuid4().hex
        upload_key = f"{job_id}/uploaded/{original_filename}"
        translated_filename = f"translated_{original_filename}"
        translated_key = f"{job_id}/translated/{translated_filename}"

        logger.info("job %s: uploading original %s", job_id, original_filename)
        self._storage.upload(upload_key, file_stream)

        try:
            uploaded_stream = self._storage.download(upload_key)
            segments = handler.collect_segments(uploaded_stream)
            logger.info("job %s: %d segments to translate", job_id, len(segments))

            translations = self._translate_all(segments, target_language)

            uploaded_stream.seek(0)
            translated_stream = handler.apply_translations(uploaded_stream, translations)
        finally:
            try:
                self._storage.delete(upload_key)
            except Exception as exc:  # noqa: BLE001 — cleanup must not fail the job
                logger.warning("job %s: failed to delete upload: %s", job_id, exc)

        logger.info("job %s: uploading translated file", job_id)
        self._storage.upload(translated_key, translated_stream)

        url = self._storage.signed_url(
            translated_key,
            expiry=self._expiry,
            download_filename=translated_filename,
        )

        return TranslationResult(
            job_id=job_id,
            download_url=url,
            download_filename=translated_filename,
        )

    def _translate_all(
        self,
        segments: List[str],
        target_language: str,
    ) -> List[str]:
        if not segments:
            return []

        def translate_one(text: str) -> str:
            if not text or not text.strip():
                return text
            return self._translate_with_retry(text, target_language)

        if self._threads == 1 or len(segments) == 1:
            return [translate_one(s) for s in segments]

        with ThreadPoolExecutor(max_workers=self._threads) as pool:
            # `map` preserves input order, which is what `apply_translations` needs.
            return list(pool.map(translate_one, segments))

    def _translate_with_retry(
        self,
        text: str,
        target_language: str,
    ) -> str:
        last_error: Optional[Exception] = None
        for attempt in range(self._max_retries + 1):
            try:
                return self._provider.translate(text, target_language)
            except TranslationError as exc:
                last_error = exc
                if attempt == self._max_retries:
                    break
                backoff = 0.5 * (2**attempt)
                logger.warning(
                    "translation attempt %d/%d failed: %s — retrying in %.1fs",
                    attempt + 1,
                    self._max_retries + 1,
                    exc,
                    backoff,
                )
                time.sleep(backoff)
        logger.error("translation gave up for segment: %s", last_error)
        # Keep the document intact; one untranslated segment is better than a
        # failed job, and the failure is recorded in logs for follow-up.
        return text
