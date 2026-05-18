"""Core domain logic — independent of Flask, HTTP, and GCP specifics."""

from .language import detect_language
from .providers import GeminiProvider, TranslationError, TranslationProvider
from .readers import DocumentHandler, get_handler, supported_extensions
from .storage import GCSBackend, ObjectNotFound, StorageBackend, StorageError
from .translation_service import (
    TranslationResult,
    TranslationService,
    UnsupportedFileType,
)

__all__ = [
    "detect_language",
    "GeminiProvider",
    "TranslationError",
    "TranslationProvider",
    "DocumentHandler",
    "get_handler",
    "supported_extensions",
    "GCSBackend",
    "ObjectNotFound",
    "StorageBackend",
    "StorageError",
    "TranslationResult",
    "TranslationService",
    "UnsupportedFileType",
]
