"""Source-language detection.

Wraps `langdetect` so callers can be swapped to a different detector later
without touching the rest of the codebase.
"""

from __future__ import annotations

import logging
from typing import Optional

from langdetect import LangDetectException, detect

logger = logging.getLogger(__name__)

_MAX_SAMPLE_CHARS = 10_000


def detect_language(text: str) -> Optional[str]:
    """Return the detected ISO 639-1 code, or None if detection fails."""
    if not text or not text.strip():
        return None
    try:
        return detect(text[:_MAX_SAMPLE_CHARS])
    except LangDetectException as exc:
        logger.warning("language detection failed: %s", exc)
        return None
