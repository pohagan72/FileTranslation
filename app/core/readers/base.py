"""Reader and translator interfaces for office document formats.

Two-phase design so segments can be translated in parallel:

1. `collect_segments(stream)` walks the document and returns the list of
   text segments to translate, in document order.
2. `apply_translations(stream, translations)` walks the document again
   and replaces each segment with its translation by index.

This lets a `ThreadPoolExecutor` translate all segments concurrently between
the two phases, without each handler needing to know about concurrency.
"""

from __future__ import annotations

import io
from abc import ABC, abstractmethod
from typing import List


class DocumentHandler(ABC):
    """Reads and writes one office file format (docx, pptx, xlsx)."""

    extension: str  # e.g. ".docx" — set on subclasses

    @abstractmethod
    def collect_segments(self, stream: io.BytesIO) -> List[str]:
        """Return every translatable segment, in stable document order."""

    @abstractmethod
    def apply_translations(self, stream: io.BytesIO, translations: List[str]) -> io.BytesIO:
        """Re-walk the document and substitute each segment with its translation.

        `translations` must have the same length and order as the list returned
        by `collect_segments` for the same `stream` contents.
        """
