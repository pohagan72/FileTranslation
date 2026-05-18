"""Registry mapping file extensions to `DocumentHandler` instances."""

from __future__ import annotations

from typing import Dict

from .base import DocumentHandler
from .docx_handler import DocxHandler
from .pptx_handler import PptxHandler
from .xlsx_handler import XlsxHandler


_HANDLERS: Dict[str, DocumentHandler] = {
    h.extension: h for h in (DocxHandler(), PptxHandler(), XlsxHandler())
}


def supported_extensions() -> tuple[str, ...]:
    return tuple(_HANDLERS.keys())


def get_handler(extension: str) -> DocumentHandler:
    """Return the handler for `extension` (e.g. ".docx"), or raise KeyError."""
    handler = _HANDLERS.get(extension.lower())
    if handler is None:
        raise KeyError(f"Unsupported file extension: {extension}")
    return handler
