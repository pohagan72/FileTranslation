"""Document readers and translators, registered by file extension."""

from .base import DocumentHandler
from .registry import get_handler, supported_extensions

__all__ = ["DocumentHandler", "get_handler", "supported_extensions"]
