"""Application logging setup."""

from __future__ import annotations

import logging
import sys


def configure_logging(level: int = logging.INFO) -> None:
    """Configure root logging once for the application.

    Idempotent: safe to call multiple times (e.g. tests + app factory).
    """
    root = logging.getLogger()
    if getattr(root, "_filetranslation_configured", False):
        return

    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    root.handlers = [handler]
    root.setLevel(level)
    root._filetranslation_configured = True  # type: ignore[attr-defined]
