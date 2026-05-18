"""Application logging setup with request-id correlation.

Every log record gets a `request_id` field populated from Flask's `g` object
when emitted inside a request. Outside a request (startup, background tasks)
the field is `-`. The format string includes it so Cloud Run logs are
correlatable end-to-end.
"""

from __future__ import annotations

import logging
import sys


class RequestIdFilter(logging.Filter):
    """Inject `request_id` into every LogRecord.

    Reads from `flask.g.request_id` if a request context is active; otherwise
    falls back to "-". Importing Flask lazily keeps this filter usable in
    tests and scripts that don't spin up an app.
    """

    def filter(self, record: logging.LogRecord) -> bool:
        request_id = "-"
        try:
            from flask import g, has_request_context

            if has_request_context():
                request_id = getattr(g, "request_id", "-") or "-"
        except Exception:  # noqa: BLE001 — logging must never raise
            pass
        record.request_id = request_id
        return True


def configure_logging(level: int = logging.INFO) -> None:
    """Configure root logging once for the application.

    Idempotent: safe to call multiple times (e.g. tests + app factory).
    """
    root = logging.getLogger()
    if getattr(root, "_filetranslation_configured", False):
        return

    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(name)s [%(request_id)s]: %(message)s")
    )
    handler.addFilter(RequestIdFilter())
    root.handlers = [handler]
    root.setLevel(level)
    root._filetranslation_configured = True  # type: ignore[attr-defined]
