"""Structured error mapping between the translation service and MCP tools.

Agents handle structured errors much better than free-text — they can branch
on a stable `code` and surface a useful message to the user. The codes here
are part of the tool contract; rename carefully.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any


@dataclass
class ToolError(Exception):
    """Raised by tool implementations; serialized into the MCP error payload.

    FastMCP converts an exception raised inside a tool into a tool result
    with `isError: true` and the exception's string form as content. To keep
    the structured `code` / `details` reachable by agents, we override
    `__str__` to emit a JSON payload — agents parsing the error get
    machine-readable fields, and humans see something still readable.
    """

    code: str
    message: str
    details: dict[str, Any] = field(default_factory=dict)

    def to_payload(self) -> dict[str, Any]:
        return {"code": self.code, "message": self.message, "details": self.details}

    def __str__(self) -> str:
        return json.dumps(self.to_payload(), default=str)


# Stable error codes — keep in sync with examples/agent_usage.md.
CODE_UNSUPPORTED_FILE_TYPE = "unsupported_file_type"
CODE_SERVICE_UNAVAILABLE = "service_unavailable"
CODE_UNAUTHORIZED = "unauthorized"
CODE_FILE_TOO_LARGE = "file_too_large"
CODE_BAD_SOURCE = "bad_source"
CODE_UPSTREAM_FAILURE = "upstream_failure"


def from_unsupported_file_type(exc: Exception, supported: list[str]) -> ToolError:
    return ToolError(
        code=CODE_UNSUPPORTED_FILE_TYPE,
        message=str(exc),
        details={"supported_extensions": supported},
    )


def from_upstream_http(status: int, body: Any) -> ToolError:
    """Translate a non-2xx response from the Flask API into a ToolError.

    Status codes here mirror `app/api/views.py` and `app/__init__.py:88`.
    """
    if status == 401:
        return ToolError(
            code=CODE_UNAUTHORIZED,
            message="Upstream API key missing or wrong. Set MCP_UPSTREAM_API_KEY.",
        )
    if status == 413:
        return ToolError(
            code=CODE_FILE_TOO_LARGE,
            message="Document exceeds the upstream upload size limit.",
            details={"upstream_body": body},
        )
    if status == 400:
        return ToolError(
            code=CODE_BAD_SOURCE,
            message="Upstream rejected the request as malformed.",
            details={"upstream_body": body},
        )
    if status == 503:
        return ToolError(
            code=CODE_SERVICE_UNAVAILABLE,
            message="Upstream translation service is degraded — check its /api/v1/health.",
        )
    return ToolError(
        code=CODE_UPSTREAM_FAILURE,
        message=f"Upstream returned HTTP {status}.",
        details={"status": status, "upstream_body": body},
    )
