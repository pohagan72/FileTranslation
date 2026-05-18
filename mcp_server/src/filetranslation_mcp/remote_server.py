"""Remote MCP server — speaks Streamable HTTP, calls the upstream Flask API.

Intended to be deployed as a sibling Cloud Run service to the FileTranslation
Flask app. It owns no domain logic; every tool call becomes an HTTP request
to `/api/v1/*` via `UpstreamClient`.

Architecture mirrors `local_server.py`: typed helpers are SDK-independent
and unit-testable; `build_server` registers FastMCP wrappers around them.
"""

# NOTE: deliberately no `from __future__ import annotations` — FastMCP
# resolves tool signatures via `inspect.signature(func, eval_str=True)`
# at registration time, and stringified annotations break that path.

import io
import logging
import os
import uuid
from typing import Any

import httpx

from .errors import CODE_BAD_SOURCE, ToolError
from .tools import (
    TOOL_DESCRIPTIONS,
    TOOL_GET_STATUS,
    TOOL_LIST_LANGUAGES,
    TOOL_TRANSLATE,
    Base64Source,
    DocumentSourceAdapter,
    GcsKeySource,
    HttpUrlSource,
    ListLanguagesOutput,
    LocalPathSource,
    StatusOutput,
    TranslateInput,
    TranslateOutput,
)
from .upstream import UpstreamClient

logger = logging.getLogger(__name__)


# -- config -------------------------------------------------------------------


def _load_env() -> tuple[str, str | None, int, int]:
    base = os.environ.get("UPSTREAM_BASE_URL")
    if not base:
        raise RuntimeError("UPSTREAM_BASE_URL is required (e.g. https://filetranslation.example.com)")
    api_key = os.environ.get("MCP_UPSTREAM_API_KEY")
    # Mirror the upstream default (`app/config.py:102`).
    expiry_minutes = int(os.environ.get("SIGNED_URL_EXPIRY_MINUTES", "5"))
    port = int(os.environ.get("PORT", "8080"))
    return base, api_key, expiry_minutes, port


# -- tool implementations (typed; SDK-independent) ---------------------------


def tool_list_languages(client: UpstreamClient, request_id: str) -> ListLanguagesOutput:
    body = client.languages(request_id=request_id)
    return ListLanguagesOutput(languages=list(body.get("languages") or []))


def tool_get_status(client: UpstreamClient, request_id: str) -> StatusOutput:
    body = client.health(request_id=request_id)
    return StatusOutput(
        status=body.get("status", "degraded"),
        gemini_configured=bool(body.get("gemini_configured")),
        gcs_configured=bool(body.get("gcs_configured")),
        supported_extensions=list(body.get("supported_extensions") or []),
    )


def tool_translate(
    client: UpstreamClient,
    payload: TranslateInput,
    request_id: str,
    expiry_minutes: int,
) -> TranslateOutput:
    file_stream, filename = _materialize_source(payload.source)
    body = client.translate(
        file_stream=file_stream,
        filename=filename,
        target_language=payload.target_language,
        request_id=request_id,
    )
    return TranslateOutput(
        job_id=body["job_id"],
        download_url=body["download_url"],
        download_filename=body["download_filename"],
        expires_in_seconds=expiry_minutes * 60,
    )


def _materialize_source(source: object) -> tuple[io.IOBase, str]:
    """Materialize an agent-supplied source into (stream, filename).

    Remote-server source policy:
      - LocalPathSource → rejected; the agent's machine ≠ this server.
      - HttpUrlSource → fetched (https only) and uploaded as multipart.
      - Base64Source → decoded; size cap is enforced by the Pydantic model.
      - GcsKeySource → not yet supported; needs a companion upstream endpoint
        (see TODO.md). Reject explicitly for now.
    """
    if isinstance(source, LocalPathSource):
        raise ToolError(
            code=CODE_BAD_SOURCE,
            message="'path' sources are local-only; supply a 'url' or 'base64' source.",
        )

    if isinstance(source, HttpUrlSource):
        if not source.url.lower().startswith("https://"):
            raise ToolError(
                code=CODE_BAD_SOURCE,
                message="only https:// URLs are accepted (mitigates SSRF).",
            )
        try:
            resp = httpx.get(source.url, timeout=30.0, follow_redirects=True)
            resp.raise_for_status()
        except httpx.HTTPError as exc:
            raise ToolError(
                code=CODE_BAD_SOURCE,
                message=f"failed to fetch source URL: {exc}",
            ) from exc
        filename = _filename_from_url(source.url)
        return io.BytesIO(resp.content), filename

    if isinstance(source, Base64Source):
        import base64

        try:
            data = base64.b64decode(source.data, validate=True)
        except Exception as exc:
            raise ToolError(code=CODE_BAD_SOURCE, message=f"invalid base64: {exc}") from exc
        return io.BytesIO(data), source.filename

    if isinstance(source, GcsKeySource):
        raise ToolError(
            code=CODE_BAD_SOURCE,
            message=(
                "'gcs_key' sources require an upstream POST /api/v1/uploads endpoint "
                "that does not yet exist — see mcp_server/TODO.md."
            ),
        )

    raise ToolError(code=CODE_BAD_SOURCE, message=f"unknown source kind: {source!r}")


def _filename_from_url(url: str) -> str:
    # Just the last path segment; ignore query strings.
    tail = url.rsplit("/", 1)[-1].split("?", 1)[0]
    return tail or "document"


# -- FastMCP wire-up ----------------------------------------------------------


def build_server(client: UpstreamClient, expiry_minutes: int) -> Any:
    """Build a configured `FastMCP` for Streamable HTTP transport."""
    from mcp.server.fastmcp import Context, FastMCP  # imported here so tests can stub

    mcp = FastMCP("filetranslation")

    @mcp.tool(name=TOOL_LIST_LANGUAGES, description=TOOL_DESCRIPTIONS[TOOL_LIST_LANGUAGES])
    def _list_languages(ctx: Context) -> ListLanguagesOutput:
        logger.info("tool=%s request_id=%s", TOOL_LIST_LANGUAGES, ctx.request_id)
        return tool_list_languages(client, request_id=ctx.request_id)

    @mcp.tool(name=TOOL_GET_STATUS, description=TOOL_DESCRIPTIONS[TOOL_GET_STATUS])
    def _get_status(ctx: Context) -> StatusOutput:
        logger.info("tool=%s request_id=%s", TOOL_GET_STATUS, ctx.request_id)
        return tool_get_status(client, request_id=ctx.request_id)

    @mcp.tool(name=TOOL_TRANSLATE, description=TOOL_DESCRIPTIONS[TOOL_TRANSLATE])
    def _translate(
        source: dict,
        target_language: str,
        ctx: Context,
    ) -> TranslateOutput:
        logger.info("tool=%s request_id=%s", TOOL_TRANSLATE, ctx.request_id)
        try:
            normalized_source = DocumentSourceAdapter.validate_python(source)
        except Exception as exc:
            raise ToolError(code=CODE_BAD_SOURCE, message=f"invalid source: {exc}") from exc
        payload = TranslateInput(source=normalized_source, target_language=target_language)
        return tool_translate(
            client,
            payload,
            request_id=ctx.request_id,
            expiry_minutes=expiry_minutes,
        )

    return mcp


# -- entry point --------------------------------------------------------------


def main() -> None:
    """Streamable HTTP entry point — runs until the process is terminated."""
    logging.basicConfig(level=logging.INFO)
    boot_id = uuid.uuid4().hex
    base, api_key, expiry_minutes, port = _load_env()
    logger.info(
        "filetranslation-mcp remote server starting (upstream=%s, auth=%s, port=%d, boot_id=%s)",
        base,
        "yes" if api_key else "no",
        port,
        boot_id,
    )

    client = UpstreamClient(base_url=base, api_key=api_key)
    try:
        mcp = build_server(client, expiry_minutes=expiry_minutes)
        # PORT is honored by FastMCP via the FASTMCP_PORT / FASTMCP_HOST env
        # vars in current SDK builds. Set them explicitly so Cloud Run's
        # PORT is respected without depending on a kwarg shape that has
        # shifted across releases.
        os.environ.setdefault("FASTMCP_HOST", "0.0.0.0")
        os.environ.setdefault("FASTMCP_PORT", str(port))
        mcp.run(transport="streamable-http")
    finally:
        client.close()


if __name__ == "__main__":
    main()
