"""Local stdio MCP server — wraps `app.core.TranslationService` in-process.

Intended for agent hosts running on the same machine as the user (Claude
Desktop, IDE plugins). The user's host is responsible for granting
filesystem permission; this server then reads the path the agent passes.

Architecture:
  - Pure tool implementations (`tool_list_languages`, `tool_get_status`,
    `tool_translate`) take typed args and are unit-testable without the SDK.
  - `main()` builds a `FastMCP` instance and registers thin wrappers that
    adapt FastMCP's per-arg signature to the typed helpers.
"""

# NOTE: deliberately no `from __future__ import annotations` — FastMCP
# resolves tool signatures via `inspect.signature(func, eval_str=True)`
# at registration time, and stringified annotations break that path.

import io
import logging
import os
import uuid
from datetime import timedelta
from pathlib import Path
from typing import Any

from .errors import (
    CODE_BAD_SOURCE,
    ToolError,
    from_unsupported_file_type,
)
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

logger = logging.getLogger(__name__)


# -- service construction -----------------------------------------------------


def _build_service() -> tuple[Any, Any]:
    """Construct a `TranslationService` and `Config` from the parent project.

    Imported lazily so this module is importable (for tests / linting) even
    in environments where the parent `app.core` isn't installed.
    """
    from app.config import load_config  # type: ignore[import-not-found]
    from app.core import (  # type: ignore[import-not-found]
        GCSBackend,
        GeminiProvider,
        TranslationService,
    )

    cfg = load_config()
    if not cfg.gemini_configured or not cfg.gcs_configured:
        raise RuntimeError(
            "Translation service is not fully configured — set GOOGLE_API_KEY, "
            "GEMINI_MODEL, GOOGLE_CLOUD_PROJECT, GCS_BUCKET_NAME before launching."
        )

    storage = GCSBackend(project=cfg.google_cloud_project, bucket_name=cfg.gcs_bucket_name)
    provider = GeminiProvider(api_key=cfg.google_api_key, model_name=cfg.gemini_model)
    service = TranslationService(
        storage=storage,
        provider=provider,
        translation_threads=cfg.translation_threads,
        signed_url_expiry=timedelta(minutes=cfg.signed_url_expiry_minutes),
    )
    return service, cfg


# -- tool implementations (typed; SDK-independent) ---------------------------


def tool_list_languages(cfg: Any) -> ListLanguagesOutput:
    return ListLanguagesOutput(languages=list(cfg.supported_languages))


def tool_get_status(service: Any, cfg: Any) -> StatusOutput:
    from app.core import TranslationService  # type: ignore[import-not-found]

    return StatusOutput(
        status="ok" if service is not None else "degraded",
        gemini_configured=cfg.gemini_configured,
        gcs_configured=cfg.gcs_configured,
        supported_extensions=list(TranslationService.supported_extensions()),
    )


def tool_translate(service: Any, cfg: Any, payload: TranslateInput) -> TranslateOutput:
    from app.core import UnsupportedFileType  # type: ignore[import-not-found]

    file_stream, filename = _open_source(payload.source)
    extension = Path(filename).suffix.lower()

    try:
        result = service.translate_document(
            file_stream=file_stream,
            original_filename=filename,
            extension=extension,
            target_language=payload.target_language,
        )
    except UnsupportedFileType as exc:
        raise from_unsupported_file_type(
            exc, list(service.supported_extensions())
        ) from exc

    return TranslateOutput(
        job_id=result.job_id,
        download_url=result.download_url,
        download_filename=result.download_filename,
        detected_language=result.detected_language,
        expires_in_seconds=cfg.signed_url_expiry_minutes * 60,
    )


def _open_source(source: Any) -> tuple[Any, str]:
    """Materialize an agent-supplied source into (stream, filename).

    The local server only honors `LocalPathSource` (and `Base64Source` for
    convenience in tests). URL and gcs_key sources require remote-mode
    machinery and are rejected here with a clear code so agents can recover.
    """
    if isinstance(source, LocalPathSource):
        path = Path(source.path)
        if not path.is_file():
            raise ToolError(
                code=CODE_BAD_SOURCE,
                message=f"path does not exist or is not a file: {source.path}",
            )
        # Streamed to avoid loading the whole file into memory.
        return path.open("rb"), path.name

    if isinstance(source, Base64Source):
        import base64

        try:
            data = base64.b64decode(source.data, validate=True)
        except Exception as exc:
            raise ToolError(code=CODE_BAD_SOURCE, message=f"invalid base64: {exc}") from exc
        return io.BytesIO(data), source.filename

    if isinstance(source, (HttpUrlSource, GcsKeySource)):
        raise ToolError(
            code=CODE_BAD_SOURCE,
            message=(
                f"source kind '{source.kind}' is not supported by the local server — "
                "use a 'path' source, or run the remote server instead."
            ),
        )

    raise ToolError(code=CODE_BAD_SOURCE, message=f"unknown source kind: {source!r}")


# -- FastMCP wire-up ----------------------------------------------------------


def build_server(service: Any, cfg: Any) -> Any:
    """Build a configured `FastMCP` instance with the three tools registered.

    Factored out of `main()` so tests can exercise registration without
    actually starting a stdio loop.

    FastMCP tool signatures take individual scalar/dict params (not a wrapper
    Pydantic model), so the wrappers below adapt to that shape and re-validate
    into typed objects via `DocumentSourceAdapter` / `TranslateInput`.
    """
    from mcp.server.fastmcp import Context, FastMCP  # imported here so tests can stub

    mcp = FastMCP("filetranslation")

    @mcp.tool(name=TOOL_LIST_LANGUAGES, description=TOOL_DESCRIPTIONS[TOOL_LIST_LANGUAGES])
    def _list_languages(ctx: Context) -> ListLanguagesOutput:
        logger.info("tool=%s request_id=%s", TOOL_LIST_LANGUAGES, ctx.request_id)
        return tool_list_languages(cfg)

    @mcp.tool(name=TOOL_GET_STATUS, description=TOOL_DESCRIPTIONS[TOOL_GET_STATUS])
    def _get_status(ctx: Context) -> StatusOutput:
        logger.info("tool=%s request_id=%s", TOOL_GET_STATUS, ctx.request_id)
        return tool_get_status(service, cfg)

    @mcp.tool(name=TOOL_TRANSLATE, description=TOOL_DESCRIPTIONS[TOOL_TRANSLATE])
    def _translate(
        source: dict,
        target_language: str,
        ctx: Context,
    ) -> TranslateOutput:
        logger.info("tool=%s request_id=%s", TOOL_TRANSLATE, ctx.request_id)
        # Validate the discriminated union manually — FastMCP accepts the raw
        # dict, but we want Pydantic to pick the right variant + run the
        # base64 size cap before any file I/O happens.
        try:
            normalized_source = DocumentSourceAdapter.validate_python(source)
        except Exception as exc:
            raise ToolError(code=CODE_BAD_SOURCE, message=f"invalid source: {exc}") from exc
        payload = TranslateInput(source=normalized_source, target_language=target_language)
        return tool_translate(service, cfg, payload)

    return mcp


# -- entry point --------------------------------------------------------------


def main() -> None:
    """stdio entry point — runs the FastMCP server until the host disconnects."""
    logging.basicConfig(level=logging.INFO)
    request_id = os.environ.get("MCP_REQUEST_ID") or uuid.uuid4().hex
    logger.info("filetranslation-mcp local stdio server starting (boot_id=%s)", request_id)

    service, cfg = _build_service()
    mcp = build_server(service, cfg)
    # FastMCP.run() defaults to stdio transport.
    mcp.run()


if __name__ == "__main__":
    main()
