"""Tests for the local stdio server's tool implementations.

These exercise the in-process dispatch using fake services that satisfy the
duck-typed interface `TranslationService` exposes — no MCP SDK, no Flask,
no Gemini, no GCS. The MCP SDK wire-up itself is tested separately once
implemented.
"""

from __future__ import annotations

from dataclasses import dataclass

import pytest

from filetranslation_mcp import local_server
from filetranslation_mcp.errors import (
    CODE_BAD_SOURCE,
    CODE_UNSUPPORTED_FILE_TYPE,
    ToolError,
)
from filetranslation_mcp.tools import (
    Base64Source,
    GcsKeySource,
    HttpUrlSource,
    LocalPathSource,
    TranslateInput,
)


@dataclass
class FakeConfig:
    supported_languages: list[str]
    gemini_configured: bool = True
    gcs_configured: bool = True
    signed_url_expiry_minutes: int = 5


class FakeResult:
    def __init__(self, **kw):
        for k, v in kw.items():
            setattr(self, k, v)


class FakeService:
    @staticmethod
    def supported_extensions() -> tuple[str, ...]:
        return (".docx", ".pptx", ".xlsx")

    def __init__(self, raise_unsupported: bool = False):
        self._raise_unsupported = raise_unsupported

    def translate_document(self, *, file_stream, original_filename, extension, target_language):
        if self._raise_unsupported:
            # Import lazily so the test only requires app.core when this branch runs.
            from app.core import UnsupportedFileType  # type: ignore[import-not-found]

            raise UnsupportedFileType(f"unsupported: {extension}")
        return FakeResult(
            job_id="job-1",
            download_url="https://example.com/signed",
            download_filename=f"translated_{original_filename}",
            detected_language="en",
        )


def test_list_languages_returns_config_list():
    cfg = FakeConfig(supported_languages=["English", "Spanish"])
    out = local_server.tool_list_languages(cfg)
    assert out.languages == ["English", "Spanish"]


def test_translate_rejects_url_source():
    payload = TranslateInput(
        source=HttpUrlSource(url="https://example.com/x.docx"),
        target_language="Spanish",
    )
    with pytest.raises(ToolError) as exc:
        local_server.tool_translate(FakeService(), FakeConfig(supported_languages=[]), payload)
    assert exc.value.code == CODE_BAD_SOURCE


def test_translate_rejects_gcs_source():
    payload = TranslateInput(
        source=GcsKeySource(gcs_key="x", original_filename="x.docx"),
        target_language="Spanish",
    )
    with pytest.raises(ToolError) as exc:
        local_server.tool_translate(FakeService(), FakeConfig(supported_languages=[]), payload)
    assert exc.value.code == CODE_BAD_SOURCE


def test_translate_rejects_missing_path(tmp_path):
    payload = TranslateInput(
        source=LocalPathSource(path=str(tmp_path / "does-not-exist.docx")),
        target_language="Spanish",
    )
    with pytest.raises(ToolError) as exc:
        local_server.tool_translate(FakeService(), FakeConfig(supported_languages=[]), payload)
    assert exc.value.code == CODE_BAD_SOURCE


@pytest.mark.skip(reason="Enable once app.core is on sys.path in CI")
def test_translate_happy_path(tmp_path):
    f = tmp_path / "input.docx"
    f.write_bytes(b"fake docx bytes")
    payload = TranslateInput(
        source=LocalPathSource(path=str(f)),
        target_language="Spanish",
    )
    out = local_server.tool_translate(
        FakeService(), FakeConfig(supported_languages=["Spanish"]), payload
    )
    assert out.job_id == "job-1"
    assert out.download_filename == "translated_input.docx"
    assert out.expires_in_seconds == 300


@pytest.mark.skip(reason="Enable once app.core is on sys.path in CI")
def test_translate_maps_unsupported_file_type(tmp_path):
    f = tmp_path / "input.pdf"
    f.write_bytes(b"x")
    payload = TranslateInput(
        source=LocalPathSource(path=str(f)),
        target_language="Spanish",
    )
    with pytest.raises(ToolError) as exc:
        local_server.tool_translate(
            FakeService(raise_unsupported=True),
            FakeConfig(supported_languages=["Spanish"]),
            payload,
        )
    assert exc.value.code == CODE_UNSUPPORTED_FILE_TYPE
    assert ".docx" in exc.value.details["supported_extensions"]


def test_build_server_registers_three_tools():
    """Smoke test: build_server returns a FastMCP with our three tools.

    Skipped automatically when the `mcp` SDK isn't installed (e.g. CI
    pipelines that only run schema tests).
    """
    pytest.importorskip("mcp.server.fastmcp")
    mcp = local_server.build_server(FakeService(), FakeConfig(supported_languages=["English"]))
    # FastMCP exposes a list_tools coroutine; we just need to confirm the
    # three names are registered. The exact accessor has shifted across
    # SDK versions, so use a structural probe.
    names = _collect_tool_names(mcp)
    assert {"list_supported_languages", "get_translation_service_status", "translate_document"} <= names


def _collect_tool_names(mcp: object) -> set[str]:
    """Best-effort tool-name extraction across FastMCP versions."""
    # Newer SDK: `mcp._tool_manager._tools` is a dict[str, Tool].
    mgr = getattr(mcp, "_tool_manager", None)
    if mgr is not None and hasattr(mgr, "_tools"):
        return set(mgr._tools.keys())
    # Fallback: introspect any list_tools coroutine via asyncio.
    import asyncio

    list_tools = getattr(mcp, "list_tools", None)
    if list_tools is not None:
        tools = asyncio.get_event_loop().run_until_complete(list_tools())
        return {t.name for t in tools}
    raise AssertionError("could not introspect FastMCP tools — SDK shape unrecognized")
