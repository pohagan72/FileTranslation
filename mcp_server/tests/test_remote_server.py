"""Tests for the remote HTTP server's tool implementations.

Source-policy tests run without network; the happy path is a stub pending
an httpx mock transport setup once the SDK wire-up lands.
"""

from __future__ import annotations

import pytest

from filetranslation_mcp import remote_server
from filetranslation_mcp.errors import CODE_BAD_SOURCE, ToolError
from filetranslation_mcp.tools import (
    Base64Source,
    GcsKeySource,
    HttpUrlSource,
    LocalPathSource,
)


def test_materialize_rejects_local_path():
    with pytest.raises(ToolError) as exc:
        remote_server._materialize_source(LocalPathSource(path="/tmp/x.docx"))
    assert exc.value.code == CODE_BAD_SOURCE


def test_materialize_rejects_http_non_tls():
    with pytest.raises(ToolError) as exc:
        remote_server._materialize_source(HttpUrlSource(url="http://example.com/x.docx"))
    assert exc.value.code == CODE_BAD_SOURCE


def test_materialize_base64_ok():
    import base64

    payload = Base64Source(data=base64.b64encode(b"hi").decode(), filename="x.docx")
    stream, name = remote_server._materialize_source(payload)
    assert stream.read() == b"hi"
    assert name == "x.docx"


def test_materialize_rejects_gcs_key_until_upstream_endpoint():
    with pytest.raises(ToolError) as exc:
        remote_server._materialize_source(
            GcsKeySource(gcs_key="uploads/abc/x.docx", original_filename="x.docx")
        )
    assert exc.value.code == CODE_BAD_SOURCE
    assert "POST /api/v1/uploads" in exc.value.message


def test_filename_from_url_strips_query():
    assert remote_server._filename_from_url("https://x.com/a/b.docx?token=xyz") == "b.docx"
    assert remote_server._filename_from_url("https://x.com/") == "document"


def test_build_server_registers_three_tools():
    pytest.importorskip("mcp.server.fastmcp")
    from filetranslation_mcp.upstream import UpstreamClient

    client = UpstreamClient(base_url="https://upstream.invalid", api_key=None)
    try:
        mcp = remote_server.build_server(client, expiry_minutes=5)
        mgr = getattr(mcp, "_tool_manager", None)
        assert mgr is not None and hasattr(mgr, "_tools"), "FastMCP shape unrecognized"
        names = set(mgr._tools.keys())
        assert {
            "list_supported_languages",
            "get_translation_service_status",
            "translate_document",
        } <= names
    finally:
        client.close()
