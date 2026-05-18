# filetranslation-mcp

Example **Model Context Protocol** server that exposes the FileTranslation
service as tools an agent can call. This package is a companion to the main
Flask app — not a replacement for it. The HTTP API at `/api/v1/*` remains the
system of record; this server is a thin adapter that demonstrates how to
surface the same domain logic to an agent host (Claude Desktop, an SDK-built
agent, etc.).

## Why this exists

The `app.core` package was built framework-independent on purpose. This server
is the proof: the same `TranslationService` that powers the Flask routes and
HTML form is wrapped here as MCP tools, without re-implementing any business
logic.

It is intended as a reference for engineers building agents who want a
"translate a Microsoft Office document" skill in their tool surface.

## Two modes

| Mode | Entry point | Transport | When to use |
|------|-------------|-----------|-------------|
| **Local** | `filetranslation-mcp-local` | stdio | Running alongside an agent host (Claude Desktop, IDE plugins). Reads files from the local filesystem and calls `TranslationService` in-process. |
| **Remote** | `filetranslation-mcp-remote` | Streamable HTTP | Deployed alongside the Cloud Run service. Calls the existing `/api/v1/*` API over HTTPS. |

Both modes share the **same tool surface and schemas** (see `tools.py`). The
only difference is how the tool implementations reach the underlying service.

## Tool surface

- `list_supported_languages` — returns the configured target languages.
- `get_translation_service_status` — health + which extensions are supported.
- `translate_document` — translate a `.docx` / `.pptx` / `.xlsx` file. Returns
  a **signed download URL**, never the file bytes. The URL expires (default
  5 minutes); fetch promptly.

A full narrative example of an agent using these tools end-to-end lives in
[`examples/agent_usage.md`](examples/agent_usage.md).

## Install & run

The package is its own pyproject; pick the extras for the mode you want.
From the repo root:

```bash
# Local stdio mode — wraps app.core in-process. Requires the parent
# project's Gemini/GCS env vars (see ../README.md → Environment variables).
pip install -e .[local] --config-settings editable_mode=strict
pip install -e ./mcp_server[local]
filetranslation-mcp-local            # starts the stdio server

# Remote HTTP mode — calls /api/v1/* on a running upstream Flask app.
pip install -e ./mcp_server[remote]
UPSTREAM_BASE_URL=https://your-cloud-run-url \
  MCP_UPSTREAM_API_KEY=$API_KEY \
  PORT=8080 \
  filetranslation-mcp-remote
```

To register the local server with Claude Desktop, copy
[`examples/claude_desktop_config.json`](examples/claude_desktop_config.json)
into your Claude Desktop config file and fill in the env block.

## Running tests

```bash
pip install -e ./mcp_server[dev]
cd mcp_server && pytest -q
```

A `tests/conftest.py` adds `src/` to `sys.path` so the suite runs without
an install. Two happy-path tests are skipped unless `app.core` is on
`PYTHONPATH` — set `PYTHONPATH=..` from `mcp_server/` to enable them.

## Status

v0.1 wired up. Both servers register the three tools against `FastMCP` and
run with their respective transports (`mcp.run()` for stdio, `mcp.run(
transport="streamable-http")` for HTTP). 20 unit tests pass; 2 happy-path
tests are skipped pending CI config that puts `app.core` on `PYTHONPATH`.
See [`TODO.md`](TODO.md) for the remaining production-readiness items.
