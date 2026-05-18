# filetranslation-mcp — implementation TODO

The MCP SDK wire-up is done: both `local_server.py` and `remote_server.py`
register the three tools against `FastMCP` and call `mcp.run()` with the
appropriate transport. The pure tool implementations are kept separate from
SDK registration so they stay unit-testable.

## Done

- [x] Tool schemas with discriminated-union `DocumentSource`.
- [x] Structured `ToolError` that survives FastMCP's exception → tool-error path.
- [x] Local stdio server: `build_server` + `main()` wired to FastMCP.
- [x] Remote HTTP server: `build_server` + `main()` wired to FastMCP
      Streamable HTTP, honoring `PORT` via `FASTMCP_PORT`.
- [x] `tests/` — 20 passing, 2 skipped (need `app.core` on `sys.path` in CI
      for the happy-path translate tests).

## Should-do before v0.1 ships

- [ ] **Real integration test of the SDK loop.** Currently we assert tool
      *registration* via FastMCP introspection. A test that connects an
      MCP client to a spawned stdio server and round-trips a call would
      catch protocol drift.
- [ ] **Upstream `POST /api/v1/uploads` endpoint** in the Flask app so the
      remote server can accept `gcs_key` sources. Today it rejects them
      with a clear error.
- [ ] **Per-call request-id propagation to the upstream API.** The remote
      server already passes `ctx.request_id` into `UpstreamClient`, but the
      Flask side honors `X-Request-Id` from `app/__init__.py:73-82` — verify
      end-to-end in a real deployment, not just in unit tests.
- [ ] **Enable the two skipped tests** in CI by adding the repo root to
      `PYTHONPATH` so `app.core` resolves.

## Nice-to-have

- [ ] OAuth flow for remote mode (current: bearer-style API key).
- [ ] MCP resource exposing `languages://supported` for clients that prefer
      resources over tool calls for static lists.
- [ ] MCP prompt template "translate-this-document".
- [ ] Cloud Run deployment manifest + Dockerfile for the remote server.
- [ ] Pin `mcp` to a tested version once the SDK reaches API stability;
      today the `>=1.2.0` floor is intentionally loose.

## Known gotchas (worth documenting if external contributors hit them)

- FastMCP resolves tool signatures with `inspect.signature(func,
  eval_str=True)`, which means `from __future__ import annotations` in a
  server module breaks tool registration with `InvalidSignature`. Both
  server modules have an explicit comment about this.
- FastMCP tool inputs cannot be a single Pydantic model; arguments must be
  individual params. We validate the discriminated `source` union manually
  via `DocumentSourceAdapter` inside the wrapper.
