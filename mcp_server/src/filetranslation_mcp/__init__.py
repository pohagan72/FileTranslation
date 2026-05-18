"""MCP server exposing the FileTranslation service as agent tools.

Two transport-bound entry points share the schemas in `tools.py`:
  - `local_server` (stdio, in-process `TranslationService`)
  - `remote_server` (Streamable HTTP, calls upstream `/api/v1/*`)
"""

__version__ = "0.1.0"
