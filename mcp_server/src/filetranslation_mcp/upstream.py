"""HTTP client wrapper around the Flask `/api/v1/*` API.

Used only by the remote MCP server. Forwards `X-API-Key` and `X-Request-Id`
so the full chain (agent → MCP → API → Gemini) shares a single trace ID
matching the request-id pattern in `app/__init__.py:73-82`.
"""

from __future__ import annotations

import logging
from typing import IO, Any, Optional

import httpx

from .errors import from_upstream_http

logger = logging.getLogger(__name__)


class UpstreamClient:
    """Thin wrapper over the Flask JSON API."""

    def __init__(
        self,
        base_url: str,
        api_key: Optional[str] = None,
        timeout_seconds: float = 60.0,
    ):
        self._base = base_url.rstrip("/")
        self._api_key = api_key
        # A single client reuses connections across tool calls within a session.
        self._http = httpx.Client(timeout=timeout_seconds)

    def close(self) -> None:
        self._http.close()

    def __enter__(self) -> "UpstreamClient":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    # -- read endpoints ------------------------------------------------------

    def health(self, request_id: str) -> dict[str, Any]:
        return self._get_json("/api/v1/health", request_id=request_id, requires_auth=False)

    def languages(self, request_id: str) -> dict[str, Any]:
        return self._get_json("/api/v1/languages", request_id=request_id, requires_auth=False)

    # -- write endpoints -----------------------------------------------------

    def translate(
        self,
        *,
        file_stream: IO[bytes],
        filename: str,
        target_language: str,
        request_id: str,
    ) -> dict[str, Any]:
        """POST a multipart upload to /api/v1/translations.

        NOTE: This requires the agent to hand us actual bytes — which is the
        whole reason the remote server prefers a `gcs_key` source (see
        TODO.md). For HttpUrlSource and Base64Source, the remote_server
        materializes bytes before calling this.
        """
        files = {"file": (filename, file_stream)}
        data = {"target_language": target_language}
        headers = self._headers(request_id=request_id, requires_auth=True)
        resp = self._http.post(
            f"{self._base}/api/v1/translations",
            files=files,
            data=data,
            headers=headers,
        )
        return self._parse(resp)

    # -- internals -----------------------------------------------------------

    def _get_json(self, path: str, *, request_id: str, requires_auth: bool) -> dict[str, Any]:
        resp = self._http.get(
            f"{self._base}{path}",
            headers=self._headers(request_id=request_id, requires_auth=requires_auth),
        )
        return self._parse(resp)

    def _headers(self, *, request_id: str, requires_auth: bool) -> dict[str, str]:
        h: dict[str, str] = {"X-Request-Id": request_id}
        if requires_auth and self._api_key:
            h["X-API-Key"] = self._api_key
        return h

    @staticmethod
    def _parse(resp: httpx.Response) -> dict[str, Any]:
        try:
            body = resp.json()
        except ValueError:
            body = {"raw": resp.text}
        if resp.status_code >= 400:
            raise from_upstream_http(resp.status_code, body)
        return body
