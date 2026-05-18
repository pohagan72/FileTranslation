"""MCP tool schemas — shared between local (stdio) and remote (HTTP) servers.

Schemas are Pydantic models so the MCP SDK can derive JSON Schema from them
and validate tool input before dispatch. The actual tool implementations live
in `local_server.py` and `remote_server.py`; this module is intentionally
free of transport or service-call logic.

Tool names are exported as constants so the two server modules can register
them without string drift.
"""

from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, Field, TypeAdapter, model_validator

# Tool names — single source of truth.
TOOL_LIST_LANGUAGES = "list_supported_languages"
TOOL_GET_STATUS = "get_translation_service_status"
TOOL_TRANSLATE = "translate_document"


# -- list_supported_languages -------------------------------------------------


class ListLanguagesInput(BaseModel):
    """No input. Returns the configured target languages."""


class ListLanguagesOutput(BaseModel):
    languages: list[str] = Field(
        ..., description="Human-readable target language names accepted by translate_document."
    )


# -- get_translation_service_status -------------------------------------------


class StatusInput(BaseModel):
    """No input. Returns health and capability of the upstream service."""


class StatusOutput(BaseModel):
    status: Literal["ok", "degraded"]
    gemini_configured: bool
    gcs_configured: bool
    supported_extensions: list[str] = Field(
        ..., description="File extensions accepted by translate_document, e.g. ['.docx', '.pptx']."
    )


# -- translate_document -------------------------------------------------------


class LocalPathSource(BaseModel):
    """Local filesystem source. Only valid for the local stdio server."""

    kind: Literal["path"] = "path"
    path: str = Field(..., description="Absolute path to a .docx, .pptx, or .xlsx file.")


class HttpUrlSource(BaseModel):
    """HTTPS URL the server fetches before translating.

    Only `https://` URLs are accepted to reduce SSRF risk. The remote server
    may further restrict this to an allow-list.
    """

    kind: Literal["url"] = "url"
    url: str = Field(..., description="https:// URL of the document to translate.")


class GcsKeySource(BaseModel):
    """A GCS object key previously written via a signed upload URL.

    Only valid for the remote server, and requires the upstream Flask API to
    expose a `POST /api/v1/uploads` companion endpoint (see TODO.md).
    """

    kind: Literal["gcs_key"] = "gcs_key"
    gcs_key: str = Field(..., description="Key in the staging bucket, returned by create_upload_url.")
    original_filename: str = Field(
        ..., description="Original filename; used for the translated download name and extension."
    )


class Base64Source(BaseModel):
    """Inline base64 source. Capped at 1 MiB; intended for tiny files and tests."""

    kind: Literal["base64"] = "base64"
    data: str = Field(..., description="Base64-encoded document bytes. Max ~1 MiB.")
    filename: str = Field(..., description="Filename including extension (.docx/.pptx/.xlsx).")

    @model_validator(mode="after")
    def _enforce_size_cap(self) -> "Base64Source":
        # Rough char→byte ratio for base64 is 4:3; 1 MiB ≈ 1_398_101 chars.
        if len(self.data) > 1_398_101:
            raise ValueError(
                "base64 source exceeds 1 MiB limit; use a url or gcs_key source instead"
            )
        return self


# Discriminated union: agents pick exactly one source shape, keyed by `kind`.
# Annotated with `discriminator="kind"` so FastMCP emits a tagged-union JSON
# Schema (better for LLM clients) and Pydantic validates against the right
# variant without trying all of them.
DocumentSource = Annotated[
    Union[LocalPathSource, HttpUrlSource, GcsKeySource, Base64Source],
    Field(discriminator="kind"),
]

# Reusable validator so FastMCP tool functions can take `source: dict` and
# normalize it into a typed variant in one line.
DocumentSourceAdapter: TypeAdapter[
    Union[LocalPathSource, HttpUrlSource, GcsKeySource, Base64Source]
] = TypeAdapter(DocumentSource)


class TranslateInput(BaseModel):
    """Translate a single Office document into `target_language`.

    Returns a signed download URL — the document bytes are never returned in
    the tool response. The URL has a short expiry (default 5 minutes); fetch
    it promptly with your host's HTTP/download tool.
    """

    source: DocumentSource = Field(
        ..., description="Where to read the document from. Exactly one of path/url/gcs_key/base64."
    )
    target_language: str = Field(
        ...,
        description="One of the names returned by list_supported_languages, e.g. 'Spanish'.",
        min_length=1,
    )


class TranslateOutput(BaseModel):
    job_id: str
    download_url: str = Field(
        ..., description="Signed URL; expires after `expires_in_seconds`. Fetch promptly."
    )
    download_filename: str
    expires_in_seconds: int = Field(
        ..., description="Lifetime of `download_url` from the moment this response was generated."
    )


# -- helper: shared tool descriptions ----------------------------------------
# Putting these here (rather than inline in the servers) so local and remote
# expose identical descriptions. Agents key off these heavily.

TOOL_DESCRIPTIONS: dict[str, str] = {
    TOOL_LIST_LANGUAGES: (
        "List the human-readable target languages this translation service accepts. "
        "Call this before translate_document to confirm the user's requested language is supported."
    ),
    TOOL_GET_STATUS: (
        "Report whether the translation service is fully configured and which file extensions "
        "it accepts. If status is 'degraded', translate_document will fail; surface that to the user."
    ),
    TOOL_TRANSLATE: (
        "Translate a Microsoft Office document (.docx, .pptx, .xlsx) into the requested language. "
        "Returns a signed download URL that expires shortly (default 5 minutes). "
        "Do NOT cache the URL across long delays — fetch the file promptly. "
        "The document bytes are never returned inline in this tool's response."
    ),
}
