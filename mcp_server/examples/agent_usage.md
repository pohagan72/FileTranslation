# Agent usage walkthrough

A narrative example of an agent using the FileTranslation MCP server. This
is the interaction *shape* — exact wire format depends on the host.

## Scenario

A localization analyst has a quarterly board deck (`Q3-board.pptx`) and asks
their agent: *"Translate this into Spanish and German and give me both files."*

The agent has the `filetranslation` MCP server registered (local stdio mode,
in this example).

## Turn-by-turn

**1. Agent confirms the request is in scope.**

Calls `list_supported_languages`:

```json
{}
```

Returns:

```json
{ "languages": ["English", "Spanish", "French", "German", "Chinese", "Japanese"] }
```

Both targets are supported — the agent proceeds. (If Spanish were missing, it
would tell the user and offer the closest supported language.)

**2. Agent confirms the service is healthy.**

Calls `get_translation_service_status`:

```json
{}
```

Returns:

```json
{
  "status": "ok",
  "gemini_configured": true,
  "gcs_configured": true,
  "supported_extensions": [".docx", ".pptx", ".xlsx"]
}
```

`.pptx` is supported. Good.

**3. Agent translates into Spanish.**

Calls `translate_document`:

```json
{
  "source": { "kind": "path", "path": "C:/Users/analyst/Documents/Q3-board.pptx" },
  "target_language": "Spanish"
}
```

Returns:

```json
{
  "job_id": "8f3a...",
  "download_url": "https://storage.googleapis.com/...",
  "download_filename": "translated_Q3-board.pptx",
  "detected_language": "en",
  "expires_in_seconds": 300
}
```

**4. Agent translates into German.**

Same shape, different `target_language`. Returns a second signed URL.

**5. Agent hands both URLs back to the user.**

Crucially, the agent does **not** put the document bytes into its own context.
It either:

- Tells the user "here are two download links, fetch them within 5 minutes",
  or
- Uses a separate file-download tool from its host to save them to a known
  location and then references those paths.

The latter is preferable for multi-turn workflows because the signed URL
expiry is short.

## Failure modes the agent should handle

| Error code | Meaning | Suggested agent behavior |
|---|---|---|
| `unsupported_file_type` | Extension not in `supported_extensions` | Tell the user, suggest converting to `.docx`/`.pptx`/`.xlsx`. |
| `service_unavailable` | Upstream booted degraded | Tell the user the service is down; do not retry in a tight loop. |
| `unauthorized` | API key missing/wrong (remote mode) | Surface to the operator — this is a config issue, not a user issue. |
| `file_too_large` | Exceeds upload limit (default 25 MiB) | Suggest splitting the document. |
| `bad_source` | Source kind not valid for this server, or file not found | Ask the user for a different path/URL. |
| `upstream_failure` | Unexpected upstream error | One retry with backoff; if it persists, surface to the user. |

## Why no streaming progress?

The upstream translation is synchronous; faking progress events would lie to
the user. When the upstream service grows an async-job mode (Cloud Tasks +
Firestore), this MCP server will add a `get_translation_job` tool and a
progress notification channel.
