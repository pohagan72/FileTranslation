# FileTranslation

AI-powered translation service for Microsoft Office documents (`.docx`,
`.pptx`, `.xlsx`) backed by Google Gemini, designed to run on Google Cloud Run
with Google Cloud Storage as the file-staging layer.

> **Status:** demo / portfolio project. The architecture and code-quality
> bars below are real, but a production deployment serving untrusted traffic
> would still benefit from an async job queue (Cloud Tasks + worker), per-
> tenant quotas, and end-to-end integration tests. See *Known limits* below.

## Features

- Translates body text, table cells, and PowerPoint text frames into one of
  several target languages.
- Source-language detection via `langdetect`.
- Per-segment translation parallelised through a `ThreadPoolExecutor`, with
  retry + exponential backoff.
- Direct-from-GCS downloads via signed URLs — the app never proxies bytes.
- Three interfaces:
  - HTML form at `/`
  - JSON API at `/api/v1/*`
  - **MCP server** (Model Context Protocol) for agent integration — see
    [`mcp_server/`](mcp_server/). Same domain logic, exposed as agent tools
    over stdio (local) or Streamable HTTP (remote).
- `/healthz` for Cloud Run probes; `/api/v1/health` reports service status.
- Optional API-key authentication on `POST /api/v1/translations`.
- Every log line and HTTP response carries an `X-Request-Id` for tracing.

## Architecture

```
app/
├── __init__.py          create_app() factory + request_id hook
├── config.py            dataclass-based env config
├── logging_setup.py     structured logging + RequestIdFilter
├── api/                 JSON API blueprint
├── web/                 HTML form blueprint
└── core/                framework-independent domain logic
    ├── language.py
    ├── translation_service.py
    ├── providers/       TranslationProvider interface + GeminiProvider
    ├── readers/         DocumentHandler interface + per-format impls
    │                    registered via app/core/readers/registry.py
    └── storage.py       StorageBackend interface + GCSBackend

mcp_server/              MCP server companion package — separate deployable;
                         imports `app.core` for the local-stdio mode and
                         calls `/api/v1/*` for the remote-HTTP mode.
```

The `core/` package has no Flask or HTTP dependencies, so each piece is unit
testable and swappable. Adding a new translation provider (OpenAI, DeepL) means
adding one class under `providers/`; adding a new file format means adding one
class under `readers/` and registering it. The same independence is what lets
[`mcp_server/`](mcp_server/) wrap `TranslationService` as agent tools without
re-implementing any business logic.

## Local development

```bash
python -m venv .venv && source .venv/bin/activate    # Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt
cp .env.example .env                                  # then fill in real values
python wsgi.py
```

### Environment variables

| Variable                     | Required? | Purpose                                                |
|------------------------------|-----------|--------------------------------------------------------|
| `SECRET_KEY`                 | yes       | Flask session signing (32+ random chars)               |
| `GOOGLE_API_KEY`             | yes       | Gemini API key                                         |
| `GEMINI_MODEL`               | yes       | e.g. `gemini-1.5-flash`                                |
| `GOOGLE_CLOUD_PROJECT`       | yes       | GCP project ID                                         |
| `GCS_BUCKET_NAME`            | yes       | Staging bucket for uploads/translations                |
| `API_KEY`                    | optional  | If set, `POST /api/v1/translations` requires `X-API-Key` |
| `MAX_CONTENT_LENGTH_BYTES`   | optional  | Upload size limit (default 25 MiB)                     |
| `TRANSLATION_THREADS`        | optional  | Concurrent Gemini calls (default 8)                    |
| `SIGNED_URL_EXPIRY_MINUTES`  | optional  | Download-link lifetime (default 5)                     |

## Running checks locally

```bash
pytest -q                              # 21 unit + route tests
black --check app tests wsgi.py        # formatting
flake8 app tests wsgi.py               # lint
mypy                                   # type check
```

CI runs the same four commands on every push and pull request via
[.github/workflows/ci.yml](.github/workflows/ci.yml).

## API

```
GET  /healthz                              → { status: ok }
GET  /api/v1/health                        → service status, supported extensions
GET  /api/v1/languages                     → { languages: [...] }
POST /api/v1/translations                  ← multipart upload
       headers:
         X-API-Key: <key>                  (required if API_KEY is set)
       body:
         file: <docx|pptx|xlsx>
         target_language: "Spanish"
       → 201 { job_id, download_url, download_filename, detected_language }
       → 401 if API key missing/wrong
       → 400 on bad input, 413 if upload exceeds size limit, 503 if degraded
```

Every response includes an `X-Request-Id` header. Clients may set their own
`X-Request-Id` to propagate trace IDs through a load balancer or CDN.

## Deployment (Cloud Run)

1. Build & push the image (Artifact Registry recommended).
2. Configure secrets in Secret Manager (`SECRET_KEY`, `GOOGLE_API_KEY`,
   `API_KEY`) and reference them from the Cloud Run service definition.
3. Configure a GCS lifecycle policy on the staging bucket to delete translated
   files after 24h. The app does not delete them inline because users may still
   hold a signed-URL link.
4. The service account attached to Cloud Run needs `roles/storage.objectAdmin`
   on the bucket and must be able to sign URLs (default Cloud Run service
   accounts can; custom ones may need `roles/iam.serviceAccountTokenCreator`).

## Known limits

Honest list of what would change for a production deployment:

- **Synchronous translation:** the request blocks until translation is done.
  A 1000-segment document holds a Gunicorn worker for the whole duration.
  Real production would use Cloud Tasks + a background worker writing job
  status to Firestore, with the API returning a `job_id` to poll.
- **No per-tenant rate limiting:** the API key is single-tenant. Multiple
  consumers would need either a per-key bucket or fronting via Cloud Endpoints.
- **No integration tests against real Gemini/GCS:** unit tests use fakes.
  Recording real provider responses with `vcrpy` would catch contract drift.
- **DOCX style preservation is best-effort:** font properties from the first
  run of a paragraph are copied to the translated text. Mixed inline styling
  (bold word inside a sentence) is lost.
