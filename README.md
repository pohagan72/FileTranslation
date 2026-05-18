# FileTranslation

AI-powered translation service for Microsoft Office documents (`.docx`, `.pptx`,
`.xlsx`) backed by Google Gemini, designed to run on Google Cloud Run with
Google Cloud Storage as the file-staging layer.

## Features

- Translates body text, table cells, and (for PowerPoint) text frames into one
  of several target languages.
- Source-language detection via `langdetect`.
- Parallel per-segment translation with retry + exponential backoff.
- Direct-from-GCS downloads via signed URLs — the app does not proxy file bytes.
- Two interfaces:
  - HTML form at `/`
  - JSON API at `/api/v1/*`
- Health endpoint at `/healthz` plus `/api/v1/health` with service status.

## Architecture

```
app/
├── __init__.py          create_app() factory
├── config.py            dataclass-based env config
├── logging.py           stdout structured logging
├── api/                 JSON API blueprint
├── web/                 HTML form blueprint
└── core/                framework-independent domain logic
    ├── language.py
    ├── translation_service.py
    ├── providers/       TranslationProvider interface + GeminiProvider
    ├── readers/         DocumentHandler interface + per-format implementations
    │                    registered via app/core/readers/registry.py
    └── storage.py       StorageBackend interface + GCSBackend
```

The `core/` package has no Flask or HTTP dependencies, so each piece is unit
testable and swappable. Adding a new translation provider (OpenAI, DeepL) means
adding one class under `providers/`; adding a new file format means adding one
class under `readers/` and registering it.

## Local development

```bash
python -m venv .venv && source .venv/bin/activate    # Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt
cp .env.example .env                                  # then fill in real values
python wsgi.py
```

Required environment variables:

| Variable                     | Purpose                                                  |
|------------------------------|----------------------------------------------------------|
| `SECRET_KEY`                 | Flask session signing (32+ random chars)                 |
| `GOOGLE_API_KEY`             | Gemini API key                                           |
| `GEMINI_MODEL`               | e.g. `gemini-1.5-flash`                                  |
| `GOOGLE_CLOUD_PROJECT`       | GCP project ID                                           |
| `GCS_BUCKET_NAME`            | Staging bucket for uploads/translations                  |
| `MAX_CONTENT_LENGTH_BYTES`   | Upload size limit (default 25 MiB)                       |
| `TRANSLATION_THREADS`        | Concurrent Gemini calls (default 8)                      |
| `SIGNED_URL_EXPIRY_MINUTES`  | Download-link lifetime (default 15)                      |

## Running tests

```bash
pytest -q          # ~10 unit + route tests
black --check .    # formatting
flake8             # lint
```

CI runs the same three commands on every push and pull request via
`.github/workflows/ci.yml`.

## Deployment (Cloud Run)

1. Build & push the image (Artifact Registry recommended).
2. Configure secrets in Secret Manager and reference them from the Cloud Run
   service definition.
3. Configure a GCS lifecycle policy on the staging bucket to delete translated
   files after 24h — the app does not delete them inline because users may
   still hold a signed-URL link.
4. The service account attached to Cloud Run needs `roles/storage.objectAdmin`
   on the bucket and must be able to sign URLs (default Cloud Run service
   accounts can; custom ones may need `roles/iam.serviceAccountTokenCreator`).

## API

```
GET  /healthz
GET  /api/v1/health
GET  /api/v1/languages
POST /api/v1/translations
       multipart/form-data:
         file: <docx|pptx|xlsx>
         target_language: "Spanish"
       → 201 { job_id, download_url, download_filename, detected_language }
```
