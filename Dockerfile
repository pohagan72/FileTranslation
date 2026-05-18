# Lightweight, supported Python base image.
FROM python:3.12-slim-bookworm

# Don't write .pyc files; flush stdout/stderr (so Cloud Run logs are realtime).
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app

# Install dependencies first so this layer caches when only code changes.
COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

# Copy the application package.
COPY app ./app
COPY templates ./templates
COPY wsgi.py ./

# Non-root user — required by some Cloud Run hardening profiles, good practice anyway.
RUN useradd --create-home --uid 10001 appuser
USER appuser

EXPOSE 8080

# Shell form so ${PORT} expands at runtime. --timeout 0 because Cloud Run
# enforces its own request timeout and Gemini calls can exceed Gunicorn's 30s default.
CMD exec gunicorn --workers 1 --threads 8 --timeout 0 --bind 0.0.0.0:${PORT:-8080} wsgi:app
