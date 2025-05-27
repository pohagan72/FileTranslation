# FileTranslation - AI-Powered Document Translation Service

This project delivers a web-based solution for translating text content within common business document formats: Microsoft Word (.docx), PowerPoint (.pptx), and Excel (.xlsx). Leveraging Google's Generative AI capabilities via the Gemini API and Google Cloud Storage for efficient file handling, this service provides a streamlined approach to document localization.

## Key Capabilities

*   **Multilingual Translation:** Supports translation of document text into several languages including English, Spanish, French, German, Chinese, and Japanese.
*   **Broad Format Support:** Processes content from `.docx`, `.pptx`, and `.xlsx` files.
*   **Scalable Architecture:** Designed for deployment on Google Cloud Run, utilizing Docker containers for portability and scalability.
*   **Secure File Handling:** Employs Google Cloud Storage for temporary file management during the translation process, enhancing security and handling large documents effectively.
*   **Streamlined User Interface:** Provides a simple web interface for easy file upload and download of translated documents.

## Technical Prerequisites

To deploy and run this application, the following technical components are required:

### Software Requirements

*   Python 3.9+
*   Docker Engine
*   Google Cloud SDK (`gcloud`)

### Google Cloud Platform Requirements

*   A Google Cloud Project.
*   Enabled GCP APIs:
    *   Generative AI API (for Gemini access)
    *   Cloud Storage API
    *   Cloud Build API (for container image builds)
    *   Cloud Run API (for service deployment)
    *   Secret Manager API (Recommended for secure secret storage)
*   A **Google Cloud Storage Bucket:** Required for temporary storage of files during translation. The Cloud Run service account will need appropriate permissions (Object Creator, Object Viewer, Object Deleter) on this bucket.
*   A **Google Gemini API Key:** Necessary for authenticating with the Generative AI service.

### Dependencies (`requirements.txt`)

All required Python libraries are specified in the `requirements.txt` file. This includes:

*   `Flask`: The web framework.
*   `python-docx`: For processing Word files.
*   `python-pptx`: For processing PowerPoint files.
*   `pandas` and `openpyxl`: For processing Excel files.
*   `google-generativeai`: The client library for accessing Gemini.
*   `google-cloud-storage`: The client library for interacting with GCS.
*   `langdetect`: For source language detection.
*   `gunicorn`: A production-ready WSGI server for deploying the Flask app.
*   Other necessary libraries (`python-dotenv`, `requests`).

It is strongly recommended to use a pinned `requirements.txt` file (generated via `pip freeze > requirements.txt`) to ensure consistent dependency versions across development and deployment environments.

## Deployment Strategy

This application is designed for deployment as a containerized service on Google Cloud Run. The deployment process involves:

1.  **Containerization:** Building a Docker image using the provided `Dockerfile`, which includes the application code and all dependencies from `requirements.txt`.
2.  **Image Hosting:** Pushing the Docker image to Google Artifact Registry (recommended) or Google Container Registry.
3.  **Secret Management:** Storing sensitive configuration values (like the Gemini API Key and Flask Secret Key) securely in Google Cloud Secret Manager.
4.  **Cloud Run Service Configuration:** Deploying the container image to a Cloud Run service, configuring environment variables (GCP Project ID, GCS Bucket Name, Gemini Model) and linking the secrets from Secret Manager. The service will run the Flask application using the Gunicorn WSGI server configured via the `Dockerfile`.
5.  **GCS Lifecycle Policy:** Configuring a Lifecycle Management policy on the GCS bucket to automatically clean up translated files after a specified duration (e.g., 24 hours).

Detailed build and deployment commands using `gcloud` are typically executed after setting up the GCP project, APIs, bucket, and secrets.

## Local Development Setup

For development and testing prior to deployment:

1.  Ensure Python, Docker, and the `gcloud` SDK are installed.
2.  Clone the repository.
3.  Set up a Python virtual environment and install dependencies using `pip install -r requirements.txt`.
4.  Create a `.env` file in the project root with placeholder values for `GOOGLE_API_KEY`, `GOOGLE_CLOUD_PROJECT`, `GCS_BUCKET_NAME`, `GEMINI_MODEL`, and `SECRET_KEY`. **Note:** This file is for local use only and is excluded from the Docker image by `.dockerignore`.
5.  Run `python main.py` to start the Flask development server.

## Usage

Access the application via its deployed Cloud Run URL (NOT YET LIVE). Use the interface to upload your document, select the target language, initiate translation, and download the resulting file.

---