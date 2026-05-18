# Use a lightweight Python image as the base image
# We use a specific version with -slim and a known distribution (bookworm) for stability
FROM python:3.9-slim-bookworm

# Set the working directory in the container to /app
WORKDIR /app

# Copy the requirements file into the working directory
# Copying requirements.txt first allows Docker to cache this layer if only the code changes
COPY requirements.txt .

# Install Python dependencies
# We upgrade pip first to ensure we have the latest version
RUN pip install --upgrade pip
# Install the dependencies from requirements.txt
RUN pip install -r requirements.txt

# Copy the rest of the application code into the working directory
# The .dockerignore file prevents unwanted files from being copied
COPY . .

# Expose the port that the container will listen on
# Cloud Run expects the application to listen on the port specified by the PORT environment variable (default is 8080)
# We expose 8080 here as a standard convention, although Gunicorn will use $PORT
EXPOSE 8080

# Define the command to run your application using Gunicorn.
# Shell form is required so ${PORT} expands at runtime (Cloud Run injects PORT).
# --timeout 0 disables Gunicorn's request timeout — Cloud Run enforces its own,
# and Gemini translation can exceed the 30s default.
CMD exec gunicorn --workers 1 --threads 8 --timeout 0 --bind 0.0.0.0:${PORT:-8080} main:app

# Optional: Switch to a non-root user for better security (good practice)
# RUN useradd appuser
# USER appuser