# Import necessary modules and functions
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import os
from docx import Document  # To handle Word documents
from pptx import Presentation  # To handle PowerPoint presentations
import pandas as pd  # To handle Excel files
import requests  # For making HTTP requests to external APIs
import io  # For handling streams of data
import json  # For handling JSON serialization and deserialization

# Create a new instance of the Flask application
app = Flask(__name__)
# Set a secret key needed by Flask to securely sign session cookies
app.secret_key = "supersecretkey"

# Configure the folder where uploaded files are stored
UPLOAD_FOLDER = "uploads"
# Ensure that the upload folder exists by creating it if necessary
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --------------------
# API Endpoint Configurations
# --------------------

# URL for the Ollama API which is used for processing certain translation tasks
OLLAMA_URL = "http://localhost:11434/api/chat"

# Azure OpenAI settings - used for internal testing (Do not edit the hardcoded credentials)
AZURE_ENDPOINT = "Endpoint_goes_here"
API_KEY = "Key_goes_here"
API_VERSION = "API_version_goes_here"
AZURE_MODEL = "Model_goes_here"

# List of supported target languages for translation.
LANGUAGES = ["English", "Spanish", "French", "German", "Chinese", "Japanese"]

# --------------------
# Utility Functions
# --------------------

def get_ollama_models():
    """
    Fetch the list of available models from the Ollama API.
    Returns a list of model names.
    If an error occurs, flashes an error to the user and returns an empty list.
    """
    try:
        # Make a GET request to fetch available models
        response = requests.get("http://localhost:11434/api/tags")
        response.raise_for_status()  # Raise an error for bad status codes
        # Parse the returned JSON to extract model names
        models = response.json().get("models", [])
        return [model["name"] for model in models]
    except requests.exceptions.RequestException as e:
        # Log error and inform the user via flashing a message
        print(f"Failed to fetch Ollama models: {e}")
        flash(f"Failed to fetch Ollama models: {e}")
        return []

def detect_language(text):
    """
    Detects the language of a given text using the 'langdetect' module.
    If the module is not installed, attempts to install it using pip.
    Returns the detected language code or None if an error occurs.
    """
    try:
        from langdetect import detect
    except ImportError:
        # Attempt to install langdetect if it is not already installed.
        try:
            import subprocess
            subprocess.check_call(["pip", "install", "langdetect"])
            from langdetect import detect
        except subprocess.CalledProcessError as e:
            print(f"Error installing langdetect: {e}")
            flash(f"Error installing langdetect. Please ensure pip is installed and try again. Error: {e}")
            return None
    try:
        # Use langdetect's detect method to guess the language
        return detect(text)
    except Exception as e:
        # Handle any error that might occur during detection
        print(f"Language detection error: {e}")
        flash(f"Language detection error: {e}")
        return None

def translate_text(text, target_lang, model_name, model_type, detected_lang=None):
    """
    Translate a given text into the target language using either the Ollama or Azure OpenAI API.
    
    Parameters:
      text         : The text that needs to be translated.
      target_lang  : The target language for the translation.
      model_name   : The name of the selected model.
      model_type   : The type of the selected API ('ollama' or 'azure').
      detected_lang: (Optional) The detected language of the input text.
    
    Returns the translated text, or if an error occurs, returns the original text.
    """
    try:
        # Build a prompt prefix message, include detected language if available
        if detected_lang:
            prompt_prefix = f"Translate the following text from {detected_lang} into {target_lang}."
        else:
            prompt_prefix = f"Translate the following text into {target_lang}."

        # If using the Ollama API
        if model_type == "ollama":
            payload = {
                "model": model_name,
                "messages": [
                    {"role": "system", "content": prompt_prefix},
                    {"role": "user", "content": text}
                ],
                "stream": False
            }
            response = requests.post(OLLAMA_URL, json=payload)
            response.raise_for_status()  # Check for errors in the response
            return response.json()["message"]["content"]

        # If using the Azure OpenAI API
        elif model_type == "azure":
            # Format the API URL with correct deployment and version information
            api_url = f"{AZURE_ENDPOINT}openai/deployments/{AZURE_MODEL}/chat/completions?api-version={API_VERSION}"
            headers = {'Content-Type': 'application/json', 'api-key': API_KEY}
            payload = json.dumps({
                "messages": [
                    {"role": "system", "content": prompt_prefix},
                    {"role": "user", "content": text}
                ]
            })
            response = requests.post(api_url, headers=headers, data=payload)
            response.raise_for_status()  # Validate response status
            return response.json()["choices"][0]["message"]["content"]
        else:
            flash(f"Invalid model type selected: {model_type}")
            return text  # Return original text if model type is not recognized

    except requests.exceptions.RequestException as e:
        # Log HTTP request errors and inform the user via flash
        print(f"{model_type.capitalize()} API error: {e}")
        flash(f"{model_type.capitalize()} API error: {e}")
        return text

    except KeyError as e:
        # Handle missing keys in the JSON response
        print(f"Unexpected JSON response from {model_type}: {e}")
        print(f"Response content: {response.content if 'response' in locals() else 'No response'}")
        flash(f"Unexpected JSON response from {model_type}. Check the server and model. Error: {e}")
        return text

def read_text_from_docx(filepath):
    """
    Reads and extracts text content from a DOCX (Word) file.
    In case of error, logs and flashes an error message and returns an empty string.
    """
    try:
        doc = Document(filepath)
        full_text = []
        # Loop through each paragraph in the document and accumulate the text
        for paragraph in doc.paragraphs:
            full_text.append(paragraph.text)
        return "\n".join(full_text)
    except Exception as e:
        print(f"Error reading DOCX: {e}")
        flash(f"Error reading DOCX file: {e}")
        return ""

def read_text_from_pptx(filepath):
    """
    Reads and extracts text content from a PPTX (PowerPoint) file.
    In case of error, logs and flashes an error message and returns an empty string.
    """
    try:
        ppt = Presentation(filepath)
        full_text = []
        # Iterate over each slide and then each shape in the slide
        for slide in ppt.slides:
            for shape in slide.shapes:
                # Check if the shape has a text attribute then accumulate it
                if hasattr(shape, "text"):
                    full_text.append(shape.text)
        return "\n".join(full_text)
    except Exception as e:
        print(f"Error reading PPTX: {e}")
        flash(f"Error reading PPTX file: {e}")
        return ""

def read_text_from_excel(filepath):
    """
    Reads and extracts all text from an Excel file by converting it into a string.
    In case of error, logs and flashes an error message and returns an empty string.
    """
    try:
        df = pd.read_excel(filepath)
        # Convert entire DataFrame to string without headers or indexes
        full_text = df.astype(str).to_string(header=False, index=False)
        return full_text
    except Exception as e:
        print(f"Error reading Excel: {e}")
        flash(f"Error reading Excel file: {e}")
        return ""

def translate_docx(filepath, target_lang, model_name, model_type, detected_lang):
    """
    Translates the text in a DOCX file by replacing each paragraph with its translation.
    Returns a Document object with the translated content or None if an error occurs.
    """
    try:
        doc = Document(filepath)
        for paragraph in doc.paragraphs:
            # Process non-empty paragraphs only
            if paragraph.text.strip():
                translated = translate_text(paragraph.text, target_lang, model_name, model_type, detected_lang)
                paragraph.text = translated  # Replace the original text with the translated text
        return doc
    except Exception as e:
        print(f"Error translating DOCX: {e}")
        flash(f"Error translating DOCX file: {e}")
        return None

def translate_pptx(filepath, target_lang, model_name, model_type, detected_lang):
    """
    Translates the text in a PPTX file by replacing text in each shape with its translation.
    Returns a Presentation object with the translated content or None if an error occurs.
    """
    try:
        ppt = Presentation(filepath)
        for slide in ppt.slides:
            for shape in slide.shapes:
                # Check that shape has a text attribute and contains non-empty text
                if hasattr(shape, "text") and shape.text.strip():
                    translated = translate_text(shape.text, target_lang, model_name, model_type, detected_lang)
                    shape.text = translated  # Replace original text with translation
        return ppt
    except Exception as e:
        print(f"Error translating PPTX: {e}")
        flash(f"Error translating PPTX file: {e}")
        return None

def translate_excel(filepath, target_lang, model_name, model_type, detected_lang):
    """
    Translates the content of an Excel file.
    For each cell that contains data, it converts the value to a string and applies translation.
    Returns the modified DataFrame, or None if an error occurs.
    """
    try:
        df = pd.read_excel(filepath)
        # Apply the translation function on every cell in each column of the DataFrame
        for col in df.columns:
            df[col] = df[col].apply(lambda x: translate_text(str(x), target_lang, model_name, model_type, detected_lang) if pd.notna(x) else x)
        return df
    except Exception as e:
        print(f"Error translating Excel: {e}")
        flash(f"Error translating Excel file: {e}")
        return None

# --------------------
# Flask Route Handlers (Views)
# --------------------

@app.route("/", methods=["GET", "POST"])
def index():
    """
    Main endpoint for the web application.
    On GET request: Renders the homepage with the upload form and selection of target language and AI model.
    On POST request: Processes the uploaded file, performs language detection, translates the file,
                     and then provides a download link for the translated file.
    """
    # Fetch available models from Ollama
    ollama_models = get_ollama_models()
    # Prepare a list of groups of AI models (one for Ollama and one hardcoded for Azure OpenAI)
    ai_models = [
        {"type": "ollama", "models": ollama_models},
        {"type": "azure", "models": ["Azure OpenAI"]}
    ]

    if request.method == "POST":
        # Get the uploaded file from the form
        file = request.files["file"]
        target_lang = request.form["target_language"]
        ai_model_selection = request.form["ai_model"]  # This can be "Azure OpenAI" or an Ollama model name

        # Check to ensure a file was actually selected
        if file.filename == "":
            flash("No file selected!")
            return redirect(url_for("index"))

        # Save the uploaded file to the designated upload folder
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        detected_language = None  # Variable to hold the detected language code

        # Attempt to detect the language using the file's content
        try:
            # Process file based on its extension type
            if file.filename.endswith(".docx"):
                file_content = read_text_from_docx(filepath)
            elif file.filename.endswith(".pptx"):
                file_content = read_text_from_pptx(filepath)
            elif file.filename.endswith(".xlsx"):
                file_content = read_text_from_excel(filepath)
            else:
                flash("Unsupported file type!")
                os.remove(filepath)  # Remove the file if it is unsupported
                return redirect(url_for("index"))

            # If content was successfully read, attempt language detection
            if file_content:
                detected_language = detect_language(file_content)
                if detected_language:
                    flash(f"Detected language: {detected_language}")
                else:
                    flash("Could not detect language.")
        except Exception as e:
            flash(f"Error detecting language: {str(e)}")
            # Continue without detected_language if an error occurred

        # Determine which model type to use based on the user selection
        if ai_model_selection == "Azure OpenAI":
            model_type = "azure"
            model_name = AZURE_MODEL  # Use the hardcoded Azure model name
        else:
            model_type = "ollama"
            model_name = ai_model_selection

        # Variables to hold the file name and path of the translated document
        translated_filename = None
        translated_filepath = None

        try:
            # Process and translate the file based on its type
            if file.filename.endswith(".docx"):
                translated_doc = translate_docx(filepath, target_lang, model_name, model_type, detected_language)
                if translated_doc:
                    translated_filename = f"translated_{file.filename}"
                    translated_filepath = os.path.join(UPLOAD_FOLDER, translated_filename)
                    translated_doc.save(translated_filepath)
            elif file.filename.endswith(".pptx"):
                translated_ppt = translate_pptx(filepath, target_lang, model_name, model_type, detected_language)
                if translated_ppt:
                    translated_filename = f"translated_{file.filename}"
                    translated_filepath = os.path.join(UPLOAD_FOLDER, translated_filename)
                    translated_ppt.save(translated_filepath)
            elif file.filename.endswith(".xlsx"):
                translated_df = translate_excel(filepath, target_lang, model_name, model_type, detected_language)
                if translated_df is not None:
                    translated_filename = f"translated_{file.filename}"
                    translated_filepath = os.path.join(UPLOAD_FOLDER, translated_filename)
                    translated_df.to_excel(translated_filepath, index=False)
            else:
                flash("Unsupported file type!")
                os.remove(filepath)  # Clean up the file if unsupported
                return redirect(url_for("index"))

            # If a translated file was successfully created, notify the user
            if translated_filename:
                flash("Translation completed successfully!")
        except Exception as e:
            flash(f"An unexpected error occurred during translation: {e}")
            print(f"An unexpected error occurred during translation: {e}")
        finally:
            # Remove the original uploaded file to save disk space
            try:
                os.remove(filepath)
            except OSError as e:
                print(f"Error deleting the uploaded file: {e}")

        # Render the template with the translated file download link if available
        return render_template("index.html", languages=LANGUAGES, ai_models=ai_models, translated_file=translated_filename)

    # For GET requests, simply render the homepage with the file upload form.
    return render_template("index.html", languages=LANGUAGES, ai_models=ai_models)

@app.route("/download/<filename>")
def download_file(filename):
    """
    Allows users to download a translated file.
    The file is sent from the uploads folder as an attachment to force download.
    """
    return send_file(os.path.join(UPLOAD_FOLDER, filename), as_attachment=True)

# Run the Flask app in debug mode if this script is executed directly.
if __name__ == "__main__":
    app.run(debug=True)
