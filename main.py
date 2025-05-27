# Import necessary modules and functions
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import os
from docx import Document  # To handle Word documents
from pptx import Presentation  # To handle PowerPoint presentations
import pandas as pd  # To handle Excel files
import requests  # For making HTTP requests (Removed specific Ollama/Azure imports)
import io  # For handling streams of data
import json  # For handling JSON serialization and deserialization
# Import libraries for .env and Google Gemini
from dotenv import load_dotenv
import google.generativeai as genai
from werkzeug.utils import secure_filename # Import for secure file handling

# Load environment variables from .env file
load_dotenv()

# Create a new instance of the Flask application
app = Flask(__name__)
# Set a secret key needed by Flask to securely sign session cookies
app.secret_key = os.getenv("SECRET_KEY", "supersecretkey") # Get secret key from .env or use default

# Configure the folder where uploaded files are stored
UPLOAD_FOLDER = "uploads"
# Ensure that the upload folder exists by creating it if necessary
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --------------------
# API Endpoint Configurations
# --------------------

# Google Gemini Configuration - loaded from .env
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL") # e.g., "gemini-2.0-flash"

# Configure Google Gemini API
if GOOGLE_API_KEY:
    try:
        genai.configure(api_key=GOOGLE_API_KEY)
        # Optional: Basic check if the model exists and is listable (can be slow)
        # models = [m.name for m in genai.list_models() if GEMINI_MODEL in m.name]
        # if GEMINI_MODEL not in models:
        #     print(f"Warning: Configured GEMINI_MODEL '{GEMINI_MODEL}' might not be available.")
        print("Google Gemini API configured successfully.")
    except Exception as e:
        print(f"Failed to configure Google Gemini API: {e}")
        GOOGLE_API_KEY = None # Disable Google option if configuration fails
        flash("Google Gemini API configuration failed. Check your API key and internet connection.")
else:
    flash("Google API Key not found in .env. Translation using Google Gemini is disabled.")


# List of supported target languages for translation.
LANGUAGES = ["English", "Spanish", "French", "German", "Chinese", "Japanese"]

# --------------------
# Utility Functions
# --------------------

# Removed get_ollama_models function

def detect_language(text):
    """
    Detects the language of a given text using the 'langdetect' module.
    If the module is not installed, attempts to install it using pip.
    Returns the detected language code (ISO 639-1) or None if an error occurs or text is empty.
    """
    if not text or not text.strip():
        return None # Cannot detect language of empty text

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
        except Exception as e:
            print(f"Unexpected error during langdetect installation: {e}")
            flash(f"Unexpected error during langdetect installation: {e}")
            return None

    try:
        # Use langdetect's detect method to guess the language
        # Only use a portion of text for detection to prevent excessive processing
        return detect(text[:2000]) # Limit text to first 2000 chars for speed
    except Exception as e:
        # Handle any error that might occur during detection (e.g., text is too short or contains no language)
        print(f"Language detection error: {e}")
        # Flash message might be annoying for expected cases like short text, handle gracefully
        # flash(f"Language detection error: {e}") # Removed flash for minor detection errors
        return None


def translate_text(text, target_lang, model_name, model_type, detected_lang=None):
    """
    Translate a given text into the target language using the Google Gemini API.

    Parameters:
      text         : The text that needs to be translated.
      target_lang  : The target language for the translation.
      model_name   : The name of the selected model (should be GEMINI_MODEL).
      model_type   : The type of the selected API ('google').
      detected_lang: (Optional) The detected language of the input text (ISO 639-1 code).

    Returns the translated text, or if an error occurs, returns the original text.
    """
    if not text or not text.strip():
        return "" # Don't translate empty strings

    # --- Google Gemini Translation ---
    if model_type == "google":
        if not GOOGLE_API_KEY or not model_name:
            flash("Google Gemini API is not configured.")
            print("Google Gemini API or model name missing for translation.")
            return text # Return original text if API is not configured

        try:
            # Build the prompt
            if detected_lang and detected_lang.lower() != target_lang.lower():
                 # Try to use the detected language if it's different from the target
                 prompt = f"Translate the following text from {detected_lang} into {target_lang}:"
            else:
                # Otherwise, just ask to translate to the target language
                prompt = f"Translate the following text into {target_lang}:"

            model = genai.GenerativeModel(model_name) # Use the GEMINI_MODEL name
            response = model.generate_content([prompt, text])

            # Check if response has text content
            if hasattr(response, 'text') and response.text is not None:
                 return response.text.strip() # Return stripped text
            else:
                 # Handle cases where the response might be empty, blocked, or contain safety issues
                 print(f"Gemini API returned no text or was blocked. Response: {response}")
                 # Check for safety ratings that might have blocked the content
                 if hasattr(response, 'prompt_feedback') and response.prompt_feedback.block_reason:
                     block_reason = response.prompt_feedback.block_reason.name
                     print(f"Gemini API blocked content due to: {block_reason}")
                     flash(f"Translation blocked by AI safety filters ({block_reason}).")
                 else:
                     flash(f"Gemini API returned no text.")
                 return text # Return original text if translation failed or was blocked

        except Exception as e:
             # Catch specific Google API errors or other issues during the call
             print(f"Google Gemini API error: {e}")
             flash(f"Google Gemini API error: {e}")
             return text # Return original text on error
    else:
        # This block should theoretically not be reached if only google model is offered
        flash(f"Invalid model type specified internally: {model_type}")
        return text


# --- File Reading Utility Functions (kept as is) ---

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
        # Note: This captures structure less well than cell-by-cell but works for detection
        full_text = df.to_string(header=True, index=False) # Include header for more context for detection
        return full_text
    except Exception as e:
        print(f"Error reading Excel: {e}")
        flash(f"Error reading Excel file: {e}")
        return ""

# --- File Translation Utility Functions (kept, but simplified model handling) ---

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
                # Preserve original run formatting if possible (simplified approach)
                # Clear existing runs
                paragraph.clear()
                # Add a single run with the translated text
                run = paragraph.add_run(translated)
                # Note: This simple approach loses original formatting (bold, italics, etc.)
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
                # Check that shape has a text frame and contains non-empty text
                if shape.has_text_frame and shape.text_frame.text.strip():
                    text_frame = shape.text_frame
                    # Translate the full text of the shape's text frame
                    original_text = text_frame.text
                    translated_text = translate_text(original_text, target_lang, model_name, model_type, detected_lang)

                    # Simple replacement: clears existing paragraphs and adds translated text
                    # This loses original paragraph/run formatting and structure within the shape
                    text_frame.clear()
                    p = text_frame.add_paragraph()
                    run = p.add_run()
                    run.text = translated_text
                    # Note: This also loses formatting.
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
        # Apply the translation function on every cell in the DataFrame
        # Using .applymap for element-wise application
        # Ensure we don't try to translate non-string types directly, convert to string first
        df = df.applymap(lambda x: translate_text(str(x) if pd.notna(x) else '', target_lang, model_name, model_type, detected_lang) if pd.notna(x) else x)
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
    # Prepare the list of AI models for the dropdown - ONLY Google Gemini
    ai_models = []
    if GOOGLE_API_KEY and GEMINI_MODEL:
        ai_models.append({"label": "Google Gemini", "models": [{"value": "google", "text": f"Google Gemini ({GEMINI_MODEL})"}]})
    else:
        # Add a placeholder/message if Google Gemini is not configured
        ai_models.append({"label": "Configuration Error", "models": [{"value": "none", "text": "Gemini not configured - check .env", "disabled": True}]})


    if request.method == "POST":
        # Get the uploaded file from the form
        file = request.files.get("file") # Use .get for safety
        target_lang = request.form.get("target_language") # Use .get for safety
        ai_model_selection = request.form.get("ai_model") # Use .get for safety

        # Basic validation
        if not file or file.filename == "":
            flash("No file selected!")
            return render_template("index.html", languages=LANGUAGES, ai_models=ai_models)
        if not target_lang:
             flash("No target language selected!")
             return render_template("index.html", languages=LANGUAGES, ai_models=ai_models)
        # If Gemini is the only option, ai_model_selection should be 'google'. Check if it's 'none' from disabled option.
        if ai_model_selection == 'none' or not GOOGLE_API_KEY or not GEMINI_MODEL:
             flash("Google Gemini model is not configured correctly.")
             return render_template("index.html", languages=LANGUAGES, ai_models=ai_models)


        # Secure the filename before saving
        safe_filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, safe_filename)

        try:
            file.save(filepath)
        except Exception as e:
            flash(f"Error saving file: {e}")
            return render_template("index.html", languages=LANGUAGES, ai_models=ai_models)

        detected_language = None
        file_content_for_detection = ""

        # Attempt to detect the language using the file's content
        try:
            # Process file based on its extension type
            if safe_filename.endswith(".docx"):
                file_content_for_detection = read_text_from_docx(filepath)
            elif safe_filename.endswith(".pptx"):
                file_content_for_detection = read_text_from_pptx(filepath)
            elif safe_filename.endswith(".xlsx"):
                file_content_for_detection = read_text_from_excel(filepath)
            else:
                flash("Unsupported file type! Only .docx, .pptx, .xlsx are supported.")
                os.remove(filepath)
                return render_template("index.html", languages=LANGUAGES, ai_models=ai_models)

            # If content was successfully read, attempt language detection
            if file_content_for_detection:
                detected_language = detect_language(file_content_for_detection) # detect_language handles text portion
                if detected_language:
                    flash(f"Detected language: {detected_language.upper()}")
                else:
                    # This might happen for files with little text or non-standard content
                    flash("Could not confidently detect language.")
        except Exception as e:
             # Catch errors during file reading for detection
            print(f"Error processing file for language detection: {str(e)}")
            flash(f"Error processing file content for language detection: {str(e)}")
            # Continue without detected_language if an error occurred


        # Determine model type and name (Hardcoded to Google Gemini as per requirement)
        model_type = "google"
        model_name = GEMINI_MODEL # Use the configured Gemini model name

        # Variables to hold the file name and path of the translated document
        translated_filename = None
        translated_filepath = None

        try:
            # Process and translate the file based on its type
            if safe_filename.endswith(".docx"):
                translated_doc = translate_docx(filepath, target_lang, model_name, model_type, detected_language)
                if translated_doc:
                    translated_filename = f"translated_{safe_filename}"
                    translated_filepath = os.path.join(UPLOAD_FOLDER, translated_filename)
                    translated_doc.save(translated_filepath)
            elif safe_filename.endswith(".pptx"):
                translated_ppt = translate_pptx(filepath, target_lang, model_name, model_type, detected_language)
                if translated_ppt:
                    translated_filename = f"translated_{safe_filename}"
                    translated_filepath = os.path.join(UPLOAD_FOLDER, translated_filename)
                    translated_ppt.save(translated_filepath)
            elif safe_filename.endswith(".xlsx"):
                translated_df = translate_excel(filepath, target_lang, model_name, model_type, detected_language)
                if translated_df is not None:
                    translated_filename = f"translated_{safe_filename}"
                    translated_filepath = os.path.join(UPLOAD_FOLDER, translated_filename)
                    translated_df.to_excel(translated_filepath, index=False)
            # No else needed here as unsupported types were caught earlier

            # If a translated file was successfully created, notify the user
            if translated_filename and os.path.exists(translated_filepath): # Double-check existence
                flash("Translation completed successfully!")
            elif not flash("Translation completed successfully.") in get_flashed_messages(): # Check if success wasn't already flashed (e.g. by translate func)
                 flash("Translation failed.") # Generic fail message if no specific error was flashed

        except Exception as e:
            # Catch any errors during the translation process itself
            flash(f"An unexpected error occurred during translation: {e}")
            print(f"An unexpected error occurred during translation: {e}")
            # Clean up potentially partially created translated file
            if translated_filepath and os.path.exists(translated_filepath):
                 try: os.remove(translated_filepath)
                 except OSError as e: print(f"Error deleting partial translated file: {e}")
                 translated_filename = None # Don't show download link if translation failed

        finally:
            # Remove the original uploaded file to save disk space
            try:
                if os.path.exists(filepath):
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
    Requires secure_filename check here too for safety.
    """
    safe_filename = secure_filename(filename)
    filepath = os.path.join(UPLOAD_FOLDER, safe_filename)

    # Check if the file exists before sending and if it's within the UPLOAD_FOLDER
    # os.path.abspath is crucial to prevent directory traversal with send_file
    if os.path.exists(filepath) and os.path.commonprefix([filepath, UPLOAD_FOLDER]) == UPLOAD_FOLDER:
        try:
            # Using as_attachment=True prompts download
            return send_file(filepath, as_attachment=True)
        except Exception as e:
            print(f"Error sending file {safe_filename}: {e}")
            flash("Error serving translated file.")
            return redirect(url_for("index"))
    else:
        print(f"Attempted download of non-existent or unauthorized file: {safe_filename}")
        flash("Translated file not found or access denied.")
        return redirect(url_for("index"))


# Run the Flask app in debug mode if this script is executed directly.
if __name__ == "__main__":
    # Ensure UPLOAD_FOLDER exists before running
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run(debug=True)