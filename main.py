# Import necessary modules and functions
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import os
from docx import Document  # To handle Word documents
from pptx import Presentation  # To handle PowerPoint presentations
import pandas as pd  # To handle Excel files
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

def detect_language(text):
    """Detects the language of a given text."""
    if not text or not text.strip():
        return None

    try:
        from langdetect import detect
    except ImportError:
        try:
            import subprocess
            import sys
            subprocess.check_call([sys.executable, "-m", "pip", "install", "langdetect"])
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
        return detect(text[:2000])
    except Exception as e:
        print(f"Language detection error: {e}")
        return None


def translate_text(text, target_lang, model_name, model_type, detected_lang=None):
    """Translate text using Google Gemini API."""
    if not text or not text.strip():
        return ""

    if model_type == "google":
        if not GOOGLE_API_KEY or not model_name:
            flash("Google Gemini API is not configured.")
            print("Google Gemini API or model name missing for translation.")
            return text

        try:
            if detected_lang and detected_lang.lower() != target_lang.lower():
                 prompt = f"Translate the following text from {detected_lang} into {target_lang}:"
            else:
                prompt = f"Translate the following text into {target_lang}:"

            model = genai.GenerativeModel(model_name)
            response = model.generate_content([prompt, text])

            if hasattr(response, 'text') and response.text is not None:
                 return response.text.strip()
            else:
                 print(f"Gemini API returned no text or was blocked. Response: {response}")
                 if hasattr(response, 'prompt_feedback') and response.prompt_feedback.block_reason:
                     block_reason = response.prompt_feedback.block_reason.name
                     print(f"Gemini API blocked content due to: {block_reason}")
                     flash(f"Translation blocked by AI safety filters ({block_reason}).")
                 else:
                     flash(f"Gemini API returned no text.")
                 return text

        except Exception as e:
             print(f"Google Gemini API error: {e}")
             flash(f"Google Gemini API error: {e}")
             return text
    else:
        flash(f"Invalid model type specified internally: {model_type}")
        return text


# --- File Reading Utility Functions ---

def read_text_from_docx(filepath):
    """Reads text content from a DOCX file."""
    try:
        doc = Document(filepath)
        full_text = []
        for paragraph in doc.paragraphs:
            full_text.append(paragraph.text)
        return "\n".join(full_text)
    except Exception as e:
        print(f"Error reading DOCX: {e}")
        flash(f"Error reading DOCX file: {e}")
        return ""

def read_text_from_pptx(filepath):
    """Reads text content from a PPTX file."""
    try:
        ppt = Presentation(filepath)
        full_text = []
        for slide in ppt.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    full_text.append(shape.text)
        return "\n".join(full_text)
    except Exception as e:
        print(f"Error reading PPTX: {e}")
        flash(f"Error reading PPTX file: {e}")
        return ""

def read_text_from_excel(filepath):
    """Reads text content from an Excel file."""
    try:
        df = pd.read_excel(filepath)
        full_text = df.to_string(header=True, index=False) # Include header for detection
        return full_text
    except Exception as e:
        print(f"Error reading Excel: {e}")
        flash(f"Error reading Excel file: {e}")
        return ""

# --- File Translation Utility Functions ---

def translate_docx(filepath, target_lang, model_name, model_type, detected_lang):
    """Translates the text in a DOCX file and returns a Document object."""
    try:
        doc = Document(filepath)
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                translated = translate_text(paragraph.text, target_lang, model_name, model_type, detected_lang)
                paragraph.clear()
                run = paragraph.add_run(translated)
        return doc
    except Exception as e:
        print(f"Error translating DOCX: {e}")
        flash(f"Error translating DOCX file: {e}")
        return None

def translate_pptx(filepath, target_lang, model_name, model_type, detected_lang):
    """Translates the text in a PPTX file and returns a Presentation object."""
    try:
        ppt = Presentation(filepath)
        for slide in ppt.slides:
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    text_frame = shape.text_frame
                    original_text = text_frame.text
                    translated_text = translate_text(original_text, target_lang, model_name, model_type, detected_lang)

                    text_frame.clear()
                    p = text_frame.add_paragraph()
                    run = p.add_run()
                    run.text = translated_text
        return ppt
    except Exception as e:
        print(f"Error translating PPTX: {e}")
        flash(f"Error translating PPTX file: {e}")
        return None

def translate_excel(filepath, target_lang, model_name, model_type, detected_lang):
    """Translates the content of an Excel file and returns the modified DataFrame."""
    try:
        df = pd.read_excel(filepath)
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
    """Main endpoint: handles file upload, translation, and download link."""
    ai_models = []
    if GOOGLE_API_KEY and GEMINI_MODEL:
        ai_models.append({"label": "Google Gemini", "models": [{"value": "google", "text": f"Google Gemini ({GEMINI_MODEL})"}]})
    else:
        ai_models.append({"label": "Configuration Error", "models": [{"value": "none", "text": "Gemini not configured - check .env", "disabled": True}]})


    if request.method == "POST":
        file = request.files.get("file")
        target_lang = request.form.get("target_language")
        ai_model_selection = request.form.get("ai_model")

        if not file or file.filename == "":
            flash("No file selected!")
            return render_template("index.html", languages=LANGUAGES, ai_models=ai_models)
        if not target_lang:
             flash("No target language selected!")
             return render_template("index.html", languages=LANGUAGES, ai_models=ai_models)
        if ai_model_selection == 'none' or not GOOGLE_API_KEY or not GEMINI_MODEL:
             flash("Google Gemini model is not configured correctly.")
             return render_template("index.html", languages=LANGUAGES, ai_models=ai_models)


        safe_filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, safe_filename)

        try:
            file.save(filepath)
        except Exception as e:
            flash(f"Error saving file: {e}")
            return render_template("index.html", languages=LANGUAGES, ai_models=ai_models)

        detected_language = None
        file_content_for_detection = ""

        try:
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

            if file_content_for_detection:
                detected_language = detect_language(file_content_for_detection)
                if detected_language:
                    flash(f"Detected language: {detected_language.upper()}")
                else:
                    flash("Could not confidently detect language.")
        except Exception as e:
            print(f"Error processing file for language detection: {str(e)}")
            flash(f"Error processing file content for language detection: {str(e)}")

        model_type = "google"
        model_name = GEMINI_MODEL

        translated_filename = None
        translated_filepath = None # used

        try:
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

            if translated_filename and os.path.exists(translated_filepath): # Verify file was created
                flash("Translation completed successfully!")
            elif not flash("Translation completed successfully.") in get_flashed_messages(): # Check if success wasn't flashed already
                 flash("Translation failed.") # Generic message if no specific error was flashed

        except Exception as e:
            flash(f"An unexpected error occurred during translation: {e}")
            print(f"An unexpected error occurred during translation: {e}")
            if translated_filepath and os.path.exists(translated_filepath):
                 try: os.remove(translated_filepath) # clean up file, if a partial was created
                 except OSError as e: print(f"Error deleting partial translated file: {e}")
                 translated_filename = None

        finally:
            # Remove the original uploaded file
            try:
                if os.path.exists(filepath):
                    os.remove(filepath) # Delete file from directory to limit storage
            except OSError as e:
                print(f"Error deleting the uploaded file: {e}") # If deletion fails

        # Render the template with the translated file download link if available
        return render_template("index.html", languages=LANGUAGES, ai_models=ai_models, translated_file=translated_filename)

    # For GET requests, simply render the homepage with the file upload form.
    return render_template("index.html", languages=LANGUAGES, ai_models=ai_models)

@app.route("/download/<filename>")
def download_file(filename):
    """Allows download of a translated file, ensuring security."""
    safe_filename = secure_filename(filename)
    filepath = os.path.join(UPLOAD_FOLDER, safe_filename)

    # Double-check the file is there AND within allowed folder before sending, and within same path for safety
    if os.path.exists(filepath) and os.path.commonprefix([filepath, UPLOAD_FOLDER]) == UPLOAD_FOLDER:
        try:
            # Trigger download
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