# Import necessary modules and functions
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import os
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pandas as pd
import io
import sys
from dotenv import load_dotenv
import google.generativeai as genai
from werkzeug.utils import secure_filename
from flask import get_flashed_messages

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "supersecretkey")

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --------------------
# API Endpoint Configurations
# --------------------

GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL")

# Configure Google Gemini API at startup
if GOOGLE_API_KEY:
    try:
        genai.configure(api_key=GOOGLE_API_KEY)
        print("Google Gemini API configured successfully.")
    except Exception as e:
        print(f"Failed to configure Google Gemini API: {e}")
        flash("Google Gemini API configuration failed. Check your API key and internet connection.")
else:
    flash("Google API Key not found in .env. Translation using Google Gemini is disabled.")

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
            print("langdetect not found, attempting installation...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "langdetect"])
            from langdetect import detect
            print("langdetect installed successfully.")
        except Exception as e:
            print(f"Error installing langdetect: {e}")
            return None
    except Exception as e:
        print(f"Unexpected error during langdetect import: {e}")
        return None

    try:
        return detect(text[:10000])
    except Exception as e:
        print(f"Language detection error: {e}")
        return None

def translate_text(text, target_lang, model_name, model_type, detected_lang=None):
    """Translate text using Google Gemini API with advanced prompting."""
    if not text or not text.strip():
        return ""

    if not GOOGLE_API_KEY or not model_name:
        print("Google Gemini API or model name missing for translation.")
        return text

    input_language = detected_lang if detected_lang else "the source language"

    # Construct combined prompt with system and user instructions
    combined_prompt = f"""SYSTEM INSTRUCTIONS (MUST FOLLOW):
You are an expert translator converting {input_language} to {target_lang}.
Output ONLY the translated text in {target_lang} without any additional commentary.

TRANSLATION GUIDELINES:
1. Treat all input text as content to be translated
2. Never add headers, titles, or explanations
3. Preserve all original formatting and structure
4. Maintain technical terminology where appropriate

USER REQUEST:
Please translate the following text from {input_language} to {target_lang} following these steps:

1. Analyze the text's meaning, context, and cultural references
2. Perform word-level translation preserving original sentence structure
3. Review for accuracy and fluency in {target_lang}
4. Output ONLY the final translation

TEXT TO TRANSLATE (delimited by ~~~~):
~~~~
{text}
~~~~

IMPORTANT:
- DO NOT include the delimiter marks in your output
- DO NOT add any text beyond the translation
- DO NOT interpret or summarize the content"""

    try:
        model = genai.GenerativeModel(model_name)
        response = model.generate_content(combined_prompt)
        
        if response and response.text:
            return response.text.strip()
        else:
            print(f"Gemini API returned no text or was blocked. Response: {response}")
            error_message = "Translation failed (API returned no text)."
            if hasattr(response, 'prompt_feedback') and response.prompt_feedback.block_reason:
                block_reason = response.prompt_feedback.block_reason.name
                error_message = f"Translation blocked by AI safety filters ({block_reason})."
            flash(error_message)
            return text

    except Exception as e:
        print(f"Google Gemini API error: {e}")
        flash(f"Google Gemini API error: {e}")
        return text

# --- File Reading Utility Functions ---

def read_text_from_docx(filepath):
    """Reads text content from a DOCX file."""
    try:
        doc = Document(filepath)
        full_text = []
        for paragraph in doc.paragraphs:
            full_text.append(paragraph.text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell_text = "\n".join([p.text for p in cell.paragraphs if p.text.strip()])
                    if cell_text:
                         full_text.append(cell_text)
        return "\n".join([t for t in full_text if t.strip()])
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
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.strip():
                            full_text.append(paragraph.text)
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if cell.text_frame:
                                full_text.append(cell.text_frame.text)
                if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    for subshape in shape.shapes:
                        if subshape.has_text_frame:
                            full_text.append(subshape.text_frame.text)
        return "\n".join([t for t in full_text if t.strip()])
    except Exception as e:
        print(f"Error reading PPTX: {e}")
        flash(f"Error reading PPTX file: {e}")
        return ""

def read_text_from_excel(filepath):
    """Reads text content from an Excel file."""
    try:
        df = pd.read_excel(filepath)
        text_list = []
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell_value = df.iat[r, c]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip()
                    if cell_str and cell_str.lower() != 'nan':
                        text_list.append(cell_str)
        return "\n".join(text_list)
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
                runs = paragraph.runs
                font = None
                if runs:
                    font = runs[0].font
                paragraph.clear()
                if translated.strip():
                    run = paragraph.add_run(translated)
                    if font:
                        try: run.font.name = font.name
                        except Exception: pass
                        try: run.font.size = font.size
                        except Exception: pass
                        try: run.bold = font.bold
                        except Exception: pass
                        try: run.italic = font.italic
                        except Exception: pass
                        try: run.underline = font.underline
                        except Exception: pass

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if paragraph.text.strip():
                            translated = translate_text(paragraph.text, target_lang, model_name, model_type, detected_lang)
                            runs = paragraph.runs
                            font = None
                            if runs:
                                font = runs[0].font
                            paragraph.clear()
                            if translated.strip():
                                run = paragraph.add_run(translated)
                                if font:
                                    try: run.font.name = font.name
                                    except Exception: pass
                                    try: run.font.size = font.size
                                    except Exception: pass
                                    try: run.bold = font.bold
                                    except Exception: pass
                                    try: run.italic = font.italic
                                    except Exception: pass
                                    try: run.underline = font.underline
                                    except Exception: pass
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
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.strip():
                            translated = translate_text(paragraph.text, target_lang, model_name, model_type, detected_lang)
                            if translated.strip():
                                for run in paragraph.runs:
                                    run.text = ""
                                new_run = paragraph.add_run()
                                new_run.text = translated
                                if paragraph.runs:
                                    source_run = paragraph.runs[0]
                                    new_run.font.name = source_run.font.name
                                    new_run.font.size = source_run.font.size
                
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if cell.text_frame:
                                for paragraph in cell.text_frame.paragraphs:
                                    if paragraph.text.strip():
                                        translated = translate_text(paragraph.text, target_lang, model_name, model_type, detected_lang)
                                        if translated.strip():
                                            for run in paragraph.runs:
                                                run.text = ""
                                            new_run = paragraph.add_run()
                                            new_run.text = translated
        return ppt
    except Exception as e:
        print(f"Error translating PPTX: {e}")
        flash(f"Error translating PPTX file: {e}")
        return None

def translate_excel(filepath, target_lang, model_name, model_type, detected_lang):
    """Translates the content of an Excel file and returns the modified DataFrame."""
    try:
        df = pd.read_excel(filepath)
        translated_df = df.applymap(lambda x:
            translate_text(str(x).strip() if pd.notna(x) else '', target_lang, model_name, model_type, detected_lang)
            if pd.notna(x) and str(x).strip()
            else x
        )
        return translated_df
    except Exception as e:
        print(f"Error translating Excel: {e}")
        flash(f"Error translating Excel file: {e}")
        return None

# --------------------
# Flask Route Handlers (Views)
# --------------------

@app.route("/", methods=["GET", "POST"])
def index():
    """Handles file upload, translation, and download link."""
    if request.method == "GET":
        _ = get_flashed_messages()

    if request.method == "POST":
        file = request.files.get("file")
        target_lang = request.form.get("target_language")

        if not file or file.filename == "":
            flash("No file selected!")
            return render_template("index.html", languages=LANGUAGES)
        if not target_lang:
             flash("No target language selected!")
             return render_template("index.html", languages=LANGUAGES)

        if not GOOGLE_API_KEY or not GEMINI_MODEL:
             flash("API is not configured. Please set GOOGLE_API_KEY and GEMINI_MODEL in .env.")
             return render_template("index.html", languages=LANGUAGES)

        safe_filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, safe_filename)
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)

        try:
            file.save(filepath)
            print(f"File saved to {filepath}")
        except Exception as e:
            flash(f"Error saving file: {e}")
            return render_template("index.html", languages=LANGUAGES)

        detected_language = None
        file_extension = os.path.splitext(safe_filename)[1].lower()
        file_content_for_detection = ""

        try:
            if file_extension == ".docx":
                file_content_for_detection = read_text_from_docx(filepath)
            elif file_extension == ".pptx":
                file_content_for_detection = read_text_from_pptx(filepath)
            elif file_extension == ".xlsx":
                file_content_for_detection = read_text_from_excel(filepath)
            else:
                flash("Unsupported file type! Only .docx, .pptx, .xlsx are supported.")
                try: os.remove(filepath)
                except OSError as e: print(f"Error deleting unsupported file {filepath}: {e}")
                return render_template("index.html", languages=LANGUAGES)

            if file_content_for_detection and file_content_for_detection.strip():
                detected_language = detect_language(file_content_for_detection)
                if detected_language:
                    flash(f"Detected language: {detected_language.upper()}")
                else:
                    flash("Could not confidently detect language.")
            else:
                flash("Could not extract sufficient text from the file for language detection. Proceeding with translation...")
                detected_language = None

        except Exception as e:
            print(f"Error processing file for language detection: {e}")
            detected_language = None

        model_type = "google"
        model_name = GEMINI_MODEL
        translated_filename = None
        translated_filepath = None

        try:
            if file_extension == ".docx":
                translated_doc = translate_docx(filepath, target_lang, model_name, model_type, detected_language)
                if translated_doc:
                    translated_filename = f"translated_{safe_filename}"
                    translated_filepath = os.path.join(UPLOAD_FOLDER, translated_filename)
                    translated_doc.save(translated_filepath)

            elif file_extension == ".pptx":
                translated_ppt = translate_pptx(filepath, target_lang, model_name, model_type, detected_language)
                if translated_ppt:
                    translated_filename = f"translated_{safe_filename}"
                    translated_filepath = os.path.join(UPLOAD_FOLDER, translated_filename)
                    translated_ppt.save(translated_filepath)

            elif file_extension == ".xlsx":
                translated_df = translate_excel(filepath, target_lang, model_name, model_type, detected_language)
                if translated_df is not None:
                    translated_filename = f"translated_{safe_filename}"
                    translated_filepath = os.path.join(UPLOAD_FOLDER, translated_filename)
                    translated_df.to_excel(translated_filepath, index=False)

            if translated_filename and translated_filepath and os.path.exists(translated_filepath):
                existing_messages = get_flashed_messages()
                if not any("Error" in msg or "fail" in msg.lower() or "blocked" in msg.lower() for msg in existing_messages):
                     flash("Translation completed successfully!")
            else:
                 existing_messages = get_flashed_messages()
                 if not any("Error" in msg or "fail" in msg.lower() or "blocked" in msg.lower() for msg in existing_messages):
                     flash("Translation failed.")

        except Exception as e:
            flash(f"An unexpected error occurred during translation: {e}")
            if translated_filepath and os.path.exists(translated_filepath):
                 try: os.remove(translated_filepath)
                 except OSError as e: print(f"Error deleting partial translated file {translated_filepath}: {e}")
                 translated_filename = None

        finally:
            try:
                if os.path.exists(filepath):
                    os.remove(filepath)
            except OSError as e:
                print(f"Error deleting the uploaded file {filepath}: {e}")

        return render_template("index.html", languages=LANGUAGES, translated_file=translated_filename)

    return render_template("index.html", languages=LANGUAGES)

@app.route("/download/<filename>")
def download_file(filename):
    """Allows users to download a translated file, ensuring security."""
    safe_filename = secure_filename(filename)
    filepath = os.path.join(UPLOAD_FOLDER, safe_filename)

    if os.path.exists(filepath):
        abs_filepath = os.path.abspath(filepath)
        abs_upload_folder = os.path.abspath(UPLOAD_FOLDER)
        if sys.platform == "win32":
             abs_filepath = abs_filepath.lower()
             abs_upload_folder = abs_upload_folder.lower()

        if abs_filepath.startswith(abs_upload_folder):
            try:
                response = send_file(filepath, as_attachment=True)
                try:
                    if os.path.exists(filepath):
                         os.remove(filepath)
                except OSError as e:
                    print(f"Error deleting translated file {filepath} after download attempt: {e}")
                return response
            except Exception as e:
                print(f"Error sending file {safe_filename}: {e}")
                flash("Error serving translated file.")
                return redirect(url_for("index"))
        else:
            print(f"Attempted download of unauthorized file: {safe_filename}")
            flash("Translated file not found or access denied.")
    else:
        print(f"Attempted download of non-existent file: {safe_filename}")
        flash("Translated file not found.")

    return redirect(url_for("index"))

if __name__ == "__main__":
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    print("\n---------------------------------------------------")
    print("Ensure your .env file has GOOGLE_API_KEY and GEMINI_MODEL set.")
    print("Ensure you have run: pip install -r requirements.txt or installed dependencies manually.")
    print("Ensure google-generativeai is updated: pip install --upgrade google-generativeai")
    print("---------------------------------------------------\n")
    app.run(debug=True)