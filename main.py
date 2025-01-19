# Document Translator Web App using Azure OpenAI or Ollama with Qwen2.5:latest
# Supports PowerPoint (.pptx, .ppt), Excel (.xlsx, .xls), and Word (.docx) files
# Designed for user-friendly web interaction with drag & drop functionality

import streamlit as st
from pptx import Presentation
import openpyxl  # For .xlsx files
import xlrd  # For .xls files
from docx import Document  # For .docx files
import requests
import json
import os
import tempfile
import time
import zipfile
import io

# Import specific exceptions for better error handling
from pptx.exceptions import PackageNotFoundError as PPTXPackageNotFoundError
from zipfile import BadZipFile as ZipBadZipFile
from openpyxl.utils.exceptions import InvalidFileException as OpenpyxlInvalidFileException
from xlrd import XLRDError
from docx.opc.exceptions import PackageNotFoundError as DOCXPackageNotFoundError

# ============================
# Azure OpenAI Configuration
# ============================

# Replace these placeholders with your actual Azure OpenAI credentials
azure_openai_endpoint = "https://ep-d-eus-aiservice-2-openai-gpt4o-mini-2.openai.azure.com/"  # Ensure it ends with '/'
azure_openai_api_key = "API_KEY_GOES_HERE"  # Replace with your actual API key
azure_openai_deployment_name = "gpt-4o-mini"  # Ensure this is a chat-compatible model like 'gpt-35-turbo'
azure_openai_api_version = "2024-02-15-preview"  # Use the latest supported API version

# ============================
# Ollama Configuration
# ============================

# Ollama API endpoint (replace with your Ollama instance's URL if needed)
ollama_url = "http://localhost:11434/api/chat"
ollama_model = "qwen2.5:latest"  # Or any other suitable model

# ==============================
# System Prompt Configuration
# ==============================

def get_system_prompt(source_language, target_language):
    return f"""
You are an expert translator. Your task is to translate text from {source_language} to {target_language}.

Guidelines:
- Only output text in {target_language}.
- Do not include any headers, footers, or additional explanations.
- If the input text is not in {source_language} or contains multiple languages, translate all content into {target_language} without exception.
- The input text will be provided within triple backticks. Use it solely for translation purposes.
"""

# =================================
# Azure OpenAI Translation Function
# =================================

def translate_with_azure_openai(text, source_language, target_language, endpoint, api_key, deployment_name, api_version):
    system_prompt = get_system_prompt(source_language, target_language)

    payload = {
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text}
        ],
        "max_tokens": 1000,  # Adjust as needed
        "temperature": 0  # Ensures deterministic output
    }

    headers = {
        "Content-Type": "application/json",
        "api-key": api_key,
    }

    # Ensure the endpoint URL is correctly formatted
    if not endpoint.endswith('/'):
        endpoint += '/'

    api_url = f"{endpoint}openai/deployments/{deployment_name}/chat/completions?api-version={api_version}"

    try:
        response = requests.post(api_url, headers=headers, json=payload, timeout=30) # Added timeout for robustness
        response.raise_for_status()  # Raises stored HTTPError, if one occurred.

        response_json = response.json()

        # Check if 'choices' is present and has at least one element
        if 'choices' in response_json and len(response_json['choices']) > 0:
            choice = response_json['choices'][0]
            # Depending on the API, content may be under 'message' or 'text'
            if 'message' in choice and 'content' in choice['message']:
                translated_text = choice['message']['content'].strip()
                return translated_text
            elif 'text' in choice: # Added for older API compatibility - might return text directly
                translated_text = choice['text'].strip()
                return translated_text
            else:
                st.error(f"Unexpected response format in choice: {choice}")
                return None  # Explicitly return None on error
        else:
            st.error(f"No translation found. Azure OpenAI Response: {json.dumps(response_json, indent=2)}")
            return None  # Explicitly return None on error

    except requests.exceptions.RequestException as e:
        st.error(f"Error communicating with Azure OpenAI: {e}")
        return None  # Explicitly return None on error
    except (KeyError, json.JSONDecodeError) as e:
        st.error(f"Error processing Azure OpenAI response: {e}. Response Content: {response.text if 'response' in locals() else 'No response'}")
        return None  # Explicitly return None on error
    except Exception as e:  # Catch any other unexpected errors
        st.error(f"An unexpected error occurred with Azure OpenAI: {e}")
        return None

# =================================
# Ollama Translation Function
# =================================

def translate_with_ollama(text, source_language, target_language):
    system_prompt = get_system_prompt(source_language, target_language)
    prompt = f"{system_prompt}\n\n```\n{text}\n```"
    data = {
        "model": ollama_model,
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "stream": False,
        "options": {
            "temperature": 0.2
        }
    }

    try:
        response = requests.post(ollama_url, data=json.dumps(data), headers={'Content-Type': 'application/json'})
        response.raise_for_status()

        response_json = response.json()
        translated_text = response_json.get("message", {}).get("content", "").strip()
        if not translated_text:
            translated_text = response_json.get("response", "").strip()

        if not translated_text:
            st.error(f"Error: Empty response from Ollama.  Response: {response_json}")
            return None  # Explicitly return None on error

        return translated_text

    except requests.exceptions.RequestException as e:
        st.error(f"Error communicating with Ollama: {e}")
        return None  # Explicitly return None on error
    except (KeyError, json.JSONDecodeError) as e:
        st.error(f"Error processing Ollama response: {e}. Response: {response.text}")
        return None  # Explicitly return None on error
    except Exception as e:  # Catch any other unexpected errors
        st.error(f"An unexpected error occurred with Ollama: {e}")
        return None

# ====================================
# PowerPoint Text Replacement Function
# ====================================

def replace_text_in_pptx(input_path, output_path, source_language, target_language, endpoint, api_key, deployment_name, api_version, ai_model):
    try:
        prs = Presentation(input_path) # pptx library should handle both .pptx and older .ppt to a reasonable extent
        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                text_frame = shape.text_frame
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        original_text = run.text.strip()
                        if original_text:
                            # Use selected AI model for translation
                            if ai_model == "Azure OpenAI 4o-mini":
                                translated_text = translate_with_azure_openai(
                                    original_text, source_language, target_language,
                                    endpoint, api_key, deployment_name, api_version
                                )
                            elif ai_model == "Ollama running Qwen2.5:latest":
                                translated_text = translate_with_ollama(
                                    original_text, source_language, target_language
                                )
                            # Check if translation was successful
                            if translated_text is not None:
                                run.text = translated_text
                            else:
                                st.warning(f"Translation failed for text: '{original_text}' in file: {os.path.basename(input_path)}. Keeping original text.")

        prs.save(output_path)
        st.success(f"Translated PowerPoint file saved: {os.path.basename(output_path)}")

    except PPTXPackageNotFoundError as e:
        st.error(f"Error: PowerPoint file {os.path.basename(input_path)} not found or corrupt. Skipping file.")
    except ZipBadZipFile as e:
        st.error(f"Error: PowerPoint file {os.path.basename(input_path)} is likely corrupt (Bad Zip File). Skipping file.")
    except Exception as e:
        st.error(f"Error processing PowerPoint file {os.path.basename(input_path)}: {e}")

# ====================================
# Excel Text Replacement Function
# ====================================

def replace_text_in_excel(input_path, output_path, source_language, target_language, endpoint, api_key, deployment_name, api_version, ai_model):
    try:
        ext = os.path.splitext(input_path)[1].lower()
        if ext == ".xlsx":
            try:
                wb = openpyxl.load_workbook(input_path) # openpyxl for modern .xlsx
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    # iter_rows is compatible across reasonable openpyxl versions
                    for row in ws.iter_rows():
                        for cell in row:
                            if cell.value and isinstance(cell.value, str):
                                original_text = cell.value.strip()
                                if original_text:
                                    # Use selected AI model for translation
                                    if ai_model == "Azure OpenAI 4o-mini":
                                        translated_text = translate_with_azure_openai(
                                            original_text, source_language, target_language,
                                            endpoint, api_key, deployment_name, api_version
                                        )
                                    elif ai_model == "Ollama running Qwen2.5:latest":
                                        translated_text = translate_with_ollama(
                                            original_text, source_language, target_language
                                        )
                                    # Check if translation was successful
                                    if translated_text is not None:
                                        cell.value = translated_text
                                    else:
                                        st.warning(f"Translation failed for text: '{original_text}' in file: {os.path.basename(input_path)}. Keeping original text.")

                wb.save(output_path) # Save as .xlsx

            except OpenpyxlInvalidFileException as e:
                st.error(f"Error: Excel (.xlsx) file {os.path.basename(input_path)} is likely corrupt or invalid. Skipping file.")
                return # Skip to the next file
            except ZipBadZipFile as e:
                st.error(f"Error: Excel (.xlsx) file {os.path.basename(input_path)} is likely corrupt (Bad Zip File). Skipping file.")
                return # Skip to the next file

        elif ext == ".xls":
            try:
                wb = xlrd.open_workbook(input_path, formatting_info=True)  # xlrd for older .xls, formatting_info to preserve formatting as much as possible
                wb_write = openpyxl.Workbook()  # openpyxl for creating the *new* .xlsx output - writing .xls is complex with minimal changes

                for sheet_index in range(wb.nsheets):
                    sheet = wb.sheet_by_index(sheet_index)
                    ws_write = wb_write.create_sheet(sheet.name)  # Create corresponding sheet

                    for row_index in range(sheet.nrows):
                        for col_index in range(sheet.ncols):
                            cell_value = sheet.cell_value(row_index, col_index) # cell_value is standard in xlrd
                            if isinstance(cell_value, str):
                                original_text = cell_value.strip()
                                if original_text:
                                    # Use selected AI model for translation
                                    if ai_model == "Azure OpenAI 4o-mini":
                                        translated_text = translate_with_azure_openai(
                                            original_text, source_language, target_language,
                                            endpoint, api_key, deployment_name, api_version
                                        )
                                    elif ai_model == "Ollama running Qwen2.5:latest":
                                        translated_text = translate_with_ollama(
                                            original_text, source_language, target_language
                                        )
                                    # Check if translation was successful
                                    if translated_text is not None:
                                        ws_write.cell(row=row_index + 1, column=col_index + 1, value=translated_text) # writing to new .xlsx
                                    else:
                                        st.warning(f"Translation failed for text: '{original_text}' in file: {os.path.basename(input_path)}. Keeping original text.")
                                        ws_write.cell(row=row_index + 1, column=col_index + 1, value=cell_value)  # Keep original if translation fails
                            else:
                                ws_write.cell(row=row_index + 1, column=col_index + 1, value=cell_value) # copy non-string values

                # Remove the default sheet created by openpyxl if it exists and is empty
                if "Sheet" in wb_write.sheetnames and not wb_write["Sheet"]["A1"].value:
                    del wb_write["Sheet"]

                wb_write.save(output_path) # Saving as .xlsx

            except XLRDError as e:
                st.error(f"Error: Excel (.xls) file {os.path.basename(input_path)} is likely corrupt or invalid or in a very old format incompatible with xlrd. Skipping file.") # More informative error
                return # Skip to the next file
            except Exception as e: # Catch any other potential issues during .xls processing
                st.error(f"Error processing Excel (.xls) file {os.path.basename(input_path)}: {e}. Skipping file.  The file might be corrupt or in a very old format.") # More informative error
                return # Skip to the next file

        else:
            st.error(f"Unsupported file type: {ext} for file {os.path.basename(input_path)}.")
            return

        st.success(f"Translated Excel file saved: {os.path.basename(output_path)}")

    except Exception as e: # Catch any other top-level errors during excel processing
        st.error(f"General error processing Excel file {os.path.basename(input_path)}: {e}. File might be corrupt, password protected, or there was an unexpected issue.") # More informative error

# ====================================
# Word Text Replacement Function
# ====================================

def replace_text_in_word(input_path, output_path, source_language, target_language, endpoint, api_key, deployment_name, api_version, ai_model):
    try:
        doc = Document(input_path)

        # Translate text in paragraphs
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                original_text = run.text.strip()
                if original_text:
                    # Use selected AI model for translation
                    if ai_model == "Azure OpenAI 4o-mini":
                        translated_text = translate_with_azure_openai(
                            original_text, source_language, target_language,
                            endpoint, api_key, deployment_name, api_version
                        )
                    elif ai_model == "Ollama running Qwen2.5:latest":
                        translated_text = translate_with_ollama(
                            original_text, source_language, target_language
                        )

                    # Check if translation was successful
                    if translated_text is not None:
                        run.text = translated_text
                    else:
                        st.warning(f"Translation failed for text: '{original_text}' in file: {os.path.basename(input_path)}. Keeping original text.")

        # Translate text in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            original_text = run.text.strip()
                            if original_text:
                                # Use selected AI model for translation
                                if ai_model == "Azure OpenAI 4o-mini":
                                    translated_text = translate_with_azure_openai(
                                        original_text, source_language, target_language,
                                        endpoint, api_key, deployment_name, api_version
                                    )
                                elif ai_model == "Ollama running Qwen2.5:latest":
                                    translated_text = translate_with_ollama(
                                        original_text, source_language, target_language
                                    )
                                # Check if translation was successful
                                if translated_text is not None:
                                    run.text = translated_text
                                else:
                                    st.warning(f"Translation failed for text: '{original_text}' in file: {os.path.basename(input_path)}. Keeping original text.")

        doc.save(output_path)
        st.success(f"Translated Word file saved: {os.path.basename(output_path)}")

    except DOCXPackageNotFoundError as e:
        st.error(f"Error: Word file {os.path.basename(input_path)} not found or corrupt. Skipping file.")
    except ZipBadZipFile as e:
        st.error(f"Error: Word file {os.path.basename(input_path)} is likely corrupt (Bad Zip File). Skipping file.")
    except Exception as e:
        st.error(f"Error processing Word file {os.path.basename(input_path)}: {e}")

# ===================
# Streamlit UI Setup
# ===================

st.set_page_config(page_title="📝 Document Translator", layout="wide")
st.title("📝 Document Translator")

# Updated languages list with the new additions
languages = ["English", "French", "German", "Chinese", "Arabic", "Italian", "Japanese", "Hindi", "Kashmiri", "Manipuri", "Gujarati"]
source_language = st.selectbox("Select the source language:", languages)
target_language = st.selectbox("Select the target language:", languages)

# AI Model Selection
ai_model = st.selectbox("Select the AI model for translation:", ["Azure OpenAI 4o-mini", "Ollama running Qwen2.5:latest"])

# Ensure that source and target languages are not the same
if source_language == target_language:
    st.warning("Source and target languages are the same. No translation needed.")

# File Uploader with drag & drop
uploaded_files = st.file_uploader(
    "Drag & drop PowerPoint, Excel, and Word files here (max 10 files or 2 GB total):",
    type=['pptx', 'ppt', 'xlsx', 'xls', 'docx'],
    accept_multiple_files=True
)

# Function to check total size
def check_total_size(files, max_size_bytes=2 * 1024 * 1024 * 1024):
    total_size = sum([file.size for file in files])
    return total_size <= max_size_bytes

if uploaded_files:
    if len(uploaded_files) > 10:
        st.error("You can upload a maximum of 10 files at a time.")
    elif not check_total_size(uploaded_files):
        st.error("Total upload size exceeds 2 GB. Please upload smaller files or reduce the number of files.")
    else:
        if st.button("🚀 Translate Files"):
            # Verify Azure OpenAI credentials if Azure is selected
            if ai_model == "Azure OpenAI 4o-mini" and (not azure_openai_endpoint or azure_openai_endpoint == "https://your-resource-name.openai.azure.com/"):
                st.error("Please configure your Azure OpenAI endpoint in the script.")
            elif ai_model == "Azure OpenAI 4o-mini" and (not azure_openai_api_key or azure_openai_api_key == "your-azure-openai-api-key"):
                st.error("Please configure your Azure OpenAI API key in the script.")
            elif ai_model == "Azure OpenAI 4o-mini" and (not azure_openai_deployment_name or azure_openai_deployment_name == "your-deployment-name"):
                st.error("Please configure your Azure OpenAI deployment name in the script.")
            else:
                with st.spinner("🔄 Preparing to translate files..."):
                    # Create a temporary directory
                    with tempfile.TemporaryDirectory() as tmpdir:
                        translated_files = []
                        total_files = len(uploaded_files)
                        start_time = time.time()

                        progress_bar = st.progress(0)
                        status_text = st.empty()

                        for i, uploaded_file in enumerate(uploaded_files, start=1):
                            status_text.text(f"Translating file {i} of {total_files}: {uploaded_file.name}")
                            # Save the uploaded file to the temp directory
                            input_path = os.path.join(tmpdir, uploaded_file.name)
                            with open(input_path, "wb") as f:
                                f.write(uploaded_file.getbuffer())

                            name, ext = os.path.splitext(uploaded_file.name)
                            output_filename = f"{name}-translated{ext if ext in ['.pptx', '.ppt', '.xlsx', '.xls', '.docx'] else '.' + ext}"
                            output_path = os.path.join(tmpdir, output_filename)

                            # Translate based on file type
                            if ext.lower() in ['.pptx', '.ppt']:
                                replace_text_in_pptx(
                                    input_path, output_path, source_language, target_language,
                                    azure_openai_endpoint, azure_openai_api_key,
                                    azure_openai_deployment_name, azure_openai_api_version,
                                    ai_model  # Pass the selected AI model
                                )
                            elif ext.lower() in ['.xlsx', '.xls']:
                                replace_text_in_excel(
                                    input_path, output_path, source_language, target_language,
                                    azure_openai_endpoint, azure_openai_api_key,
                                    azure_openai_deployment_name, azure_openai_api_version,
                                    ai_model  # Pass the selected AI model
                                )
                            elif ext.lower() == '.docx':
                                replace_text_in_word(
                                    input_path, output_path, source_language, target_language,
                                    azure_openai_endpoint, azure_openai_api_key,
                                    azure_openai_deployment_name, azure_openai_api_version,
                                    ai_model  # Pass the selected AI model
                                )
                            else:
                                st.error(f"Unsupported file type: {ext} for file {uploaded_file.name}. Skipping file.")
                                continue

                            # Collect the translated file
                            if os.path.exists(output_path):
                                with open(output_path, "rb") as f:
                                    file_bytes = f.read()
                                    translated_files.append((output_filename, file_bytes))

                            # Update progress bar
                            progress = i / total_files
                            progress_bar.progress(progress)

                        end_time = time.time()
                        total_time = end_time - start_time
                        average_time_per_file = total_time / len(translated_files) if translated_files else 0

                        status_text.text("🎉 Translation complete!")
                        st.write(f"Translated **{len(translated_files)}** file(s) in **{total_time:.2f} seconds**.")
                        if translated_files:
                            st.write(f"Average time per file: **{average_time_per_file:.2f} seconds**.")

                        if translated_files:
                            # Create a ZIP archive in-memory
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                                for filename, data in translated_files:
                                    zipf.writestr(filename, data)
                            zip_buffer.seek(0)

                            # Provide a download button
                            st.download_button(
                                label="📥 Download Translated Files (ZIP)",
                                data=zip_buffer,
                                file_name="translated_files.zip",
                                mime="application/zip"
                            )
