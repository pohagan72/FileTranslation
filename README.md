# FileTranslation

This project is for translating text in PowerPoint, Excel and Word files. The goal is to provide a simple way for users to have the text in those files translated. 

The index.html file is expected to be in a folder named "templates" in the same location as the python file. /templates/index.html

## Project Dependencies:
``` pip install Flask python-docx python-pptx pandas requests langdetect openpyxl ```

## LLM Requirements
The app allows you to select from a model installed through Ollama or an Azure model. For Azure the API key needs to be added. 
