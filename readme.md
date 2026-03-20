# Zensar AI Resume Formatter

A next-generation AI-powered resume parser and formatter. This tool extracts unstructured text from PDF resumes using OpenAI's `gpt-4o` model, intelligently categorizes skills and experiences, and perfectly formats the data into a standardized Microsoft Word (`.docx`) template.

## Features
- **Smart Data Extraction:** Uses AI to comprehensively extract work experience, education, skills, and achievements without dropping data.
- **Automated Templating:** Generates a perfectly formatted Word document using `docxtpl`.
- **Fast Web Interface:** A clean, intuitive upload portal built with FastAPI and Jinja2.
- **Network Sharing:** Easily share the local web app with colleagues on the same Wi-Fi/VPN.

## Prerequisites
Before you begin, ensure you have **Python 3.8+** installed on your machine. 

## Installation

1. Open a terminal or command prompt in this project folder.
2. Install the required Python dependencies by running:
```bash
pip install fastapi uvicorn pdfplumber openai docxtpl python-docx python-multipart jinja2
```

## Setup
1. **API Key Setup:** Open `index.py` and verify that your `OPENAI_API_KEY` is configured correctly.
2. **Template Document:** The system uses `somu_fixed.docx` as the base template (configured in `index.py`). Ensure this file exists in the directory. You can easily switch templates by changing `TEMPLATE_PATH = "somu_fixed.docx"` in `index.py`.

## How to Run

1. Open your terminal in the project folder (`resume-converter`).
2. Run the share script:
```bash
python share.py
```
3. The server will start and provide you with two links in the console:
   - **Local Access:** `http://localhost:8000` (click this to open it on your own machine)
   - **Network Access:** `http://<YOUR_IP>:8000` (share this link with colleagues to let them upload resumes directly from their computers!)

## Usage
1. Open the provided link in your browser.
2. Upload a candidate's PDF resume using the web interface.
3. Wait for the AI multi-pass extraction to process the resume (this usually takes 15-30 seconds).
4. **Download** the final generated `.docx` file when the generation is complete!

## Troubleshooting

- **`ModuleNotFoundError`**: If you get an error saying a module is missing when starting the server, make sure you ran the `pip install` command from the Installation section.
- **`[Errno 98] Address already in use`**: Port 8000 is currently occupied by another app. Open `share.py` and change `port = 8000` to `port = 8080`.
- **Word Formatting Issues / Blank Gaps**: If you edit the `docx` templates manually, MS Word may secretly insert invisible `<w:p>` paragraph tags when you hit Enter. This can cause the Jinja parser to leave huge blank gaps or swallow conditionally-rendered text. Always try to keep `{%p if ... %}` tags tightly packed, or use the `somu_fixed.docx` which was scrubbed programmatically!
