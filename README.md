<div align="center">

# AI Excel Bridge

### Gemini Vision AI + Native Excel COM Integration for Batch Document Automation

![Python](https://img.shields.io/badge/Python-3.12-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Gemini](https://img.shields.io/badge/Google_Gemini-1.5_Flash-4285F4?style=for-the-badge&logo=google&logoColor=white)
![Windows](https://img.shields.io/badge/Windows_COM-win32com-0078D4?style=for-the-badge&logo=windows&logoColor=white)
![Excel](https://img.shields.io/badge/Microsoft_Excel-Interop-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)

**Turns a 45-minute manual data-entry task into a 5-second automated batch process.**

</div>

---

## The Problem

An HR department receives batches of physical vehicle insurance certificates (photographed Arabic documents) that must be manually transcribed into a master Excel ERP spreadsheet. This was:
- Taking **45+ minutes per batch** of 15-20 documents
- Prone to human transcription errors affecting payroll and compliance
- Requiring interaction with a complex ERP spreadsheet with protected formulas

## The Solution

A standalone Windows desktop application that uses **Google Gemini Vision AI** to read the certificates and writes the extracted data back into the ERP Excel file using **native Windows COM interop** — preserving all existing macros, formulas, and formatting.

---

## Key Features

| Feature | Description |
|---|---|
| **Concurrent AI Extraction** | Fires up to 10 images simultaneously to Gemini Vision AI using ThreadPoolExecutor, achieving a **10x speed improvement** over sequential processing |
| **Zero Formula Corruption** | Uses win32com.client to drive the actual Microsoft Excel application silently in the background, instead of rewriting the file directly |
| **Live Diff Preview** | Before committing, shows a field-by-field comparison of old vs. new values for each certificate |
| **Visual Audit Trail** | Updated rows and cells are highlighted in bright yellow in the master Excel file for easy supervisor review |
| **Smart Batch Receipt Date** | Changing the receipt date on any image propagates instantly to all items in the batch (one-click for the whole delivery) |
| **Standalone EXE** | Packaged via PyInstaller into a portable .exe that runs on any Windows machine with Excel installed — no Python needed |

---

## Architecture

`
┌──────────────────────────────────────────────────────────┐
│                    gui_app.py (Tkinter UI)                │
│  ┌────────────────────────────────────────────────────┐   │
│  │         ThreadPoolExecutor (max_workers=10)         │   │
│  │  [Image 1] [Image 2] ... [Image 15]  ──► Gemini   │   │
│  └────────────────────────────────────────────────────┘   │
│                          │                                 │
│                    ocr_engine.py                          │
│              (Gemini 1.5 Flash Vision AI)                 │
│                          │                                 │
│                    excel_handler.py                        │
│              (win32com.client → Excel COM)                 │
└──────────────────────────────────────────────────────────┘
`

---

## Tech Stack

- **UI:** Python Tkinter
- **AI Vision:** Google Gemini 1.5 Flash (Free Tier — 1,500 req/day)
- **Image Processing:** OpenCV, Pillow, EasyOCR (fallback)
- **Excel Interop:** pywin32 / win32com.client
- **Packaging:** PyInstaller

---

## Setup

> **Prerequisite:** Microsoft Excel must be installed on the host machine (required for the COM interop layer).

**1. Clone the repository:**
`ash
git clone https://github.com/Anas9906/ai-excel-bridge.git
cd ai-excel-bridge
`

**2. Install dependencies:**
`ash
pip install -r requirements.txt
`

**3. Create your .env file:**
`env
GEMINI_API_KEY=your_api_key_here
`
Get a free key at [Google AI Studio](https://aistudio.google.com).

**4. Run the application:**
`ash
python gui_app.py
`

**5. Or build the standalone .exe:**
`ash
python build_exe.py
`

---

## Engineering Journey

For a detailed walkthrough of the architectural decisions, failed attempts, and iterative improvements that led to this final design, see **[ENGINEERING_JOURNEY.md](ENGINEERING_JOURNEY.md)**.

---

## License

MIT
