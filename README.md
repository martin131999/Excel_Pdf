# Excel → PDF (Python)

A small **Python + Streamlit** app that converts an uploaded Excel file into a downloadable **PDF**.

## What you need installed

- **Python 3.10+** (recommended 3.11)
- Cursor (you already have it)

## Install Python (Windows)

1. Install Python from [python.org](https://www.python.org/downloads/windows/)
2. During install, **check**: “Add python.exe to PATH”
3. After installation, open a *new* terminal and verify:

```powershell
python --version
pip --version
```

If Windows opens the Microsoft Store instead, disable the aliases:
Settings → Apps → Advanced app settings → **App execution aliases** → turn off `python.exe` and `python3.exe`.

## Setup (Windows / PowerShell)

Open Cursor terminal in this folder:

`D:\Martin Workspace\Projects\excel-to-pdf`

Create and activate a virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Install dependencies:

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Run the app:

```powershell
streamlit run app.py
```

Then open the URL Streamlit prints (usually `http://localhost:8501`).

## How it works

- Upload Excel (`.xlsx`, `.xlsm`, `.xls`)
- Select sheets to include
- Preview the data
- Click **Generate PDF** → download

## Notes / limits

- Very large sheets can make PDFs huge; by default the PDF output truncates to **250 rows** and **20 columns** per sheet (you can change this in the sidebar).
- The PDF is generated using **ReportLab** for good Windows compatibility.

