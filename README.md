# WildDoc

WildDoc is a Windows desktop tool for batch-generating Word documents from a
Word template and Excel data. It provides a GUI for mapping fields, validating
data, and exporting multiple documents in one run.

## Features
- Placeholder fields in Word templates
- Excel import and field mapping
- Data validation for required, number, and date fields
- Batch document generation

## Requirements
- Windows
- Python 3.8+ (recommended)
- Dependencies in `requirements.txt`

## Setup
```bash
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

## Run
```bash
python start_fixed.py
```

## Docs
See `docs/user_guide.md` for a detailed walkthrough.
