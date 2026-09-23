# Input File Validator

A Streamlit application for validating, comparing, and placing clinical study data files against template specifications. Designed for pharmaceutical/clinical data management workflows.

## Features

### Compare & Validate (Primary)
- **Intelligent file matching** — matches placed files to templates using multi-signal similarity scoring (weighted Jaccard, containment, LCS ratio)
- **Comprehensive validation** — checks filenames, file types (.csv vs .xlsx), sheet names, column headers (extra/missing/renamed), data types, and date formats
- **Auto-fix capabilities** — rename files, convert CSV→XLSX, rename sheets/columns, remove extra columns, reformat date columns
- **Place files** — copy validated and fixed files to an output location

### Extract (Secondary)
- Generate a `CPT_files.xlsx` metadata summary from selected Excel files
- Extracts file classification, paths, sheet names, header row ranges, subject/status columns, and output dataset naming

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
streamlit run app.py
```

Then open the URL shown in your terminal (usually http://localhost:8501).

## Deployment (Posit Connect)

This app is configured for Posit Connect deployment with:
- `manifest.json` — deployment manifest
- `runtime.txt` — Python version specification
- `.streamlit/config.toml` — server configuration

## Project Structure

```
├── app.py                  # Main entry point
├── utils.py                # Shared utility functions
├── pages/
│   ├── compare.py          # Compare & Validate page
│   └── extract.py          # CPT metadata extraction page
├── .streamlit/
│   └── config.toml         # Streamlit theme & server config
├── study_config.json       # Study/path configuration
├── manifest.json           # Posit Connect deployment manifest
├── requirements.txt        # Python dependencies
└── runtime.txt             # Python version for Posit Connect
```

## Requirements

- Python 3.13+
- streamlit
- pandas
- openpyxl
- numpy
- xlrd (for legacy .xls files)
