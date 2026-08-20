# Financial Statement Generator

Python 3.11+ FastAPI service that will convert uploaded financial statement documents into an Excel workbook following **ICAI Division I – Non-Ind AS Schedule III**.

This repository currently contains the **project skeleton only**. PDF extraction, classification, mapping, validation, and Excel generation are not implemented yet.

## Requirements

- Python 3.11 or later

## Setup

```bash
cd financial_statement_generator
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Place these reference files before implementing generation:

- `templates/Financial Statements_Sample.xlsx` — Division I Schedule III Excel template
- `reference/ICAI_GN_Div_I_Sch_III.pdf` — ICAI Guidance Note on Division I of Schedule III

## Run

From the `financial_statement_generator` directory:

```bash
uvicorn app.main:app --reload
```

Then open:

- API root: http://127.0.0.1:8000/
- Health: http://127.0.0.1:8000/health
- Docs: http://127.0.0.1:8000/docs

## Configuration

Paths are defined in `app/config.py` and can be overridden with environment variables or `.env`:

| Variable | Default | Purpose |
| --- | --- | --- |
| `INPUT_DIR` | `input` | Uploaded source documents |
| `REFERENCE_DIR` | `reference` | ICAI guidance note and other references |
| `TEMPLATE_DIR` | `templates` | Schedule III Excel template |
| `OUTPUT_DIR` | `output` | Generated workbooks |

## Project layout

```
financial_statement_generator/
├── app/                  # FastAPI application
│   ├── api/              # HTTP routes
│   ├── extraction/       # PDF / table extraction (stub)
│   ├── classification/   # Document type classification (stub)
│   ├── mapping/          # Schedule III mapping (stub)
│   ├── validation/       # Financial checks (stub)
│   ├── excel/            # Workbook generation (stub)
│   ├── models/           # Pydantic models
│   └── config.py         # Path and runtime settings
├── templates/            # Excel template
├── reference/            # ICAI Division I Schedule III guidance
├── output/               # Generated workbooks
└── tests/
```
