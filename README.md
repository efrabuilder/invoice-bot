# 🤖 Invoice Automation Bot

Python bot that extracts invoice data from PDFs, validates it, detects duplicates and generates Excel/CSV/PDF reports automatically.

Built by **Efraín Rojas Artavia**

---

## Features

- ✅ Extracts **invoice number, date, vendor and total** from any PDF
- ✅ **Validates** completeness, date validity, and number format (handles both `1,234.56` and `1.234,56`)
- ✅ **Detects duplicate** invoice numbers, both within a single run and **across separate runs** (persisted to disk)
- ✅ Flags PDFs with **no extractable text** (e.g. scanned images) instead of silently marking them "invalid"
- ✅ Exports results to **Excel** (color-coded) and **CSV**
- ✅ Generates a **PDF summary report**
- ✅ Processes entire folders of PDFs in one run, with configurable input/output folders via CLI flags
- ✅ Full **logging** with automatic rotation — every action is recorded
- ✅ Covered by an automated **test suite** (pytest) and a CI workflow (GitHub Actions)

---

## Requirements

- Python 3.10+

## Quick Start

```bash
# 1. Clone
git clone https://github.com/efrabuilder/invoice-bot.git
cd invoice-bot

# 2. (Recommended) create a virtual environment
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Generate sample invoices for testing
python generate_samples.py

# 5. Run the bot
python invoice_bot.py
```

Results appear in the `output/` folder.

---

## Usage

```bash
python invoice_bot.py [--input INVOICES_DIR] [--output OUTPUT_DIR] [--no-history] [--verbose]
```

- `--input` / `-i`: folder with invoice PDFs (default: `invoices/`)
- `--output` / `-o`: folder to write reports to (default: `output/`)
- `--no-history`: only check duplicates within this run, skip the persisted cross-run history
- `--verbose` / `-v`: debug-level logging

1. Place your invoice PDFs in the `invoices/` folder
2. Run `python invoice_bot.py`
3. Check `output/` for:
   - `invoices_YYYYMMDD_HHMMSS.xlsx` — color-coded Excel report
   - `invoices_YYYYMMDD_HHMMSS.csv`  — raw CSV data
   - `report_YYYYMMDD_HHMMSS.pdf`    — formatted PDF summary
   - `seen_invoices.json` — persisted invoice-number history used for cross-run duplicate detection
   - `invoice_bot.log` — rotating log (max 1MB × 3 backups)

---

## Running the tests

```bash
pytest tests/ -v
```

Tests cover amount parsing (US and EU/LatAm decimal formats), date validation, field extraction, validation rules, and both in-run and cross-run duplicate detection. A GitHub Actions workflow (`.github/workflows/tests.yml`) runs the suite on every push and pull request.

---

## Output Color Coding (Excel)

| Color | Meaning |
|-------|---------|
| 🟢 Green | Valid invoice |
| 🔴 Red | Invalid — missing fields, bad data, or unreadable PDF |
| 🟡 Yellow | Duplicate invoice number |

---

## Sample Output

```
2026-03-15 08:01 [INFO] Found 5 PDF(s) to process.
2026-03-15 08:01 [INFO] Processing: invoice_001.pdf
2026-03-15 08:01 [INFO]   → VALID | Invoice: INV-2026-001 | Total: 925.00
2026-03-15 08:01 [INFO] Processing: invoice_004_duplicate.pdf
2026-03-15 08:01 [WARNING] Duplicate invoice number: INV-2026-001
2026-03-15 08:01 [INFO] Results: 3 valid | 1 invalid | 1 duplicate
```

---

## Tech Stack

| Tool | Purpose |
|------|---------|
| `pdfplumber` | PDF text extraction |
| `pandas` | Data manipulation |
| `openpyxl` | Excel report generation |
| `fpdf2` | PDF report generation |
| `pytest` | Automated testing |

---

## License
MIT
