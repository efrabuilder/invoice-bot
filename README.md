# 🤖 Invoice Automation Bot

Python bot that extracts invoice data from PDFs, validates it, detects duplicates and generates Excel/CSV/PDF reports automatically.

Built by **Efraín Rojas Artavia**

---

## Features

- ✅ Extracts **invoice number, date, vendor and total** from any PDF
- ✅ **Validates** completeness and data format
- ✅ **Detects duplicate** invoice numbers automatically
- ✅ Exports results to **Excel** (color-coded) and **CSV**
- ✅ Generates a **PDF summary report**
- ✅ Processes entire folders of PDFs in one run
- ✅ Full **logging** — every action is recorded

---

## Quick Start

```bash
# 1. Clone
git clone https://github.com/efrabuilder/invoice-bot.git
cd invoice-bot

# 2. Install dependencies
py -m pip install -r requirements.txt

# 3. Generate sample invoices for testing
py generate_samples.py

# 4. Run the bot
py invoice_bot.py
```

Results appear in the `output/` folder.

---

## Usage

1. Place your invoice PDFs in the `invoices/` folder
2. Run `py invoice_bot.py`
3. Check `output/` for:
   - `invoices_YYYYMMDD_HHMMSS.xlsx` — color-coded Excel report
   - `invoices_YYYYMMDD_HHMMSS.csv`  — raw CSV data
   - `report_YYYYMMDD_HHMMSS.pdf`    — formatted PDF summary

---

## Output Color Coding (Excel)

| Color | Meaning |
|-------|---------|
| 🟢 Green | Valid invoice |
| 🔴 Red | Invalid — missing fields or bad data |
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

---

## License
MIT
