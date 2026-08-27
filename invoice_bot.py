"""
invoice_bot.py — Efraín Rojas Artavia
Extracts invoice data from PDFs, validates it, detects duplicates,
saves results to Excel/CSV and generates a summary report.
"""

import argparse
import logging
import json
import re
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path

import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from fpdf import FPDF

log = logging.getLogger(__name__)


def setup_logging(log_file: Path, verbose: bool = False) -> None:
    """Configure logging with a rotating file handler (max 1MB x 3 backups)."""
    handlers = [
        RotatingFileHandler(log_file, maxBytes=1_000_000, backupCount=3, encoding="utf-8"),
        logging.StreamHandler(),
    ]
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=handlers,
        force=True,
    )


# ── Amount / date parsing helpers ───────────────────────────────────────────────
def clean_amount(raw: str) -> float:
    """
    Parse an amount string into a float, supporting both:
      - US/UK style: 1,234.56  (comma = thousands, dot = decimal)
      - EU/LatAm style: 1.234,56  (dot = thousands, comma = decimal)
    Raises ValueError if it can't be parsed as a number.
    """
    if raw is None:
        raise ValueError("empty amount")
    s = raw.strip().replace("$", "").replace("₡", "").strip()
    if not s:
        raise ValueError("empty amount")

    has_comma = "," in s
    has_dot = "." in s

    if has_comma and has_dot:
        # Whichever separator appears last is the decimal separator.
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif has_comma and not has_dot:
        # Could be "1,234" (thousands) or "1234,56" (EU decimal).
        # Treat as decimal only if exactly 2 digits follow the comma.
        if re.match(r"^\d+,\d{2}$", s):
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    # dot-only or no separators: already valid Python float format

    return float(s)


DATE_FORMATS = ["%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y",
                 "%d.%m.%Y", "%d/%m/%y", "%m/%d/%y"]


def parse_date(raw: str):
    """Try to parse a date string against common formats. Returns a datetime or None."""
    if not raw:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


# ── PDF Extraction ─────────────────────────────────────────────────────────────
MIN_TEXT_LEN_FOR_SCAN_CHECK = 20  # below this, we assume the PDF has no real text layer


def parse_invoice_text(text: str, filename: str = "") -> dict:
    """
    Pure text -> fields parser (no file I/O), so this can be unit-tested
    directly with plain strings instead of real PDFs.
    """

    def find(patterns, txt, require_digit=False):
        for pat in patterns:
            for m in re.finditer(pat, txt, re.IGNORECASE):
                val = m.group(1).strip()
                if require_digit and not re.search(r"\d", val):
                    continue  # skip matches like the bare word "Invoice" with no number in it
                return val
        return ""

    # NOTE: "[:\-]" is required (not optional) here so the pattern doesn't match
    # a lone header word like "INVOICE" with nothing after it. require_digit=True
    # is a second safety net since real invoice numbers always contain a digit.
    # The "N°" pattern is deliberately last and requires a digit immediately
    # after the symbol, since on its own it's too generic (e.g. "Página N° 2").
    invoice_number = find([
        r"invoice\s*#\s*[:\-]?\s*([A-Z0-9\-]+)",
        r"invoice\s*[:\-]\s*([A-Z0-9\-]+)",
        r"factura\s*#?\s*[:\-]\s*([A-Z0-9\-]+)",
        r"inv\s*[:\-]\s*([A-Z0-9\-]+)",
        r"n[°º]\s*[:\-]?\s*(\d[A-Z0-9\-]*)",
    ], text, require_digit=True)

    date_raw = find([
        r"date\s*[:\-]\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
        r"fecha\s*[:\-]\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
        r"issued\s*[:\-]\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
        r"(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4})",
    ], text)
    parsed_date = parse_date(date_raw)

    vendor = find([
        r"from\s*[:\-]\s*(.+)",
        r"vendor\s*[:\-]\s*(.+)",
        r"proveedor\s*[:\-]\s*(.+)",
        r"bill\s*from\s*[:\-]\s*(.+)",
        r"company\s*[:\-]\s*(.+)",
    ], text)
    # Normalize: trim stray trailing commas/semicolons/whitespace picked up
    # from the regex capturing to end-of-line. Trailing periods are kept on
    # purpose since they're often part of the name itself (e.g. "S.A.").
    vendor = re.sub(r"[\s,;:]+$", "", vendor.strip())

    total_raw = find([
        r"total\s*[:\-]\s*\$?\s*([\d.,]+)",
        r"amount\s*due\s*[:\-]\s*\$?\s*([\d.,]+)",
        r"grand\s*total\s*[:\-]\s*\$?\s*([\d.,]+)",
        r"total\s*a\s*pagar\s*[:\-]\s*\$?\s*([\d.,]+)",
        r"\$\s*([\d.,]+)",
    ], text)

    currency = "USD"
    if re.search(r"€|EUR", text): currency = "EUR"
    elif re.search(r"₡|CRC|colones", text, re.IGNORECASE): currency = "CRC"

    return {
        "file":            filename,
        "invoice_number":  invoice_number,
        "date":            date_raw,
        "date_valid":      parsed_date is not None if date_raw else None,
        "vendor":          vendor[:60] if vendor else "",
        "total":           total_raw,
        "currency":        currency,
        "raw_text_len":    len(text),
        "scanned_no_text": len(text.strip()) < MIN_TEXT_LEN_FOR_SCAN_CHECK,
    }


def extract_invoice_data(pdf_path: Path) -> dict:
    """Extract key fields from an invoice PDF using regex patterns."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return parse_invoice_text(text, filename=pdf_path.name)


# ── Validation ─────────────────────────────────────────────────────────────────
REQUIRED_FIELDS = ["invoice_number", "date", "vendor", "total"]

def validate(record: dict) -> tuple:
    """Validate extracted invoice data. Returns (is_valid, list_of_errors)."""
    errors = []

    if record.get("scanned_no_text"):
        errors.append("Sin texto extraible (posible PDF escaneado / solo imagen)")

    for field in REQUIRED_FIELDS:
        if not record.get(field):
            errors.append(f"Missing: {field}")

    if record.get("date") and record.get("date_valid") is False:
        errors.append(f"Fecha invalida: {record['date']}")

    if record.get("total"):
        try:
            val = clean_amount(record["total"])
            if val <= 0:
                errors.append("Total must be > 0")
        except ValueError:
            errors.append(f"Invalid total format: {record['total']}")

    return (len(errors) == 0, errors)


# ── Duplicate Detection ────────────────────────────────────────────────────────
def detect_duplicates(records: list) -> list:
    """Mark duplicate invoice numbers within the current run."""
    seen = {}
    for rec in records:
        inv = rec.get("invoice_number", "").strip().upper()
        if not inv:
            rec["duplicate"] = False
            continue
        if inv in seen:
            rec["duplicate"] = True
            seen[inv]["duplicate"] = True
            log.warning(f"Duplicate invoice number: {inv} in {rec['file']}")
        else:
            seen[inv] = rec
            rec["duplicate"] = False
    return records


def apply_history(records: list, history_path: Path) -> list:
    """
    Cross-run duplicate detection: compares invoice numbers against a JSON
    file of previously-seen numbers (persisted across executions), then
    updates that file with the current run's valid invoice numbers.
    """
    history = {}
    if history_path.exists():
        try:
            history = json.loads(history_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as e:
            log.warning(f"Could not read history file ({e}); starting fresh.")

    # Bug fix: an invoice that was ALSO a duplicate within this same run
    # (detect_duplicates already flagged it) must still be recorded here
    # under its first-seen file, or it silently drops out of history and
    # stops being detected as a duplicate in later runs. We only skip
    # writing a NEW history entry for the second+ occurrence within this run.
    seen_this_run = set()
    for rec in records:
        inv = rec.get("invoice_number", "").strip().upper()
        if not inv:
            continue
        if inv in history and history[inv] != rec.get("file"):
            if not rec.get("duplicate"):
                log.warning(f"Duplicate invoice number: {inv} in {rec['file']} (already seen in {history[inv]})")
            rec["duplicate"] = True
        elif inv not in seen_this_run:
            history[inv] = rec.get("file")
        seen_this_run.add(inv)

    try:
        history_path.write_text(json.dumps(history, indent=2, ensure_ascii=False), encoding="utf-8")
    except OSError as e:
        log.warning(f"Could not save history file: {e}")

    return records


# ── Excel Export ───────────────────────────────────────────────────────────────
HEADER_FILL  = PatternFill("solid", fgColor="FF6B35")
HEADER_FONT  = Font(bold=True, color="FFFFFF", size=11)
ALT_FILL     = PatternFill("solid", fgColor="1E1E1E")
VALID_FILL   = PatternFill("solid", fgColor="0D3320")
INVALID_FILL = PatternFill("solid", fgColor="3B0D0D")
DUP_FILL     = PatternFill("solid", fgColor="3B2A00")

COLUMNS = ["file", "invoice_number", "date", "vendor", "total", "currency",
           "valid", "duplicate", "errors"]

def export_excel(records: list, output_dir: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    ws.sheet_view.showGridLines = False

    # Header
    for ci, col in enumerate(COLUMNS, 1):
        cell = ws.cell(1, ci, col.replace("_", " ").title())
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(ci)].width = max(len(col) + 6, 16)

    # Rows
    for ri, rec in enumerate(records, 2):
        row_fill = INVALID_FILL if not rec.get("valid") else (DUP_FILL if rec.get("duplicate") else (ALT_FILL if ri % 2 == 0 else None))
        for ci, col in enumerate(COLUMNS, 1):
            val = rec.get(col, "")
            if isinstance(val, list): val = "; ".join(val)
            elif isinstance(val, bool): val = "Yes" if val else "No"
            cell = ws.cell(ri, ci, val)
            cell.font = Font(color="F0EEE8", size=10)
            cell.alignment = Alignment(horizontal="center")
            if row_fill: cell.fill = row_fill

    path = output_dir / f"invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(path)
    log.info(f"Excel saved: {path.name}")
    return path


# ── CSV Export ─────────────────────────────────────────────────────────────────
def export_csv(records: list, output_dir: Path) -> Path:
    df = pd.DataFrame(records)[COLUMNS]
    df["errors"] = df["errors"].apply(lambda x: "; ".join(x) if isinstance(x, list) else x)
    path = output_dir / f"invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    df.to_csv(path, index=False)
    log.info(f"CSV saved: {path.name}")
    return path


# ── PDF Report ─────────────────────────────────────────────────────────────────
class ReportPDF(FPDF):
    def header(self):
        self.set_fill_color(13, 13, 13)
        self.rect(0, 0, 210, 18, "F")
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(255, 107, 53)
        self.cell(0, 18, "Invoice Processing Report", align="L", ln=True)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(136, 136, 136)
        self.cell(0, 10, f"Page {self.page_no()} | Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}", align="C")


def export_pdf_report(records: list, stats: dict, output_dir: Path) -> Path:
    pdf = ReportPDF()
    pdf.set_auto_page_break(True, margin=15)
    pdf.add_page()
    pdf.set_fill_color(13, 13, 13)
    pdf.rect(0, 0, 210, 297, "F")

    # Title
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_text_color(255, 107, 53)
    pdf.ln(8)
    pdf.cell(0, 10, "Invoice Processing Report", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(136, 136, 136)
    pdf.cell(0, 6, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Total PDFs: {stats['total']}", ln=True)
    pdf.ln(6)

    # Summary stats
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(240, 238, 232)
    pdf.cell(0, 8, "Summary", ln=True)
    pdf.ln(2)

    stat_items = [
        ("Total invoices",  stats["total"],      (255,107,53)),
        ("Valid",           stats["valid"],       (16,185,129)),
        ("Invalid",         stats["invalid"],     (239,68,68)),
        ("Duplicates",      stats["duplicates"],  (245,158,11)),
        ("Total amount",    f"${stats['total_amount']:,.2f}", (59,130,246)),
    ]

    for label, val, color in stat_items:
        pdf.set_fill_color(20, 20, 20)
        pdf.set_text_color(*color)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(80, 8, label, fill=True)
        pdf.set_text_color(240, 238, 232)
        pdf.set_font("Helvetica", "", 11)
        pdf.cell(0, 8, str(val), fill=True, ln=True)
        pdf.ln(1)

    pdf.ln(8)

    # Invoice table
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(240, 238, 232)
    pdf.cell(0, 8, "Invoice Details", ln=True)
    pdf.ln(2)

    headers = ["File", "Invoice #", "Date", "Vendor", "Total", "Status"]
    widths  = [35, 25, 22, 55, 22, 22]

    pdf.set_fill_color(255, 107, 53)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8)
    for h, w in zip(headers, widths):
        pdf.cell(w, 7, h, fill=True)
    pdf.ln()

    pdf.set_font("Helvetica", "", 8)
    for i, rec in enumerate(records):
        if rec.get("duplicate"):
            pdf.set_fill_color(59, 42, 0)
        elif not rec.get("valid"):
            pdf.set_fill_color(59, 13, 13)
        elif i % 2 == 0:
            pdf.set_fill_color(20, 20, 20)
        else:
            pdf.set_fill_color(26, 26, 26)

        # Plain ASCII labels: the default Helvetica core font can't encode
        # symbols like ⚠ ✓ ✗, which crashed PDF generation whenever there
        # was a duplicate or invalid record.
        status = "DUP" if rec.get("duplicate") else ("OK" if rec.get("valid") else "ERR")
        pdf.set_text_color(240, 238, 232)
        row = [
            rec.get("file", "")[:18],
            rec.get("invoice_number", "")[:12],
            rec.get("date", "")[:10],
            rec.get("vendor", "")[:28],
            f"${rec.get('total','0')}",
            status,
        ]
        for val, w in zip(row, widths):
            pdf.cell(w, 6, str(val), fill=True)
        pdf.ln()

    path = output_dir / f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    pdf.output(str(path))
    log.info(f"PDF report saved: {path.name}")
    return path


# ── CLI ──────────────────────────────────────────────────────────────────────
def parse_args():
    parser = argparse.ArgumentParser(description="Extract, validate and report on invoice PDFs.")
    parser.add_argument("--input", "-i", type=Path, default=Path("invoices"),
                        help="Folder with invoice PDFs to process (default: invoices/)")
    parser.add_argument("--output", "-o", type=Path, default=Path("output"),
                        help="Folder to write reports to (default: output/)")
    parser.add_argument("--no-history", action="store_true",
                        help="Disable cross-run duplicate detection (only checks duplicates within this run)")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Enable debug-level logging")
    return parser.parse_args()


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    args = parse_args()
    args.input.mkdir(exist_ok=True)
    args.output.mkdir(exist_ok=True)
    setup_logging(args.output / "invoice_bot.log", verbose=args.verbose)

    log.info("=== Invoice Bot started ===")

    pdf_files = list(args.input.glob("*.pdf"))
    if not pdf_files:
        log.warning(f"No PDF files found in '{args.input}/'. Add invoice PDFs and run again.")
        return

    log.info(f"Found {len(pdf_files)} PDF(s) to process.")
    records = []

    for pdf_path in pdf_files:
        log.info(f"Processing: {pdf_path.name}")
        try:
            data = extract_invoice_data(pdf_path)
            is_valid, errors = validate(data)
            data["valid"]  = is_valid
            data["errors"] = errors
            records.append(data)
            log.info(f"  → {'VALID' if is_valid else 'INVALID'} | Invoice: {data['invoice_number']} | Total: {data['total']}")
        except Exception as e:
            log.error(f"  → Failed to process {pdf_path.name}: {e}")
            records.append({
                "file": pdf_path.name, "invoice_number": "", "date": "",
                "vendor": "", "total": "", "currency": "", "valid": False,
                "duplicate": False, "errors": [str(e)], "raw_text_len": 0,
            })

    records = detect_duplicates(records)
    if not args.no_history:
        records = apply_history(records, args.output / "seen_invoices.json")

    # Stats
    total_amount = 0.0
    for rec in records:
        try:
            total_amount += clean_amount(rec.get("total", ""))
        except ValueError:
            pass

    stats = {
        "total":        len(records),
        "valid":        sum(1 for r in records if r.get("valid")),
        "invalid":      sum(1 for r in records if not r.get("valid")),
        "duplicates":   sum(1 for r in records if r.get("duplicate")),
        "total_amount": total_amount,
    }

    log.info(f"Results: {stats['valid']} valid | {stats['invalid']} invalid | {stats['duplicates']} duplicates")

    export_excel(records, args.output)
    export_csv(records, args.output)
    export_pdf_report(records, stats, args.output)

    log.info("=== Invoice Bot completed. Check output/ folder. ===")


if __name__ == "__main__":
    main()
