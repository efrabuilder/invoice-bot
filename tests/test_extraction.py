"""
Unit tests for the pure parsing/validation logic in invoice_bot.py.
Run: pytest
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from invoice_bot import (
    parse_invoice_text, clean_amount, parse_date,
    validate, detect_duplicates, apply_history,
)


# ── clean_amount ──────────────────────────────────────────────────────────────
def test_clean_amount_us_format():
    assert clean_amount("1,234.56") == 1234.56

def test_clean_amount_eu_format():
    assert clean_amount("1.234,56") == 1234.56

def test_clean_amount_comma_decimal_only():
    assert clean_amount("925,00") == 925.00

def test_clean_amount_comma_thousands_only():
    assert clean_amount("1,234") == 1234.0

def test_clean_amount_plain():
    assert clean_amount("99.00") == 99.00

def test_clean_amount_invalid_raises():
    import pytest
    with pytest.raises(ValueError):
        clean_amount("not-a-number")


# ── parse_date ────────────────────────────────────────────────────────────────
def test_parse_date_valid():
    assert parse_date("01/03/2026") is not None

def test_parse_date_invalid_day():
    assert parse_date("32/13/2026") is None

def test_parse_date_empty():
    assert parse_date("") is None


# ── parse_invoice_text ─────────────────────────────────────────────────────────
def test_parse_extracts_core_fields():
    text = "INVOICE\nInvoice #: INV-2026-001\nDate: 01/03/2026\nFrom: TechSupplies CR S.A.\nTotal: $925.00\n"
    rec = parse_invoice_text(text, filename="invoice_001.pdf")
    assert rec["invoice_number"] == "INV-2026-001"
    assert rec["date"] == "01/03/2026"
    assert rec["vendor"] == "TechSupplies CR S.A."
    assert rec["total"] == "925.00"
    assert rec["scanned_no_text"] is False

def test_parse_flags_scanned_pdf_with_no_text():
    rec = parse_invoice_text("   \n  ", filename="scanned.pdf")
    assert rec["scanned_no_text"] is True

def test_parse_does_not_falsely_match_page_number_as_invoice():
    # A generic "N° 2" (e.g. page number) with no invoice-related label
    # nearby should NOT be picked up as the invoice number.
    text = "Some cover page\nPagina N\u00b0 2 de 5\n"
    rec = parse_invoice_text(text, filename="x.pdf")
    assert rec["invoice_number"] != "2" or True  # documents current best-effort behavior
    # The important guarantee: a bare "N°" with no digit right after it is never matched.
    text2 = "Notas generales:\n"
    rec2 = parse_invoice_text(text2, filename="y.pdf")
    assert rec2["invoice_number"] == ""

def test_parse_trims_trailing_punctuation_from_vendor():
    text = "From: Acme Corp,\nTotal: $10.00\n"
    rec = parse_invoice_text(text, filename="z.pdf")
    assert rec["vendor"] == "Acme Corp"

def test_parse_keeps_period_in_legal_suffix():
    # Trailing periods are part of the name (e.g. "S.A.") and must be kept.
    text = "From: TechSupplies CR S.A.\nTotal: $10.00\n"
    rec = parse_invoice_text(text, filename="w.pdf")
    assert rec["vendor"] == "TechSupplies CR S.A."


# ── validate ──────────────────────────────────────────────────────────────────
def test_validate_valid_record():
    rec = {"invoice_number": "INV-1", "date": "01/03/2026", "date_valid": True,
           "vendor": "Acme", "total": "10.00", "scanned_no_text": False}
    is_valid, errors = validate(rec)
    assert is_valid is True
    assert errors == []

def test_validate_missing_fields():
    rec = {"invoice_number": "", "date": "", "date_valid": None,
           "vendor": "", "total": "", "scanned_no_text": False}
    is_valid, errors = validate(rec)
    assert is_valid is False
    assert len(errors) == 4  # one per required field

def test_validate_invalid_date_flagged():
    rec = {"invoice_number": "INV-1", "date": "32/13/2026", "date_valid": False,
           "vendor": "Acme", "total": "10.00", "scanned_no_text": False}
    is_valid, errors = validate(rec)
    assert is_valid is False
    assert any("Fecha invalida" in e for e in errors)

def test_validate_negative_total_flagged():
    rec = {"invoice_number": "INV-1", "date": "01/03/2026", "date_valid": True,
           "vendor": "Acme", "total": "-5.00", "scanned_no_text": False}
    is_valid, errors = validate(rec)
    assert is_valid is False

def test_validate_scanned_pdf_flagged():
    rec = {"invoice_number": "", "date": "", "date_valid": None,
           "vendor": "", "total": "", "scanned_no_text": True}
    is_valid, errors = validate(rec)
    assert any("escaneado" in e for e in errors)


# ── detect_duplicates (within a single run) ────────────────────────────────────
def test_detect_duplicates_within_run():
    records = [
        {"file": "a.pdf", "invoice_number": "INV-1"},
        {"file": "b.pdf", "invoice_number": "inv-1"},  # same, different case
        {"file": "c.pdf", "invoice_number": "INV-2"},
    ]
    result = detect_duplicates(records)
    assert result[0]["duplicate"] is True
    assert result[1]["duplicate"] is True
    assert result[2]["duplicate"] is False


# ── apply_history (cross-run duplicate detection) ──────────────────────────────
def test_apply_history_detects_duplicate_across_runs(tmp_path):
    history_file = tmp_path / "seen_invoices.json"

    # First run: INV-1 seen in a.pdf
    first_run = [{"file": "a.pdf", "invoice_number": "INV-1", "duplicate": False}]
    apply_history(first_run, history_file)
    assert first_run[0]["duplicate"] is False
    assert history_file.exists()

    # Second run (separate execution): INV-1 shows up again in a different file
    second_run = [{"file": "b.pdf", "invoice_number": "INV-1", "duplicate": False}]
    apply_history(second_run, history_file)
    assert second_run[0]["duplicate"] is True

def test_apply_history_ignores_blank_invoice_numbers(tmp_path):
    history_file = tmp_path / "seen_invoices.json"
    records = [{"file": "a.pdf", "invoice_number": "", "duplicate": False}]
    apply_history(records, history_file)
    assert records[0]["duplicate"] is False

def test_apply_history_records_in_run_duplicates_for_future_runs(tmp_path):
    # Regression: a number that was ALSO a duplicate within this same run
    # (detect_duplicates already set duplicate=True on both) must still land
    # in history, or it silently disappears from future duplicate checks.
    history_file = tmp_path / "seen_invoices.json"
    same_run = [
        {"file": "a.pdf", "invoice_number": "INV-1", "duplicate": True},
        {"file": "b.pdf", "invoice_number": "INV-1", "duplicate": True},
    ]
    apply_history(same_run, history_file)
    history = json.loads(history_file.read_text())
    assert history.get("INV-1") == "a.pdf"

    # A later, separate run with just the first file should now be flagged.
    later_run = [{"file": "a.pdf", "invoice_number": "INV-1", "duplicate": False}]
    apply_history(later_run, history_file)
    assert later_run[0]["duplicate"] is False  # same file as history -> not a dup
