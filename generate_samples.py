"""
generate_samples.py — generates sample invoice PDFs for testing.
Run: py generate_samples.py
"""
from fpdf import FPDF
from pathlib import Path

Path("invoices").mkdir(exist_ok=True)

invoices = [
    {"number": "INV-2026-001", "date": "01/03/2026", "vendor": "TechSupplies CR S.A.", "items": [("Laptop Dell Latitude", 1, 850.00), ("Wireless Mouse", 3, 25.00)], "file": "invoice_001.pdf"},
    {"number": "INV-2026-002", "date": "05/03/2026", "vendor": "Office Depot Costa Rica", "items": [("Paper A4 500 sheets", 10, 8.50), ("Stapler", 2, 12.00)], "file": "invoice_002.pdf"},
    {"number": "INV-2026-003", "date": "10/03/2026", "vendor": "CloudServices Inc.", "items": [("AWS Hosting March", 1, 320.00), ("SSL Certificate", 1, 45.00)], "file": "invoice_003.pdf"},
    {"number": "INV-2026-001", "date": "11/03/2026", "vendor": "TechSupplies CR S.A.", "items": [("USB Hub", 5, 18.00)], "file": "invoice_004_duplicate.pdf"},
    {"number": "", "date": "", "vendor": "", "items": [("Unknown item", 1, 99.00)], "file": "invoice_005_invalid.pdf"},
]

for inv in invoices:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_fill_color(13, 13, 13)
    pdf.rect(0, 0, 210, 297, "F")

    # Header
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(255, 107, 53)
    pdf.ln(12)
    pdf.cell(0, 12, "INVOICE", ln=True, align="C")

    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(200, 200, 200)
    if inv["number"]:
        pdf.cell(0, 7, f"Invoice #: {inv['number']}", ln=True, align="C")
    if inv["date"]:
        pdf.cell(0, 7, f"Date: {inv['date']}", ln=True, align="C")
    pdf.ln(8)

    # Vendor
    if inv["vendor"]:
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(255, 107, 53)
        pdf.cell(0, 7, "From:", ln=True)
        pdf.set_font("Helvetica", "", 11)
        pdf.set_text_color(240, 238, 232)
        pdf.cell(0, 7, inv["vendor"], ln=True)
    pdf.ln(8)

    # Items table
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_fill_color(255, 107, 53)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(90, 8, "Description", fill=True)
    pdf.cell(30, 8, "Qty", fill=True, align="C")
    pdf.cell(35, 8, "Unit Price", fill=True, align="C")
    pdf.cell(35, 8, "Amount", fill=True, align="C")
    pdf.ln()

    total = 0
    pdf.set_font("Helvetica", "", 10)
    for i, (desc, qty, price) in enumerate(inv["items"]):
        amount = qty * price
        total += amount
        pdf.set_fill_color(25, 25, 25) if i % 2 == 0 else pdf.set_fill_color(20, 20, 20)
        pdf.set_text_color(240, 238, 232)
        pdf.cell(90, 7, desc, fill=True)
        pdf.cell(30, 7, str(qty), fill=True, align="C")
        pdf.cell(35, 7, f"${price:.2f}", fill=True, align="C")
        pdf.cell(35, 7, f"${amount:.2f}", fill=True, align="C")
        pdf.ln()

    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(255, 107, 53)
    pdf.cell(155, 9, "TOTAL:", align="R")
    pdf.set_text_color(240, 238, 232)
    pdf.cell(35, 9, f"${total:.2f}", align="C")
    pdf.ln()

    pdf.output(f"invoices/{inv['file']}")
    print(f"✅ Generated: invoices/{inv['file']}")

print("\nDone! Run: py invoice_bot.py")
