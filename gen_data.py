"""Generate synthetic, realistic bank-statement PDFs (+ deliberate edge cases)."""
import os, random
from datetime import date, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph,
                                Spacer)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from pypdf import PdfReader, PdfWriter

random.seed(7)
INBOX = os.path.join(os.path.dirname(__file__), "inbox")
os.makedirs(INBOX, exist_ok=True)
styles = getSampleStyleSheet()

# ---- transaction vocabulary (description -> implied category, not embedded in PDF) ----
DEBITS = [
    ("UPI/DR/SWIGGY BANGALORE", 180, 900),
    ("UPI/DR/ZOMATO ONLINE", 220, 1100),
    ("POS/DMART FARIDABAD", 900, 4200),
    ("UPI/DR/BIGBASKET", 600, 3000),
    ("NEFT/DR/HDFC HOME LOAN EMI", 28500, 28500),
    ("ACH/DR/HDFC LIFE INSURANCE", 3200, 3200),
    ("UPI/DR/UBER INDIA", 120, 700),
    ("POS/INDIAN OIL FUEL", 1000, 3500),
    ("UPI/DR/AMAZON PAY", 350, 5200),
    ("UPI/DR/FLIPKART", 500, 6000),
    ("ATM/CASH WITHDRAWAL", 2000, 10000),
    ("BILLPAY/AIRTEL MOBILE", 399, 799),
    ("BILLPAY/BSES ELECTRICITY", 900, 3200),
    ("UPI/DR/NETFLIX SUBSCRIPTION", 199, 649),
    ("UPI/DR/RENT TO LANDLORD", 22000, 22000),
]
CREDITS = [
    ("NEFT/CR/ACME CORP SALARY", 95000, 145000),
    ("UPI/CR/REFUND AMAZON", 350, 2200),
    ("INT.CR/SAVINGS INTEREST", 210, 640),
    ("UPI/CR/FRIEND SETTLEMENT", 500, 4000),
]

def rupees(x): return f"{x:,.2f}"

def build_txns(start: date, n: int, opening: float):
    """Return list of dicts with running balance computed correctly."""
    rows, bal, d = [], opening, start
    for i in range(n):
        d = d + timedelta(days=random.randint(0, 2))
        if random.random() < 0.30:
            desc, lo, hi = random.choice(CREDITS)
            amt = round(random.uniform(lo, hi), 2)
            bal += amt
            rows.append({"date": d, "desc": desc, "dr": 0.0, "cr": amt, "bal": round(bal, 2)})
        else:
            desc, lo, hi = random.choice(DEBITS)
            amt = round(random.uniform(lo, hi), 2)
            if amt > bal:  # keep it plausible
                amt = round(bal * 0.2, 2)
            bal -= amt
            rows.append({"date": d, "desc": desc, "dr": amt, "cr": 0.0, "bal": round(bal, 2)})
    return rows

def header_block(bank, addr, holder, acct, extra_label, extra_val, period):
    h = ParagraphStyle("h", parent=styles["Title"], fontSize=16, spaceAfter=2,
                       textColor=colors.HexColor("#0C1C2C"))
    small = ParagraphStyle("s", parent=styles["Normal"], fontSize=8.5, leading=11,
                           textColor=colors.HexColor("#3a4a58"))
    b = ParagraphStyle("b", parent=styles["Normal"], fontSize=9.5, leading=13)
    els = [Paragraph(bank, h), Paragraph(addr, small), Spacer(1, 6),
           Paragraph(f"<b>Account Holder:</b> {holder}", b),
           Paragraph(f"<b>Account Number:</b> {acct}", b),
           Paragraph(f"<b>{extra_label}:</b> {extra_val}", b),
           Paragraph(f"<b>Statement Period:</b> {period}", b),
           Spacer(1, 10)]
    return els

def make_statement(fname, bank, addr, holder, acct, extra_label, extra_val,
                   headers, rows, date_fmt, opening, closing):
    """headers: list of column titles in this bank's own words."""
    path = os.path.join(INBOX, fname)
    doc = SimpleDocTemplate(path, pagesize=A4, topMargin=15*mm, bottomMargin=15*mm,
                            leftMargin=14*mm, rightMargin=14*mm, title="Statement")
    period = f"{rows[0]['date'].strftime(date_fmt)} to {rows[-1]['date'].strftime(date_fmt)}"
    els = header_block(bank, addr, holder, acct, extra_label, extra_val, period)
    els.append(Paragraph(f"<b>Opening Balance:</b> INR {rupees(opening)} &nbsp;&nbsp; "
                         f"<b>Closing Balance:</b> INR {rupees(closing)}",
                         ParagraphStyle("ob", parent=styles["Normal"], fontSize=9.5,
                                        spaceAfter=8)))
    data = [headers]
    for r in rows:
        data.append([r["date"].strftime(date_fmt), r["desc"],
                     "" if r["dr"] == 0 else rupees(r["dr"]),
                     "" if r["cr"] == 0 else rupees(r["cr"]),
                     rupees(r["bal"])])
    col_w = [24*mm, 74*mm, 26*mm, 26*mm, 30*mm]
    t = Table(data, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0C1C2C")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 7.6),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F4F6")]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#C9D2DA")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]))
    els.append(t)
    els.append(Spacer(1, 8))
    els.append(Paragraph("This is a computer-generated statement and does not require "
                         "a signature. SYNTHETIC DATA — for demonstration only.",
                         ParagraphStyle("f", parent=styles["Normal"], fontSize=7.5,
                                        textColor=colors.grey)))
    doc.build(els)
    return path

# ---------- Statement 1: HDFC-like, clean, dd/mm/yy ----------
r1 = build_txns(date(2026, 4, 1), 22, opening=84000.0)
make_statement("HDFC_Statement_Apr2026.pdf",
    "HDFC BANK", "Ground Floor, Sector 16, Faridabad, Haryana 121002",
    "AMISH KUMAR", "5010 0123 4567 8901", "IFSC", "HDFC0001234",
    ["Date", "Narration", "Withdrawal (Dr)", "Deposit (Cr)", "Closing Balance"],
    r1, "%d/%m/%y", 84000.0, r1[-1]["bal"])

# ---------- Statement 2: ICICI-like, multi-page, DD-MON-YYYY, with anomalies ----------
r2 = build_txns(date(2026, 5, 1), 40, opening=r1[-1]["bal"])
# EDGE CASE A: duplicate transaction (exact repeat of a row) -> dedup should flag
dup = dict(r2[10]); r2.insert(11, dup)
# recompute balances after insertion so only the injected corruption breaks reconciliation
bal = 96000.0 if False else None
# EDGE CASE B: corrupt ONE balance value -> reconciliation must catch it
r2b = []
running = r1[-1]["bal"]
for r in r2:
    running = running - r["dr"] + r["cr"]
    r = dict(r); r["bal"] = round(running, 2)
    r2b.append(r)
r2b[20]["bal"] = round(r2b[20]["bal"] + 5000.0, 2)   # <-- injected error
make_statement("ICICI_Statement_May2026.pdf",
    "ICICI BANK", "Nariman Point, Mumbai, Maharashtra 400021",
    "AMISH KUMAR", "6220 9988 7766 5544", "IFSC", "ICIC0006220",
    ["Txn Date", "Particulars", "Debit", "Credit", "Balance"],
    r2b, "%d-%b-%Y", r1[-1]["bal"], r2b[-1]["bal"])

# ---------- EDGE CASE C: empty / no-text PDF ----------
from reportlab.pdfgen import canvas
c = canvas.Canvas(os.path.join(INBOX, "EMPTY_scan_blank.pdf"), pagesize=A4); c.showPage(); c.save()

# ---------- EDGE CASE D: password-protected PDF ----------
tmp = os.path.join(INBOX, "_tmp_locked.pdf")
c = canvas.Canvas(tmp, pagesize=A4)
c.drawString(80, 760, "LOCKED STATEMENT - requires password"); c.showPage(); c.save()
reader = PdfReader(tmp); writer = PdfWriter()
for p in reader.pages: writer.add_page(p)
writer.encrypt("secret123")
with open(os.path.join(INBOX, "LOCKED_Statement_Jun2026.pdf"), "wb") as f:
    writer.write(f)
os.remove(tmp)

print("Generated files in inbox/:")
for f in sorted(os.listdir(INBOX)):
    print("  ", f, os.path.getsize(os.path.join(INBOX, f)), "bytes")
