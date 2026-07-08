# Bank Statement Automation — MVP

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/rohit-srivastava-ist/bank-statement-automation/blob/main/demo.ipynb)

**Gmail → Drive → Parse → Analyse → Visualise → Cross-check → Excel**

Fetch bank-statement PDFs from Gmail, store them systematically in Drive, parse, analyse, visualise, cross-check, mask PII (optional), and handle edge cases — all automated.

Built and demonstrated on **synthetic** statements. Everything runs end to end.

---

## Quick demo (one click)

Click the **Open in Colab** badge above → **Runtime → Run all**. Takes ~30 seconds. You'll see:
- 4 synthetic PDFs processed (including edge cases)
- 63 transactions parsed from 2 banks
- Charts rendered inline
- Validation flags catching deliberate anomalies
- A production-ready Excel workbook generated

---

## Architecture

| Stage | Tool | Where it runs |
|-------|------|---------------|
| **Fetch + Store** | `FetchStatements.gs` (Google Apps Script) | Inside your Google account, on a schedule — zero credential setup |
| **Parse → Analyse → Visualise → Cross-check → Mask → Export** | `statement_pipeline.py` (Python) | Google Colab or local |

Why this split: Apps Script has **native, permission-simple** access to Gmail and Drive — no OAuth JSON, no Cloud project. PDF parsing and analysis are far stronger in Python (`pdfplumber` handles statement tables best). Colab keeps the Python side inside Google too.

---

## Part 1 — The Gmail → Drive fetcher (Apps Script) · 5 min setup

1. Open **script.google.com** → **New project**. Paste `FetchStatements.gs`.
2. Edit the `CONFIG` block — set `SEARCH_QUERY` to match your statement emails.
3. Run **`fetchBankStatements`** once. Approve the permission prompt.
4. Check Drive: a **`Bank Statements/<year>/<year-month>/`** tree now holds your PDFs, renamed systematically (`2026-05-31__icicibank__May_Statement.pdf`), with a `_StatementLog` sheet for audit.
5. Run **`installTrigger`** once for automatic hourly execution.

Built-in safeguards: duplicate prevention, a `statement-saved` Gmail label so threads are never processed twice, and a Drive log sheet for full audit trail.

## Part 2 — The analysis pipeline (Python) · 30 seconds

```python
# In Colab:
from google.colab import drive
drive.mount('/content/drive')

!pip -q install pdfplumber pandas matplotlib openpyxl python-dateutil pypdf

import statement_pipeline as sp
sp.INBOX = '/content/drive/MyDrive/Bank Statements'
sp.run()
```

Or locally: drop PDFs in `inbox/` and run `python statement_pipeline.py`.

---

## Pipeline stages

1. **Fetch** — Apps Script (prod) / read `inbox/` (demo)
2. **Parse** — `pdfplumber`; header-mapping layer recognises each bank's column names (`Narration`/`Particulars` → description, `Withdrawal (Dr)`/`Debit` → debit, etc.), stitches multi-page tables, skips repeated headers
3. **Analyse** — rule-based categorisation (Salary, Rent, Groceries, Dining, etc.) + per-category/month summaries
4. **Visualise** — three charts: spend by category, income vs expense, balance over time. Embedded in Excel Dashboard sheet
5. **Cross-check** — running-balance reconciliation, duplicate detection, date-order checks, zero-value flags → Validation Flags sheet
6. **Mask PII** — optional (`MASK_PII = True`) to redact account numbers and IFSC codes
7. **Edge cases** — encrypted PDFs, blank/scanned PDFs, multi-page statements all handled gracefully

## Edge cases demonstrated

| File | Status | Behaviour |
|------|--------|-----------|
| `HDFC_Statement_Apr2026.pdf` | PARSED | 22 txns, single page |
| `ICICI_Statement_May2026.pdf` | PARSED | 41 txns across 2 pages; deliberate balance error + duplicate flagged |
| `EMPTY_scan_blank.pdf` | EMPTY | no text layer → flagged, skipped cleanly |
| `LOCKED_Statement_Jun2026.pdf` | ENCRYPTED | password-protected → skipped with clear reason |

---

## Output

```
outputs/
  Bank_Statement_Analysis.xlsx   — Dashboard + Transactions + Summaries + Validation Flags + File Status
  transactions_clean.csv         — flat export
  charts/*.png                   — spend by category, income vs expense, balance trend
  run_log.txt                    — full execution audit trail
```

---

## Files

```
FetchStatements.gs        Gmail → Drive automation (Apps Script)
statement_pipeline.py     Full 6-stage pipeline (Python)
gen_data.py               Generates synthetic test PDFs + edge cases
demo.ipynb                One-click Colab demo notebook
inbox/                    Synthetic input PDFs
outputs/                  Generated outputs
```

## Requirements

```
pdfplumber
pandas
matplotlib
openpyxl
python-dateutil
pypdf
reportlab  (only needed for gen_data.py)
```
