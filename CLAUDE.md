# Project: COMBIMATCH — Rebuild

---

## About Me

I am a CA/GST consultant, not a programmer. I build tools through
conversation with AI. Explain technical decisions in plain English.
Ask before making architectural choices. Never assume I understand
technical jargon — explain it on first use.

## Tech Stack

- Python 3.x
- PyQt5 for desktop interface
- openpyxl for Excel handling
- [any other libraries you've decided on]

## Reference Documents

- `.build-docs/PROJECT_AUDIT_COMBIMATCH.md` — Full audit of the original tool
- `.build-docs/REBUILD_BRIEF_COMBIMATCH.md` — My decisions for the rebuild
- `.build-docs/DECISIONS_LOG.md` — Running log of every discussion and decision
- `.build-docs/context/` — Additional reference files (samples, templates, specs)

## Current Phase

Phase 1: [Update this as you progress]

---

# CODING STANDARDS — MANDATORY FOR ALL CODE

These standards apply to EVERY file and EVERY function in this project.
No exceptions. No shortcuts. These exist so that a non-programmer can
read the code six months later and understand what everything does,
why it exists, how it connects to everything else, and what assumptions
were made.

---

## Standard 1: Modular Folder Structure From Day One

NEVER build in a single file. NEVER start flat and split later.
Always use folder-based structure from the very first line of code:

```
my-tool/
├── main.py                      ← Entry point ONLY. Starts the app.
│
├── config/                      ← All settings, constants, mappings
│   ├── __init__.py
│   ├── settings.py              ← Paths, thresholds, environment config
│   ├── constants.py             ← Fixed values (tax rates, state codes, limits)
│   └── mappings.py              ← Lookup tables (state code → name, HSN → rate)
│
├── core/                        ← Business logic, one file per domain area
│   ├── __init__.py
│   ├── gstin_validator.py
│   ├── tax_calculator.py
│   ├── invoice_validator.py
│   └── return_classifier.py
│
├── models/                      ← Data structures / object definitions
│   ├── __init__.py
│   ├── invoice.py               ← What an Invoice object looks like
│   ├── error_entry.py           ← What an Error record looks like
│   └── client.py                ← What a Client record looks like
│
├── readers/                     ← Input handling, one file per format
│   ├── __init__.py
│   ├── excel_reader.py
│   ├── csv_reader.py
│   └── json_reader.py
│
├── writers/                     ← Output generation, one file per type
│   ├── __init__.py
│   ├── excel_writer.py
│   ├── report_writer.py
│   └── json_writer.py
│
├── gui/                         ← Interface, one file per screen/dialog
│   ├── __init__.py
│   ├── main_window.py
│   ├── settings_dialog.py
│   ├── error_viewer.py
│   └── progress_bar.py
│
└── utils/                       ← Shared helpers, split by concern
    ├── __init__.py
    ├── file_helpers.py
    └── format_helpers.py
```

Not every project needs every folder. Propose only what's relevant.
But NEVER combine responsibilities into one file.

Each file does ONE thing. If you can't describe its purpose in one
sentence, split it. Before creating any new file or reorganising
the structure, tell me and get approval.

---

## Standard 2: Naming Conventions

### Files and Folders

- All lowercase with underscores: `gstin_validator.py`, `excel_reader.py`
- Folder names short and descriptive: `core/`, `gui/`, `config/`
- NO abbreviations unless universally known: `gst` is fine, `inv_val` is not
  (use `invoice_validator` instead)

### Functions and Variables

- All lowercase with underscores (snake_case):
  `validate_gstin()`, `tax_rate`, `input_file_path`
- Function names start with a VERB: `validate_`, `calculate_`, `read_`,
  `write_`, `export_`, `parse_`, `format_`, `check_`, `get_`, `set_`
- Boolean variables start with `is_` or `has_`:
  `is_valid`, `has_errors`, `is_interstate`
- NEVER use single-letter variables except `i`, `j` in short loops
- NEVER use vague names: `data`, `result`, `temp`, `info`, `stuff`
  Instead: `invoice_data`, `validation_result`, `temp_file_path`

### Classes

- PascalCase (capitalise each word): `InvoiceValidator`, `ExcelReader`,
  `MainWindow`, `GSTINResult`

### Constants

- ALL_UPPERCASE with underscores: `MAX_FILE_SIZE`, `GST_RATES`,
  `STATE_CODES`, `B2CL_THRESHOLD`
- ALL constants live in `config/constants.py` — NEVER hardcode values
  directly in logic files

---

## Standard 3: File-Level Documentation

Every Python file must start with this header block:

```python
"""
FILE: core/gstin_validator.py

PURPOSE: Validates GSTIN (Goods and Services Tax Identification Number)
         format, structure, and check digit.

CONTAINS:
- validate_gstin()     — Main validation, returns valid/invalid with details
- extract_components() — Breaks GSTIN into state, PAN, entity, check digit
- verify_checksum()    — Runs the government-specified check digit algorithm

DEPENDS ON:
- config/constants.py  → STATE_CODES, GSTIN_LENGTH, CHECKSUM_CHARS
- utils/format_helpers.py → clean_string()

USED BY:
- core/invoice_validator.py → calls validate_gstin() for each invoice row
- gui/error_viewer.py → displays validation error messages

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| DD-MM-YYYY | Created — initial validation rules        | Section 25 CGST Act compliance   |
| DD-MM-YYYY | Added old state code "00" handling         | Client data contained code "00"  |
| DD-MM-YYYY | Added composition taxpayer identification  | Needed for GSTR-4 projects       |
"""
```

---

## Standard 4: Function-Level Documentation

Every function MUST have a documentation block with ALL of these fields:

```python
def validate_gstin(gstin_string, strict_mode=False):
    """
    WHAT:
        Validates a GSTIN string against format rules and checksum.
        Returns a result dict with valid/invalid status, error details,
        and extracted components (state, PAN, entity type).

    WHY ADDED:
        Clients frequently submit invoices with malformed GSTINs —
        wrong length, invalid state codes, typos in PAN portion.
        This catches errors before they reach the GST Portal upload.
        Legal basis: Section 25 CGST Act, Rule 10 CGST Rules.

    CALLED BY:
        → core/invoice_validator.py → validate_invoice_row()
        → gui/main_window.py → on_single_gstin_check()

    CALLS:
        → config/constants.py → STATE_CODES, GSTIN_LENGTH
        → utils/format_helpers.py → clean_string()
        → self → verify_checksum() (private helper in this file)

    EDGE CASES HANDLED:
        - None or empty input → returns invalid, message: "GSTIN is empty"
        - Leading/trailing spaces → auto-trimmed before validation
        - Lowercase input → auto-converted to uppercase
        - Old state code "00" → flagged: "use state-specific code"
        - Valid format but wrong check digit → specific error message
        - Non-string input (int, float) → converted to string first

    ASSUMPTIONS:
        - State code list is current as of [date]. New codes require
          updating config/constants.py → STATE_CODES
        - Check digit algorithm follows government specification v1.0.
          If algorithm changes, this function needs manual update.
        *** ASSUMPTION: strict_mode defaults to False because most
            client data has minor formatting issues that should be
            auto-corrected, not rejected. Decided in Session 2. ***

    PARAMETERS:
        gstin_string (str): The GSTIN to validate. May contain spaces,
                            lowercase, or other formatting issues.
        strict_mode (bool): If True, rejects auto-correctable issues
                            instead of fixing them. Default: False.

    RETURNS:
        dict: {
            "valid": bool,          — True if GSTIN passes all checks
            "error": str or None,   — Error message if invalid
            "original": str,        — Input as received
            "cleaned": str,         — Input after auto-correction
            "state_code": str,      — First 2 characters
            "state_name": str,      — Full state name from lookup
            "pan": str,             — Characters 3-12
            "entity_code": str,     — Character 13
            "entity_type": str,     — "Regular"/"Composition"/"Government"/etc.
            "check_digit": str      — Last character
        }
    """
```

For small helper functions (under 5 lines), a shortened format is acceptable:

```python
def clean_string(text):
    """
    WHAT: Strips whitespace and converts to uppercase.
    CALLED BY: core/gstin_validator.py, core/invoice_validator.py
    RETURNS: Cleaned string, or empty string if input is None.
    """
    if text is None:
        return ""
    return str(text).strip().upper()
```

---

## Standard 5: In-Line Code Comments

Comments inside the code follow these rules:

```python
# GOOD — explains WHY (not obvious from the code)
# GST Portal rejects amounts with more than 2 decimal places
amount = round(amount, 2)

# GOOD — cites business rule
# B2CL threshold: Section 37 read with Table 5 of GSTR-1
# Invoices above ₹2,50,000 to unregistered persons → B2CL
if amount > B2CL_THRESHOLD and not recipient_gstin:
    table = "B2CL"

# GOOD — warns about non-obvious behaviour
# WARNING: openpyxl reads merged cells as None for all but the
# top-left cell. We handle this in _unmerge_and_fill() below.

# BAD — describes WHAT (obvious from the code, adds no value)
# increment counter by 1
counter += 1

# BAD — vague
# fix the thing
data = clean_data(data)
```

### Special Comment Tags (searchable markers):

```python
# TODO(Sidhant, DD-MM-YYYY): Add support for GSTR-2B reconciliation
#   Context: Client requested this, deferred to Phase 2.
#   See Parking Lot entry #3 in Rebuild Brief.

# FIXME(DD-MM-YYYY): This breaks if Excel file has merged headers.
#   Workaround: manually unmerge before processing.
#   Proper fix: implement auto-unmerge in readers/excel_reader.py

# HACK(DD-MM-YYYY): Using string comparison for dates because
#   openpyxl returns mixed types (datetime vs string) depending
#   on Excel cell formatting. Proper fix: normalise in reader.

# ASSUMPTION(DD-MM-YYYY): Assuming all input files use UTF-8.
#   If client sends ANSI-encoded files, this will break.
#   Decided in Session 3 — revisit if encoding errors appear.

# LEGAL(Section 16(2) CGST Act): ITC available only if supplier
#   has filed GSTR-1 and tax is reflected in recipient's GSTR-2B.
```

---

## Standard 6: No Magic Numbers

NEVER hardcode values directly in logic. Always use named constants
from `config/constants.py`:

```python
# BAD — magic number, no one knows what 250000 means
if amount > 250000:
    table = "B2CL"

# GOOD — named constant with clear meaning
# In config/constants.py:
B2CL_THRESHOLD = 250000  # ₹2,50,000 — Section 37 / GSTR-1 Table 5

# In core/return_classifier.py:
from config.constants import B2CL_THRESHOLD
if amount > B2CL_THRESHOLD:
    table = "B2CL"
```

This applies to ALL values: tax rates, thresholds, file size limits,
retry counts, timeout values, column names, sheet names, date formats,
state codes — everything. If a value might ever change or needs
explanation, it belongs in config/.

---

## Standard 7: Consistent Error Handling

Every function that can fail follows this pattern:

```python
def read_excel_file(file_path):
    """
    [Standard docstring as per Standard 4]
    """
    # --- Input validation ---
    if not file_path:
        return {"success": False, "error": "No file path provided", "data": None}

    if not os.path.exists(file_path):
        return {"success": False, "error": f"File not found: {file_path}", "data": None}

    if not file_path.endswith(('.xlsx', '.xls')):
        return {"success": False, "error": f"Not an Excel file: {file_path}", "data": None}

    # --- Main logic ---
    try:
        workbook = openpyxl.load_workbook(file_path)
        # ... processing ...
        return {"success": True, "error": None, "data": processed_data}

    except PermissionError:
        # File is open in another program
        return {"success": False,
                "error": f"Cannot read file — it may be open in Excel: {file_path}",
                "data": None}

    except Exception as e:
        # Unexpected error — log details, show friendly message
        return {"success": False,
                "error": f"Unexpected error reading {file_path}: {str(e)}",
                "data": None}
```

Rules:

- NEVER let exceptions crash the application
- ALWAYS return a consistent result format: `{"success": bool, "error": str, "data": ...}`
- ALWAYS validate inputs at the top of the function, before any processing
- ALWAYS catch specific exceptions first (PermissionError, FileNotFoundError),
  then catch general Exception as a safety net
- Error messages must be in PLAIN ENGLISH — the user is not a programmer.
  "File not found" not "FileNotFoundError". "File may be open in Excel" not
  "PermissionError: [Errno 13]"

---

## Standard 8: Type Hints

Every function signature must include type hints. These make the code
self-documenting — even a non-programmer can see what goes in and what
comes out:

```python
# Without type hints — unclear what this expects or returns
def validate_gstin(gstin_string, strict_mode):
    ...

# With type hints — immediately clear
def validate_gstin(gstin_string: str, strict_mode: bool = False) -> dict:
    ...

# More examples:
def read_excel_file(file_path: str) -> dict:
def calculate_tax(amount: float, rate: float, is_interstate: bool) -> dict:
def format_date(date_value: str, output_format: str = "DD-MM-YYYY") -> str:
def get_state_name(state_code: str) -> str | None:
```

---

## Standard 9: Import Organisation

Every file organises imports in this order, with blank lines between groups:

```python
# Group 1: Python standard library
import os
import sys
from datetime import datetime
from pathlib import Path

# Group 2: Third-party libraries
import openpyxl
from PyQt5.QtWidgets import QMainWindow, QPushButton

# Group 3: This project's modules
from config.constants import STATE_CODES, B2CL_THRESHOLD
from config.settings import OUTPUT_PATH
from core.gstin_validator import validate_gstin
from utils.format_helpers import clean_string
```

NEVER use `from module import *` — always import specific names.
NEVER use relative imports (`.module`) — always use full paths from
the project root.

---

## Standard 10: Git Commit Messages

Every git commit follows this format:

```
[TYPE] Short description (what changed)

WHY: Why this change was made
AFFECTS: Which files/features are impacted
EDGE CASES: Any new edge cases handled (if applicable)
DECIDED: Reference to decision (if applicable)
```

Types:

- `[FEATURE]` — New feature added
- `[FIX]` — Bug fix
- `[REFACTOR]` — Code restructure without behaviour change
- `[CONFIG]` — Settings, constants, dependency changes
- `[DOCS]` — Documentation only
- `[STYLE]` — Formatting, naming (no logic change)

Example:

```
[FEATURE] Add GSTIN checksum validation

WHY: Format-only validation missed GSTINs with valid structure but wrong
     check digit. 3 client files had this issue in January batch.
AFFECTS: core/gstin_validator.py (new verify_checksum function),
         config/constants.py (added CHECKSUM_CHARS mapping)
EDGE CASES: Handles old-format GSTINs from pre-2019 registrations
DECIDED: Session 2, Message 14 — chose algorithmic check over API lookup
```

---

## Standard 11: Assumptions Must Be Visible

If Claude Code makes ANY assumption — about data format, user behaviour,
business rules, file structure, edge case handling, or anything else —
it MUST appear in THREE places:

1. **In the code** — in the ASSUMPTIONS field of the function docstring,
   or as an `# ASSUMPTION(date):` inline comment
2. **In the DECISIONS_LOG** — in that message's log entry
3. **Verbally in the chat** — "I'm assuming X because Y. Is that correct?"

No silent assumptions. Ever. If the Brief doesn't cover something and
Claude Code has to make a judgment call, that judgment must be visible
and traceable. Silent assumptions become invisible bugs.

---
