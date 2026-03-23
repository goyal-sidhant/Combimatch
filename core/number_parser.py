"""
FILE: core/number_parser.py

PURPOSE: Parses raw text input into NumberItem objects. Handles two input modes:
         line-separated (one number per line) and semicolon-separated (numbers
         separated by semicolons). Both modes strip thousand-separator commas
         before parsing.

CONTAINS:
- parse_numbers_line_separated()      — Parses line-separated text input
- parse_numbers_semicolon_separated() — Parses semicolon-separated text input
- _clean_and_parse_value()            — Shared helper: strips commas, parses float

DEPENDS ON:
- config/constants.py → THOUSAND_SEPARATOR, ROUNDING_PRECISION
- models/number_item.py → NumberItem

USED BY:
- gui/input_panel.py → calls parse_numbers_line_separated() or
                        parse_numbers_semicolon_separated() based on
                        the selected input mode

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — number parsing for both modes   | Sub-phase 1A foundation          |
| 22-03-2026 | Skip zero values silently                  | Zeros don't contribute to sums   |
"""

# Group 1: Python standard library
from typing import List, Tuple

# Group 3: This project's modules
from config.constants import THOUSAND_SEPARATOR
from models.number_item import NumberItem


def _clean_and_parse_value(raw_text: str) -> float:
    """
    WHAT:
        Strips thousand-separator commas and whitespace from a raw text
        value, then converts to float.

    WHY ADDED:
        Both line-separated and semicolon-separated modes need the same
        cleaning logic. Office culture uses accounting format in Excel,
        so copy-pasting brings commas (e.g., "1,00,000.50"). These must
        be stripped before float conversion.

    CALLED BY:
        → parse_numbers_line_separated()
        → parse_numbers_semicolon_separated()

    CALLS:
        → str.replace(), str.strip(), float()

    EDGE CASES HANDLED:
        - "1,00,000" (Indian format) → "100000" → 100000.0
        - "1,000.50" (international format) → "1000.50" → 1000.50
        - "  -500  " → "-500" → -500.0 (credit notes)
        - Spaces inside the number are stripped

    ASSUMPTIONS:
        - Commas are ALWAYS thousand separators, never decimal separators.
          Indian accounting uses period (.) for decimals.

    PARAMETERS:
        raw_text (str): The raw text to parse. May contain commas, spaces.

    RETURNS:
        float: The parsed numeric value.

    RAISES:
        ValueError: If the text cannot be converted to a float after cleaning.
    """
    # Strip whitespace, then remove all thousand-separator commas and spaces
    cleaned = raw_text.strip()
    cleaned = cleaned.replace(THOUSAND_SEPARATOR, "")
    cleaned = cleaned.replace(" ", "")

    if not cleaned:
        raise ValueError("Empty value")

    return float(cleaned)


def parse_numbers_line_separated(text: str) -> dict:
    """
    WHAT:
        Parses a block of text where each line contains one number.
        Returns a dict with the list of successfully parsed NumberItems
        and a list of error messages for lines that failed to parse.

    WHY ADDED:
        Primary input mode — users paste columns of numbers from Excel
        or type them manually, one per line.

    CALLED BY:
        → gui/input_panel.py → when user clicks "Load Numbers" in
          "Line Separated" mode

    CALLS:
        → _clean_and_parse_value() → for each non-empty line
        → NumberItem() → to create the data object

    EDGE CASES HANDLED:
        - Empty lines → skipped silently (common when pasting from Excel)
        - Lines with only whitespace → skipped silently
        - Non-numeric lines → error message with line number
        - None input → returns empty results with error
        - Negative numbers → parsed correctly (credit notes in GST work)
        - "1,00,000" → commas stripped, parsed as 100000
        - Zero values → skipped silently (don't contribute to sums)

    ASSUMPTIONS:
        - Each line contains exactly one number (or is empty/invalid).
        - Commas within a line are always thousand separators.
        - Index is sequential among valid items (not the original line number).

    PARAMETERS:
        text (str): The raw text block to parse.

    RETURNS:
        dict: {
            "success": bool,          — True if at least one number was parsed
            "error": str or None,     — Summary error if no input provided
            "data": {
                "items": List[NumberItem],   — Successfully parsed numbers
                "errors": List[str]          — Error messages for failed lines
            }
        }
    """
    if not text or not text.strip():
        return {
            "success": False,
            "error": "No numbers entered. Please paste or type numbers, one per line.",
            "data": {"items": [], "errors": []}
        }

    items: List[NumberItem] = []
    errors: List[str] = []
    lines = text.split("\n")

    for line_number, line in enumerate(lines, start=1):
        stripped = line.strip()

        # Skip empty lines silently
        if not stripped:
            continue

        try:
            value = _clean_and_parse_value(stripped)
            # Skip zeros — they don't contribute to sums
            if value == 0.0:
                continue
            item = NumberItem(
                value=value,
                index=len(items),
                original_text=stripped,
            )
            items.append(item)
        except ValueError:
            errors.append(
                f"Line {line_number}: '{stripped}' is not a valid number"
            )

    if not items and not errors:
        return {
            "success": False,
            "error": "No numbers entered. Please paste or type numbers, one per line.",
            "data": {"items": [], "errors": []}
        }

    if not items and errors:
        return {
            "success": False,
            "error": "No valid numbers found. Check the errors below.",
            "data": {"items": [], "errors": errors}
        }

    return {
        "success": True,
        "error": None,
        "data": {"items": items, "errors": errors}
    }


def parse_numbers_semicolon_separated(text: str) -> dict:
    """
    WHAT:
        Parses a block of text where numbers are separated by semicolons (;).
        Returns a dict with successfully parsed NumberItems and error messages.

    WHY ADDED:
        Alternative input mode for compact data entry. Uses semicolons
        instead of commas as the delimiter to eliminate the ambiguity
        between "1,000" (one thousand) and "1" + "000" (two values).
        This was a deliberate change from the original tool which used
        commas as delimiters.

    CALLED BY:
        → gui/input_panel.py → when user clicks "Load Numbers" in
          "Semicolon Separated" mode

    CALLS:
        → _clean_and_parse_value() → for each non-empty part
        → NumberItem() → to create the data object

    EDGE CASES HANDLED:
        - Empty parts between semicolons → skipped (e.g., "100;;200")
        - Parts with only whitespace → skipped
        - Non-numeric parts → error with position number
        - None input → returns empty results with error
        - "1,000;2,000" → commas are thousand separators, not delimiters
        - Negative numbers → parsed correctly
        - Zero values → skipped silently (don't contribute to sums)

    ASSUMPTIONS:
        - Semicolons are the ONLY delimiter. Commas within values are
          always thousand separators.
        - Newlines within the text are treated as part of the same input
          (not as separate entries). The text is split on semicolons only.

    PARAMETERS:
        text (str): The raw text to parse.

    RETURNS:
        dict: Same format as parse_numbers_line_separated().
    """
    if not text or not text.strip():
        return {
            "success": False,
            "error": "No numbers entered. Please enter numbers separated by semicolons (;).",
            "data": {"items": [], "errors": []}
        }

    items: List[NumberItem] = []
    errors: List[str] = []

    # Replace newlines with semicolons so multi-line semicolon input works
    normalized = text.replace("\n", ";")
    parts = normalized.split(";")

    for part_number, part in enumerate(parts, start=1):
        stripped = part.strip()

        # Skip empty parts silently
        if not stripped:
            continue

        try:
            value = _clean_and_parse_value(stripped)
            # Skip zeros — they don't contribute to sums
            if value == 0.0:
                continue
            item = NumberItem(
                value=value,
                index=len(items),
                original_text=stripped,
            )
            items.append(item)
        except ValueError:
            errors.append(
                f"Value {part_number}: '{stripped}' is not a valid number"
            )

    if not items and not errors:
        return {
            "success": False,
            "error": "No numbers entered. Please enter numbers separated by semicolons (;).",
            "data": {"items": [], "errors": []}
        }

    if not items and errors:
        return {
            "success": False,
            "error": "No valid numbers found. Check the errors below.",
            "data": {"items": [], "errors": errors}
        }

    return {
        "success": True,
        "error": None,
        "data": {"items": items, "errors": errors}
    }
