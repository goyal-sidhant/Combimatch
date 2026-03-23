"""
FILE: core/target_parser.py

PURPOSE: Parses and validates the target sum and buffer (tolerance) values
         entered by the user. Strips thousand-separator commas so users
         can paste values like "1,00,000" from Excel.

CONTAINS:
- parse_target() — Parses the target sum field
- parse_buffer() — Parses the buffer (tolerance) field

DEPENDS ON:
- config/constants.py → THOUSAND_SEPARATOR

USED BY:
- core/parameter_validator.py → calls parse_target() and parse_buffer()
                                 as part of full parameter validation
- gui/input_panel.py → may call directly for real-time validation hints

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — target/buffer parsing            | Sub-phase 1A foundation          |
"""

# Group 3: This project's modules
from config.constants import THOUSAND_SEPARATOR


def parse_target(text: str) -> dict:
    """
    WHAT:
        Parses the target sum field. Strips thousand-separator commas
        and whitespace before converting to float.

    WHY ADDED:
        Users copy-paste target amounts from Excel in accounting format
        (e.g., "1,00,000" or "25,000.50"). The original tool rejected
        these, forcing manual re-typing. This parser accepts them.

    CALLED BY:
        → core/parameter_validator.py → validate_search_parameters()

    CALLS:
        → str.replace(), float()

    EDGE CASES HANDLED:
        - "1,00,000" → 100000.0 (Indian accounting format)
        - "25,000.50" → 25000.50
        - "  500  " → 500.0 (whitespace stripped)
        - "" or None → error: "Target sum is required"
        - "abc" → error: "Target sum is not a valid number"
        - Negative targets → allowed (valid for some use cases)

    ASSUMPTIONS:
        - Target of zero is technically valid (would find items summing
          to zero). Not blocked because edge cases exist in real data.

    PARAMETERS:
        text (str): Raw text from the target input field.

    RETURNS:
        dict: {
            "success": bool,
            "error": str or None,
            "data": float or None — the parsed target value
        }
    """
    if not text or not text.strip():
        return {
            "success": False,
            "error": "Target sum is required. Enter the amount you want to match.",
            "data": None
        }

    cleaned = text.strip().replace(THOUSAND_SEPARATOR, "").replace(" ", "")

    try:
        value = float(cleaned)
        return {"success": True, "error": None, "data": value}
    except ValueError:
        return {
            "success": False,
            "error": f"Target sum '{text.strip()}' is not a valid number.",
            "data": None
        }


def parse_buffer(text: str) -> dict:
    """
    WHAT:
        Parses the buffer (tolerance) field. Strips thousand-separator
        commas, converts to float, and forces the value to be non-negative
        (absolute value applied silently).

    WHY ADDED:
        Buffer allows approximate matches within target ± buffer.
        Like target, users may paste comma-formatted values.

    CALLED BY:
        → core/parameter_validator.py → validate_search_parameters()

    CALLS:
        → str.replace(), float(), abs()

    EDGE CASES HANDLED:
        - "" or None → defaults to 0.0 (exact matches only)
        - "1,000" → 1000.0
        - "-50" → 50.0 (negative silently converted to absolute)
        - "abc" → error message

    ASSUMPTIONS:
        - Buffer is always absolute (not percentage-based). This is a
          deliberate design decision confirmed in the Rebuild Brief.
        - Empty buffer defaults to 0, not an error — most users want
          exact matches and leave the buffer field empty.

    PARAMETERS:
        text (str): Raw text from the buffer input field.

    RETURNS:
        dict: {
            "success": bool,
            "error": str or None,
            "data": float or None — the parsed buffer value (always >= 0)
        }
    """
    if not text or not text.strip():
        # Empty buffer = exact match mode (buffer = 0)
        return {"success": True, "error": None, "data": 0.0}

    cleaned = text.strip().replace(THOUSAND_SEPARATOR, "").replace(" ", "")

    try:
        value = abs(float(cleaned))
        return {"success": True, "error": None, "data": value}
    except ValueError:
        return {
            "success": False,
            "error": f"Buffer '{text.strip()}' is not a valid number.",
            "data": None
        }
