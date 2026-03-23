"""
FILE: utils/format_helpers.py

PURPOSE: Shared formatting and cleaning helper functions used across
         the application. Handles string cleaning, number display
         formatting (Indian accounting format), and difference display.

CONTAINS:
- clean_string()             — Strips whitespace and converts to uppercase
- format_number_display()    — Formats a number in Indian accounting format
- format_difference()        — Formats a difference value with +/- sign

DEPENDS ON:
- config/constants.py → EXACT_MATCH_THRESHOLD

USED BY:
- core/number_parser.py → (could use clean_string if needed)
- gui/source_panel.py → format_number_display() for source list items
- gui/combo_info_panel.py → format_difference() for combo detail display
- gui/results_panel.py → format_number_display(), format_difference()
- gui/summary_tab.py → format_number_display(), format_difference()
- writers/report_writer.py → format_number_display() for report values

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — formatting helpers              | Sub-phase 1A foundation          |
"""

# Group 1: Python standard library
from typing import Optional

# Group 3: This project's modules
from config.constants import EXACT_MATCH_THRESHOLD


def clean_string(text: Optional[str]) -> str:
    """
    WHAT: Strips whitespace and converts to uppercase.
    CALLED BY: Various modules for cleaning user input.
    RETURNS: Cleaned string, or empty string if input is None.
    """
    if text is None:
        return ""
    return str(text).strip().upper()


def format_number_indian(value: float) -> str:
    """
    WHAT:
        Formats a number in Indian accounting format with commas.
        e.g., 100000.50 → "1,00,000.50"
             -25000 → "-25,000.00"

    WHY ADDED:
        Indian accounting convention groups digits as:
        last 3 digits, then groups of 2 (e.g., 1,00,00,000).
        This matches the format used in the office and makes
        numbers immediately readable for CA practitioners.

    CALLED BY:
        → gui/source_panel.py → displaying loaded numbers
        → gui/combo_info_panel.py → showing combination sums
        → gui/results_panel.py → displaying combination details
        → gui/summary_tab.py → finalization card values
        → writers/report_writer.py → report number formatting

    CALLS:
        → Built-in string operations only.

    EDGE CASES HANDLED:
        - Zero → "0.00"
        - Negative numbers → "-25,000.00" (hyphen prefix)
        - Very small decimals → "0.50"
        - Large numbers → "1,00,00,00,000.00"
        - Already rounded to 2dp by NumberItem, but this function
          also ensures 2dp display.

    ASSUMPTIONS:
        - Always displays exactly 2 decimal places (standard for
          financial/accounting display).
        - Uses Indian grouping: last 3 digits, then groups of 2.

    PARAMETERS:
        value (float): The number to format.

    RETURNS:
        str: The formatted number string with Indian commas and 2 decimal places.
    """
    # Handle negative numbers
    is_negative = value < 0
    abs_value = abs(value)

    # Split into integer and decimal parts
    integer_part = int(abs_value)
    decimal_part = round(abs_value - integer_part, 2)
    decimal_str = f"{decimal_part:.2f}"[1:]  # ".50" — includes the dot

    # Format integer part with Indian grouping
    int_str = str(integer_part)

    if len(int_str) <= 3:
        # No grouping needed for numbers up to 999
        formatted_int = int_str
    else:
        # Last 3 digits
        last_three = int_str[-3:]
        remaining = int_str[:-3]

        # Group remaining digits in pairs from right
        groups = []
        while remaining:
            groups.append(remaining[-2:])
            remaining = remaining[:-2]

        groups.reverse()
        formatted_int = ",".join(groups) + "," + last_three

    result = formatted_int + decimal_str

    if is_negative:
        result = "-" + result

    return result


def format_difference(difference: float) -> str:
    """
    WHAT:
        Formats a combination's difference from the target with a +/- sign.
        Exact matches (difference < threshold) show as "Exact".

    WHY ADDED:
        Differences need clear visual formatting — positive means over
        target, negative means under. Exact matches are called out
        explicitly since they're the ideal result.

    CALLED BY:
        → gui/combo_info_panel.py → difference display
        → gui/results_panel.py → approximate match display
        → gui/summary_tab.py → finalization card display

    EDGE CASES HANDLED:
        - Exact match (< 0.001) → "Exact"
        - Positive difference → "+100.50"
        - Negative difference → "-100.50"
        - Zero difference → "Exact"

    PARAMETERS:
        difference (float): The combo's sum minus the target.

    RETURNS:
        str: Formatted difference string.
    """
    if abs(difference) < EXACT_MATCH_THRESHOLD:
        return "Exact"

    if difference > 0:
        return f"+{format_number_indian(difference)}"
    else:
        return format_number_indian(difference)
