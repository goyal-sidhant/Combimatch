"""
FILE: config/constants.py

PURPOSE: All fixed values used across the application. Every magic number,
         threshold, default, color, marker, and delimiter lives here.
         No other file should hardcode values that belong in this file.

CONTAINS:
- ROUNDING_PRECISION         — Decimal places for rounding loaded numbers
- EXACT_MATCH_THRESHOLD      — Difference below which a match is "exact"
- DEFAULT_BUFFER             — Default buffer (tolerance) value
- DEFAULT_MIN_SIZE           — Default minimum combination size
- DEFAULT_MAX_SIZE           — Default maximum combination size
- DEFAULT_MAX_RESULTS        — Default maximum number of results to show
- PROGRESS_CHECK_INTERVAL    — How often the solver checks stop flag / reports progress
- HIGHLIGHT_COLORS           — 20 predefined soft colors for finalization
- UNMATCHED_COLOR            — Grey color for unmatched numbers
- FINALIZED_MARKER           — "✓" prefix for finalized items in source list
- SELECTED_MARKER            — "▶" prefix for selected items in source list
- SEED_MARKER                — "[SEED]" prefix for seed items
- SEMICOLON_DELIMITER        — ";" used as delimiter in semicolon-separated mode
- THOUSAND_SEPARATOR         — "," stripped from numbers during parsing
- MAX_RESULTS_UPPER_LIMIT    — Upper bound for the max results spinbox
- SEARCH_ORDER_SMALLEST      — Constant for smallest-first search order
- SEARCH_ORDER_LARGEST       — Constant for largest-first search order
- WINDOW_MIN_WIDTH           — Minimum window width in pixels
- WINDOW_MIN_HEIGHT          — Minimum window height in pixels
- SPLITTER_SIZES             — Initial panel widths [left, middle, right]
- COLOR_COUNT                — Number of highlight colors (20)
- SESSION_AUTOSAVE_INTERVAL  — Milliseconds between auto-saves

DEPENDS ON:
- Nothing — this is a leaf module with no imports from this project.

USED BY:
- Nearly every module in the project imports constants from here.

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — all constants for rebuild       | Project skeleton setup (Phase 1A)|
"""


# ---------------------------------------------------------------------------
# Rounding and Matching
# ---------------------------------------------------------------------------

# All loaded numbers are rounded to this many decimal places.
# Real-world invoice data is always 2 decimal places.
ROUNDING_PRECISION = 2

# Any combination whose absolute difference from the target is below this
# value is considered an "exact" match. Handles floating-point rounding
# so paisa-level differences don't create false "approximate" results.
EXACT_MATCH_THRESHOLD = 0.001


# ---------------------------------------------------------------------------
# Search Parameter Defaults
# ---------------------------------------------------------------------------

# Default buffer (tolerance): 0 means exact matches only by default.
DEFAULT_BUFFER = 0.0

# Default minimum combination size. A single invoice can be a valid match.
DEFAULT_MIN_SIZE = 1

# Default maximum combination size. Reasonable starting point; user adjusts
# based on their data. No longer capped at 100 — upper limit is dynamically
# set to the number of loaded items.
DEFAULT_MAX_SIZE = 10

# Default maximum number of results to display. Changed from 100 to 25
# based on user feedback — 100 was too many for practical review.
DEFAULT_MAX_RESULTS = 25

# Upper bound for the max results spinbox. User can type up to this value.
MAX_RESULTS_UPPER_LIMIT = 10000


# ---------------------------------------------------------------------------
# Solver Behaviour
# ---------------------------------------------------------------------------

# How often the solver checks the stop flag and reports progress.
# Original was every 1000 iterations — too slow for responsive stop.
# Increased to 1000 from 100 to reduce callback overhead on large searches.
# Progress signal throttling in SolverThread ensures UI stays responsive.
PROGRESS_CHECK_INTERVAL = 1000

# Minimum seconds between progress signal emissions to the UI.
# Prevents signal flooding on searches with millions of iterations.
# 4 updates per second is smooth enough for the progress label.
PROGRESS_SIGNAL_MIN_INTERVAL = 0.25

# Search order options
SEARCH_ORDER_SMALLEST = "smallest_first"
SEARCH_ORDER_LARGEST = "largest_first"

# If the estimated total combinations exceeds this limit, show a warning
# dialog before starting the search. Prevents the app from hanging on
# impossibly large search spaces like C(100, 50) ≈ 10^29.
SEARCH_SPACE_WARNING_LIMIT = 10_000_000


# ---------------------------------------------------------------------------
# Input Parsing
# ---------------------------------------------------------------------------

# Delimiter for semicolon-separated input mode.
# Changed from comma (,) to semicolon (;) to eliminate the ambiguity
# between "1,000" (one thousand) and "1" + "000" (two separate values).
SEMICOLON_DELIMITER = ";"

# Character treated as thousand separator in ALL input modes.
# Stripped before parsing so "1,00,000" becomes "100000".
THOUSAND_SEPARATOR = ","


# ---------------------------------------------------------------------------
# UI Markers (text prefixes in source list)
# ---------------------------------------------------------------------------

# Prefix added to finalized items in the source numbers list.
FINALIZED_MARKER = "✓"

# Prefix added to items that are part of the currently selected combination.
SELECTED_MARKER = "▶"

# Prefix added to items pinned as seed numbers.
SEED_MARKER = "[SEED]"


# ---------------------------------------------------------------------------
# Window Layout
# ---------------------------------------------------------------------------

# Minimum window dimensions in pixels.
WINDOW_MIN_WIDTH = 1000
WINDOW_MIN_HEIGHT = 700

# Initial widths for the three-panel horizontal splitter in the Find tab.
# [left panel (input), middle panel (results), right panel (source/info)]
SPLITTER_SIZES = [280, 380, 320]


# ---------------------------------------------------------------------------
# Highlight Colors — 20 predefined soft colors for finalization
# ---------------------------------------------------------------------------

# Each entry is ((R, G, B), "Human-Readable Name").
# These cycle: finalization #1 gets color #0, #21 wraps back to #0.
# Colors chosen to be soft/pastel so text remains readable on top.
HIGHLIGHT_COLORS = [
    ((173, 216, 230), "Light Blue"),
    ((144, 238, 144), "Light Green"),
    ((255, 182, 193), "Light Pink"),
    ((255, 255, 153), "Light Yellow"),
    ((216, 191, 216), "Thistle"),
    ((255, 218, 185), "Peach Puff"),
    ((176, 224, 230), "Powder Blue"),
    ((152, 251, 152), "Pale Green"),
    ((255, 228, 196), "Bisque"),
    ((221, 160, 221), "Plum"),
    ((175, 238, 238), "Pale Turquoise"),
    ((245, 222, 179), "Wheat"),
    ((230, 230, 250), "Lavender"),
    ((188, 143, 143), "Rosy Brown"),
    ((135, 206, 235), "Sky Blue"),
    ((219, 112, 147), "Pale Violet Red"),
    ((240, 230, 140), "Khaki"),
    ((127, 255, 212), "Aquamarine"),
    ((244, 164, 96),  "Sandy Brown"),
    ((186, 85, 211),  "Medium Orchid"),
]

# Total number of highlight colors. Wraps around after this many finalizations.
COLOR_COUNT = len(HIGHLIGHT_COLORS)

# Color for unmatched numbers (light grey). Visually distinct from all
# 20 highlight colors. Match colors always override this when a previously
# unmatched item gets finalized.
UNMATCHED_COLOR = ((220, 220, 220), "Light Grey")


# ---------------------------------------------------------------------------
# Session Persistence
# ---------------------------------------------------------------------------

# Milliseconds between auto-saves. 60 seconds = 60000 ms.
# Session also saves immediately after every finalization and undo action.
SESSION_AUTOSAVE_INTERVAL = 60000


# ---------------------------------------------------------------------------
# Excel Monitor
# ---------------------------------------------------------------------------

# Milliseconds between Excel connection checks. 5 seconds = 5000 ms.
EXCEL_MONITOR_INTERVAL = 5000


# ---------------------------------------------------------------------------
# Batch UI Updates
# ---------------------------------------------------------------------------

# Maximum number of combinations per UI update batch.
# Smaller batches keep the UI responsive during search.
# With incremental insertion (not full rebuild), 10 is plenty.
BATCH_SIZE = 10

# Maximum seconds between UI update batches.
# Ensures results appear smoothly even if fewer than BATCH_SIZE are found.
BATCH_INTERVAL = 0.1
