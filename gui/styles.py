"""
FILE: gui/styles.py

PURPOSE: Defines the QSS (Qt Style Sheet) for the entire application.
         All colors, fonts, spacing, and visual styling are centralized
         here. No other file should hardcode visual properties.

CONTAINS:
- APP_STYLESHEET          — Complete QSS string applied to QApplication
- Color constants          — Named color values used in the stylesheet
- get_stylesheet()         — Returns the stylesheet (for dynamic themes later)

DEPENDS ON:
- Nothing — this is a leaf module with no project imports.

USED BY:
- main.py → applies get_stylesheet() to the QApplication
- gui/main_window.py → may reference color constants for programmatic styling

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — application stylesheet          | Sub-phase 1C app shell           |
| 22-03-2026 | Scaled fonts, soft cream background        | UI fixes before Phase 3          |
"""


# ---------------------------------------------------------------------------
# Named Color Constants (used in the stylesheet and available for import)
# ---------------------------------------------------------------------------

# Primary palette — soft warm tones, professional and easy on the eyes
COLOR_BACKGROUND = "#F7F5F0"         # Main window background (warm cream)
COLOR_PANEL_BG = "#FDFCF9"           # Panel/widget background (off-white cream)
COLOR_BORDER = "#D1D5DB"             # Borders and separators
COLOR_BORDER_FOCUS = "#6B7FD7"       # Border when focused/active

# Text colors
COLOR_TEXT_PRIMARY = "#1F2937"        # Main text
COLOR_TEXT_SECONDARY = "#6B7280"      # Secondary/hint text
COLOR_TEXT_DISABLED = "#9CA3AF"       # Disabled controls

# Accent colors
COLOR_ACCENT = "#6B7FD7"             # Primary accent (buttons, active tab)
COLOR_ACCENT_HOVER = "#5A6DC4"       # Accent on hover
COLOR_ACCENT_PRESSED = "#4A5CB0"     # Accent when pressed
COLOR_ACCENT_LIGHT = "#E8EBFA"       # Light accent for selections

# Status colors
COLOR_SUCCESS = "#10B981"             # Green for success/exact match
COLOR_WARNING = "#F59E0B"             # Amber for warnings
COLOR_ERROR = "#EF4444"               # Red for errors/no solution

# Tab and header colors
COLOR_TAB_ACTIVE_BG = "#FDFCF9"       # Active tab background (warm off-white, no pure white)
COLOR_TAB_INACTIVE_BG = "#E5E7EB"     # Inactive tab background
COLOR_HEADER_BG = "#EEF0F6"          # Section headers, group labels

# Source list highlight (orange for selected combo items)
COLOR_SELECTED_HIGHLIGHT = "#FFF3E0"  # Light orange background
COLOR_SELECTED_BORDER = "#FF9800"     # Orange border

# Progress bar
COLOR_PROGRESS_BG = "#E5E7EB"        # Progress bar track
COLOR_PROGRESS_FILL = "#6B7FD7"      # Progress bar fill


# ---------------------------------------------------------------------------
# Font Scaling
# ---------------------------------------------------------------------------

# Default font scale factor — updated by compute_font_scale() after
# QApplication is created and screen geometry is available.
_font_scale: float = 1.0


def compute_font_scale() -> float:
    """
    WHAT:
        Computes a font scale factor based on the primary screen's height.
        Reference: 1.0x at 1080px height. Smaller screens get smaller text,
        larger screens get larger text. Clamped between 0.80 and 1.30.

    WHY ADDED:
        Fixed pixel font sizes look too small on smaller screens and
        waste space on larger ones. Scaling to screen height gives a
        comfortable reading size on every display.

    CALLED BY:
        → main.py → after QApplication is created, before get_stylesheet()

    RETURNS:
        float: Scale factor (e.g. 0.90, 1.0, 1.15).
    """
    global _font_scale

    try:
        from PyQt5.QtWidgets import QApplication
        app = QApplication.instance()
        if app is not None:
            screen = app.primaryScreen()
            if screen is not None:
                height = screen.availableGeometry().height()
                # Reference: 1.0x at 1080px height
                raw_scale = height / 1080.0
                _font_scale = max(0.80, min(1.30, raw_scale))
    except Exception:
        _font_scale = 1.0

    return _font_scale


def scaled_size(base_px: int) -> int:
    """
    WHAT: Scales a base pixel size by the current font scale factor.
    CALLED BY: _build_stylesheet() — for all font-size values.

    PARAMETERS:
        base_px (int): Base size in pixels (designed for 1080p).

    RETURNS:
        int: Scaled size in pixels, minimum 9px.
    """
    return max(9, int(base_px * _font_scale))


# ---------------------------------------------------------------------------
# The Stylesheet
# ---------------------------------------------------------------------------

def _build_stylesheet() -> str:
    """
    WHAT:
        Builds the complete QSS stylesheet with scaled font sizes.
        Uses the current _font_scale value set by compute_font_scale().

    CALLED BY:
        → get_stylesheet()

    RETURNS:
        str: Complete QSS stylesheet string.
    """
    return f"""
/* === Global === */
QWidget {{
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: {scaled_size(15)}px;
    color: {COLOR_TEXT_PRIMARY};
    background-color: {COLOR_BACKGROUND};
}}

/* === Main Window === */
QMainWindow {{
    background-color: {COLOR_BACKGROUND};
}}

/* === Tab Widget === */
QTabWidget::pane {{
    border: 1px solid {COLOR_BORDER};
    background-color: {COLOR_PANEL_BG};
    border-radius: 4px;
    top: -1px;
}}

QTabBar::tab {{
    background-color: {COLOR_TAB_INACTIVE_BG};
    color: {COLOR_TEXT_SECONDARY};
    border: 1px solid {COLOR_BORDER};
    border-bottom: none;
    padding: 10px 24px;
    margin-right: 2px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    min-width: 90px;
    font-size: {scaled_size(15)}px;
}}

QTabBar::tab:selected {{
    background-color: {COLOR_TAB_ACTIVE_BG};
    color: {COLOR_ACCENT};
    font-weight: bold;
    border-bottom: 2px solid {COLOR_ACCENT};
}}

QTabBar::tab:hover:!selected {{
    background-color: {COLOR_ACCENT_LIGHT};
    color: {COLOR_TEXT_PRIMARY};
}}

/* === Buttons === */
QPushButton {{
    background-color: {COLOR_ACCENT};
    color: white;
    border: none;
    padding: 10px 20px;
    border-radius: 5px;
    font-weight: bold;
    min-height: 24px;
    font-size: {scaled_size(15)}px;
}}

QPushButton:hover {{
    background-color: {COLOR_ACCENT_HOVER};
}}

QPushButton:pressed {{
    background-color: {COLOR_ACCENT_PRESSED};
}}

QPushButton:disabled {{
    background-color: {COLOR_BORDER};
    color: {COLOR_TEXT_DISABLED};
}}

/* Secondary buttons (flat style) */
QPushButton[flat="true"] {{
    background-color: transparent;
    color: {COLOR_ACCENT};
    border: 1px solid {COLOR_ACCENT};
}}

QPushButton[flat="true"]:hover {{
    background-color: {COLOR_ACCENT_LIGHT};
}}

/* === Text Inputs === */
QLineEdit, QTextEdit, QPlainTextEdit {{
    background-color: #FFFFFF;
    border: 1px solid {COLOR_BORDER};
    border-radius: 4px;
    padding: 7px 10px;
    color: {COLOR_TEXT_PRIMARY};
    font-size: {scaled_size(15)}px;
}}

QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {{
    border: 1px solid {COLOR_BORDER_FOCUS};
}}

QLineEdit:disabled, QTextEdit:disabled {{
    background-color: {COLOR_BACKGROUND};
    color: {COLOR_TEXT_DISABLED};
}}

/* === Spinbox === */
QSpinBox {{
    background-color: #FFFFFF;
    border: 1px solid {COLOR_BORDER};
    border-radius: 4px;
    padding: 6px 10px;
    min-width: 70px;
    font-size: {scaled_size(15)}px;
}}

QSpinBox:focus {{
    border: 1px solid {COLOR_BORDER_FOCUS};
}}

QSpinBox::up-button, QSpinBox::down-button {{
    width: 18px;
    border: none;
}}

/* === ComboBox === */
QComboBox {{
    background-color: #FFFFFF;
    border: 1px solid {COLOR_BORDER};
    border-radius: 4px;
    padding: 6px 10px;
    min-width: 100px;
    font-size: {scaled_size(15)}px;
}}

QComboBox:focus {{
    border: 1px solid {COLOR_BORDER_FOCUS};
}}

QComboBox::drop-down {{
    border: none;
    width: 22px;
}}

QComboBox QAbstractItemView {{
    background-color: #FFFFFF;
    border: 1px solid {COLOR_BORDER};
    selection-background-color: {COLOR_ACCENT_LIGHT};
    selection-color: {COLOR_TEXT_PRIMARY};
}}

/* === List Widget === */
QListWidget {{
    background-color: {COLOR_PANEL_BG};
    border: 1px solid {COLOR_BORDER};
    border-radius: 4px;
    padding: 4px;
    outline: none;
    font-size: {scaled_size(15)}px;
}}

QListWidget::item {{
    padding: 7px 10px;
    border-bottom: 1px solid {COLOR_BACKGROUND};
}}

QListWidget::item:selected {{
    background-color: {COLOR_ACCENT_LIGHT};
    color: {COLOR_TEXT_PRIMARY};
}}

QListWidget::item:hover:!selected {{
    background-color: {COLOR_BACKGROUND};
}}

/* === Group Box === */
QGroupBox {{
    font-weight: bold;
    border: 1px solid {COLOR_BORDER};
    border-radius: 4px;
    margin-top: 10px;
    padding-top: 20px;
    font-size: {scaled_size(15)}px;
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 10px;
    padding: 0 6px;
    color: {COLOR_ACCENT};
}}

/* === Labels === */
QLabel {{
    background-color: transparent;
    padding: 2px;
    font-size: {scaled_size(15)}px;
}}

/* === Splitter === */
QSplitter::handle {{
    background-color: {COLOR_BORDER};
    width: 3px;
    margin: 2px 1px;
}}

QSplitter::handle:hover {{
    background-color: {COLOR_ACCENT};
}}

/* === Progress Bar === */
QProgressBar {{
    background-color: {COLOR_PROGRESS_BG};
    border: none;
    border-radius: 4px;
    text-align: center;
    height: 22px;
    font-size: {scaled_size(13)}px;
    color: {COLOR_TEXT_SECONDARY};
}}

QProgressBar::chunk {{
    background-color: {COLOR_PROGRESS_FILL};
    border-radius: 4px;
}}

/* === Scroll Bars === */
QScrollBar:vertical {{
    background-color: {COLOR_BACKGROUND};
    width: 10px;
    margin: 0;
    border: none;
}}

QScrollBar::handle:vertical {{
    background-color: {COLOR_BORDER};
    min-height: 30px;
    border-radius: 5px;
    margin: 2px;
}}

QScrollBar::handle:vertical:hover {{
    background-color: {COLOR_TEXT_DISABLED};
}}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}

QScrollBar:horizontal {{
    background-color: {COLOR_BACKGROUND};
    height: 10px;
    margin: 0;
    border: none;
}}

QScrollBar::handle:horizontal {{
    background-color: {COLOR_BORDER};
    min-width: 30px;
    border-radius: 5px;
    margin: 2px;
}}

QScrollBar::handle:horizontal:hover {{
    background-color: {COLOR_TEXT_DISABLED};
}}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0;
}}

/* === Status Bar === */
QStatusBar {{
    background-color: {COLOR_HEADER_BG};
    border-top: 1px solid {COLOR_BORDER};
    color: {COLOR_TEXT_SECONDARY};
    font-size: {scaled_size(13)}px;
    padding: 4px;
}}

/* === Tool Tips === */
QToolTip {{
    background-color: {COLOR_TEXT_PRIMARY};
    color: white;
    border: none;
    padding: 6px 10px;
    border-radius: 3px;
    font-size: {scaled_size(13)}px;
}}

/* === Form Layout === */
QFormLayout {{
    spacing: 10px;
}}
"""


def get_stylesheet() -> str:
    """
    WHAT:
        Returns the application stylesheet string with font sizes
        scaled to the current screen. Call compute_font_scale() first
        to set the scale factor.

    CALLED BY:
        → main.py → applied to QApplication.

    RETURNS:
        str — complete QSS stylesheet.
    """
    return _build_stylesheet()
