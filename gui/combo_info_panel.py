"""
FILE: gui/combo_info_panel.py

PURPOSE: Shows details of the currently selected combination: sum,
         difference from target, item count, and individual values.
         Sits below the source list in the right panel.

CONTAINS:
- ComboInfoPanel — QWidget showing selected combination details

DEPENDS ON:
- models/combination.py → Combination
- utils/format_helpers.py → format_number_indian, format_difference
- gui/styles.py → color constants

USED BY:
- gui/source_panel.py → embeds ComboInfoPanel below the source list

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — combination detail display      | Sub-phase 1D input + source      |
| 23-03-2026 | Added scroll area with max height         | Large combos pushed source away  |
"""

# Group 2: Third-party libraries
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QGroupBox, QScrollArea,
)
from PyQt5.QtCore import Qt

# Group 3: This project's modules
from models.combination import Combination
from utils.format_helpers import format_number_indian, format_difference
from gui.styles import COLOR_TEXT_SECONDARY, COLOR_SUCCESS, COLOR_ACCENT, scaled_size, scaled_px


# Maximum height for the combo info panel before scrolling kicks in
# (base value — scaled by scaled_px at runtime)
COMBO_INFO_MAX_HEIGHT_BASE = 150


class ComboInfoPanel(QWidget):
    """
    WHAT:
        Displays detailed information about the currently selected
        combination from the results list. Shows sum, difference,
        count, and lists all item values.

    WHY ADDED:
        When a user clicks a combination in the results list, they
        need to see the details — what numbers are in it, what the
        sum is, and how close it is to the target. This panel provides
        that at-a-glance view.

    CALLED BY:
        → gui/source_panel.py → embeds this below the source list
        → gui/find_tab.py → calls show_combination() when selection changes

    ASSUMPTIONS:
        - Only one combination can be selected at a time.
        - show_combination(None) clears the panel.
        *** ASSUMPTION: Max height of 150px chosen to keep the panel
            compact. With ~14px font, this fits about 7-8 lines of
            values before scrolling. Enough for most combos, scrollable
            for large ones. ***
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        """
        WHAT: Creates the combination info display with a scrollable area
              so large combinations don't push the source list out of view.
        CALLED BY: __init__()
        """
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # --- Group box ---
        self._group = QGroupBox("Selected Combination")
        group_layout = QVBoxLayout(self._group)
        m = scaled_px(8)
        group_layout.setContentsMargins(m, m, m, m)
        group_layout.setSpacing(scaled_px(4))

        # Summary line (sum, difference, count) — always visible, outside scroll
        self._summary_label = QLabel("No combination selected")
        self._summary_label.setStyleSheet(f"color: {COLOR_TEXT_SECONDARY};")
        self._summary_label.setWordWrap(True)
        group_layout.addWidget(self._summary_label)

        # Scrollable area for item values (can grow large)
        self._scroll_area = QScrollArea()
        self._scroll_area.setWidgetResizable(True)
        self._scroll_area.setFrameShape(QScrollArea.NoFrame)
        self._scroll_area.setMaximumHeight(scaled_px(COMBO_INFO_MAX_HEIGHT_BASE))
        self._scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        # Container widget inside the scroll area
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setContentsMargins(0, 0, 0, 0)
        scroll_layout.setSpacing(0)

        self._values_label = QLabel("")
        self._values_label.setStyleSheet(
            f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(14)}px;"
        )
        self._values_label.setWordWrap(True)
        self._values_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        scroll_layout.addWidget(self._values_label)

        self._scroll_area.setWidget(scroll_content)
        self._scroll_area.setVisible(False)  # Hidden until a combo is selected
        group_layout.addWidget(self._scroll_area)

        layout.addWidget(self._group)

    def show_combination(self, combo):
        """
        WHAT:
            Updates the panel to show details of the given combination.
            Pass None to clear the display.

        CALLED BY:
            → gui/find_tab.py → when user selects a combo in results

        PARAMETERS:
            combo (Combination or None): The combination to display.
        """
        if combo is None:
            self._summary_label.setText("No combination selected")
            self._summary_label.setStyleSheet(f"color: {COLOR_TEXT_SECONDARY};")
            self._values_label.setText("")
            self._scroll_area.setVisible(False)
            return

        # Summary: "Sum: 1,00,000.00 | Diff: Exact | 3 items"
        sum_str = format_number_indian(combo.sum_value)
        diff_str = format_difference(combo.difference)
        count = combo.size

        diff_color = COLOR_SUCCESS if diff_str == "Exact" else COLOR_ACCENT
        self._summary_label.setText(
            f"Sum: {sum_str}  |  Diff: {diff_str}  |  "
            f"{count} item{'s' if count != 1 else ''}"
        )
        self._summary_label.setStyleSheet(f"font-weight: bold;")

        # Individual values — one per line for readability in scroll area
        value_lines = [format_number_indian(item.value) for item in combo.items]
        self._values_label.setText("Values: " + ", ".join(value_lines))
        self._values_label.setStyleSheet(
            f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(14)}px;"
        )
        self._scroll_area.setVisible(True)

    def clear(self):
        """
        WHAT: Resets the panel to empty state.
        CALLED BY: gui/find_tab.py → on clear all or new search.
        """
        self.show_combination(None)
