"""
FILE: gui/source_panel.py

PURPOSE: The right panel of the Find tab. Shows all loaded numbers
         in a scrollable list. Highlights items belonging to the
         currently selected combination. Shows count and total.

CONTAINS:
- SourcePanel — QWidget with the source numbers list and combo info

DEPENDS ON:
- config/constants.py → SELECTED_MARKER, FINALIZED_MARKER, SEED_MARKER
- models/number_item.py → NumberItem
- models/combination.py → Combination
- utils/format_helpers.py → format_number_indian
- gui/combo_info_panel.py → ComboInfoPanel
- gui/styles.py → color constants

USED BY:
- gui/find_tab.py → embeds SourcePanel in the right section of the splitter

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — source numbers list panel       | Sub-phase 1D input + source      |
| 22-03-2026 | Added refresh_display() for finalization  | Sub-phase 2B finalization UI     |
| 22-03-2026 | Bold + orange fill for highlighted items  | UI fixes before Phase 3          |
| 22-03-2026 | Added add_items() for accumulative grab   | Sub-phase 3A Excel integration   |
| 23-03-2026 | Added seed pin/unpin via double-click     | Phase 4A seed numbers            |
"""

# Group 1: Python standard library
from typing import List, Optional, Set

# Group 2: Third-party libraries
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QListWidget,
    QListWidgetItem, QAbstractItemView, QStyledItemDelegate,
)
from PyQt5.QtCore import Qt, pyqtSignal, QRect
from PyQt5.QtGui import QColor, QBrush, QFont, QPainter, QPen

# Group 3: This project's modules
from config.constants import SELECTED_MARKER, FINALIZED_MARKER, SEED_MARKER
from models.number_item import NumberItem
from models.combination import Combination
from utils.format_helpers import format_number_indian
from gui.combo_info_panel import ComboInfoPanel
from gui.styles import (
    COLOR_SELECTED_HIGHLIGHT, COLOR_SELECTED_BORDER,
    COLOR_TEXT_SECONDARY, COLOR_ACCENT, COLOR_TEXT_PRIMARY,
    scaled_size,
)


class SourceItemDelegate(QStyledItemDelegate):
    """
    WHAT:
        Custom delegate that paints item backgrounds directly,
        bypassing QSS which overrides programmatic QBrush styling.
        This is the ONLY reliable way to show colored backgrounds
        on QListWidgetItem when QSS is active.

    WHY ADDED:
        Qt's QSS for QListWidget::item overrides setBackground()
        calls on individual items. This bug existed in the original
        tool and was never fixed. Using a delegate that paints
        backgrounds before text bypasses QSS entirely.

    CALLED BY:
        → SourcePanel._setup_ui() → set as delegate on the list widget
    """

    def paint(self, painter, option, index):
        """
        WHAT: Paints background, then text, bypassing QSS item styling.
        """
        painter.save()

        # Draw background from item's BackgroundRole
        bg_brush = index.data(Qt.BackgroundRole)
        if bg_brush is not None and isinstance(bg_brush, QBrush) and bg_brush.color().alpha() > 0:
            painter.fillRect(option.rect, bg_brush)

        # Draw font if set
        font = index.data(Qt.FontRole)
        if font is not None and isinstance(font, QFont):
            painter.setFont(font)
        else:
            painter.setFont(option.font)

        # Draw text with foreground color
        fg_brush = index.data(Qt.ForegroundRole)
        if fg_brush is not None and isinstance(fg_brush, QBrush):
            painter.setPen(fg_brush.color())
        else:
            painter.setPen(QColor(COLOR_TEXT_PRIMARY))

        text = index.data(Qt.DisplayRole)
        if text:
            text_rect = option.rect.adjusted(10, 0, -4, 0)
            painter.drawText(text_rect, Qt.AlignLeft | Qt.AlignVCenter, text)

        painter.restore()

    def sizeHint(self, option, index):
        """
        WHAT: Returns item height with comfortable padding.
        """
        hint = super().sizeHint(option, index)
        hint.setHeight(max(hint.height(), 30))
        return hint


class ScrollbarMarkerOverlay(QWidget):
    """
    WHAT:
        A thin transparent widget overlaid on the scrollbar that draws
        small colored markers showing where highlighted items are in
        the list. Similar to find-in-page markers in web browsers.

    WHY ADDED:
        When data is large (800+ numbers), users need to see at a
        glance where highlighted items are without scrolling through
        the entire list.

    CALLED BY:
        → SourcePanel._setup_ui() → created and positioned over scrollbar
    """

    def __init__(self, list_widget: QListWidget, parent=None):
        super().__init__(parent)
        self._list_widget = list_widget
        self._marker_positions: List[float] = []
        self._marker_color = QColor("#FF9800")  # Orange
        self._finalized_markers: List[tuple] = []  # (position, r, g, b)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedWidth(14)

    def set_highlighted_positions(self, indices: Set[int], total_items: int):
        """
        WHAT: Calculates marker positions from item indices.
        CALLED BY: SourcePanel.highlight_combination() (legacy)
        """
        if total_items == 0:
            self._marker_positions = []
        else:
            self._marker_positions = [idx / total_items for idx in indices]
        self.update()

    def set_highlighted_positions_direct(self, positions: List[float]):
        """
        WHAT: Sets marker positions from pre-computed row fractions.
        CALLED BY: SourcePanel._update_scrollbar_markers()
        """
        self._marker_positions = positions
        self.update()

    def set_finalized_positions(self, positions: List[tuple]):
        """
        WHAT: Sets positions and colors for finalized item markers.
        CALLED BY: SourcePanel.refresh_display()
        """
        self._finalized_markers = positions
        self.update()

    def clear(self):
        """WHAT: Removes all markers."""
        self._marker_positions = []
        self._finalized_markers = []
        self.update()

    def paintEvent(self, event):
        """WHAT: Draws colored markers on the overlay."""
        if not self._marker_positions and not self._finalized_markers:
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        h = self.height()
        w = self.width()
        # Marker height — thick enough to be visible at a glance
        marker_h = max(4, int(h * 0.005))

        # Draw finalized markers (color from finalization)
        for pos, r, g, b in self._finalized_markers:
            y = max(0, min(int(pos * h), h - marker_h))
            color = QColor(r, g, b)
            color.setAlpha(200)
            painter.fillRect(0, y, w, marker_h, color)

        # Draw highlighted markers (bright orange, on top, larger)
        highlight_h = max(5, int(h * 0.007))
        for pos in self._marker_positions:
            y = max(0, min(int(pos * h), h - highlight_h))
            painter.fillRect(0, y, w, highlight_h, QColor("#FF6D00"))
            # Bright border for contrast
            painter.setPen(QPen(QColor("#E65100"), 1))
            painter.drawRect(0, y, w - 1, highlight_h - 1)

        painter.end()


# Qt data role for storing the NumberItem index on each list row.
# Using UserRole so the actual NumberItem index is always accessible
# regardless of visual position — this prevents the combo removal
# index mismatch bug from the original tool.
ITEM_INDEX_ROLE = Qt.UserRole


class SourcePanel(QWidget):
    """
    WHAT:
        Shows all loaded numbers in a list, with their index and
        formatted value. Highlights items that belong to the currently
        selected combination in orange. Displays count and total at
        the top. Includes the ComboInfoPanel below the list.
        Supports double-click to toggle seed (pin/unpin) on items.

    WHY ADDED:
        Users need to see all their loaded numbers and visually
        identify which ones are part of a selected combination.
        The original tool did this but the orange highlight never
        worked (QBrush bug) — this version fixes that.

    CALLED BY:
        → gui/find_tab.py → creates and embeds this panel

    SIGNALS EMITTED:
        → seeds_changed() — Emitted when a seed is pinned or unpinned.
          FindTab listens to recompute smart bounds hints.

    ASSUMPTIONS:
        - Items are displayed in their original load order.
        - Each list row stores the NumberItem.index in ITEM_INDEX_ROLE
          for reliable identification (not row position).
        *** ASSUMPTION: The source list is read-only — users cannot
            edit values directly. All modifications (finalization,
            seeds) happen through other controls. ***
    """

    seeds_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._items: List[NumberItem] = []
        self._items_by_index: dict = {}  # index → NumberItem for O(1) lookup
        self._highlighted_indices: Set[int] = set()
        self._setup_ui()

    def _setup_ui(self):
        """
        WHAT: Creates the source list, header, and combo info panel.
        CALLED BY: __init__()
        """
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # --- Header with count and total ---
        header_layout = QHBoxLayout()
        self._count_label = QLabel("No numbers loaded")
        self._count_label.setStyleSheet(f"font-weight: bold; color: {COLOR_ACCENT};")
        self._total_label = QLabel("")
        self._total_label.setStyleSheet(f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(13)}px;")
        self._total_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        header_layout.addWidget(self._count_label)
        header_layout.addWidget(self._total_label, 1)
        layout.addLayout(header_layout)

        # --- Sticky source header (shows current group when scrolling) ---
        self._sticky_header = QLabel("")
        self._sticky_header.setStyleSheet(
            f"font-weight: bold; font-size: {scaled_size(12)}px; "
            f"color: {COLOR_ACCENT}; background-color: #EEF2F7; "
            f"padding: 4px 8px; border: 1px solid #D1D5DB; border-radius: 3px;"
        )
        self._sticky_header.setWordWrap(True)
        self._sticky_header.hide()
        layout.addWidget(self._sticky_header)

        # --- Source list ---
        self._list_widget = QListWidget()
        self._list_widget.setSelectionMode(QAbstractItemView.NoSelection)
        self._list_widget.setFocusPolicy(Qt.NoFocus)
        # Custom delegate bypasses QSS to render programmatic backgrounds
        self._list_widget.setItemDelegate(SourceItemDelegate(self._list_widget))
        self._list_widget.setWordWrap(True)
        layout.addWidget(self._list_widget, 1)

        # --- Scrollbar marker overlay ---
        self._scrollbar_overlay = ScrollbarMarkerOverlay(self._list_widget, self)
        self._scrollbar_overlay.hide()

        # --- Double-click to toggle seed pin/unpin ---
        self._list_widget.itemDoubleClicked.connect(self._on_item_double_clicked)

        # --- Connect scroll to update sticky header ---
        self._list_widget.verticalScrollBar().valueChanged.connect(
            self._update_sticky_header
        )

        # --- Combo info panel ---
        self._combo_info = ComboInfoPanel()
        layout.addWidget(self._combo_info)

    # -----------------------------------------------------------------------
    # Public Methods
    # -----------------------------------------------------------------------

    def load_items(self, items: List[NumberItem]):
        """
        WHAT:
            Populates the source list with the given NumberItem list.
            Clears any existing items first.

        CALLED BY:
            → gui/find_tab.py → when numbers are loaded

        PARAMETERS:
            items (List[NumberItem]): The parsed number items to display.
        """
        self._items = items
        self._items_by_index = {item.index: item for item in items}
        self._highlighted_indices.clear()
        self._rebuild_list()
        self._update_header()
        self._update_sticky_header()

    def add_items(self, new_items: List[NumberItem]):
        """
        WHAT:
            Appends new items to the existing source list without replacing.
            Used for accumulative grab — "Grab from Excel" adds to the
            current list instead of clearing it.

        CALLED BY:
            → gui/find_tab.py → on grab from Excel (accumulative)

        PARAMETERS:
            new_items (List[NumberItem]): The new items to append.
        """
        self._items.extend(new_items)
        for item in new_items:
            self._items_by_index[item.index] = item
        self._highlighted_indices.clear()
        self._rebuild_list()
        self._update_header()
        self._update_sticky_header()

    def highlight_combination(self, combo: Optional[Combination]):
        """
        WHAT:
            Highlights the items belonging to the given combination
            in the source list (orange background). Clears previous
            highlight first. Pass None to clear all highlights.

        CALLED BY:
            → gui/find_tab.py → when user selects a combo in results

        PARAMETERS:
            combo (Combination or None): The combination to highlight.

        EDGE CASES HANDLED:
            - combo is None → clears all highlights
            - combo items not in source list → silently skipped
        """
        # Clear previous highlights
        old_indices = self._highlighted_indices.copy()
        self._highlighted_indices.clear()

        if combo is not None:
            self._highlighted_indices = combo.item_indices.copy()

        # Update changed rows only (efficient for large lists)
        changed_indices = old_indices.symmetric_difference(self._highlighted_indices)
        for row in range(self._list_widget.count()):
            list_item = self._list_widget.item(row)
            item_index = list_item.data(ITEM_INDEX_ROLE)
            if item_index == -1:
                continue  # Skip source group headers
            if item_index in changed_indices:
                self._style_list_item(list_item, item_index)

        # Update combo info panel
        self._combo_info.show_combination(combo)

        # Update scrollbar markers
        self._update_scrollbar_markers()

        # Scroll to first highlighted item
        if self._highlighted_indices:
            self._scroll_to_index(min(self._highlighted_indices))

    def clear(self):
        """
        WHAT: Removes all items from the source list and resets state.
        CALLED BY: gui/find_tab.py → on clear_all_requested.
        """
        self._items = []
        self._items_by_index = {}
        self._highlighted_indices.clear()
        self._list_widget.clear()
        self._count_label.setText("No numbers loaded")
        self._total_label.setText("")
        self._combo_info.clear()
        self._scrollbar_overlay.clear()
        self._scrollbar_overlay.hide()
        self._sticky_header.hide()

    def get_item_count(self) -> int:
        """
        WHAT: Returns the number of loaded items (excludes finalized).
        CALLED BY: gui/find_tab.py → for parameter validation.
        """
        return len([item for item in self._items if not item.is_finalized])

    def get_available_items(self) -> List[NumberItem]:
        """
        WHAT: Returns non-finalized items for the solver.
        CALLED BY: gui/find_tab.py → when starting a search.
        """
        return [item for item in self._items if not item.is_finalized]

    def get_all_items(self) -> List[NumberItem]:
        """
        WHAT: Returns all loaded items (including finalized).
        CALLED BY: gui/find_tab.py → for session save.
        """
        return list(self._items)

    def get_seed_indices(self) -> list:
        """
        WHAT:
            Returns the indices of all items currently pinned as seeds.
            Only returns non-finalized seed items (finalized items are
            excluded from the search anyway).

        CALLED BY:
            → gui/find_tab.py → when building SearchParameters for the solver

        RETURNS:
            list: List of NumberItem.index values for pinned seeds.
        """
        return [
            item.index for item in self._items
            if item.is_seed and not item.is_finalized
        ]

    def get_seed_count(self) -> int:
        """
        WHAT: Returns the number of active (non-finalized) seeds.
        CALLED BY: gui/find_tab.py → for display purposes.
        """
        return len([
            item for item in self._items
            if item.is_seed and not item.is_finalized
        ])

    def clear_all_seeds(self):
        """
        WHAT: Unpins all seed items. Called on Clear All.
        CALLED BY: gui/find_tab.py → on clear_all_requested.
        """
        for item in self._items:
            item.is_seed = False

    def refresh_display(self):
        """
        WHAT:
            Rebuilds the entire list display and updates the header.
            Called after finalization or undo to reflect the updated
            is_finalized / finalized_color state of items.

        CALLED BY:
            → gui/find_tab.py → after finalize or undo operations
        """
        self._highlighted_indices.clear()
        self._rebuild_list()
        self._update_header()
        self._combo_info.clear()
        self._update_scrollbar_markers()

    # -----------------------------------------------------------------------
    # Internal Methods
    # -----------------------------------------------------------------------

    def _on_item_double_clicked(self, list_item: QListWidgetItem):
        """
        WHAT:
            Toggles the seed (pin/unpin) state of a double-clicked item.
            Finalized items cannot be pinned as seeds. Updates the display
            and emits seeds_changed signal.

        CALLED BY:
            → QListWidget.itemDoubleClicked signal

        EDGE CASES HANDLED:
            - Double-click on group header → ignored (ITEM_INDEX_ROLE == -1)
            - Double-click on finalized item → ignored (already locked)

        PARAMETERS:
            list_item (QListWidgetItem): The list item that was double-clicked.
        """
        item_index = list_item.data(ITEM_INDEX_ROLE)
        if item_index is None or item_index == -1:
            return  # Skip group headers

        number_item = self._items_by_index.get(item_index)
        if number_item is None:
            return

        # Cannot seed a finalized item
        if number_item.is_finalized:
            return

        # Toggle seed state
        number_item.is_seed = not number_item.is_seed

        # Update just this row's display
        self._style_list_item(list_item, item_index)

        # Update header to show seed count
        self._update_header()

        # Notify FindTab to recompute bounds
        self.seeds_changed.emit()

    def _rebuild_list(self):
        """
        WHAT:
            Clears and repopulates the QListWidget from self._items.
            If items have SourceTag metadata, groups them by
            workbook → sheet with non-selectable header rows.
            Uses setUpdatesEnabled(False) during rebuild to prevent
            per-item repaints (important for 800+ items).

        CALLED BY: load_items(), add_items(), refresh_display()
        """
        self._list_widget.setUpdatesEnabled(False)
        self._list_widget.clear()

        # Check if any items have source tags (Excel grab)
        has_sources = any(
            item.source is not None for item in self._items
        )

        if has_sources:
            self._rebuild_list_grouped()
        else:
            self._rebuild_list_flat()

        self._list_widget.setUpdatesEnabled(True)

    def _rebuild_list_flat(self):
        """
        WHAT: Populates the list without grouping (manual input mode).
        CALLED BY: _rebuild_list() → when no SourceTag metadata.
        """
        for item in self._items:
            list_item = QListWidgetItem()
            list_item.setData(ITEM_INDEX_ROLE, item.index)

            text = self._format_item_text(item)
            list_item.setText(text)

            self._style_list_item(list_item, item.index)
            self._list_widget.addItem(list_item)

    def _rebuild_list_grouped(self):
        """
        WHAT:
            Populates the list grouped by workbook → sheet.
            Adds non-selectable header rows like:
              [Book1.xlsx → Sheet1]
            Followed by items with cell references:
              1. A1: 245.70

            Items without SourceTag are grouped under "Manual Input".

        CALLED BY: _rebuild_list() → when SourceTag metadata exists.
        """
        # Group items by (workbook, sheet)
        groups = {}
        manual_items = []

        for item in self._items:
            if item.source is not None:
                key = (item.source.workbook_name, item.source.sheet_name)
                if key not in groups:
                    groups[key] = []
                groups[key].append(item)
            else:
                manual_items.append(item)

        # Add manual items first (if any)
        if manual_items:
            self._add_source_header("Manual Input")
            for item in manual_items:
                self._add_item_row(item)

        # Add grouped items
        for (wb_name, sheet_name), items in groups.items():
            self._add_source_header(f"{wb_name} → {sheet_name}")
            for item in items:
                self._add_item_row(item)

    def _add_source_header(self, header_text: str):
        """
        WHAT: Adds a non-selectable group header row to the source list.
        CALLED BY: _rebuild_list_grouped()
        """
        header_item = QListWidgetItem(f"  [{header_text}]")
        header_item.setFlags(Qt.NoItemFlags)  # Non-selectable, non-clickable
        header_font = QFont()
        header_font.setBold(True)
        header_font.setPointSize(11)
        header_item.setFont(header_font)
        header_item.setForeground(QBrush(QColor(COLOR_ACCENT)))
        header_item.setBackground(QBrush(QColor("#EEF2F7")))
        # Store a sentinel so highlight_combination skips it
        header_item.setData(ITEM_INDEX_ROLE, -1)
        self._list_widget.addItem(header_item)

    def _add_item_row(self, item: NumberItem):
        """
        WHAT: Adds a single NumberItem row to the source list.
        CALLED BY: _rebuild_list_grouped(), _rebuild_list_flat() (indirectly)
        """
        list_item = QListWidgetItem()
        list_item.setData(ITEM_INDEX_ROLE, item.index)

        text = self._format_item_text(item)
        list_item.setText(text)

        self._style_list_item(list_item, item.index)
        self._list_widget.addItem(list_item)

    def _format_item_text(self, item: NumberItem) -> str:
        """
        WHAT: Formats a NumberItem for display in the source list.
        CALLED BY: _rebuild_list(), _style_list_item()

        FORMAT:
            "  1.  1,00,000.00"
            "✓ 2.  25,000.50"     (finalized)
            "▶ 3.  -500.00"       (selected/highlighted)
            "[SEED] 4.  200.00"   (seed)
        """
        prefix = ""
        if item.is_finalized:
            prefix = f"{FINALIZED_MARKER} "
        elif item.index in self._highlighted_indices:
            prefix = f"{SELECTED_MARKER} "
        elif item.is_seed:
            prefix = f"{SEED_MARKER} "
        else:
            prefix = "  "

        formatted_value = format_number_indian(item.value)
        # Right-align index numbers with padding
        index_str = f"{item.index + 1}."

        # Include cell reference if item has SourceTag
        if item.source is not None and item.source.cell_address:
            return f"{prefix}{index_str:>4}  {item.source.cell_address}: {formatted_value}"
        return f"{prefix}{index_str:>4}  {formatted_value}"

    def _style_list_item(self, list_item: QListWidgetItem, item_index: int):
        """
        WHAT:
            Applies visual styling to a list item based on its state
            (highlighted, finalized, normal).

        CALLED BY:
            → _rebuild_list(), highlight_combination()

        EDGE CASES HANDLED:
            - Highlighted (selected combo) → orange background
            - Finalized → colored background matching finalization color
            - Normal → default background
        """
        # Find the NumberItem for this index using the lookup dict
        item = self._items_by_index.get(item_index)
        if item is None:
            return

        # Update text (prefix may have changed)
        list_item.setText(self._format_item_text(item))

        # Reset font to normal first
        normal_font = QFont()
        list_item.setFont(normal_font)

        if item_index in self._highlighted_indices and not item.is_finalized:
            # Orange highlight for selected combination items:
            # bold text + solid orange background fill + dark text
            bold_font = QFont()
            bold_font.setBold(True)
            list_item.setFont(bold_font)
            list_item.setBackground(QBrush(QColor("#FFE0B2")))  # Solid orange fill
            list_item.setForeground(QBrush(QColor("#BF360C")))  # Dark orange text
        elif item.is_finalized and item.finalized_color:
            # Finalized item — show in its assigned color
            r, g, b = item.finalized_color
            list_item.setBackground(QBrush(QColor(r, g, b)))
            list_item.setForeground(QBrush(QColor("#1F2937")))
        elif item.is_seed:
            # Seed item — light blue background with bold text
            bold_font = QFont()
            bold_font.setBold(True)
            list_item.setFont(bold_font)
            list_item.setBackground(QBrush(QColor("#BBDEFB")))  # Light blue fill
            list_item.setForeground(QBrush(QColor("#0D47A1")))  # Dark blue text
        else:
            # Normal item
            list_item.setBackground(QBrush(QColor(255, 255, 255, 0)))
            list_item.setForeground(QBrush(QColor("#1F2937")))

    def _update_header(self):
        """
        WHAT: Updates the count and total labels above the list.
              Shows seed count when seeds are pinned.
        CALLED BY: load_items(), _on_item_double_clicked()
        """
        available = [item for item in self._items if not item.is_finalized]
        count = len(available)
        total = sum(item.value for item in available)
        seed_count = self.get_seed_count()

        count_text = f"{count} number{'s' if count != 1 else ''}"
        if seed_count > 0:
            count_text += f" ({seed_count} pinned)"

        self._count_label.setText(count_text)
        self._total_label.setText(
            f"Total: {format_number_indian(total)}"
        )

    def _update_sticky_header(self):
        """
        WHAT:
            Updates the sticky header label to show the current source
            group name based on scroll position. Scans upward from the
            first visible item to find the most recent header.

        CALLED BY:
            → QListWidget.verticalScrollBar().valueChanged signal
        """
        # Only relevant for grouped display
        has_sources = any(item.source is not None for item in self._items)
        if not has_sources:
            self._sticky_header.hide()
            return

        # Find the first visible item
        first_visible = self._list_widget.itemAt(2, 2)
        if first_visible is None:
            self._sticky_header.hide()
            return

        first_row = self._list_widget.row(first_visible)

        # Scan backward from first visible to find the most recent header
        for row in range(first_row, -1, -1):
            item = self._list_widget.item(row)
            if item is not None and item.data(ITEM_INDEX_ROLE) == -1:
                header_text = item.text().strip().strip("[]")
                self._sticky_header.setText(header_text)
                self._sticky_header.show()
                return

        self._sticky_header.hide()

    def _update_scrollbar_markers(self):
        """
        WHAT:
            Updates the scrollbar overlay markers to show where
            highlighted and finalized items are in the list.
            Uses actual row positions in the list widget (not item
            indices) so markers align correctly even when the list
            has group header rows.

        CALLED BY:
            → highlight_combination(), refresh_display()
        """
        total_rows = self._list_widget.count()
        if total_rows == 0:
            self._scrollbar_overlay.hide()
            return

        # Build row-position markers by scanning the list widget
        highlight_positions = []
        fin_markers = []

        for row in range(total_rows):
            list_item = self._list_widget.item(row)
            item_index = list_item.data(ITEM_INDEX_ROLE)
            if item_index == -1:
                continue  # Skip group headers

            row_fraction = row / total_rows

            # Check highlighted
            if item_index in self._highlighted_indices:
                highlight_positions.append(row_fraction)

            # Check finalized
            number_item = self._items_by_index.get(item_index)
            if number_item and number_item.is_finalized and number_item.finalized_color:
                r, g, b = number_item.finalized_color
                fin_markers.append((row_fraction, r, g, b))

        self._scrollbar_overlay.set_highlighted_positions_direct(highlight_positions)
        self._scrollbar_overlay.set_finalized_positions(fin_markers)

        # Position overlay directly on the scrollbar
        scrollbar = self._list_widget.verticalScrollBar()
        if scrollbar.isVisible():
            self._scrollbar_overlay.setFixedHeight(scrollbar.height())
            # Overlay on top of the scrollbar itself
            sb_pos = scrollbar.mapTo(self, scrollbar.rect().topLeft())
            self._scrollbar_overlay.move(sb_pos.x(), sb_pos.y())
            self._scrollbar_overlay.show()
        elif highlight_positions or fin_markers:
            # Scrollbar not visible but we have markers — show at right edge
            self._scrollbar_overlay.setFixedHeight(self._list_widget.height())
            list_right = self._list_widget.x() + self._list_widget.width()
            self._scrollbar_overlay.move(list_right - 12, self._list_widget.y())
            self._scrollbar_overlay.show()
        else:
            self._scrollbar_overlay.hide()

    def resizeEvent(self, event):
        """
        WHAT: Repositions the scrollbar overlay when the panel is resized.
        """
        super().resizeEvent(event)
        if self._scrollbar_overlay.isVisible():
            self._update_scrollbar_markers()

    def _scroll_to_index(self, item_index: int):
        """
        WHAT: Scrolls the list to make the given item index visible.
        CALLED BY: highlight_combination()
        """
        for row in range(self._list_widget.count()):
            list_item = self._list_widget.item(row)
            if list_item.data(ITEM_INDEX_ROLE) == item_index:
                self._list_widget.scrollToItem(
                    list_item, QAbstractItemView.PositionAtCenter
                )
                break
