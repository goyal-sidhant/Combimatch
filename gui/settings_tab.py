"""
FILE: gui/settings_tab.py

PURPOSE: The Settings tab. Provides Excel connection controls,
         workbook/sheet selection with checkboxes (N8 multi-grab),
         color palette preview, and About section.

CONTAINS:
- SettingsTab — QWidget with Excel connection and workbook/sheet tree

DEPENDS ON:
- config/constants.py → HIGHLIGHT_COLORS
- readers/excel_workbook_manager.py → WorkbookManager
- gui/styles.py → color constants

USED BY:
- gui/main_window.py → embeds SettingsTab as the third tab

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — settings tab with Excel + tree | Sub-phase 3A Excel integration   |
"""

# Group 1: Python standard library
# (none)

# Group 2: Third-party libraries
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QTreeWidget, QTreeWidgetItem, QFrame,
    QGridLayout, QScrollArea,
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor

# Group 3: This project's modules
from config.constants import HIGHLIGHT_COLORS
from readers.excel_workbook_manager import WorkbookManager
from gui.styles import (
    COLOR_ACCENT, COLOR_TEXT_SECONDARY, COLOR_TEXT_PRIMARY,
    COLOR_SUCCESS, COLOR_ERROR, COLOR_PANEL_BG, COLOR_BORDER,
    scaled_size, scaled_px,
)


class SettingsTab(QWidget):
    """
    WHAT:
        The Settings tab provides Excel connection management and
        workbook/sheet selection for multi-source grab (N8). Also shows
        a color palette preview and About section.

    WHY ADDED:
        Users need a dedicated place to connect to Excel, choose which
        workbooks and sheets to grab from, and preview the finalization
        color palette. Separating this from the Find tab keeps both tabs
        focused on their own responsibilities.

    CALLED BY:
        → gui/main_window.py → added as the "Settings" tab

    SIGNALS EMITTED:
        → excel_connected() — after successful connection
        → excel_disconnected() — after disconnection

    ASSUMPTIONS:
        - WorkbookManager is shared with FindTab (via MainWindow).
        - The tree uses checkboxes at both workbook and sheet levels.
          Toggling a workbook toggles all its sheets.
        *** ASSUMPTION: The user connects to Excel here first, then
            switches to the Find tab to grab numbers. The connection
            persists across tab switches. ***
    """

    excel_connected = pyqtSignal()
    excel_disconnected = pyqtSignal()

    def __init__(self, workbook_manager: WorkbookManager, parent=None):
        """
        WHAT: Creates the Settings tab with Excel controls and color preview.

        PARAMETERS:
            workbook_manager (WorkbookManager): Shared manager for Excel operations.
        """
        super().__init__(parent)
        self._manager = workbook_manager
        self._setup_ui()

    def _setup_ui(self):
        """
        WHAT: Creates the settings layout with connection, tree, colors, about.
        CALLED BY: __init__()
        """
        layout = QVBoxLayout(self)
        m = scaled_px(12)
        layout.setContentsMargins(m, m, m, m)
        layout.setSpacing(m)

        # --- Excel Connection Group ---
        conn_group = QGroupBox("Excel Connection")
        conn_layout = QVBoxLayout(conn_group)
        conn_layout.setSpacing(scaled_px(10))

        # Status indicator
        status_layout = QHBoxLayout()
        status_label = QLabel("Status:")
        status_label.setStyleSheet(f"font-weight: bold;")
        self._status_text = QLabel("Not connected")
        self._status_text.setStyleSheet(
            f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(15)}px;"
        )
        status_layout.addWidget(status_label)
        status_layout.addWidget(self._status_text, 1)
        conn_layout.addLayout(status_layout)

        # Connect / Disconnect buttons
        btn_layout = QHBoxLayout()
        self._connect_btn = QPushButton("Connect to Excel")
        self._connect_btn.clicked.connect(self._on_connect)
        btn_layout.addWidget(self._connect_btn)

        self._disconnect_btn = QPushButton("Disconnect")
        self._disconnect_btn.setEnabled(False)
        self._disconnect_btn.setStyleSheet(
            f"QPushButton {{ background-color: {COLOR_ERROR}; }}"
            f"QPushButton:hover {{ background-color: #DC2626; }}"
            f"QPushButton:disabled {{ background-color: #D1D5DB; color: #9CA3AF; }}"
        )
        self._disconnect_btn.clicked.connect(self._on_disconnect)
        btn_layout.addWidget(self._disconnect_btn)

        conn_layout.addLayout(btn_layout)
        layout.addWidget(conn_group)

        # --- Workbooks & Sheets Group ---
        wb_group = QGroupBox("Workbooks && Sheets")
        wb_layout = QVBoxLayout(wb_group)
        wb_layout.setSpacing(scaled_px(8))

        # Refresh button
        refresh_layout = QHBoxLayout()
        self._refresh_btn = QPushButton("Refresh")
        self._refresh_btn.setEnabled(False)
        self._refresh_btn.clicked.connect(self._on_refresh)
        refresh_layout.addWidget(self._refresh_btn)
        refresh_layout.addStretch()

        checked_label = QLabel("")
        self._checked_count_label = checked_label
        checked_label.setStyleSheet(f"color: {COLOR_TEXT_SECONDARY};")
        refresh_layout.addWidget(checked_label)

        wb_layout.addLayout(refresh_layout)

        # Workbook/sheet tree with checkboxes
        self._tree = QTreeWidget()
        self._tree.setHeaderHidden(True)
        self._tree.setRootIsDecorated(True)
        self._tree.setAnimated(True)
        self._tree.setMinimumHeight(scaled_px(150))
        self._tree.itemChanged.connect(self._on_tree_item_changed)
        wb_layout.addWidget(self._tree, 1)

        # Hint
        hint = QLabel(
            "Check the sheets you want to grab from, then select cells\n"
            "on each sheet in Excel and use \"Grab from Excel\" in the Find tab."
        )
        hint.setStyleSheet(
            f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(13)}px; "
            f"font-style: italic; padding: {scaled_px(4)}px;"
        )
        hint.setWordWrap(True)
        wb_layout.addWidget(hint)

        layout.addWidget(wb_group, 1)

        # --- Color Palette Group ---
        colors_group = QGroupBox("Finalization Colors")
        colors_layout = QVBoxLayout(colors_group)

        palette_grid = QGridLayout()
        palette_grid.setSpacing(scaled_px(4))

        for idx, (rgb, name) in enumerate(HIGHLIGHT_COLORS):
            row = idx // 10
            col = idx % 10
            swatch = QFrame()
            swatch.setFixedSize(scaled_px(28), scaled_px(28))
            r, g, b = rgb
            swatch.setStyleSheet(
                f"background-color: rgb({r}, {g}, {b}); "
                f"border: 1px solid {COLOR_BORDER}; border-radius: 3px;"
            )
            swatch.setToolTip(f"#{idx + 1}: {name}")
            palette_grid.addWidget(swatch, row, col)

        colors_layout.addLayout(palette_grid)

        palette_hint = QLabel(
            "20 colors cycle during finalization. Hover for color names."
        )
        palette_hint.setStyleSheet(
            f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(12)}px;"
        )
        colors_layout.addWidget(palette_hint)

        layout.addWidget(colors_group)

        # --- About Group ---
        about_group = QGroupBox("About")
        about_layout = QVBoxLayout(about_group)

        about_text = QLabel(
            "CombiMatch v2.0\n"
            "Subset-sum solver for invoice reconciliation\n"
            "Rebuilt March 2026"
        )
        about_text.setStyleSheet(
            f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(13)}px;"
        )
        about_layout.addWidget(about_text)

        layout.addWidget(about_group)

    # -------------------------------------------------------------------
    # Connection Handlers
    # -------------------------------------------------------------------

    def _on_connect(self):
        """
        WHAT:
            Connects to Excel via the WorkbookManager's ExcelHandler.
            On success, refreshes workbook list and updates UI.

        CALLED BY:
            → Connect button click

        EDGE CASES HANDLED:
            - pywin32 not installed → error message
            - Excel not running → error message
            - Already connected → refreshes instead
        """
        handler = self._manager.excel_handler

        if handler.is_connected:
            # Already connected — just refresh
            self._on_refresh()
            return

        result = handler.connect()

        if result["success"]:
            self._set_connected_state(True)
            self._on_refresh()
            self.excel_connected.emit()
        else:
            self._status_text.setText(result["error"])
            self._status_text.setStyleSheet(
                f"color: {COLOR_ERROR}; font-size: {scaled_size(15)}px;"
            )

    def _on_disconnect(self):
        """
        WHAT: Disconnects from Excel and resets UI.
        CALLED BY: Disconnect button click.
        """
        self._manager.excel_handler.disconnect()
        self._manager.clear()
        self._set_connected_state(False)
        self._tree.clear()
        self._update_checked_count()
        self.excel_disconnected.emit()

    def _on_refresh(self):
        """
        WHAT:
            Refreshes the workbook/sheet list from the live Excel instance.
            Preserves existing check states for sheets that are still open.

        CALLED BY:
            → Refresh button click
            → _on_connect() after successful connection
        """
        result = self._manager.refresh_workbooks()

        if result["success"]:
            self._rebuild_tree()
            self._status_text.setText(
                f"Connected — {len(result['data']['selections'])} workbook(s)"
            )
            self._status_text.setStyleSheet(
                f"color: {COLOR_SUCCESS}; font-size: {scaled_size(15)}px; "
                f"font-weight: bold;"
            )
        else:
            self._status_text.setText(result["error"])
            self._status_text.setStyleSheet(
                f"color: {COLOR_ERROR}; font-size: {scaled_size(15)}px;"
            )

    # -------------------------------------------------------------------
    # Tree Management
    # -------------------------------------------------------------------

    def _rebuild_tree(self):
        """
        WHAT:
            Clears and rebuilds the workbook/sheet tree from the
            WorkbookManager's current selections state.

        CALLED BY:
            → _on_refresh()

        WHY:
            Full rebuild is simpler than incremental updates and fast
            enough for the expected number of workbooks (typically < 10).
        """
        self._tree.blockSignals(True)
        self._tree.clear()

        selections = self._manager.get_selections()

        for wb_name, sheets in selections.items():
            # Workbook node
            wb_item = QTreeWidgetItem(self._tree, [wb_name])
            wb_item.setFlags(
                wb_item.flags() | Qt.ItemIsUserCheckable
            )
            wb_item.setData(0, Qt.UserRole, "workbook")
            wb_item.setData(0, Qt.UserRole + 1, wb_name)

            # Determine workbook check state from children
            all_checked = all(sheets.values())
            any_checked = any(sheets.values())

            if all_checked and sheets:
                wb_item.setCheckState(0, Qt.Checked)
            elif any_checked:
                wb_item.setCheckState(0, Qt.PartiallyChecked)
            else:
                wb_item.setCheckState(0, Qt.Unchecked)

            # Sheet nodes
            for sheet_name, is_checked in sheets.items():
                sheet_item = QTreeWidgetItem(wb_item, [sheet_name])
                sheet_item.setFlags(
                    sheet_item.flags() | Qt.ItemIsUserCheckable
                )
                sheet_item.setCheckState(
                    0, Qt.Checked if is_checked else Qt.Unchecked,
                )
                sheet_item.setData(0, Qt.UserRole, "sheet")
                sheet_item.setData(0, Qt.UserRole + 1, sheet_name)

            wb_item.setExpanded(True)

        self._tree.blockSignals(False)
        self._update_checked_count()

    def _on_tree_item_changed(self, item, column):
        """
        WHAT:
            Handles checkbox changes in the tree. Updates the
            WorkbookManager's selection state. If a workbook item
            is toggled, updates all its child sheets.

        CALLED BY:
            → QTreeWidget.itemChanged signal
        """
        item_type = item.data(0, Qt.UserRole)
        name = item.data(0, Qt.UserRole + 1)
        check_state = item.checkState(0)

        if item_type == "workbook":
            # Only act on explicit Checked/Unchecked — ignore PartiallyChecked
            # (PartiallyChecked is only set programmatically by _update_parent_check_state)
            if check_state == Qt.PartiallyChecked:
                return
            is_checked = check_state == Qt.Checked
            # Toggle all sheets under this workbook
            self._manager.set_workbook_checked(name, is_checked)
            # Update child items visually
            self._tree.blockSignals(True)
            for i in range(item.childCount()):
                child = item.child(i)
                child.setCheckState(
                    0, Qt.Checked if is_checked else Qt.Unchecked,
                )
            self._tree.blockSignals(False)

        elif item_type == "sheet":
            # Get parent workbook name
            is_checked = check_state == Qt.Checked
            parent = item.parent()
            if parent:
                wb_name = parent.data(0, Qt.UserRole + 1)
                self._manager.set_sheet_checked(wb_name, name, is_checked)
                # Update parent tri-state
                self._update_parent_check_state(parent)

        self._update_checked_count()

    def _update_parent_check_state(self, parent_item: QTreeWidgetItem):
        """
        WHAT: Updates a workbook item's check state based on its children.
        CALLED BY: _on_tree_item_changed() → after a sheet toggle.
        """
        self._tree.blockSignals(True)

        checked_count = 0
        total = parent_item.childCount()

        for i in range(total):
            if parent_item.child(i).checkState(0) == Qt.Checked:
                checked_count += 1

        if checked_count == 0:
            parent_item.setCheckState(0, Qt.Unchecked)
        elif checked_count == total:
            parent_item.setCheckState(0, Qt.Checked)
        else:
            parent_item.setCheckState(0, Qt.PartiallyChecked)

        self._tree.blockSignals(False)

    def _update_checked_count(self):
        """
        WHAT: Updates the label showing how many sheets are checked.
        CALLED BY: _rebuild_tree(), _on_tree_item_changed()
        """
        sources = self._manager.get_checked_sources()
        count = len(sources)
        if count == 0:
            self._checked_count_label.setText("")
        else:
            self._checked_count_label.setText(
                f"{count} sheet{'s' if count != 1 else ''} selected"
            )
            self._checked_count_label.setStyleSheet(
                f"color: {COLOR_SUCCESS}; font-weight: bold;"
            )

    # -------------------------------------------------------------------
    # UI State Helpers
    # -------------------------------------------------------------------

    def _set_connected_state(self, connected: bool):
        """
        WHAT: Toggles the UI between connected and disconnected states.
        CALLED BY: _on_connect(), _on_disconnect()
        """
        self._connect_btn.setEnabled(not connected)
        self._disconnect_btn.setEnabled(connected)
        self._refresh_btn.setEnabled(connected)

        if connected:
            self._status_text.setText("Connected to Excel")
            self._status_text.setStyleSheet(
                f"color: {COLOR_SUCCESS}; font-size: {scaled_size(15)}px; "
                f"font-weight: bold;"
            )
        else:
            self._status_text.setText("Not connected")
            self._status_text.setStyleSheet(
                f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(15)}px;"
            )

    # -------------------------------------------------------------------
    # Public Methods
    # -------------------------------------------------------------------

    def check_connection_alive(self) -> bool:
        """
        WHAT:
            Checks if the Excel connection is still alive. If it dropped
            (user closed Excel), updates the UI to disconnected state.

        CALLED BY:
            → gui/main_window.py → Excel monitor timer

        RETURNS:
            bool: True if still connected, False if dropped.
        """
        handler = self._manager.excel_handler
        try:
            is_alive = handler.is_connected
        except Exception:
            is_alive = False

        if self._disconnect_btn.isEnabled() and not is_alive:
            # Connection dropped — update UI
            self._set_connected_state(False)
            self._tree.clear()
            self._manager.clear()
            self._update_checked_count()
            return False
        return is_alive
