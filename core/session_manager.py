"""
FILE: core/session_manager.py

PURPOSE: Handles saving and loading application state to/from a JSON file
         in the userdata/ folder. Provides auto-save, crash recovery detection,
         and conversion between live objects and serializable dicts.

CONTAINS:
- SessionManager — Stateless utility class for save/load operations

DEPENDS ON:
- config/settings.py → get_session_file_path()
- models/session_state.py → SessionState
- models/number_item.py → NumberItem
- models/source_tag.py → SourceTag
- models/combination.py → Combination
- models/finalized_combination.py → FinalizedCombination

USED BY:
- gui/main_window.py → auto-save timer, startup restore, close-event save

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 23-03-2026 | Created — session save/restore            | Phase 4C session persistence     |
"""

# Group 1: Python standard library
import json
import os
from datetime import datetime
from typing import Any, Dict, List, Optional

# Group 3: This project's modules
from config.settings import get_session_file_path
from models.session_state import SessionState
from models.number_item import NumberItem
from models.source_tag import SourceTag
from models.combination import Combination
from models.finalized_combination import FinalizedCombination


class SessionManager:
    """
    WHAT:
        Handles all session persistence: converting live objects to JSON-safe
        dicts, writing them to userdata/session.json, and reading them back.
        Also detects crash recovery scenarios (incomplete saves).

    WHY ADDED:
        Users finalize 10+ combinations over an hour. If the app crashes,
        Excel closes, or they need to resume tomorrow, all that work should
        be recoverable from disk automatically.

    CALLED BY:
        → gui/main_window.py → auto-save timer, startup restore, close save

    ASSUMPTIONS:
        - The session file is a single JSON object written atomically
          (write to temp file, then rename). This prevents corrupted saves
          if the app crashes mid-write.
        - Only one session file exists at a time. There is no multi-session
          history — each save overwrites the previous one.
        *** ASSUMPTION: Undo stack is NOT saved across sessions. When the
            user restores a session, they've implicitly accepted that state.
            Undo starts fresh after restore. ***
    """

    # -------------------------------------------------------------------
    # Serialization — Live Objects → Dicts
    # -------------------------------------------------------------------

    @staticmethod
    def _serialize_source_tag(tag: Optional[SourceTag]) -> Optional[Dict[str, Any]]:
        """
        WHAT: Converts a SourceTag to a JSON-safe dict.
        CALLED BY: _serialize_number_item()
        """
        if tag is None:
            return None
        return {
            "workbook_name": tag.workbook_name,
            "sheet_name": tag.sheet_name,
            "cell_address": tag.cell_address,
            "row": tag.row,
            "column": tag.column,
        }

    @staticmethod
    def _serialize_number_item(item: NumberItem) -> Dict[str, Any]:
        """
        WHAT: Converts a NumberItem to a JSON-safe dict.
        CALLED BY: save_session()
        """
        return {
            "value": item.value,
            "index": item.index,
            "source": SessionManager._serialize_source_tag(item.source),
            "is_finalized": item.is_finalized,
            "finalized_color": list(item.finalized_color) if item.finalized_color else None,
            "is_seed": item.is_seed,
            "original_text": item.original_text,
        }

    @staticmethod
    def _serialize_combination(combo: Combination) -> Dict[str, Any]:
        """
        WHAT: Converts a Combination to a JSON-safe dict.
        CALLED BY: _serialize_finalized_combination()
        """
        return {
            "items": [SessionManager._serialize_number_item(item) for item in combo.items],
            "target": combo.target,
        }

    @staticmethod
    def _serialize_finalized_combination(fc: FinalizedCombination) -> Dict[str, Any]:
        """
        WHAT: Converts a FinalizedCombination to a JSON-safe dict.
        CALLED BY: save_session()
        """
        return {
            "combination": SessionManager._serialize_combination(fc.combination),
            "color_rgb": list(fc.color_rgb),
            "color_name": fc.color_name,
            "label": fc.label,
            "combo_number": fc.combo_number,
            "timestamp": fc.timestamp,
        }

    # -------------------------------------------------------------------
    # Deserialization — Dicts → Live Objects
    # -------------------------------------------------------------------

    @staticmethod
    def _deserialize_source_tag(data: Optional[Dict[str, Any]]) -> Optional[SourceTag]:
        """
        WHAT: Converts a dict back to a SourceTag.
        CALLED BY: _deserialize_number_item()
        """
        if data is None:
            return None
        return SourceTag(
            workbook_name=data.get("workbook_name", ""),
            sheet_name=data.get("sheet_name", ""),
            cell_address=data.get("cell_address", ""),
            row=data.get("row"),
            column=data.get("column"),
        )

    @staticmethod
    def _deserialize_number_item(data: Dict[str, Any]) -> NumberItem:
        """
        WHAT: Converts a dict back to a NumberItem.
        CALLED BY: load_session()
        """
        color = data.get("finalized_color")
        if color is not None:
            color = tuple(color)
        return NumberItem(
            value=data.get("value", 0.0),
            index=data.get("index", 0),
            source=SessionManager._deserialize_source_tag(data.get("source")),
            is_finalized=data.get("is_finalized", False),
            finalized_color=color,
            is_seed=data.get("is_seed", False),
            original_text=data.get("original_text", ""),
        )

    @staticmethod
    def _deserialize_combination(data: Dict[str, Any]) -> Combination:
        """
        WHAT: Converts a dict back to a Combination.
        CALLED BY: _deserialize_finalized_combination()
        """
        items = [
            SessionManager._deserialize_number_item(item_data)
            for item_data in data.get("items", [])
        ]
        return Combination(
            items=items,
            target=data.get("target", 0.0),
        )

    @staticmethod
    def _deserialize_finalized_combination(data: Dict[str, Any]) -> FinalizedCombination:
        """
        WHAT: Converts a dict back to a FinalizedCombination.
        CALLED BY: load_session()
        """
        return FinalizedCombination(
            combination=SessionManager._deserialize_combination(data.get("combination", {})),
            color_rgb=tuple(data.get("color_rgb", (0, 0, 0))),
            color_name=data.get("color_name", ""),
            label=data.get("label", ""),
            combo_number=data.get("combo_number", 0),
            timestamp=data.get("timestamp", ""),
        )

    # -------------------------------------------------------------------
    # Save / Load
    # -------------------------------------------------------------------

    @staticmethod
    def save_session(
        items: List[NumberItem],
        finalized_list: List[FinalizedCombination],
        next_color_index: int,
        next_combo_number: int,
        search_params: Dict[str, Any],
    ) -> dict:
        """
        WHAT:
            Saves the current application state to userdata/session.json.
            Uses atomic write (temp file + rename) to prevent corruption
            if the app crashes mid-write.

        WHY ADDED:
            Auto-save every 60 seconds + immediate save on finalization/undo
            ensures minimal data loss on crash.

        CALLED BY:
            → gui/main_window.py → auto-save timer, finalization, close event

        PARAMETERS:
            items (List[NumberItem]): All loaded number items.
            finalized_list (List[FinalizedCombination]): All finalized combos.
            next_color_index (int): Next color to assign.
            next_combo_number (int): Next finalization sequence number.
            search_params (Dict): Last-used search parameters.

        RETURNS:
            dict: {"success": bool, "error": str or None}
        """
        try:
            state = {
                "version": 1,
                "timestamp": datetime.now().isoformat(),
                "numbers": [
                    SessionManager._serialize_number_item(item)
                    for item in items
                ],
                "finalized": [
                    SessionManager._serialize_finalized_combination(fc)
                    for fc in finalized_list
                ],
                "next_color_index": next_color_index,
                "next_combo_number": next_combo_number,
                "search_params": search_params,
            }

            session_path = get_session_file_path()
            temp_path = session_path + ".tmp"

            # Write to temp file first (atomic write pattern)
            with open(temp_path, "w", encoding="utf-8") as f:
                json.dump(state, f, indent=2, ensure_ascii=False)

            # Replace the real file atomically
            # On Windows, os.replace is atomic within the same volume
            os.replace(temp_path, session_path)

            return {"success": True, "error": None}

        except Exception as e:
            return {"success": False, "error": f"Could not save session: {str(e)}"}

    @staticmethod
    def load_session() -> dict:
        """
        WHAT:
            Loads session state from userdata/session.json. Returns the
            deserialized objects ready for the app to restore.

        CALLED BY:
            → gui/main_window.py → on startup, before showing the window

        RETURNS:
            dict: {
                "success": bool,
                "error": str or None,
                "data": {
                    "items": List[NumberItem],
                    "finalized": List[FinalizedCombination],
                    "next_color_index": int,
                    "next_combo_number": int,
                    "search_params": dict,
                    "timestamp": str,
                } or None
            }
        """
        session_path = get_session_file_path()

        if not os.path.exists(session_path):
            return {"success": False, "error": "No session file found", "data": None}

        try:
            with open(session_path, "r", encoding="utf-8") as f:
                state = json.load(f)

            # Deserialize numbers
            items = [
                SessionManager._deserialize_number_item(item_data)
                for item_data in state.get("numbers", [])
            ]

            # Deserialize finalized combinations
            finalized = [
                SessionManager._deserialize_finalized_combination(fc_data)
                for fc_data in state.get("finalized", [])
            ]

            return {
                "success": True,
                "error": None,
                "data": {
                    "items": items,
                    "finalized": finalized,
                    "next_color_index": state.get("next_color_index", 0),
                    "next_combo_number": state.get("next_combo_number", 1),
                    "search_params": state.get("search_params", {}),
                    "timestamp": state.get("timestamp", ""),
                },
            }

        except json.JSONDecodeError as e:
            return {
                "success": False,
                "error": f"Session file is corrupted: {str(e)}",
                "data": None,
            }
        except Exception as e:
            return {
                "success": False,
                "error": f"Could not load session: {str(e)}",
                "data": None,
            }

    @staticmethod
    def has_saved_session() -> bool:
        """
        WHAT: Checks whether a session file exists on disk.
        CALLED BY: gui/main_window.py → to decide whether to offer restore.
        """
        return os.path.exists(get_session_file_path())

    @staticmethod
    def get_session_summary() -> Optional[Dict[str, Any]]:
        """
        WHAT:
            Reads the session file and returns a brief summary (timestamp,
            item count, finalized count) without fully deserializing.
            Used for the restore prompt dialog.

        CALLED BY:
            → gui/main_window.py → to show session info in restore dialog

        RETURNS:
            dict or None: {"timestamp": str, "item_count": int,
                           "finalized_count": int} or None if no session.
        """
        session_path = get_session_file_path()
        if not os.path.exists(session_path):
            return None

        try:
            with open(session_path, "r", encoding="utf-8") as f:
                state = json.load(f)

            return {
                "timestamp": state.get("timestamp", "Unknown"),
                "item_count": len(state.get("numbers", [])),
                "finalized_count": len(state.get("finalized", [])),
            }
        except Exception:
            return None

    @staticmethod
    def delete_session() -> dict:
        """
        WHAT: Deletes the session file from disk.
        CALLED BY: gui/main_window.py → after successful restore or on Clear All.

        RETURNS:
            dict: {"success": bool, "error": str or None}
        """
        session_path = get_session_file_path()
        try:
            if os.path.exists(session_path):
                os.remove(session_path)
            return {"success": True, "error": None}
        except Exception as e:
            return {"success": False, "error": f"Could not delete session file: {str(e)}"}
