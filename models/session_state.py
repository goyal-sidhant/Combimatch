"""
FILE: models/session_state.py

PURPOSE: Defines the SessionState dataclass — a fully serializable snapshot
         of the application state for save/restore and crash recovery.

CONTAINS:
- SessionState — dataclass holding all state needed to resume a session

DEPENDS ON:
- Nothing from this project (pure data container, serialized to/from JSON
  by session_manager.py which handles the conversion).

USED BY:
- core/session_manager.py → creates SessionState for saving,
                             reads SessionState after loading

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — session persistence data model  | Project skeleton setup (Phase 1A)|
"""

# Group 1: Python standard library
from dataclasses import dataclass, field
from typing import Any, Dict, List


@dataclass
class SessionState:
    """
    WHAT:
        A complete snapshot of the app's state, designed to be serialized
        to JSON for persistence. Contains loaded numbers, finalized
        combinations, search parameters, and undo stack.

    WHY ADDED:
        Session save/restore (N1) requires capturing all state so users
        can resume after crashes, restarts, or lunch breaks. This
        dataclass defines exactly what gets saved.

    ASSUMPTIONS:
        - All fields are JSON-serializable (no Qt objects, no COM references).
        - session_manager.py converts NumberItem/FinalizedCombination objects
          to/from plain dicts during serialization. This dataclass holds the
          intermediate dict form.
        - The "saving" flag in the JSON file indicates an incomplete save
          (crash detection). If present on load, the previous session may
          be recoverable.

    PARAMETERS:
        numbers (List[Dict]): Serialized NumberItem objects.
        finalized (List[Dict]): Serialized FinalizedCombination objects.
        search_params (Dict): Last-used search parameters.
        undo_stack (List[Dict]): Stack of undoable finalization records.
        excel_info (Dict): Last-known Excel connection info (workbook paths).
        next_color_index (int): Which highlight color to assign next.
        next_combo_number (int): Next finalization sequence number.
        timestamp (str): When this state was saved.
    """

    # Serialized NumberItem objects (list of dicts).
    numbers: List[Dict[str, Any]] = field(default_factory=list)

    # Serialized FinalizedCombination objects (list of dicts).
    finalized: List[Dict[str, Any]] = field(default_factory=list)

    # Last-used search parameters (dict form).
    search_params: Dict[str, Any] = field(default_factory=dict)

    # Stack of undoable finalization records (most recent last).
    undo_stack: List[Dict[str, Any]] = field(default_factory=list)

    # Last-known Excel connection info (workbook names/paths).
    excel_info: Dict[str, Any] = field(default_factory=dict)

    # Which highlight color index to assign next (0-19, wraps).
    next_color_index: int = 0

    # Next finalization sequence number (1-based).
    next_combo_number: int = 1

    # ISO format timestamp of when this state was saved.
    timestamp: str = ""
