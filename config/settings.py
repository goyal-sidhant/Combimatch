"""
FILE: config/settings.py

PURPOSE: Runtime paths, directory locations, and environment-specific
         configuration. Unlike constants.py (which holds fixed values),
         this file holds values that may change based on where the app
         is installed or how it's deployed.

CONTAINS:
- get_app_directory()        — Returns the folder where main.py lives
- get_userdata_directory()   — Returns the userdata/ folder path (creates if missing)
- get_session_file_path()    — Returns the full path to the session save file
- get_dll_path()             — Returns the expected path to solver.dll

DEPENDS ON:
- Nothing from this project (uses only Python standard library).

USED BY:
- core/session_manager.py  → get_session_file_path(), get_userdata_directory()
- core/solver_manager.py   → get_dll_path()
- main.py                  → get_app_directory()

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — paths and directory helpers     | Project skeleton setup (Phase 1A)|
"""

# Group 1: Python standard library
import os
import sys


def get_app_directory() -> str:
    """
    WHAT:
        Returns the absolute path to the folder containing main.py.
        This is the root of the application, used as the base for all
        relative paths (userdata/, csolver/, solver.dll).

    WHY ADDED:
        The app may be launched from different working directories
        (e.g., via a desktop shortcut). Using __file__-relative paths
        ensures consistency regardless of where the shortcut points.

    CALLED BY:
        → get_userdata_directory()
        → get_dll_path()
        → main.py (if needed for working directory setup)

    CALLS:
        → os.path.dirname(), os.path.abspath()

    EDGE CASES HANDLED:
        - Frozen executable (PyInstaller): uses sys.executable directory
        - Normal Python run: uses this file's grandparent (config/ → project root)

    ASSUMPTIONS:
        - This file lives at config/settings.py, one level below project root.

    RETURNS:
        str: Absolute path to the project root directory.
    """
    # When running as a frozen executable (e.g., via PyInstaller),
    # sys.executable points to the .exe, so use its directory.
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)

    # Normal Python run: this file is at config/settings.py,
    # so go up one level to reach the project root.
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_userdata_directory() -> str:
    """
    WHAT:
        Returns the path to the userdata/ subfolder within the app directory.
        Creates the folder if it doesn't exist.

    WHY ADDED:
        Session files and logs are stored in userdata/ to keep them
        separate from source code. This folder is git-ignored.

    CALLED BY:
        → get_session_file_path()
        → core/session_manager.py

    CALLS:
        → get_app_directory()
        → os.makedirs()

    EDGE CASES HANDLED:
        - Folder doesn't exist yet: created automatically.
        - Folder already exists: os.makedirs with exist_ok=True, no error.

    ASSUMPTIONS:
        - The app has write permission to its own directory.
          On nComputing setups, the app runs from a shared location
          but each node has its own Windows session with write access.

    RETURNS:
        str: Absolute path to the userdata/ directory.
    """
    userdata_path = os.path.join(get_app_directory(), "userdata")
    os.makedirs(userdata_path, exist_ok=True)
    return userdata_path


def get_session_file_path() -> str:
    """
    WHAT:
        Returns the full path to the session save file (session.json).

    WHY ADDED:
        Centralized path so session_manager.py doesn't hardcode filenames.

    CALLED BY:
        → core/session_manager.py → save_session(), load_session()

    CALLS:
        → get_userdata_directory()

    RETURNS:
        str: Absolute path to userdata/session.json.
    """
    return os.path.join(get_userdata_directory(), "session.json")


def get_dll_path() -> str:
    """
    WHAT:
        Returns the expected path to solver.dll (the compiled C solver).

    WHY ADDED:
        The DLL is tracked in git and lives in the project root.
        solver_manager.py checks this path at startup.

    CALLED BY:
        → core/solver_manager.py → checks if DLL exists, loads it

    CALLS:
        → get_app_directory()

    RETURNS:
        str: Absolute path to solver.dll in the project root.
    """
    return os.path.join(get_app_directory(), "solver.dll")
