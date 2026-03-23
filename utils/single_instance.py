"""
FILE: utils/single_instance.py

PURPOSE: Prevents multiple instances of CombiMatch from running on the
         same Windows session. Uses a Windows named mutex that is
         per-session (not per-machine) so nComputing nodes can each
         run their own instance.

CONTAINS:
- SingleInstanceGuard — Context manager that acquires/releases a mutex

DEPENDS ON:
- Nothing from this project (uses ctypes for Windows API).

USED BY:
- main.py → wraps app startup in SingleInstanceGuard

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — per-session single instance     | Sub-phase 1C app shell           |
"""

# Group 1: Python standard library
import ctypes
import ctypes.wintypes


# Windows API constants
ERROR_ALREADY_EXISTS = 183


class SingleInstanceGuard:
    """
    WHAT:
        Context manager that creates a Windows named mutex to ensure
        only one instance of CombiMatch runs per Windows session.
        If another instance is already running in the same session,
        is_running returns True.

    WHY ADDED:
        Articles and interns sometimes accidentally launch the app
        twice, causing confusion when both instances connect to Excel.
        The mutex prevents this. It is per-session (Local\\ prefix)
        so different nComputing nodes — which are separate Windows
        sessions on the same physical machine — can each run one
        instance independently.

    CALLED BY:
        → main.py → at startup, before creating QApplication

    ASSUMPTIONS:
        - Windows only. This uses ctypes to call the Windows API.
          On non-Windows systems, the guard always allows running
          (is_running returns False) as a graceful fallback.
        *** ASSUMPTION: The mutex name "Local\\CombiMatch_SingleInstance"
            uses the Local\\ prefix which scopes to the current Windows
            session. This is correct for nComputing where each node is
            a separate session. "Global\\" would be per-machine and
            would block other nodes. ***

    USAGE:
        guard = SingleInstanceGuard()
        if guard.is_running:
            # Another instance exists — show message and exit
            sys.exit(1)
        # ... run the app ...
        guard.release()
    """

    # Mutex name scoped to the current Windows session (Local\\)
    MUTEX_NAME = "Local\\CombiMatch_SingleInstance"

    def __init__(self):
        """
        WHAT: Creates the mutex and checks if another instance exists.
        """
        self._mutex_handle = None
        self._is_running = False

        try:
            # CreateMutexW(lpMutexAttributes, bInitialOwner, lpName)
            kernel32 = ctypes.windll.kernel32
            self._mutex_handle = kernel32.CreateMutexW(
                None,   # Default security
                False,  # Don't take ownership immediately
                self.MUTEX_NAME,
            )

            # If the mutex already existed, another instance is running
            last_error = kernel32.GetLastError()
            if last_error == ERROR_ALREADY_EXISTS:
                self._is_running = True

        except (OSError, AttributeError):
            # Non-Windows system or ctypes issue — allow running
            self._is_running = False

    @property
    def is_running(self) -> bool:
        """
        WHAT: Returns True if another instance is already running.
        CALLED BY: main.py → to decide whether to show error and exit.
        """
        return self._is_running

    def release(self):
        """
        WHAT: Releases the mutex handle when the app exits.
        CALLED BY: main.py → at shutdown.
        """
        if self._mutex_handle is not None:
            try:
                ctypes.windll.kernel32.CloseHandle(self._mutex_handle)
            except (OSError, AttributeError):
                pass
            self._mutex_handle = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.release()
        return False
