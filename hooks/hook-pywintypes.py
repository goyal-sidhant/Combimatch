# hooks/hook-pywintypes.py
# ──────────────────────────────────────────────────────────────────────
# PURPOSE: Custom PyInstaller hook for pywintypes.
#          On many pywin32 installs, pywintypesXX.dll is not found
#          automatically because it lives in pywin32_system32/ rather
#          than the standard DLL search path.
#
# WHY NEEDED:
#     Without this hook, the frozen exe may crash at startup with:
#         ImportError: DLL load failed while importing pywintypes
#     This hook locates the DLL and tells PyInstaller to bundle it.
# ──────────────────────────────────────────────────────────────────────

import os
import sys
import glob

# Find pywintypes DLL in the Python environment
hiddenimports = ['pywintypes']
binaries = []

# Search common locations for pywintypesXX.dll
python_dir = os.path.dirname(sys.executable)
search_paths = [
    os.path.join(python_dir, 'Lib', 'site-packages', 'pywin32_system32'),
    os.path.join(python_dir, 'Lib', 'site-packages', 'win32'),
]

for search_path in search_paths:
    for dll in glob.glob(os.path.join(search_path, 'pywintypes*.dll')):
        binaries.append((dll, '.'))
