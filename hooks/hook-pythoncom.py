# hooks/hook-pythoncom.py
# ──────────────────────────────────────────────────────────────────────
# PURPOSE: Custom PyInstaller hook for pythoncom.
#          Similar to pywintypes, the pythoncomXX.dll often lives in
#          pywin32_system32/ and is missed by the standard analysis.
#          This DLL is required for COM automation (win32com.client).
#
# WHY NEEDED:
#     The Excel COM connection in readers/excel_reader.py uses
#     win32com.client.Dispatch() which internally loads pythoncom.
#     Without this DLL, the frozen exe crashes when connecting to Excel.
# ──────────────────────────────────────────────────────────────────────

import os
import sys
import glob

hiddenimports = ['pythoncom']
binaries = []

# Search common locations for pythoncomXX.dll
python_dir = os.path.dirname(sys.executable)
search_paths = [
    os.path.join(python_dir, 'Lib', 'site-packages', 'pywin32_system32'),
    os.path.join(python_dir, 'Lib', 'site-packages', 'win32'),
]

for search_path in search_paths:
    for dll in glob.glob(os.path.join(search_path, 'pythoncom*.dll')):
        binaries.append((dll, '.'))
