# hooks/hook-win32com.py
# ──────────────────────────────────────────────────────────────────────
# PURPOSE: Custom PyInstaller hook for win32com.
#          Ensures all win32com submodules (especially win32com.client)
#          and the gen_py cache directory are collected properly.
#
# WHY NEEDED:
#     PyInstaller often misses dynamically-imported COM submodules.
#     The Excel COM automation in readers/excel_reader.py uses
#     win32com.client which the standard analysis can't always trace.
# ──────────────────────────────────────────────────────────────────────

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

hiddenimports = collect_submodules('win32com')
datas = collect_data_files('win32com')
