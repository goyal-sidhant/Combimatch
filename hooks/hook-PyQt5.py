# hooks/hook-PyQt5.py
# ──────────────────────────────────────────────────────────────────────
# PURPOSE: Custom PyInstaller hook for PyQt5.
#          Ensures Qt platform plugins (qwindows.dll), image format
#          plugins, and style plugins are collected. Without these,
#          the app either shows a blank window or crashes on launch
#          with "could not find or load the Qt platform plugin 'windows'".
#
# WHY NEEDED:
#     Standard PyInstaller hooks for PyQt5 sometimes miss platform
#     plugins, especially on machines where Qt was installed via pip
#     rather than the official installer.
# ──────────────────────────────────────────────────────────────────────

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

hiddenimports = collect_submodules('PyQt5')
datas = collect_data_files('PyQt5', include_py_files=True)
