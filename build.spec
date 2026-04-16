# build.spec
# ──────────────────────────────────────────────────────────────────────
# PyInstaller spec file for CombiMatch
#
# PURPOSE: Packages the CombiMatch PyQt5 desktop application into a
#          single .exe with all dependencies, including solver.dll.
#
# HOW TO USE:
#     pyinstaller build.spec
#
# OUTPUT:
#     dist/CombiMatch.exe   — standalone executable (no terminal window)
#
# CHANGE LOG:
# | Date       | Change                                 | Why                           |
# |------------|----------------------------------------|-------------------------------|
# | 16-04-2026 | Created — initial PyInstaller spec     | Package app for distribution  |
# ──────────────────────────────────────────────────────────────────────

import sys
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# ── Collect all assets for problematic packages ──────────────────────
pyqt5_datas, pyqt5_binaries, pyqt5_hiddenimports = collect_all('PyQt5')

a = Analysis(
    ['main.py'],                        # ← entry point
    pathex=['.'],
    binaries=[
        *pyqt5_binaries,
        # solver.dll — the compiled C solver, must sit next to the exe
        ('solver.dll', '.'),
    ],
    datas=[
        *pyqt5_datas,
        # userdata folder — created at runtime, but include the empty
        # directory so the exe can locate it relative to itself
        ('userdata', 'userdata'),
    ],
    hiddenimports=[
        # ── PyQt5 essentials ─────────────────────────────────────────
        *pyqt5_hiddenimports,
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'PyQt5.sip',

        # PyQt5 platform plugins (fixes blank window on unknown PCs)
        'PyQt5.Qt',
        'PyQt5.QtPrintSupport',

        # ── pywin32 — needed for Excel COM automation ────────────────
        'win32api',
        'win32con',
        'win32gui',
        'win32ui',
        'win32process',
        'win32security',
        'win32service',
        'pywintypes',
        'winerror',
        *collect_submodules('win32com'),
        *collect_submodules('win32'),

        # ── ctypes — used by single_instance.py and solver_c.py ─────
        'ctypes',
        'ctypes.util',
        'ctypes.wintypes',

        # ── stdlib extras commonly missed ────────────────────────────
        'encodings',
        'encodings.utf_8',
        'encodings.ascii',
        'json',
        'subprocess',

        # ── project modules — explicit list to prevent misses ────────
        'config',
        'config.constants',
        'config.mappings',
        'config.settings',
        'core',
        'core.finalization_manager',
        'core.number_parser',
        'core.parameter_validator',
        'core.session_manager',
        'core.smart_bounds',
        'core.solver_c',
        'core.solver_manager',
        'core.solver_python',
        'core.target_parser',
        'gui',
        'gui.combo_info_panel',
        'gui.dialogs',
        'gui.find_tab',
        'gui.input_panel',
        'gui.main_window',
        'gui.results_panel',
        'gui.settings_tab',
        'gui.source_panel',
        'gui.styles',
        'gui.summary_tab',
        'models',
        'models.combination',
        'models.finalized_combination',
        'models.number_item',
        'models.search_parameters',
        'models.session_state',
        'models.source_tag',
        'readers',
        'readers.excel_monitor',
        'readers.excel_reader',
        'readers.excel_workbook_manager',
        'utils',
        'utils.format_helpers',
        'utils.single_instance',
        'writers',
        'writers.excel_highlighter',
    ],
    hookspath=['./hooks'],              # ← custom hooks folder
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',                      # not needed, saves space
        'matplotlib',
        'numpy',
        'pandas',
        'pytest',
        'unittest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CombiMatch',                  # ← output exe name
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,                           # compresses the exe
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                      # no terminal window (GUI app)
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='CombiMatch.ico',            # ← uncomment when you have an icon
)
