# PyInstaller spec for CLMRS Data Analyzer.
# Build with:  pyinstaller build/clmrs_analyzer.spec
# (this is run automatically by the GitHub Actions workflow — see
#  .github/workflows/build-windows.yml — you should not need to run it by hand)

import sys
import os

block_cipher = None
project_root = os.path.abspath(os.path.join(os.path.dirname(SPEC), ".."))

a = Analysis(
    [os.path.join(project_root, "main.py")],
    pathex=[project_root],
    binaries=[],
    datas=[],
    hiddenimports=[
        "pandas", "openpyxl", "dateutil", "dateutil.relativedelta",
        "engine", "engine.analysis", "engine.pipeline", "engine.data_loader",
        "engine.column_mapping", "engine.quarters", "engine.excel_export",
        "config", "config.programmes", "gui", "gui.app",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
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
    name="CLMRS_Data_Analyzer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
