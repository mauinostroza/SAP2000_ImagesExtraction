# -*- mode: python ; coding: utf-8 -*-
#
# Variante CI del build del CLI modular sap_capture.
# console=True para que stdout/stderr aparezcan en GitHub Actions.

from pathlib import Path

project_root = Path(__file__).resolve().parent
entry_point = project_root / "main.py"
module_dir = project_root

hiddenimports = [
    "comtypes",
    "comtypes.client",
    "comtypes.gen",
    "openpyxl",
    "pythoncom",
    "pywintypes",
    "win32con",
    "win32gui",
    "win32ui",
    "PIL.Image",
]

block_cipher = None

a = Analysis(
    [str(entry_point)],
    pathex=[str(module_dir)],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "numpy", "pandas", "pyautogui", "tkinter", "xlwings"],
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
    name="sap_capture_ci",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
