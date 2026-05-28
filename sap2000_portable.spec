# -*- mode: python ; coding: utf-8 -*-
#
# Build portable del nuevo CLI modular sap_capture.
# Entrada: main.py
# Salida: dist/sap_capture.exe

from pathlib import Path

spec_file = Path(globals().get("SPEC", "sap2000_portable.spec")).resolve()
project_root = Path(globals().get("SPECPATH", spec_file.parent)).resolve()
entry_point = project_root / "main.py"
module_dir = project_root

hiddenimports = [
    "sap2000_gui",
    "comtypes",
    "comtypes.client",
    "comtypes.gen",
    "openpyxl",
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "tkinter.scrolledtext",
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
    excludes=["matplotlib", "numpy", "pandas", "pyautogui", "xlwings"],
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
    name="sap_capture",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
