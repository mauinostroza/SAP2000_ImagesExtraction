# -*- mode: python ; coding: utf-8 -*-
#
# Variante CI de sap2000_portable.spec — console=True para capturar
# stdout/stderr en GitHub Actions y evitar bloqueo por MessageBoxW.
# Se usa SOLO en CI; los usuarios finales reciben el EXE de sap2000_portable.spec.

hiddenimports = [
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
    "win32gui",
    "win32con",
    "win32ui",
    "win32api",
    "win32process",
    "PIL.ImageGrab",
    "PIL.Image",
    "PIL.ImageGrab",
    "mouseinfo",
    "pygetwindow",
    "pyscreeze",
    "pytweening",
    "pyrect",
    "pymsgbox",
    "sap_imagenes",
]

block_cipher = None

a = Analysis(
    ["sap2000_gui.py"],
    pathex=[],
    binaries=[],
    datas=[("sap_imagenes.py", ".")],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "numpy", "pandas", "xlwings"],
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
    name="sap2000_capture",
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
