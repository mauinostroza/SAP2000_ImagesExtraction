# -*- mode: python ; coding: utf-8 -*-
#
# Genera un SOLO ejecutable: dist/sap2000_capture.exe
# Sin carpeta ni archivos extra — el .exe se auto-extrae al ejecutarse.
# Punto de entrada directo: sap2000_gui.py

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
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
