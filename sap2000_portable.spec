# -*- mode: python ; coding: utf-8 -*-

hiddenimports = [
    "comtypes",
    "comtypes.client",
    "comtypes.gen",
    "pythoncom",
    "pywintypes",
    "win32gui",
    "win32con",
    "win32ui",
    "win32api",
    "win32timezone",
    "PIL.Image",
    "mouseinfo",
    "pygetwindow",
    "pyscreeze",
    "pytweening",
    "pyrect",
    "pymsgbox",
]

block_cipher = None


a = Analysis(
    ["sap2000_portable.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["xlwings", "matplotlib", "numpy", "pandas", "tkinter"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
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
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="sap2000_capture",
)
