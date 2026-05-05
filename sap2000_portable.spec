# -*- mode: python ; coding: utf-8 -*-

hiddenimports = [
    "comtypes",
    "comtypes.client",
    "comtypes.gen",
    "sap2000_gui",
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


def _pick_script(scripts, script_name):
    for entry in scripts:
        if entry[0] == script_name:
            return [entry]
    raise ValueError(f"No se encontro el script {script_name!r} en Analysis.scripts")


a = Analysis(
    ["sap2000_portable.py", "sap2000_gui.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["xlwings", "matplotlib", "numpy", "pandas"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
cli_exe = EXE(
    pyz,
    _pick_script(a.scripts, "sap2000_portable"),
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
gui_exe = EXE(
    pyz,
    _pick_script(a.scripts, "sap2000_gui"),
    [],
    exclude_binaries=True,
    name="sap2000_capture_gui",
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
coll = COLLECT(
    cli_exe,
    gui_exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="sap2000_capture",
)
