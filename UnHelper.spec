# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


project_root = Path.cwd()
asset_dir = project_root / "assets"

datas = [
    (str(project_root / "UPDATE_HISTORY.txt"), "."),
    (str(project_root / "chromedriver.exe"), "."),
]
token_path = project_root / "bug_report_token.dat"
if token_path.is_file():
    datas.append((str(token_path), "."))
if asset_dir.is_dir():
    for asset in sorted(asset_dir.rglob("*")):
        if asset.is_file():
            datas.append((str(asset), str(asset.parent.relative_to(project_root))))

icon_path = asset_dir / "app-icon.ico"
icon = str(icon_path) if icon_path.is_file() else None

a = Analysis(
    ["App.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=[
        "PySide6.QtCore",
        "PySide6.QtGui",
        "PySide6.QtWidgets",
        "selenium",
        "pythoncom",
        "pywintypes",
        "win32crypt",
        "win32timezone",
        "win32com.client",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="UnHelper",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="UnHelper",
)
