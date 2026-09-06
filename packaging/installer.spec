# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the standalone MRSI DNT installer.

Build it with ``python packaging/build.py installer``. This is the small
app end users download first; it fetches a release from GitHub and unpacks
the tool itself. Single file on both platforms, since there is no in-place
updater to worry about.
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(SPECPATH).parent
ASSETS = PROJECT_ROOT / "assets" / "images"

IS_MACOS = sys.platform == "darwin"

APP_NAME = "MRSI_DNT_Installer"
BUNDLE_ID = "ca.mcmaster.mrsi.dnt.installer"

icon = str(ASSETS / ("logo.icns" if IS_MACOS else "logo.ico"))

a = Analysis(
    [str(PROJECT_ROOT / "src" / "utils" / "installer" / "installer.py")],
    pathex=[str(PROJECT_ROOT / "src")],
    binaries=[],
    datas=[(str(ASSETS), "assets/images")],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=(
        str(PROJECT_ROOT / "packaging" / "entitlements.plist") if IS_MACOS else None
    ),
    icon=icon,
)

if IS_MACOS:
    app = BUNDLE(
        exe,
        name=f"{APP_NAME}.app",
        icon=icon,
        bundle_identifier=BUNDLE_ID,
    )
