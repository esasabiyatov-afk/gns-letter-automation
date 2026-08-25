# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules


project_root = Path(SPECPATH).resolve().parent

datas = [
    (str(project_root / "src" / "gns_app" / "templates"), "gns_app/templates"),
    (str(project_root / "src" / "gns_app" / "static"), "gns_app/static"),
    (str(project_root / "src" / "gns_app" / "data"), "gns_app/data"),
    (str(project_root / "models"), "models"),
    (
        str(project_root / "УГНС" / "шаблон ответа одиночный.docx"),
        "УГНС",
    ),
    (
        str(project_root / "УГНС" / "шаблон ответа много.docx"),
        "УГНС",
    ),
]
binaries = []
hiddenimports = [
    "gns_app.main",
    "gns_app.services.pdf_service",
    "gns_app.services.outlook_service",
    "gns_app.services.scanner_service",
    "pythoncom",
    "pywintypes",
    "win32timezone",
    "win32com.client",
]

for package_name in (
    "tesserocr",
    "pypdfium2",
    "pyzbar",
    "pymorphy3_dicts_ru",
):
    package_datas, package_binaries, package_hidden = collect_all(package_name)
    datas += package_datas
    binaries += package_binaries
    hiddenimports += package_hidden

datas += collect_data_files("certifi")
hiddenimports += collect_submodules("win32com")

a = Analysis(
    [str(project_root / "src" / "gns_app" / "launcher.py")],
    pathex=[str(project_root / "src")],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[str(project_root / "packaging" / "office_runtime_hook.py")],
    excludes=["pytest", "py7zr"],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="GNS-Test-Win8.1",
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
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="GNS-Test-Win8.1",
)
