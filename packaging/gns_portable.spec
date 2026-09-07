# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules


project_root = Path(SPECPATH).resolve().parent
app_name = "GNS-Portable"
app_icon = project_root / "packaging" / "assets" / "gns-document-seal.ico"

datas = [
    (str(project_root / "src" / "gns_app" / "templates"), "gns_app/templates"),
    (str(project_root / "src" / "gns_app" / "static"), "gns_app/static"),
    (str(project_root / "src" / "gns_app" / "data"), "gns_app/data"),
    (str(project_root / "models"), "models"),
    (str(project_root / "УГНС" / "шаблон ответа одиночный.docx"), "УГНС"),
    (str(project_root / "УГНС" / "шаблон ответа много.docx"), "УГНС"),
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

# OCR/QR и реестры используют нативные либо динамически загружаемые модули.
for package_name in (
    "tesserocr",
    "pypdfium2",
    "pyzbar",
    "pymorphy3_dicts_ru",
    "cloudscraper",
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
    runtime_hooks=[str(project_root / "packaging" / "portable_runtime_hook.py")],
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
    name=app_name,
    icon=str(app_icon),
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=True,
)

# OCR/PDF, Outlook COM и WIA обмениваются данными через stdin/stdout. Для них
# нужен console bootloader, но процесс всегда запускается приложением с
# CREATE_NO_WINDOW, поэтому отдельное окно терминала не появляется.
worker_exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="GNS-Worker",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
)

coll = COLLECT(
    exe,
    worker_exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=app_name,
)
