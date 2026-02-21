# -*- mode: python ; coding: utf-8 -*-

import importlib.util
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

project_root = Path(SPECPATH).resolve()
version_path = project_root / "version.py"
spec = importlib.util.spec_from_file_location("release_version", version_path)
if spec is None or spec.loader is None:
    raise RuntimeError(f"version.py を読み込めません: {version_path}")

version_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(version_module)

APP_NAME = version_module.APP_NAME
VERSION = version_module.VERSION
exe_name = f"{APP_NAME}-{VERSION}"

block_cipher = None

added_datas = [
    (str(project_root / "core"), "core"),
    (str(project_root / "services"), "services"),
    (str(project_root / "ui"), "ui"),
    (str(project_root / "utils"), "utils"),
    (str(project_root / "settings.json"), "."),
]

added_datas += collect_data_files("selenium")
added_datas += collect_data_files("webdriver_manager")

hidden_imports = []
hidden_imports += collect_submodules("selenium")
hidden_imports += collect_submodules("webdriver_manager")


a = Analysis(
    ["app.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=added_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name=exe_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
)
