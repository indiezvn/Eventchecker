# -*- mode: python ; coding: utf-8 -*-
import os

from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_submodules

datas = [('Log_checker.py', '.'), ('remote_update.py', '.'), ('Default event + Default Params.xlsx', '.'), ('remote_update_config_v250.json', '.'), ('services_checker/app.py', 'services_checker'), ('services_checker/axml_fallback.py', 'services_checker'), ('services_checker/bundletool-all-1.18.1.jar', 'services_checker'), ('services_checker/my-key.keystore', 'services_checker')]
hiddenimports = ['engineio.async_drivers.threading', 'androguard.core.axml']
datas += collect_data_files('androguard')
hiddenimports += collect_submodules('androguard.core')


a = Analysis(
    ['desktop_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['markupsafe._speedups'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='EventInspector',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    # Keep the spec usable on both Apple Silicon and Intel.  The build script
    # supplies MACOS_TARGET_ARCH when a specific target is requested; otherwise
    # PyInstaller follows the architecture of the host running the build.
    target_arch=os.getenv('MACOS_TARGET_ARCH') or None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['assets/app.icns'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='EventInspector',
)
app = BUNDLE(
    coll,
    name='EventInspector.app',
    icon='assets/app.icns',
    bundle_identifier=None,
)
