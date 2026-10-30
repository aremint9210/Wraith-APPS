# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:/Users/aremi/Desktop/REPORT CBM/WRAITH APP/src/app_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('C:/Users/aremi/Desktop/REPORT CBM/WRAITH APP/src', 'src'), ('C:/Users/aremi/Desktop/REPORT CBM/WRAITH APP/templates', 'templates'), ('C:/Users/aremi/Desktop/REPORT CBM/WRAITH APP/src/app_icon.ico', 'src')],
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
    name='WRAITH_APP_SingleFile',
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
    entitlements_file=None,
    icon=['C:/Users/aremi/Desktop/REPORT CBM/WRAITH APP/src/app_icon.ico'],
)
