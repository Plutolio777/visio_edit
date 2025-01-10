# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['window.py'],
    pathex=[],
    binaries=[
        ("C:\\Users\\liuyijun\\AppData\\Local\\Programs\\Python\\Python311\\python3.dll", "."),
        ("C:\\Users\\liuyijun\\AppData\\Local\\Programs\\Python\\Python311\\python311.dll", ".")
    ],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=True,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [('v', None, 'OPTION')],
    name='window',
    debug=True,
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
    noarchive=True
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    a.scripts,
    a.zipfiles,
    a.pure,
    name='dist',
)