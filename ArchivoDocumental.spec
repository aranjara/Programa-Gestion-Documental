# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

block_cipher = None

PROJECT_ROOT = Path.cwd()

a = Analysis(
    ['app.py'],
    pathex=[str(PROJECT_ROOT)],
    binaries=[],
    datas=[
        ('activos', 'activos'),
    ],
    hiddenimports=[
        'openpyxl',
        'docx',
        'PIL',
        'PIL.Image',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ArchivoDocumental',
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
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ArchivoDocumental',
)
