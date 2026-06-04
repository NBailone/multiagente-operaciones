# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec para ui_app.py
# Incluye los dorsos PDF y el ícono embebidos en el ejecutable.

block_cipher = None

a = Analysis(
    ['ui_app.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # Dorsos PDF
        ('DORSO MIC.pdf', '.'),
        ('DORSO CRT.pdf', '.'),
        ('dorso PE.pdf', '.'),
        # Ícono
        ('icono.ico', '.'),
        # Módulos internos
        ('constants', 'constants'),
        ('utils', 'utils'),
        # Config (si existe en dist)
        ('ui_config.json', '.'),
    ],
    hiddenimports=[
        'customtkinter',
        'openpyxl',
        'xlrd',
        'win32com',
        'win32com.client',
        'win32print',
        'pywintypes',
        'pythoncom',
        'email',
        'imaplib',
        'dotenv',
    ],
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
    name='Multiagente',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # Sin ventana de consola
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icono.ico',
    version=None,
    uac_admin=False,
    uac_uiaccess=False,
)
