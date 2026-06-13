# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec para Multiagente — MODO ONEFOLDER (portable)

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
        # Config
        ('ui_config.json', '.'),
        # Engines OCR (ocr_helper.py + tesseract)
        ('engines', 'engines'),
        # Poppler (pdf2image)
        ('poppler', 'poppler'),
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
        'pytesseract',
        'pdf2image',
        'PIL',
        'cv2',
        'numpy',
        'fitz',
        'procesar_tickets',
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
    console=False,
    disable_windowed_traceback=True,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icono.ico',
    version=None,
    uac_admin=False,
    uac_uiaccess=False,
)

# COLLECT: onefolder mode — datos al lado del exe
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Multiagente',
)
