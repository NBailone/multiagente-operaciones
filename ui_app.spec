# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec para Multiagente — MODO ONEFOLDER (portable)

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # Dorsos PDF
        ('DORSO MIC.pdf', '.'),
        ('DORSO CRT.pdf', '.'),
        ('dorso PE.pdf', '.'),
        # Ícono
        ('icono.ico', '.'),
        # Config
        ('ui_config.json', '.'),
        # Módulos internos
        ('constants', 'constants'),
        ('utils', 'utils'),
        ('panels', 'panels'),
        ('engines', 'engines'),
        # Poppler (pdf2image)
        ('poppler', 'poppler'),
        # Assets (iconos, etc.)
        ('assets', 'assets'),
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
        # pbkdf2_hmac vive en _hashlib (OpenSSL); sin él, el exe falla al cifrar
        '_hashlib',
        # win32com.client.DispatchEx importa esto dinámicamente en runtime
        'win32timezone',
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
    name='Sistema_Automatizacion',
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

# COLLECT: onefolder mode — todo dentro de _internal/ (PyInstaller 6.x standard)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Sistema_Automatizacion',
)
