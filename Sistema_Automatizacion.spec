# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('DORSO MIC.pdf', '.'), ('DORSO CRT.pdf', '.'), ('dorso PE.pdf', '.')],
    hiddenimports=['openpyxl', 'xlrd', 'xlutils', 'xlwt', 'customtkinter', 'PIL', 'requests', 'google.auth', 'googleapiclient', 'google_auth_oauthlib', 'smtplib', 'email.mime.text', 'win32com', 'pythoncom'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
a.datas += Tree('engines\\tesseract', 'engines/tesseract')
a.datas += Tree('engines\\paddleocr', 'engines/paddleocr')
a.datas += Tree('poppler', 'poppler')
a.datas += Tree('constants', 'constants')
a.datas += Tree('utils', 'utils')
a.datas += Tree('panels', 'panels')
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Sistema_Automatizacion',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icono.ico'],
)
