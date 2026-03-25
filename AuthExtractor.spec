# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\00 Pace Auths- FOR COLE\\auth_extractor.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\00 Pace Auths- FOR COLE\\auth_radar_logo.png', '.'), ('C:\\00 Pace Auths- FOR COLE\\poppler\\poppler-24.08.0\\Library\\bin', 'poppler'), ('C:\\00 Pace Auths- FOR COLE\\credentials.json', '.')],
    hiddenimports=['pdfplumber', 'pdfplumber.page', 'pdfminer', 'pdfminer.pdfparser', 'pdfminer.pdfdocument', 'pdfminer.pdfpage', 'PIL', 'PIL.Image', 'pdf2image', 'pdf2image.pdf2image', 'pytesseract', 'pandas', 'openpyxl', 'google.oauth2.credentials', 'google_auth_oauthlib.flow', 'google.auth.transport.requests', 'googleapiclient.discovery', 'googleapiclient.errors'],
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
    [],
    exclude_binaries=True,
    name='AuthExtractor',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AuthExtractor',
)
