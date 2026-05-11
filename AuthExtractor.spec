# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('C:\\00 Pace Auths- FOR COLE\\Auth Radar Logo.png', '.'), ('C:\\00 Pace Auths- FOR COLE\\.env', '.'), ('C:\\00 Pace Auths- FOR COLE\\integrations', 'integrations'), ('C:\\00 Pace Auths- FOR COLE\\extraction', 'extraction'), ('C:\\00 Pace Auths- FOR COLE\\audit', 'audit'), ('C:\\00 Pace Auths- FOR COLE\\services', 'services'), ('C:\\00 Pace Auths- FOR COLE\\review', 'review'), ('C:\\00 Pace Auths- FOR COLE\\db', 'db'), ('C:\\00 Pace Auths- FOR COLE\\config.py', '.'), ('C:\\00 Pace Auths- FOR COLE\\poppler\\poppler-24.08.0\\Library\\bin', 'poppler'), ('C:\\00 Pace Auths- FOR COLE\\credentials.json', '.')]
binaries = []
hiddenimports = ['pdfplumber', 'pdfplumber.page', 'pdfminer', 'pdfminer.pdfparser', 'pdfminer.pdfdocument', 'pdfminer.pdfpage', 'PIL', 'PIL.Image', 'pdf2image', 'pdf2image.pdf2image', 'pytesseract', 'pandas', 'openpyxl', 'google.oauth2.credentials', 'google_auth_oauthlib.flow', 'google.auth.transport.requests', 'googleapiclient.discovery', 'googleapiclient.errors', 'dotenv', 'dropbox', 'dropbox.files', 'dropbox.oauth', 'integrations', 'integrations.dropbox_service', 'extraction', 'extraction.schema', 'extraction.router', 'extraction.pdf_text_extractor', 'extraction.ocr_extractor', 'extraction.structured_extractor', 'extraction.excel_extractor', 'audit', 'audit.logger', 'services', 'services.pdf_unlock_service', 'services.excel_export_service', 'services.page_detection_service', 'fitz', 'fitz._fitz', 'config']
tmp_ret = collect_all('fitz')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['C:\\00 Pace Auths- FOR COLE\\auth_extractor.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
