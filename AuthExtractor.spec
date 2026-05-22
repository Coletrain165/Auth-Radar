# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('C:\\Auth Radar\\Auth Radar Logo.png', '.'), ('C:\\Auth Radar\\.env', '.'), ('C:\\Auth Radar\\integrations', 'integrations'), ('C:\\Auth Radar\\extraction', 'extraction'), ('C:\\Auth Radar\\audit', 'audit'), ('C:\\Auth Radar\\services', 'services'), ('C:\\Auth Radar\\review', 'review'), ('C:\\Auth Radar\\db', 'db'), ('C:\\Auth Radar\\config.py', '.'), ('C:\\Auth Radar\\poppler\\poppler-24.08.0\\Library\\bin', 'poppler'), ('C:\\Auth Radar\\credentials.json', '.')]
binaries = []
hiddenimports = ['pdfplumber', 'pdfplumber.page', 'pdfminer', 'pdfminer.pdfparser', 'pdfminer.pdfdocument', 'pdfminer.pdfpage', 'PIL', 'PIL.Image', 'pdf2image', 'pdf2image.pdf2image', 'pytesseract', 'pandas', 'openpyxl', 'google.oauth2.credentials', 'google_auth_oauthlib.flow', 'google.auth.transport.requests', 'googleapiclient.discovery', 'googleapiclient.errors', 'dotenv', 'dropbox', 'dropbox.files', 'dropbox.oauth', 'integrations', 'integrations.dropbox_service', 'extraction', 'extraction.schema', 'extraction.router', 'extraction.pdf_text_extractor', 'extraction.ocr_extractor', 'extraction.structured_extractor', 'extraction.excel_extractor', 'audit', 'audit.logger', 'services', 'services.pdf_unlock_service', 'services.excel_export_service', 'services.page_detection_service', 'fitz', 'fitz._fitz', 'config']
tmp_ret = collect_all('fitz')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['C:\\Auth Radar\\auth_extractor.py'],
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
