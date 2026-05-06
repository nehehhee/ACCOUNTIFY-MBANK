# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_all

pdf_datas,   pdf_bins,   pdf_hidden   = collect_all('pdfplumber')
miner_datas, miner_bins, miner_hidden = collect_all('pdfminer')

a = Analysis(
    ['mbank_konvertor.py'],
    pathex=['.'],
    binaries=pdf_bins + miner_bins,
    datas=[
        ('logo.png', '.'),
        ('icon.ico', '.'),
        ('banks',    'banks'),
        ('core',     'core'),
    ] + pdf_datas + miner_datas,
    hiddenimports=[
        'pdfplumber',
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
        'pdfminer.converter',
        'pdfminer.pdfpage',
        'pdfminer.pdfdevice',
        'pdfminer.pdfdocument',
        'pdfminer.pdfinterp',
        'pdfminer.pdfparser',
        'pdfminer.pdftypes',
        'pdfminer.utils',
        'pandas',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.worksheet.table',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
    ] + pdf_hidden + miner_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Bank Converter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.icns',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='Bank Converter',
)

app = BUNDLE(
    coll,
    name='Bank Converter.app',
    icon='icon.icns',
    bundle_identifier='com.bankconverter.app',
    info_plist={
        'NSHighResolutionCapable': True,
        'CFBundleShortVersionString': '1.0.0',
    },
)
