# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['financial_chatbot.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pandas', 'numpy', 'matplotlib', 'matplotlib.pyplot', 'pyttsx3', 'pyttsx3.drivers', 'pyttsx3.drivers.sapi5', 'tkinter', 'reportlab', 'reportlab.pdfgen', 'reportlab.lib', 'PIL', 'yfinance', 'requests', 'openpyxl'],
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
    a.binaries,
    a.datas,
    [],
    name='financial_chatbot',
    debug=False,
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
)
