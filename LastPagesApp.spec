# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

# Gemini SDK ve bagimliklarini topla
datas_g, binaries_g, hiddenimports_g = collect_all('google.generativeai')
datas_a, binaries_a, hiddenimports_a = collect_all('google.ai.generativelanguage')
try:
    datas_c, binaries_c, hiddenimports_c = collect_all('google.api_core')
except Exception:
    datas_c, binaries_c, hiddenimports_c = ([], [], [])

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries_g + binaries_a + binaries_c,
    datas=datas_g + datas_a + datas_c,
    hiddenimports=(
        ['win32com', 'win32com.client', 'pythoncom', 'pywintypes',
         'fitz', 'PIL', 'PIL.Image']
        + hiddenimports_g + hiddenimports_a + hiddenimports_c
    ),
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
    name='LastPagesApp',
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
