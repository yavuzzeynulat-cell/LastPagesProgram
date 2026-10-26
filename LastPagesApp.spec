# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

# Yerel AI okuyucu: kodu "local reading/kod"da; hizli model (ONNX, ~21 MB) ve kalip listesi exe'ye gomulur.
# torch/transformers GEREKMEZ (sadece onnxruntime + opencv).
KOD = r'..\local reading\kod'
MODEL = r'..\local reading\modeller\pp_hepsi_onnx'
yerel_moduller = ['yerel_okuma', 'okuyucu', 'ocr', 'cells', 'core', 'grid', 'zincir', 'bicim', 'kontrol',
                  'isaretle', 'listeler', 'gemini_yerine', 'degerlendir', 'cevap_anahtari', 'cv2', 'numpy',
                  'onnxruntime']
datas_y = [(MODEL + r'\model.onnx', 'modeller/pp_hepsi_onnx'), (MODEL + r'\chars.json', 'modeller/pp_hepsi_onnx'),
           (KOD + r'\kaliplar.txt', '.')]
datas_o, binaries_o, hiddenimports_o = collect_all('onnxruntime')
# onnxruntime.tools/transformers alt paketleri numba+scipy (~160 MB) cekiyordu; okuyucu sadece capi kullanir
hiddenimports_o = [m for m in hiddenimports_o if not m.startswith(('onnxruntime.tools', 'onnxruntime.transformers', 'onnxruntime.quantization'))]

# Gemini SDK ve bagimliklarini topla
datas_g, binaries_g, hiddenimports_g = collect_all('google.generativeai')
datas_a, binaries_a, hiddenimports_a = collect_all('google.ai.generativelanguage')
try:
    datas_c, binaries_c, hiddenimports_c = collect_all('google.api_core')
except Exception:
    datas_c, binaries_c, hiddenimports_c = ([], [], [])

a = Analysis(
    ['app.py'],
    pathex=[KOD],
    binaries=binaries_g + binaries_a + binaries_c + binaries_o,
    datas=datas_g + datas_a + datas_c + datas_o + datas_y,
    hiddenimports=(
        ['win32com', 'win32com.client', 'pythoncom', 'pywintypes',
         'fitz', 'PIL', 'PIL.Image']
        + hiddenimports_g + hiddenimports_a + hiddenimports_c + hiddenimports_o + yerel_moduller
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'transformers', 'tensorflow', 'jax', 'matplotlib', 'pandas', 'IPython',
              'numba', 'llvmlite', 'scipy', 'onnxruntime.tools', 'onnxruntime.transformers', 'onnxruntime.quantization'],
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
