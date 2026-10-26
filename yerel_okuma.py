"""Yerel (internetsiz) okuyucu ile PDF -> blok listesi. DENEME: Gemini yolunun yerine değil, yanında.

Okuyucunun kodu "local reading/kod" klasöründe; paketlenmiş exe'de modüller exe'nin içinde.
Model ve kalıp listesi exe'nin yanındaki "modeller/kat_hepsi" ve "kaliplar.txt"ten okunur.
Sonuç Gemini'nin döndürdüğü biçimdedir; app.py aynı _normalize_block'tan geçirir.
Gemini önbelleğine hiçbir şey yazılmaz.
"""
import os
import re
import sys
import time
from pathlib import Path


def _app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _ensure_path():
    if getattr(sys, "frozen", False):
        return
    kod = os.environ.get("YEREL_KOD") or str(_app_dir().parent / "local reading" / "kod")
    if kod not in sys.path:
        sys.path.insert(0, kod)


def read_blocks(pdf_path, log, progress_cb=None):
    """PDF'i yerel modelle okur, Gemini biçiminde ham blok listesi döndürür."""
    _ensure_path()
    import cv2
    import fitz
    import numpy as np

    import cells
    import core
    import gemini_yerine
    import isaretle
    import listeler
    import ocr
    import okuyucu

    model_dir = _app_dir() / "modeller" / "kat_hepsi"
    model = str(model_dir) if (model_dir / "config.json").exists() else okuyucu.MODEL
    moulds = _app_dir() / "kaliplar.txt"
    if moulds.exists():  # yeni eklenen kalıplar exe'nin yanında kalıcı olsun
        listeler.MOULD_FILE = moulds
        listeler.MOULDS[:] = listeler.load_moulds()
    log(f"Yerel model yukleniyor: {model}")
    t0 = time.time()
    eng = ocr.TrOCR(model)
    log(f"  model hazir ({time.time() - t0:.0f} sn). Kalip listesi: {len(listeler.MOULDS)} kalip")

    doc = fitz.open(pdf_path)
    pages = []
    for i, page in enumerate(doc):
        if progress_cb:
            progress_cb(i, len(doc), "page")
        imgs = page.get_images()
        if len(imgs) == 1:  # taranmış sayfa: gömülü görüntüyü kayıpsız al
            data = doc.extract_image(imgs[0][0])["image"]
        else:
            data = page.get_pixmap(dpi=200).tobytes("png")
        img = cv2.imdecode(np.frombuffer(data, np.uint8), cv2.IMREAD_COLOR)
        name = f"sayfa{i + 1}"
        t = time.time()
        try:  # okumadan önce form tipi: Core formu bu programda okunmaz
            cells.extract(img)
        except cells.FormUyusmuyor:
            try:
                core.extract(img)
                log(f"  [sayfa {i + 1}] Core (shotcrete) formu -> bu programda atlandi")
                continue
            except cells.FormUyusmuyor:
                pass  # read_page açık hata verecek
        try:
            pred, conf, notes = okuyucu.read_page(eng, img, name)
        except okuyucu.TabloBulunamadi as e:
            log(f"  [sayfa {i + 1}] OKUNAMADI ({e}) -> bu sayfayi Gemini ile okuyun")
            continue
        flags = isaretle.flags_for({name: pred}, {name: notes}, {name: conf}).get(name, {})
        pages.append((name, pred, flags))
        log(f"  [sayfa {i + 1}/{len(doc)}] {time.time() - t:.0f} sn")

    blocks = gemini_yerine.to_lastpages(pages)["blocks"]
    for b in blocks:
        if b.get("_belirsiz"):
            log(f"  KONTROL ET cube {b.get('cube_no')}: {', '.join(b['_belirsiz'])}")
            # satır işaretleri satırın kendisine: normalize satır düşürünce indeksler kayar
            blok = []
            for u in b["_belirsiz"]:
                m = re.fullmatch(r"rows\[(\d+)\]\.(\w+)", u)
                if m:
                    b["rows"][int(m.group(1))].setdefault("_belirsiz", []).append(m.group(2))
                else:
                    blok.append(u)
            b["_belirsiz"] = blok
    return blocks
