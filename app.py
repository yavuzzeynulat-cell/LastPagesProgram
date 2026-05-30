"""
LastPagesApp - Cube block creator GUI (self-contained).

Tek dosya: writer + GUI + blok verisi.
PyInstaller ile .exe yapilir, Python kurmaya gerek yok.
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
import urllib.request
from copy import copy
from datetime import date, datetime, timedelta

import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox

import openpyxl
from openpyxl.styles import Alignment, Font


__version__ = "0.1.8"
GITHUB_REPO = "yavuzzeynulat-cell/LastPagesProgram"
RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
UPDATE_ASSET_NAME = "LastPagesApp.exe"


# ---------------- BLOCK DATA (inline; ileride GitHub'tan cekilebilir) ----------------

BLOCKS = [
    {
        "cube_no": 713,
        "sample_mark": "G26-CON-0759",
        "supplier": "S2A BP",
        "cmd_code": "7015",
        "site_location": "KM 8+166 Bridge B0080 Pier Lift 14 at Pier 57 on Bridge B0080",
        "section": "2A",
        "batch_ticket": 15059,
        "c_grade": "C35/45",
        "sampling_date": "26.05.26",
        "sampled_by": "B.I",
        "rows": [
            {"mould": 5,   "row_type": "7d",  "age": 7},
            {"mould": 69,  "row_type": "7d",  "age": 7},
            {"mould": 80,  "row_type": "cmd", "age": 7},
            {"mould": 132, "row_type": "28d", "age": 28},
            {"mould": 113, "row_type": "28d", "age": 28},
            {"mould": 96,  "row_type": "28d", "age": 28},
        ],
    },
    {
        "cube_no": 714,
        "sample_mark": "G26-CON-0761",
        "supplier": "S2A BP",
        "cmd_code": "7008",
        "site_location": "KM 8+166 Bridge B0080 Precast Girders - L2G on Bridge B0080",
        "section": "2A",
        "batch_ticket": 15062,
        "c_grade": "C40/50",
        "sampling_date": "26.05.26",
        "sampled_by": "S.B",
        "rows": [
            {"mould": 171, "row_type": "7d",   "age": 7},
            {"mould": 92,  "row_type": "7d",   "age": 7},
            {"mould": 117, "row_type": "cmd",  "age": 7},
            {"mould": 2,   "row_type": "28d",  "age": 28},
            {"mould": 16,  "row_type": "28d",  "age": 28},
            {"mould": 186, "row_type": "28d",  "age": 28},
            {"mould": 48,  "row_type": "wp",   "age": "WP"},
            {"mould": 44,  "row_type": "wp",   "age": "WP"},
            {"mould": 31,  "row_type": "wp",   "age": "WP"},
            {"mould": 60,  "row_type": "site", "age": 2, "testing_date": "28.05.26"},
            {"mould": 8,   "row_type": "site"},
            {"mould": 121, "row_type": "site"},
            {"mould": None,"row_type": "ft",   "age": "F/T", "ft_note": "(15 Adet)"},
        ],
    },
    {
        "cube_no": 715,
        "sample_mark": "G26-CON-0762",
        "supplier": "S2A BP",
        "cmd_code": "7302",
        "site_location": "KM 5+050 Bridge B0051 Pile No.60,57 at Pier S4R on Bridge B0051",
        "section": "2A",
        "batch_ticket": 15070,
        "c_grade": "C30/37",
        "sampling_date": "26.05.26",
        "sampled_by": "S.B",
        "rows": [
            {"mould": 139, "row_type": "7d",  "age": 7},
            {"mould": 90,  "row_type": "7d",  "age": 7},
            {"mould": 157, "row_type": "cmd", "age": 7},
            {"mould": 136, "row_type": "28d", "age": 28},
            {"mould": 71,  "row_type": "28d", "age": 28},
            {"mould": 48,  "row_type": "28d", "age": 28},
            {"mould": 180, "row_type": "wp",  "age": "WP"},
            {"mould": 29,  "row_type": "wp",  "age": "WP"},
            {"mould": 130, "row_type": "wp",  "age": "WP"},
        ],
    },
    {
        "cube_no": 716,
        "sample_mark": "G26-CON-0763",
        "supplier": "S2A BP",
        "cmd_code": "7410",
        "site_location": "KM:35+935 - KM 37+960 Irrigation Channel Relocation Works Radiovce Bistrica Open Irrigation Channel",
        "section": 1,
        "batch_ticket": 15080,
        "c_grade": "C25/30",
        "sampling_date": "26.05.26",
        "sampled_by": "Y.A",
        "rows": [
            {"mould": 93,  "row_type": "7d",  "age": 7},
            {"mould": 138, "row_type": "7d",  "age": 7},
            {"mould": 24,  "row_type": "cmd", "age": 7},
            {"mould": 75,  "row_type": "28d", "age": 28},
            {"mould": 63,  "row_type": "28d", "age": 28},
            {"mould": 74,  "row_type": "28d", "age": 28},
        ],
    },
]


# ---------------- WRITER ----------------

COL = {
    'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'F': 6, 'G': 7, 'H': 8,
    'I': 9, 'J': 10, 'K': 11, 'L': 12, 'M': 13, 'N': 14, 'O': 15,
    'P': 16, 'Q': 17,
}
WHOLE_BLOCK_MERGES = ['A', 'B', 'E', 'F', 'H', 'J', 'P']  # Q is per-row (site rows write 'Site' there)
DATE_COLS = ['I', 'K']
DATE_FORMAT = 'DD/MM/YYYY'


def parse_date(s):
    if isinstance(s, datetime):
        return s.date()
    if isinstance(s, date):
        return s
    s = str(s).replace('/', '.').replace('-', '.')
    parts = s.split('.')
    d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
    if y < 100:
        y += 2000
    return date(y, m, d)


def compute_testing_date(sampling, row_type, custom):
    if custom:
        return parse_date(custom)
    days = {'7d': 7, 'cmd': 7, '28d': 28, 'wp': 28, 'ft': 28,
            '1day': 1, '2day': 2}.get(row_type)
    return sampling + timedelta(days=days) if days else None


def find_last_data_row(ws):
    for r in range(ws.max_row, 0, -1):
        for col in (1, 2, 3):
            if ws.cell(row=r, column=col).value not in (None, ''):
                return r
    return 0


def find_donor_block(ws, before_row):
    end = before_row - 1
    while end > 0:
        for col in (1, 2, 3):
            if ws.cell(row=end, column=col).value not in (None, ''):
                break
        else:
            end -= 1
            continue
        break
    if end <= 0:
        return None
    start = end
    for r in range(end, 0, -1):
        if ws.cell(row=r, column=1).value not in (None, ''):
            start = r
            for rng in ws.merged_cells.ranges:
                if rng.min_col == 1 and rng.max_col == 1 and rng.min_row == r:
                    end = max(end, rng.max_row)
                    break
            break
    return (start, end)


def get_donor_styles(ws, start, end):
    n = end - start + 1
    styles = {}
    for col_letter, col_idx in COL.items():
        top = ws.cell(row=start, column=col_idx)
        mid = ws.cell(row=start + n // 2, column=col_idx)
        bot = ws.cell(row=end, column=col_idx)
        styles[col_letter] = {'top': top, 'mid': mid, 'bot': bot}
    return styles


def get_donor_formula(ws, start, end, col_letter):
    col_idx = COL[col_letter]
    for r in range(start, end + 1):
        val = ws.cell(row=r, column=col_idx).value
        if isinstance(val, str) and val.startswith('='):
            return (val, r)
    return (None, None)


def shift_formula(formula, src_row, dst_row):
    diff = dst_row - src_row
    if diff == 0:
        return formula
    return re.sub(r'(\$?[A-Z]+\$?)(\d+)',
                  lambda m: f"{m.group(1)}{int(m.group(2)) + diff}",
                  formula)


def copy_style(src_cell, dst_cell):
    # Copy the StyleArray INDEX directly instead of duplicating each
    # style component (font/fill/border/alignment). Otherwise openpyxl
    # creates a fresh style entry per cell -> styles.xml explodes ->
    # Excel becomes very slow on big sheets.
    if not src_cell.has_style:
        return
    dst_cell._style = copy(src_cell._style)


def style_position(i, n):
    if i == 0:
        return 'top'
    if i == n - 1:
        return 'bot'
    return 'mid'


def unmerge_overlapping(ws, start_row, end_row):
    to_remove = [str(r) for r in list(ws.merged_cells.ranges)
                 if r.min_row <= end_row and r.max_row >= start_row]
    for r in to_remove:
        ws.unmerge_cells(r)


def write_block(ws, block, start_row, donor_styles, donor_formulas):
    rows = block['rows']
    n = len(rows)
    end_row = start_row + n - 1
    unmerge_overlapping(ws, start_row, end_row)
    sampling = parse_date(block['sampling_date'])
    cmd_idx = next((i for i, r in enumerate(rows) if r.get('row_type') == 'cmd'), None)

    batch_tickets = block.get('batch_tickets')
    if not batch_tickets:
        bt = block.get('batch_ticket')
        batch_tickets = [{'ticket': bt, 'rows': [0, n - 1]}] if bt is not None else []

    for i in range(n):
        r = start_row + i
        pos = style_position(i, n)
        for col_letter, col_idx in COL.items():
            copy_style(donor_styles[col_letter][pos], ws.cell(row=r, column=col_idx))
        for dc in DATE_COLS:
            ws.cell(row=r, column=COL[dc]).number_format = DATE_FORMAT

    if n > 1:
        for col_letter in WHOLE_BLOCK_MERGES:
            ws.merge_cells(start_row=start_row, end_row=end_row,
                           start_column=COL[col_letter], end_column=COL[col_letter])

    # D column: SINGLE merge across whole block with multi-line content
    # ("S2A BP\nCMD-XXXX"). Earlier we split into two merges, but the user
    # confirmed it should be one cell with two lines.
    if n > 1:
        ws.merge_cells(start_row=start_row, end_row=end_row,
                       start_column=COL['D'], end_column=COL['D'])

    for bt in batch_tickets:
        a, b = bt['rows']
        if b > a:
            ws.merge_cells(start_row=start_row + a, end_row=start_row + b,
                           start_column=COL['G'], end_column=COL['G'])

    ws.cell(row=start_row, column=COL['A'], value=block['cube_no'])
    ws.cell(row=start_row, column=COL['B'], value=block['sample_mark'])
    # D column: single merged cell with multi-line content (supplier + CMD-XXX)
    d_value = block.get('supplier', 'S2A BP')
    if block.get('cmd_code'):
        d_value = f"{d_value}\nCMD-{block['cmd_code']}"
    d_cell = ws.cell(row=start_row, column=COL['D'])
    d_cell.value = d_value
    d_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    e_cell = ws.cell(row=start_row, column=COL['E'], value=block.get('site_location', ''))
    # E column text is RED for newly added blocks so user can spot them at a glance
    src_font = e_cell.font
    e_cell.font = Font(
        name=src_font.name, size=src_font.size,
        bold=src_font.bold, italic=src_font.italic,
        family=src_font.family,
        color='FFFF0000',
    )
    if block.get('section') is not None:
        ws.cell(row=start_row, column=COL['F'], value=block['section'])
    for bt in batch_tickets:
        ws.cell(row=start_row + bt['rows'][0], column=COL['G'], value=bt['ticket'])
    ws.cell(row=start_row, column=COL['H'], value=block.get('c_grade', ''))
    ws.cell(row=start_row, column=COL['J'], value=block.get('sampled_by', ''))

    for i, row in enumerate(rows):
        r = start_row + i
        if row.get('mould') is not None:
            ws.cell(row=r, column=COL['C'], value=row['mould'])
        ws.cell(row=r, column=COL['I'], value=sampling)
        td = compute_testing_date(sampling, row.get('row_type'), row.get('testing_date'))
        if td is not None:
            ws.cell(row=r, column=COL['K'], value=td)
        # Age (L): formula =K-I when both sampling + testing date are real numbers.
        # For WP / F-T / Site without date, write the literal label.
        age_val = row.get('age')
        if td is not None and isinstance(age_val, (int, float)):
            ws.cell(row=r, column=COL['L'], value=f"=K{r}-I{r}")
        elif age_val is not None:
            ws.cell(row=r, column=COL['L'], value=age_val)
        if row.get('row_type') == 'ft' and row.get('ft_note'):
            ws.cell(row=r, column=COL['N'], value=row['ft_note'])
        if row.get('row_type') == 'site':
            ws.cell(row=r, column=COL['Q'], value='Site')
        f, src_r = donor_formulas.get('O', (None, None))
        if f:
            ws.cell(row=r, column=COL['O'], value=shift_formula(f, src_r, r))

    return end_row + 1


def _excel_recalc_and_save(path, log):
    """Open the file in Excel via COM, force full recalc, save, close.
    This restores Excel's formula calculation cache (openpyxl strips it),
    making the file open at original speed instead of recalculating
    25k+ rows on every open."""
    try:
        import pythoncom
        from win32com.client import DispatchEx
    except ImportError:
        log("Uyari: pywin32 yok, recalc adimi atlandi (dosya yine yazildi).")
        return
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        wb = excel.Workbooks.Open(os.path.abspath(path))
        try:
            wb.Application.CalculateFull()
            wb.Save()
        finally:
            wb.Close(SaveChanges=False)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def run_writer(excel_path, log):
    log(f"Excel aciliyor: {excel_path}")
    wb = openpyxl.load_workbook(excel_path)
    if 'Concrete' not in wb.sheetnames:
        raise RuntimeError(f"'Concrete' sheet bulunamadi. Mevcut: {wb.sheetnames}")
    ws = wb['Concrete']

    start_row = find_last_data_row(ws) + 1
    log(f"Yazma baslangici: satir {start_row}")

    donor = find_donor_block(ws, start_row)
    if not donor:
        raise RuntimeError("Donor block bulunamadi - Concrete sheet bos mu?")
    donor_start, donor_end = donor
    log(f"Donor block: satir {donor_start}-{donor_end}")

    donor_styles = get_donor_styles(ws, donor_start, donor_end)
    donor_formulas = {'O': get_donor_formula(ws, donor_start, donor_end, 'O')}
    if donor_formulas['O'][0]:
        log(f"Donor O formul: {donor_formulas['O'][0]} (@ satir {donor_formulas['O'][1]})")

    next_row = start_row
    for block in BLOCKS:
        log(f"  -> cube {block['cube_no']} @ satir {next_row} ({len(block['rows'])} satir)")
        next_row = write_block(ws, block, next_row, donor_styles, donor_formulas)

    log(f"Kaydediliyor: {excel_path}")
    wb.save(excel_path)
    log("Excel ile recalc + kaydet (dosyanin hizini geri getirir)...")
    try:
        _excel_recalc_and_save(excel_path, log)
        log("Recalc tamam - dosya artik hizli aciliyor.")
    except Exception as e:
        log(f"Uyari: recalc basarisiz ({e}). Veri yazildi ama ilk acilis biraz yavas olabilir.")
    log("Tamamlandi.")


# ---------------- AUTO-UPDATER ----------------

def _parse_version(s):
    try:
        return tuple(int(p) for p in s.lstrip('v').strip().split('.'))
    except (ValueError, AttributeError):
        return None


def _is_newer(latest, current):
    a, b = _parse_version(latest), _parse_version(current)
    return a is not None and b is not None and a > b


def _get_latest_release():
    """Return (tag_no_v, asset_url) or None on any failure."""
    try:
        req = urllib.request.Request(
            RELEASES_API,
            headers={'User-Agent': f'LastPagesApp/{__version__}'},
        )
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.load(resp)
        tag = data.get('tag_name', '').lstrip('v').strip()
        if not tag:
            return None
        for a in data.get('assets', []):
            if a.get('name') == UPDATE_ASSET_NAME:
                return (tag, a.get('browser_download_url'))
        return None
    except Exception:
        return None


def _download_file(url, dest_path, progress_cb=None):
    """Download to dest_path; raise if size mismatch (Content-Length vs written)."""
    req = urllib.request.Request(url, headers={'User-Agent': f'LastPagesApp/{__version__}'})
    with urllib.request.urlopen(req, timeout=60) as resp:
        total = int(resp.headers.get('Content-Length') or 0)
        downloaded = 0
        with open(dest_path, 'wb') as f:
            while True:
                chunk = resp.read(64 * 1024)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)
                if progress_cb:
                    progress_cb(downloaded, total)
    actual = os.path.getsize(dest_path)
    if total and actual != total:
        raise IOError(f"Indirme eksik: {actual} / {total} bayt")
    if actual < 1024 * 1024:  # exe should be >1MB
        raise IOError(f"Indirilen dosya cok kucuk: {actual} bayt (bozuk olabilir)")


PS_UPDATER_TEMPLATE = r'''$ErrorActionPreference = "Stop"
$log = "__LOG__"
function Log($msg) { Add-Content -Path $log -Value "$(Get-Date -Format o) $msg" }
Log "updater started, waiting for pid __PID__"
$waited = 0
while (Get-Process -Id __PID__ -ErrorAction SilentlyContinue) {
    Start-Sleep -Milliseconds 500
    $waited += 0.5
    if ($waited -gt 20) { Log "timeout waiting for app"; break }
}
Start-Sleep -Seconds 2  # let Windows fully release file handles
$newSize = (Get-Item -LiteralPath "__NEW_EXE__").Length
Log "new exe size: $newSize bytes"
$copied = $false
for ($i = 1; $i -le 5; $i++) {
    try {
        Copy-Item -LiteralPath "__NEW_EXE__" -Destination "__TARGET_EXE__" -Force
        $destSize = (Get-Item -LiteralPath "__TARGET_EXE__").Length
        if ($destSize -ne $newSize) {
            throw "size mismatch after copy: $destSize vs $newSize"
        }
        Log "copied (attempt $i): $destSize bytes -> __TARGET_EXE__"
        $copied = $true
        break
    } catch {
        Log "copy attempt $i failed: $_"
        Start-Sleep -Seconds 2
    }
}
if (-not $copied) {
    Log "ABORT: could not replace exe after 5 attempts; restarting OLD version"
    Start-Process -FilePath "__TARGET_EXE__"
    exit 1
}
try {
    Start-Process -FilePath "__TARGET_EXE__"
    Log "restarted"
} catch {
    Log "ERROR starting new exe: $_"
}
'''


def _write_powershell_updater(temp_dir, new_exe, target_exe, log_path):
    ps_path = os.path.join(temp_dir, 'update.ps1')
    script = (PS_UPDATER_TEMPLATE
              .replace('__LOG__', log_path)
              .replace('__PID__', str(os.getpid()))
              .replace('__NEW_EXE__', new_exe)
              .replace('__TARGET_EXE__', target_exe))
    with open(ps_path, 'w', encoding='utf-8') as f:
        f.write(script)
    return ps_path


def _launch_powershell_updater(ps_path):
    # CREATE_NO_WINDOW = 0x08000000 — helper runs silently
    subprocess.Popen(
        ['powershell', '-ExecutionPolicy', 'Bypass',
         '-WindowStyle', 'Hidden', '-File', ps_path],
        creationflags=0x08000000,
    )


def _do_update(root, asset_url, new_tag):
    """Tk thread: download into temp, launch helper, exit app."""
    target_exe = sys.executable if getattr(sys, 'frozen', False) else None
    if not target_exe or not target_exe.lower().endswith('.exe'):
        messagebox.showinfo("Guncelleme",
            "Otomatik guncelleme sadece .exe surumde calisir.\n"
            "Mevcut versiyonla devam ediliyor.")
        return

    temp_dir = os.path.join(tempfile.gettempdir(), 'lastpages_update')
    try:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
        os.makedirs(temp_dir, exist_ok=True)
    except Exception as e:
        messagebox.showerror("Guncelleme hatasi", f"Temp klasor olusturulamadi: {e}")
        return

    new_exe = os.path.join(temp_dir, UPDATE_ASSET_NAME)
    log_path = os.path.join(temp_dir, 'update.log')

    win = tk.Toplevel(root)
    win.title("Guncelleniyor")
    win.geometry("420x110")
    win.transient(root)
    win.grab_set()
    win.protocol("WM_DELETE_WINDOW", lambda: None)
    ttk.Label(win, text=f"v{new_tag} indiriliyor...", padding=10).pack()
    bar = ttk.Progressbar(win, length=380, mode='determinate', maximum=100)
    bar.pack(padx=20, pady=5)
    pct = ttk.Label(win, text="0%")
    pct.pack()

    state = {'err': None}

    def progress(d, t):
        def update():
            if t:
                bar['maximum'] = t
                bar['value'] = d
                pct.config(text=f"{int(d * 100 / t)}%  ({d // 1024} / {t // 1024} KB)")
            else:
                pct.config(text=f"{d // 1024} KB")
        root.after(0, update)

    def worker():
        try:
            _download_file(asset_url, new_exe, progress)
        except Exception as e:
            state['err'] = str(e)
        root.after(0, finish)

    def finish():
        try:
            win.destroy()
        except Exception:
            pass
        if state['err']:
            messagebox.showerror("Guncelleme hatasi",
                f"Indirme basarisiz:\n{state['err']}\n\nMevcut versiyonla devam ediliyor.")
            return
        ps_path = _write_powershell_updater(temp_dir, new_exe, target_exe, log_path)
        _launch_powershell_updater(ps_path)
        root.destroy()
        sys.exit(0)

    threading.Thread(target=worker, daemon=True).start()


def check_for_update(root):
    """Background-check GitHub, prompt on Tk thread if newer."""
    def bg():
        result = _get_latest_release()
        if not result:
            return
        latest_tag, asset_url = result
        if not asset_url or not _is_newer(latest_tag, __version__):
            return
        def prompt():
            if messagebox.askyesno("Guncelleme mevcut",
                f"Yeni versiyon v{latest_tag} mevcut.\n"
                f"Mevcut: v{__version__}\n\n"
                "Simdi guncellemek ister misin?"):
                _do_update(root, asset_url, latest_tag)
        root.after(0, prompt)
    threading.Thread(target=bg, daemon=True).start()


# ---------------- GUI ----------------

APP_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
CONFIG = os.path.join(APP_DIR, "config.txt")


class App:
    def __init__(self, root):
        self.root = root
        root.title(f"Last Pages - Cube Block Creator  (v{__version__})")
        root.geometry("900x650")

        top = ttk.Frame(root, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Excel:").pack(side="left")
        self.excel_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.excel_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(top, text="Sec...", command=self.browse).pack(side="left")

        btns = ttk.Frame(root, padding=(10, 0))
        btns.pack(fill="x")
        self.run_btn = ttk.Button(btns, text="EXCEL'E YAZ", command=self.start_run)
        self.run_btn.pack(side="left")
        ttk.Button(btns, text="Log temizle", command=self.clear).pack(side="left", padx=5)

        info = ttk.Label(root, padding=(10, 5),
                         text=f"Yazilacak {len(BLOCKS)} blok: " +
                              ", ".join(str(b['cube_no']) for b in BLOCKS))
        info.pack(fill="x")

        self.output = scrolledtext.ScrolledText(root, wrap="word", font=("Consolas", 10))
        self.output.pack(fill="both", expand=True, padx=10, pady=10)

        self.status = ttk.Label(root, text="Hazir.", anchor="w", padding=5)
        self.status.pack(fill="x")

        if os.path.exists(CONFIG):
            try:
                with open(CONFIG, encoding="utf-8") as f:
                    self.excel_var.set(f.read().strip())
            except Exception:
                pass

        self.log(f"Calisma klasoru: {APP_DIR}")
        self.log(f"Python: {sys.version.split()[0]}")
        self.log("")
        self.log("Yazilacak bloklar: " + ", ".join(str(b['cube_no']) for b in BLOCKS))
        self.log("Excel dosyasini sec, 'EXCEL'E YAZ' bas.")
        self.log("ONEMLI: yazilan satirlar Concrete sheet'in sonuna eklenir,")
        self.log("        son bloktan stil + formul kopyalanir.")
        self.log("")

    def log(self, msg=""):
        self.output.insert("end", msg + "\n")
        self.output.see("end")
        self.root.update_idletasks()

    def clear(self):
        self.output.delete("1.0", "end")

    def browse(self):
        path = filedialog.askopenfilename(
            title="Excel sec",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")],
        )
        if path:
            self.excel_var.set(path)

    def set_status(self, txt):
        self.status.config(text=txt)
        self.root.update_idletasks()

    def start_run(self):
        excel = self.excel_var.get().strip().strip('"')
        if not excel or not os.path.exists(excel):
            messagebox.showerror("Hata", "Excel dosyasi bulunamadi.")
            return
        if not messagebox.askokcancel("Onay",
                f"{len(BLOCKS)} blok orijinal Excel'in UZERINE yazilacak.\n\n"
                "Yedek aldigindan emin misin?\n\nDosya: " + excel):
            return
        try:
            with open(CONFIG, "w", encoding="utf-8") as f:
                f.write(excel)
        except Exception:
            pass
        self.run_btn.config(state="disabled")
        threading.Thread(target=self._worker, args=(excel,), daemon=True).start()

    def _worker(self, excel):
        try:
            self.log("=" * 60)
            self.set_status("Yaziliyor...")
            run_writer(excel, self.log)
            self.log("")
            self.log("*** BASARILI ***")
            self.set_status("Bitti.")
            messagebox.showinfo("Bitti",
                f"Excel guncellendi:\n{excel}\n\n"
                "Excel'i acip 713-716 bloklarini kontrol et.")
        except Exception as e:
            self.log("")
            self.log(f"*** HATA: {e} ***")
            self.log(traceback.format_exc())
            self.set_status("Hata.")
            messagebox.showerror("Hata", str(e))
        finally:
            self.run_btn.config(state="normal")


def main():
    root = tk.Tk()
    App(root)
    root.after(800, lambda: check_for_update(root))
    root.mainloop()


if __name__ == "__main__":
    main()
