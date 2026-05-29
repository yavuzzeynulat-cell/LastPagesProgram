"""
LastPagesApp - Cube block creator GUI (self-contained).

Tek dosya: writer + GUI + blok verisi.
PyInstaller ile .exe yapilir, Python kurmaya gerek yok.
"""
import os
import re
import sys
import threading
import traceback
from copy import copy
from datetime import date, datetime, timedelta

import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox

import openpyxl
from openpyxl.styles import Alignment


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
            {"mould": 60,  "row_type": "2day", "age": 2},
            {"mould": 8,   "row_type": "site"},
            {"mould": 121, "row_type": "site"},
            {"mould": None,"row_type": "site"},
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
WHOLE_BLOCK_MERGES = ['A', 'B', 'E', 'F', 'H', 'J']  # P handled separately (site rows need individual cells)
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
    if not src_cell.has_style:
        return
    dst_cell.font = copy(src_cell.font)
    dst_cell.fill = copy(src_cell.fill)
    dst_cell.border = copy(src_cell.border)
    dst_cell.alignment = copy(src_cell.alignment)
    dst_cell.number_format = src_cell.number_format
    dst_cell.protection = copy(src_cell.protection)


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

    # P column: merge only over rows BEFORE the first site row.
    # Site rows need individual P cells to write 'Site'.
    first_site_idx = next((i for i, r in enumerate(rows) if r.get('row_type') == 'site'), n)
    if first_site_idx >= 2:
        ws.merge_cells(start_row=start_row, end_row=start_row + first_site_idx - 1,
                       start_column=COL['P'], end_column=COL['P'])

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
    ws.cell(row=start_row, column=COL['E'], value=block.get('site_location', ''))
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
        # Age (L): formula =K-I for numeric ages; literal text for WP/F/T/etc.
        rt = row.get('row_type')
        if rt in ('7d', 'cmd', '28d', '1day', '2day') and td is not None:
            ws.cell(row=r, column=COL['L'], value=f"=K{r}-I{r}")
        elif row.get('age') is not None:
            ws.cell(row=r, column=COL['L'], value=row['age'])
        if row.get('row_type') == 'ft' and row.get('ft_note'):
            ws.cell(row=r, column=COL['N'], value=row['ft_note'])
        if row.get('row_type') == 'site':
            ws.cell(row=r, column=COL['P'], value='Site')
        f, src_r = donor_formulas.get('O', (None, None))
        if f:
            ws.cell(row=r, column=COL['O'], value=shift_formula(f, src_r, r))

    return end_row + 1


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
    log("Tamamlandi.")


# ---------------- GUI ----------------

APP_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
CONFIG = os.path.join(APP_DIR, "config.txt")


class App:
    def __init__(self, root):
        self.root = root
        root.title("Last Pages - Cube Block Creator")
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
    root.mainloop()


if __name__ == "__main__":
    main()
