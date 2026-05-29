"""
launcher.py - GUI bootstrap for Last Pages Program.

Tek dosyada bir tkinter penceresi acar:
  - Excel dosyasini sec
  - "Calistir" butonu ile guncelle + calistir
  - Tum cikti penceredeki text area'da gozukur
  - Hata olsa bile pencere ACIK kalir
"""
import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox
from urllib.request import urlretrieve
from urllib.error import URLError

REPO = "https://raw.githubusercontent.com/yavuzzeynulat-cell/LastPagesProgram/main"
HERE = os.path.dirname(os.path.abspath(sys.argv[0]))
CONFIG = os.path.join(HERE, "config.txt")


class App:
    def __init__(self, root):
        self.root = root
        root.title("Last Pages - Cube Block Creator")
        root.geometry("900x650")

        top = ttk.Frame(root, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Excel dosyasi:").pack(side="left")
        self.excel_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.excel_var).pack(
            side="left", fill="x", expand=True, padx=5
        )
        ttk.Button(top, text="Sec...", command=self.browse).pack(side="left")

        btns = ttk.Frame(root, padding=(10, 0))
        btns.pack(fill="x")
        self.run_btn = ttk.Button(
            btns, text="GUNCELLE VE CALISTIR", command=self.start_run
        )
        self.run_btn.pack(side="left")
        ttk.Button(btns, text="Log temizle", command=self.clear).pack(
            side="left", padx=5
        )
        ttk.Button(btns, text="Klasoru ac", command=self.open_folder).pack(
            side="left"
        )

        self.output = scrolledtext.ScrolledText(
            root, wrap="word", font=("Consolas", 10)
        )
        self.output.pack(fill="both", expand=True, padx=10, pady=10)

        self.status = ttk.Label(root, text="Hazir.", anchor="w", padding=5)
        self.status.pack(fill="x")

        if os.path.exists(CONFIG):
            try:
                with open(CONFIG, encoding="utf-8") as f:
                    self.excel_var.set(f.read().strip())
            except Exception:
                pass

        self.log(f"Calisma klasoru: {HERE}")
        self.log(f"Python: {sys.version.split()[0]}")
        self.log("")
        self.log("Excel dosyasini sec, 'GUNCELLE VE CALISTIR' bas.")
        self.log("Hata olursa bu pencere acik kalacak, mesajlar burada cikacak.")
        self.log("")

    def log(self, msg=""):
        self.output.insert("end", msg + "\n")
        self.output.see("end")
        self.root.update_idletasks()

    def clear(self):
        self.output.delete("1.0", "end")

    def open_folder(self):
        os.startfile(HERE)

    def browse(self):
        path = filedialog.askopenfilename(
            title="Excel dosyasini sec",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tum dosyalar", "*.*")],
        )
        if path:
            self.excel_var.set(path)

    def set_status(self, txt):
        self.status.config(text=txt)
        self.root.update_idletasks()

    def start_run(self):
        excel = self.excel_var.get().strip().strip('"')
        if not excel:
            messagebox.showerror("Hata", "Excel dosyasi sec.")
            return
        if not os.path.exists(excel):
            messagebox.showerror("Hata", f"Excel bulunamadi:\n{excel}")
            return
        try:
            with open(CONFIG, "w", encoding="utf-8") as f:
                f.write(excel)
        except Exception as e:
            self.log(f"config.txt kaydedilemedi: {e}")

        self.run_btn.config(state="disabled")
        threading.Thread(target=self._run_worker, args=(excel,), daemon=True).start()

    def _run_worker(self, excel):
        try:
            self.log("=" * 60)
            self.log("[1] Guncel dosyalar indiriliyor...")
            self.set_status("Indiriliyor...")
            for fn in ("create_blocks.py", "blocks_data.json"):
                url = f"{REPO}/{fn}"
                dst = os.path.join(HERE, fn)
                try:
                    urlretrieve(url, dst)
                    self.log(f"    OK: {fn}")
                except URLError as e:
                    self.log(f"    HATA: {fn} indirilemedi: {e}")
                    if not os.path.exists(dst):
                        self.log("    Yerel kopya da yok, devam edemiyoruz.")
                        return
                    self.log("    Yerel kopya kullanilacak.")

            self.log("")
            self.log("[2] openpyxl kontrol...")
            self.set_status("openpyxl kontrol...")
            try:
                import openpyxl  # noqa: F401
                self.log("    Zaten kurulu.")
            except ImportError:
                self.log("    Kuruluyor (pip)...")
                r = subprocess.run(
                    [sys.executable, "-m", "pip", "install", "openpyxl"],
                    capture_output=True, text=True,
                )
                self.log(r.stdout)
                if r.returncode != 0:
                    self.log(f"    HATA: {r.stderr}")
                    return
                self.log("    OK.")

            self.log("")
            self.log("[3] Script calistiriliyor...")
            self.log(f"    Excel: {excel}")
            self.set_status("Calisiyor...")
            script = os.path.join(HERE, "create_blocks.py")
            r = subprocess.run(
                [sys.executable, script, excel],
                capture_output=True, text=True, cwd=HERE,
            )
            if r.stdout:
                self.log(r.stdout.rstrip())
            if r.stderr:
                self.log("--- STDERR ---")
                self.log(r.stderr.rstrip())
            self.log("")
            if r.returncode == 0:
                self.log("*** BASARILI ***")
                self.set_status("Bitti.")
                messagebox.showinfo(
                    "Bitti",
                    f"Excel guncellendi:\n{excel}\n\n"
                    "Excel'i acip blok 713-716 yerli yerinde mi kontrol et.",
                )
            else:
                self.log(f"*** HATA (rc={r.returncode}) ***")
                self.set_status("Hata.")
        except Exception as e:
            self.log(f"BEKLENMEDIK HATA: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.set_status("Hata.")
        finally:
            self.run_btn.config(state="normal")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
