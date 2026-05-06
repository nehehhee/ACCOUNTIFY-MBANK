"""
Bank PDF → Excel Konvertor — offline Python desktop app
Spusti: python mbank_konvertor.py
"""

import os
import sys
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

_DND_AVAILABLE = False

import pandas as pd
from PIL import Image, ImageTk

from banks import get_parser, supported_banks
from core import accounts
from core.exporter import write_excel

def resource(filename):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, filename)

def open_file(path: str):
    """Otvorí súbor alebo priečinok natívnou aplikáciou — funguje na Mac aj Windows."""
    try:
        if sys.platform == "darwin":
            import subprocess
            subprocess.call(["open", path])
        elif sys.platform == "win32":
            open_file(path)
        else:
            import subprocess
            subprocess.call(["xdg-open", path])
    except Exception:
        pass

# ── Farby ──────────────────────────────────────────────────────────────────────
BG        = "#1a1a1a"
BG2       = "#242424"
BG3       = "#1e1e1e"
BORDER    = "#3a3a3a"
RED       = "#CC1F1A"
RED2      = "#e02020"
ORANGE    = "#F5A623"
BLUE_MB   = "#1455C0"
GREEN_MB  = "#2E8B4A"
GREEN     = "#4ade80"
TEXT      = "#e2e8f0"
TEXT_DIM  = "#888888"
TEXT_DARK = "#555555"
W         = 480

# ── Pomocné widgety ────────────────────────────────────────────────────────────

def make_logo(parent):
    try:
        logo_path = resource("logo.png")
        img = Image.open(logo_path)
        orig_w, orig_h = img.size
        logo_h = int(W * orig_h / orig_w)
        img = img.resize((W, logo_h), Image.LANCZOS)
        photo = ImageTk.PhotoImage(img)
        lbl = tk.Label(parent, image=photo, bd=0, highlightthickness=0)
        lbl.image = photo
        return lbl
    except Exception:
        return tk.Label(parent, text="Bank Converter", bg=BG2, fg=TEXT,
                        font=("Segoe UI", 16, "bold"), pady=16)

def make_dropzone(parent, click_cmd):
    outer = tk.Frame(parent, bg=BORDER, padx=2, pady=2)
    inner = tk.Frame(outer, bg=BG3, cursor="hand2")
    inner.pack(fill="both", expand=True)
    icon_lbl = tk.Label(inner, text="🗋", bg=BG3, fg=TEXT_DARK, font=("Segoe UI", 36))
    icon_lbl.pack(pady=(24, 4))
    title_lbl = tk.Label(inner, text="Pretiahnite súbor sem",
                         bg=BG3, fg=TEXT, font=("Segoe UI", 13, "bold"))
    title_lbl.pack()
    sub_lbl = tk.Label(inner, text="alebo kliknite pre výber",
                       bg=BG3, fg=TEXT_DARK, font=("Segoe UI", 11))
    sub_lbl.pack(pady=(3, 24))
    for w in [outer, inner, icon_lbl, title_lbl, sub_lbl]:
        w.bind("<Button-1>", lambda e: click_cmd())
    return outer

def make_stat_card(parent, label, color, col, row):
    cell = tk.Frame(parent, bg=BG3, highlightbackground=BORDER, highlightthickness=1)
    cell.grid(row=row, column=col, padx=(0, 8) if col == 0 else (0, 0),
              pady=(0, 8), sticky="nsew")
    tk.Label(cell, text=label, bg=BG3, fg=TEXT_DARK,
             font=("Segoe UI", 10)).pack(anchor="w", padx=14, pady=(12, 2))
    var = tk.StringVar(value="—")
    tk.Label(cell, textvariable=var, bg=BG3, fg=color,
             font=("Segoe UI", 20, "bold")).pack(anchor="w", padx=14, pady=(0, 12))
    return var

def make_rainbow(parent):
    c = tk.Canvas(parent, width=W, height=5, highlightthickness=0, bd=0)
    seg_w = W // 4
    for i, col in enumerate([RED, ORANGE, BLUE_MB, GREEN_MB]):
        c.create_rectangle(i * seg_w, 0, (i + 1) * seg_w + 2, 5, fill=col, outline="")
    return c

# ── Hlavná appka ───────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bank Converter")
        self.resizable(False, False)
        self.configure(bg=BG2)
        self._pdf_paths = []
        self._df = None
        self._individual_dfs = []
        self._iban = ""
        self._peek_iban = ""
        self._confirmed_name = ""
        icon_path = resource("icon.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        self._build()
        self._center()
        # macOS: reaguje na pretiahnutie súboru na ikonu v Docku / "Otvoriť pomocou"
        if sys.platform == "darwin":
            self.createcommand("::tk::mac::OpenDocument", self._on_open_document)

    def _center(self):
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"+{x}+{y}")

    def _build(self):
        make_logo(self).pack(fill="x")

        self._body = tk.Frame(self, bg=BG2, padx=24, pady=22)
        self._body.pack(fill="x")

        # Drop zóna
        self._drop = make_dropzone(self._body, self._pick_file)
        self._drop.pack(fill="x", pady=(0, 14))
        if _DND_AVAILABLE:
            self._drop.drop_target_register(DND_FILES)
            self._drop.dnd_bind('<<Drop>>', self._on_drop)
            for w in self._drop.winfo_children():
                w.drop_target_register(DND_FILES)
                w.dnd_bind('<<Drop>>', self._on_drop)

        # Scanning indikátor (skrytý)
        self._scan_lbl = tk.Label(self._body, text="⏳  Rozpoznávam banku...",
                                  bg=BG2, fg=TEXT_DIM, font=("Segoe UI", 11), anchor="w")

        # Confirm panel (skrytý)
        self._confirm = self._build_confirm_panel(self._body)

        # Progress
        self._prog_lbl = tk.Label(self._body, text="", bg=BG2, fg=TEXT_DIM,
                                  font=("Segoe UI", 10), anchor="w")
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("mb.Horizontal.TProgressbar",
                        troughcolor=BG3, background=RED, thickness=5, borderwidth=0)
        self._prog = ttk.Progressbar(self._body, style="mb.Horizontal.TProgressbar",
                                     orient="horizontal", mode="determinate")

        # Stats 2×2 (skrytý)
        self._stats = tk.Frame(self._body, bg=BG2)
        self._stats.columnconfigure(0, weight=1)
        self._stats.columnconfigure(1, weight=1)
        self._stat_vars = {
            "total":  make_stat_card(self._stats, "Celkom transakcií", TEXT,       0, 0),
            "iban":   make_stat_card(self._stats, "S IBAN protiúčtu",  TEXT,       1, 0),
            "credit": make_stat_card(self._stats, "Príjmy  (Credit)",  GREEN,      0, 1),
            "debit":  make_stat_card(self._stats, "Výdaje  (Debit)",   "#f87171",  1, 1),
        }

        # Uložiť button (skrytý)
        self._btn_save = tk.Button(
            self._body, text="⬇   Uložiť Excel",
            bg="#1a2e1a", fg=GREEN,
            font=("Segoe UI", 13, "bold"), bd=1, relief="solid",
            highlightbackground="#2d4a2d",
            activebackground="#1f381f", activeforeground=GREEN,
            cursor="hand2", pady=13,
            command=self._save_excel,
        )

        make_rainbow(self).pack(fill="x", side="bottom")

    def _build_confirm_panel(self, parent):
        frame = tk.Frame(parent, bg=BG3, highlightbackground=BORDER, highlightthickness=1)

        # Hlavička — banka + IBAN
        header = tk.Frame(frame, bg=BG3)
        header.pack(fill="x", padx=16, pady=(14, 0))

        tk.Label(header, text="✓", bg=BG3, fg=GREEN,
                 font=("Segoe UI", 12, "bold")).pack(side="left")
        self._cf_bank = tk.Label(header, text="", bg=BG3, fg=TEXT,
                                 font=("Segoe UI", 12, "bold"))
        self._cf_bank.pack(side="left", padx=(6, 0))

        self._cf_iban = tk.Label(frame, text="", bg=BG3, fg=TEXT_DARK,
                                 font=("Segoe UI", 10))
        self._cf_iban.pack(anchor="w", padx=16, pady=(3, 10))

        # Oddeľovač
        tk.Frame(frame, bg=BORDER, height=1).pack(fill="x")

        # Pole pre názov účtu
        lbl_row = tk.Frame(frame, bg=BG3)
        lbl_row.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(lbl_row, text="Názov účtu", bg=BG3, fg=TEXT_DIM,
                 font=("Segoe UI", 10)).pack(side="left")
        self._cf_hint = tk.Label(lbl_row, text="", bg=BG3, fg=ORANGE,
                                 font=("Segoe UI", 9))
        self._cf_hint.pack(side="left", padx=(8, 0))

        self._nazov_var = tk.StringVar()
        self._cf_entry = tk.Entry(
            frame, textvariable=self._nazov_var,
            bg="#2a2a2a", fg=TEXT, insertbackground=TEXT,
            relief="flat", bd=0, font=("Segoe UI", 12),
            highlightbackground=BORDER, highlightthickness=1,
        )
        self._cf_entry.pack(fill="x", padx=16, pady=(0, 14), ipady=8)

        # Tlačidlá
        btn_row = tk.Frame(frame, bg=BG3)
        btn_row.pack(fill="x", padx=16, pady=(0, 14))
        btn_row.columnconfigure(0, weight=1)
        btn_row.columnconfigure(1, weight=2)

        tk.Button(
            btn_row, text="Preskočiť",
            bg=BG3, fg=TEXT_DARK, font=("Segoe UI", 11),
            bd=0, relief="flat", activebackground=BG2, activeforeground=TEXT,
            cursor="hand2", pady=10, command=self._on_skip,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))

        tk.Button(
            btn_row, text="Konvertovať  →",
            bg=RED, fg="white", font=("Segoe UI", 12, "bold"),
            bd=0, relief="flat", activebackground=RED2, activeforeground="white",
            cursor="hand2", pady=10, command=self._on_confirm,
        ).grid(row=0, column=1, sticky="ew")

        return frame

    # ── Akcie ──────────────────────────────────────────────────────────────────

    def _on_open_document(self, *files):
        """macOS: súbory pretiahnuité na ikonu alebo otvorené cez Finder."""
        pdfs = [f for f in files if f.lower().endswith('.pdf')]
        if pdfs:
            self._load_files(pdfs)

    def _on_drop(self, event):
        raw = event.data.strip()
        paths = re.findall(r'\{([^}]+)\}|(\S+)', raw)
        paths = [a or b for a, b in paths]
        pdfs = [p for p in paths if p.lower().endswith('.pdf')]
        if pdfs:
            self._load_files(pdfs)
        else:
            messagebox.showwarning("Nesprávny súbor", "Pretiahni prosím PDF súbor(y).")

    def _pick_file(self):
        paths = filedialog.askopenfilenames(
            title="Vybrať bankový PDF výpis",
            filetypes=[("PDF súbory", "*.pdf"), ("Všetky súbory", "*.*")],
        )
        if paths:
            self._load_files(list(paths))

    def _load_files(self, paths):
        self._pdf_paths = paths
        self._df = None
        self._individual_dfs = []

        # Schovaj všetko, ukáž scanning indikátor
        self._confirm.pack_forget()
        self._stats.pack_forget()
        self._btn_save.pack_forget()
        self._prog_lbl.pack_forget()
        self._prog.pack_forget()
        self._drop.pack_forget()

        self._scan_lbl.pack(fill="x", pady=(0, 8))
        self._center()

        threading.Thread(target=self._peek_worker, daemon=True).start()

    def _peek_worker(self):
        try:
            if not self._pdf_paths:
                return
            path = self._pdf_paths[0]
            parser = get_parser(path)
            if parser is None:
                banks = ", ".join(supported_banks())
                raise RuntimeError(
                    f'Neznáma banka.\nPodporované banky: {banks}'
                )
            result = parser.peek(path)
            self.after(0, lambda r=result: self._show_confirm(r))
        except Exception as e:
            self.after(0, lambda err=str(e): self._on_peek_error(err))

    def _show_confirm(self, peek_result):
        self._scan_lbl.pack_forget()

        bank  = peek_result.get("bank", "")
        iban  = peek_result.get("iban", "")
        nazov = peek_result.get("nazov", "")

        self._peek_iban = iban  # uložíme pre _on_confirm
        self._cf_bank.config(text=bank)
        self._cf_iban.config(text=iban if iban else "IBAN sa nenašiel")

        if nazov:
            self._cf_hint.config(text="(nájdené v PDF)")
            self._nazov_var.set(nazov)
        else:
            self._cf_hint.config(text="(nenašlo sa — zadajte ručne alebo preskočte)")
            self._nazov_var.set("")

        self._confirm.pack(fill="x", pady=(0, 4))
        self._cf_entry.focus_set()
        self._center()

    def _on_confirm(self):
        name = self._nazov_var.get().strip()
        if name and self._peek_iban:
            accounts.save(self._peek_iban, name)
        self._confirmed_name = name
        self._start_convert()

    def _on_skip(self):
        self._confirmed_name = ""
        self._start_convert()

    def _start_convert(self):
        self._confirm.pack_forget()
        self._prog_lbl.config(text="Spracovávam...")
        self._prog_lbl.pack(fill="x", pady=(8, 2))
        self._prog.pack(fill="x", pady=(0, 4))
        self._center()
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            individual = []
            total = len(self._pdf_paths)
            for i, path in enumerate(self._pdf_paths):
                def safe_cb(pct, text, i=i, total=total):
                    overall = int((i / total) * 100 + pct / total)
                    label = f"[{i+1}/{total}] {text}"
                    self.after(0, lambda p=overall, t=label: self._set_progress(p, t))
                try:
                    parser = get_parser(path)
                    if parser is None:
                        banks = ", ".join(supported_banks())
                        raise RuntimeError(f'Neznáma banka.\nPodporované banky: {banks}')
                    iban, df = parser.parse(path, nazov_uctu=self._confirmed_name,
                                            progress_cb=safe_cb)
                except RuntimeError:
                    raise
                except Exception as e:
                    raise RuntimeError(
                        f'Chyba pri spracovaní "{os.path.basename(path)}":\n{e}'
                    ) from e
                individual.append((path, df))

            self._individual_dfs = individual
            self._df = pd.concat([df for _, df in individual], ignore_index=True)
            self._iban = iban
            self.after(0, self._on_done)
        except Exception as e:
            self.after(0, lambda err=str(e): self._on_error(err))

    def _set_progress(self, pct, text):
        self._prog_lbl.config(text=text)
        self._prog["value"] = pct
        self.update_idletasks()

    def _on_done(self):
        df = self._df
        self._stat_vars["total"].set(str(len(df)))
        self._stat_vars["credit"].set(str(int((df["cd"] == "C").sum())))
        self._stat_vars["debit"].set(str(int((df["cd"] == "D").sum())))
        self._stat_vars["iban"].set(str(int((df["iban_protiuctu"] != "").sum())))

        self._set_progress(100, f"✓  Hotovo — {len(df)} transakcií")
        self._stats.pack(fill="x", pady=(14, 0))
        self._btn_save.pack(fill="x", pady=(14, 0))
        self._center()

    def _on_error(self, msg):
        self._prog_lbl.pack_forget()
        self._prog.pack_forget()
        self._drop.pack(fill="x", pady=(0, 14))
        self._center()
        messagebox.showerror(
            "Chyba konverzie",
            f"{msg}\n\nUisti sa že:\n"
            "• súbor je platný bankový PDF výpis\n"
            "• súbor nie je otvorený v inom programe\n"
            "• súbor nie je poškodený"
        )

    def _on_peek_error(self, msg):
        self._scan_lbl.pack_forget()
        self._drop.pack(fill="x", pady=(0, 14))
        self._center()
        messagebox.showerror("Chyba rozpoznania", msg)

    def _save_excel(self):
        if self._df is None:
            return
        if len(self._individual_dfs) == 1:
            src_path, df = self._individual_dfs[0]
            base = os.path.splitext(os.path.basename(src_path))[0]
            out_xlsx = filedialog.asksaveasfilename(
                title="Uložiť Excel",
                defaultextension=".xlsx",
                initialfile=f"{base} - skonvertované.xlsx",
                filetypes=[("Excel súbory", "*.xlsx")],
            )
            if not out_xlsx:
                return
            write_excel(df, out_xlsx)
            messagebox.showinfo(
                "Uložené",
                f"Excel:  {os.path.basename(out_xlsx)}\n\n"
                f"Priečinok: {os.path.dirname(out_xlsx)}"
            )
            open_file(out_xlsx)
        else:
            folder = filedialog.askdirectory(title="Vybrať priečinok pre uloženie")
            if not folder:
                return
            saved = []
            for src_path, df in self._individual_dfs:
                base = os.path.splitext(os.path.basename(src_path))[0]
                out_xlsx = os.path.join(folder, f"{base} - skonvertované.xlsx")
                write_excel(df, out_xlsx)
                saved.append(os.path.basename(out_xlsx))
            messagebox.showinfo(
                "Uložené",
                f"Uložených {len(saved)} súborov (Excel + HTML graf) do:\n{folder}\n\n"
                + "\n".join(saved)
            )
            open_file(folder)


if __name__ == "__main__":
    app = App()
    app.mainloop()
