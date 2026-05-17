#!/usr/bin/env python3
"""
app.py — GUI for SPT Extractor (PPN & PPh 23).
"""

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import pph23_extractor
import ppn_extractor


# ===================================================================
# ABOUT / BRANDING
# ===================================================================

ABOUT_TITLE = "SPT Extractor  —  Version 1.0"
ABOUT_SUBTITLE = (
    "Designed and developed by Achmad Maulana\n"
    "Tax Consultant | Automation Specialist\n"
    "\"Turning tax filing into a single click.\""
)
ABOUT_CONTACT = (
    "achmadmaulana199902@gmail.com  |  LinkedIn: Achmad Maulana  |  GitHub: Forcedbalanced\n\n"
    "© 2026 Achmad Maulana — Maulana Automation Consultant. Not for resale."
)

_LEGACY_INSCRIPTION = """
This engine was hand-forged by Achmad Maulana in the year of the Dragon.
To those reading this: the mind that shaped these numbers has departed.

"The Moving Finger writes; and, having writ,
Moves on: nor all thy Piety nor Wit
Shall lure it back to cancel half a Line."
— Omar Khayyám (The Rubáiyát)
"""


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('SPT Extractor')
        self.resizable(False, False)
        self._build()

    def _build(self):
        pad = {'padx': 16}

        # ── Extractor type ────────────────────────────────────────────────────
        type_frame = tk.Frame(self, **pad)
        type_frame.grid(row=0, column=0, sticky='w', pady=(14, 4))

        tk.Label(type_frame, text='Jenis SPT:', anchor='w').pack(side='left', padx=(0, 12))
        self.spt_var = tk.StringVar(value='PPN')
        for label in ('PPN', 'PPh 23'):
            tk.Radiobutton(
                type_frame, text=label, variable=self.spt_var, value=label
            ).pack(side='left', padx=4)

        # ── File inputs ───────────────────────────────────────────────────────
        inp_frame = tk.Frame(self, **pad)
        inp_frame.grid(row=1, column=0, sticky='ew')

        tk.Label(inp_frame, text='Folder PDF:', anchor='w').grid(
            row=0, column=0, sticky='w', pady=(0, 6)
        )
        self.folder_var = tk.StringVar()
        tk.Entry(inp_frame, textvariable=self.folder_var, width=55).grid(
            row=0, column=1, padx=(8, 4)
        )
        tk.Button(inp_frame, text='Browse', command=self._browse_folder).grid(row=0, column=2)

        tk.Label(inp_frame, text='Output Excel:', anchor='w').grid(
            row=1, column=0, sticky='w', pady=(0, 4)
        )
        self.output_var = tk.StringVar()
        tk.Entry(inp_frame, textvariable=self.output_var, width=55).grid(
            row=1, column=1, padx=(8, 4)
        )
        tk.Button(inp_frame, text='Browse', command=self._browse_output).grid(row=1, column=2)

        # ── Buttons ───────────────────────────────────────────────────────────
        btn_frame = tk.Frame(inp_frame)
        btn_frame.grid(row=2, column=0, columnspan=3, pady=(16, 8))

        self.btn = tk.Button(
            btn_frame, text='Extract', bg='#2563eb', fg='white',
            font=('', 11, 'bold'), padx=16, pady=6,
            command=self._run
        )
        self.btn.pack(side='left', padx=(0, 12))

        tk.Button(
            btn_frame, text='About',
            font=('', 9), padx=10, pady=2,
            command=self._show_about
        ).pack(side='left')

        # ── Log box ───────────────────────────────────────────────────────────
        self.log_box = tk.Text(self, height=12, width=80, state='disabled', bg='#f1f5f9')
        self.log_box.grid(row=2, column=0, padx=8, pady=(0, 8))

    # ── Browse helpers ────────────────────────────────────────────────────────

    def _browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_var.set(path)
            # Auto-suggest output filename next to selected folder
            if not self.output_var.get():
                spt = self.spt_var.get().lower().replace(' ', '_')
                default_out = str(Path(path).parent / f"hasil_ekstraksi_{spt}.xlsx")
                self.output_var.set(default_out)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx')]
        )
        if path:
            self.output_var.set(path)

    # ── About dialog ─────────────────────────────────────────────────────────

    def _show_about(self):
        persian = (
            'انگشت تحریر قضا در کار است\n'
            'تا چند کنی حیله که این ناپایدار است\n'
            'بر لوح جهان هر آنچه نقش ازل است\n'
            'نتوان ستردن اگر چه نقش نگار است'
        )
        messagebox.showinfo(
            'About',
            f'{ABOUT_TITLE}\n\n'
            f'{ABOUT_SUBTITLE}\n\n'
            f'{ABOUT_CONTACT}\n\n'
            f'{persian}'
        )

    # ── Log helper ────────────────────────────────────────────────────────────

    def _log(self, msg):
        self.log_box.configure(state='normal')
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box.configure(state='disabled')

    # ── Run ───────────────────────────────────────────────────────────────────

    def _run(self):
        folder = self.folder_var.get().strip()
        output = self.output_var.get().strip()
        spt = self.spt_var.get()

        if not folder:
            messagebox.showerror('Error', 'Pilih folder PDF terlebih dahulu.')
            return
        if not Path(folder).is_dir():
            messagebox.showerror('Error', f'Folder tidak ditemukan:\n{folder}')
            return
        if not output:
            messagebox.showerror('Error', 'Tentukan path output Excel.')
            return

        self.btn.configure(state='disabled', text='Processing...')
        self.log_box.configure(state='normal')
        self.log_box.delete('1.0', 'end')
        self.log_box.configure(state='disabled')
        self._log(f'Memproses folder: {folder}')
        self._log(f'Jenis SPT: {spt}\n')

        def task():
            try:
                if spt == 'PPN':
                    results = ppn_extractor.process_folder(folder)
                    ppn_extractor.export_to_excel(results, output)
                else:
                    results = pph23_extractor.process_folder(folder)
                    pph23_extractor.export_to_excel(results, output)

                count = len(results)
                self._log(f'\nSelesai — {count} file diproses.')
                self._log(f'Output: {output}')
                messagebox.showinfo(
                    'Selesai',
                    f'{count} file berhasil diekstrak.\nDisimpan di:\n{output}'
                )
            except Exception as e:
                self._log(f'\nError: {e}')
                messagebox.showerror('Error', str(e))
            finally:
                self.btn.configure(state='normal', text='Extract')

        threading.Thread(target=task, daemon=True).start()


if __name__ == '__main__':
    App().mainloop()
