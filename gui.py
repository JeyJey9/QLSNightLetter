import os
import threading
import datetime
import pdfplumber
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor

from processing import process_pdf
from excel_updater import update_masters
from languages import LANG_CODES, TRANSLATIONS

def tr(lang, key):
    return TRANSLATIONS.get(lang, TRANSLATIONS["en_US"]).get(key, key)

def get_flag_path(lang_name):
    fname = lang_name.replace(" ", "_") + ".png"
    return os.path.join("flags", fname)

def show_language_picker(parent, set_lang_callback, current_lang):
    picker = tk.Toplevel(parent)
    picker.title(tr(current_lang, "Select language"))
    picker.resizable(False, False)
    frame = ttk.Frame(picker, padding=14)
    frame.pack(fill="both", expand=True)
    row, col = 0, 0
    max_cols = 6
    flag_images = {}
    for lang_key, lang_code in LANG_CODES.items():
        flag_path = get_flag_path(lang_key)
        try:
            img = tk.PhotoImage(file=flag_path)
        except Exception:
            img = None
        flag_images[lang_key] = img
        def make_select_handler(lc=lang_code):
            def handler(event=None):
                set_lang_callback(lc)
                picker.destroy()
            return handler
        btn = tk.Button(
            frame,
            image=img, text=tr(current_lang, lang_key),
            compound="top", width=100, height=70,
            font=("Segoe UI", 9, "bold"),
            relief="raised", bd=2, bg="#fff", fg="#111",
            command=make_select_handler()
        )
        btn.grid(row=row, column=col, padx=8, pady=10)
        col += 1
        if col >= max_cols:
            col = 0
            row += 1

def launch_gui():
    

    root = tk.Tk()
    root.title("QLS Night Letter")

    # --- Language Handling ---
    current_lang_code = tk.StringVar(value="en_US")
    
    frm = ttk.Frame(root, padding=10)
    frm.grid(row=0, column=0, sticky="nsew")

    # Language BUTTON at the top
    lang_button = tk.Button(frm, text=tr("en_US", "Language"), width=12, font=("Segoe UI", 11, "bold"))
    lang_button.grid(row=0, column=0, sticky="nw", pady=(0,8))

    # All controls that need translation are stored here
    translatable_widgets = [(lang_button, "Language")]

    # --- Variable fields ---
    root_folder    = tk.StringVar(value=os.getcwd())
    mapping_file_v = tk.StringVar(value="Sticker_Mapping.xlsx")
    last_monthly   = tk.BooleanVar(value=True)
    last_weekly    = tk.BooleanVar(value=True)
    last_daily     = tk.BooleanVar(value=True)
    prev_monthly   = tk.BooleanVar(value=False)
    prev_weekly    = tk.BooleanVar(value=False)
    prev_daily     = tk.BooleanVar(value=False)

    shift_monthly  = tk.BooleanVar(value=False)
    shift_weekly   = tk.BooleanVar(value=False)
    shift_daily    = tk.BooleanVar(value=False)

    # --- Path selectors ---
    lbl_root = ttk.Label(frm, text=tr("en_US", "Root Folder:"))
    lbl_root.grid(row=1, column=0, sticky="w")
    translatable_widgets.append((lbl_root, "Root Folder:"))

    entry_root = ttk.Entry(frm, textvariable=root_folder, width=50)
    entry_root.grid(row=1, column=1)
    btn_root = ttk.Button(frm, text=tr("en_US", "Browse"), command=lambda: root_folder.set(filedialog.askdirectory()))
    btn_root.grid(row=1, column=2)
    translatable_widgets.append((btn_root, "Browse"))

    lbl_map = ttk.Label(frm, text=tr("en_US", "Mapping File:"))
    lbl_map.grid(row=2, column=0, sticky="w")
    translatable_widgets.append((lbl_map, "Mapping File:"))

    entry_map = ttk.Entry(frm, textvariable=mapping_file_v, width=50)
    entry_map.grid(row=2, column=1)
    btn_map = ttk.Button(frm, text=tr("en_US", "Browse"), command=lambda: mapping_file_v.set(filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])))
    btn_map.grid(row=2, column=2)
    translatable_widgets.append((btn_map, "Browse"))

    # --- Extraction toggles ---
    cb_last_monthly = ttk.Checkbutton(frm, text=tr("en_US", "Last Monthly  (F)"), variable=last_monthly)
    cb_last_monthly.grid(row=3, column=0)
    cb_last_weekly = ttk.Checkbutton(frm, text=tr("en_US", "Last Weekly   (M)"), variable=last_weekly)
    cb_last_weekly.grid(row=3, column=1)
    cb_last_daily = ttk.Checkbutton(frm, text=tr("en_US", "Last Daily    (Q)"), variable=last_daily)
    cb_last_daily.grid(row=3, column=2)

    cb_prev_monthly = ttk.Checkbutton(frm, text=tr("en_US", "Prev Monthly  (E)"), variable=prev_monthly)
    cb_prev_monthly.grid(row=4, column=0)
    cb_prev_weekly = ttk.Checkbutton(frm, text=tr("en_US", "Prev Weekly   (L)"), variable=prev_weekly)
    cb_prev_weekly.grid(row=4, column=1)
    cb_prev_daily = ttk.Checkbutton(frm, text=tr("en_US", "Prev Daily    (P)"), variable=prev_daily)
    cb_prev_daily.grid(row=4, column=2)

    for cb, key in [
        (cb_last_monthly, "Last Monthly  (F)"),
        (cb_last_weekly, "Last Weekly   (M)"),
        (cb_last_daily, "Last Daily    (Q)"),
        (cb_prev_monthly, "Prev Monthly  (E)"),
        (cb_prev_weekly, "Prev Weekly   (L)"),
        (cb_prev_daily, "Prev Daily    (P)")
    ]:
        translatable_widgets.append((cb, key))

    # --- Shift toggles ---
    lbl_shift = ttk.Label(frm, text=tr("en_US", "Shift Values:"))
    lbl_shift.grid(row=5, column=0, sticky="w")
    translatable_widgets.append((lbl_shift, "Shift Values:"))

    cb_shift_monthly = ttk.Checkbutton(frm, text=tr("en_US", "Monthly (D→C–F→E)"), variable=shift_monthly)
    cb_shift_monthly.grid(row=5, column=1)
    cb_shift_weekly = ttk.Checkbutton(frm, text=tr("en_US", "Weekly  (I→H–M→L)"), variable=shift_weekly)
    cb_shift_weekly.grid(row=5, column=2)
    cb_shift_daily = ttk.Checkbutton(frm, text=tr("en_US", "Daily   (P→O–Q→P)"), variable=shift_daily)
    cb_shift_daily.grid(row=5, column=3)

    for cb, key in [
        (cb_shift_monthly, "Monthly (D→C–F→E)"),
        (cb_shift_weekly, "Weekly  (I→H–M→L)"),
        (cb_shift_daily, "Daily   (P→O–Q→P)")
    ]:
        translatable_widgets.append((cb, key))

    # --- Progress bar and log text box ---
    progress = ttk.Progressbar(frm, length=400, mode='determinate')
    progress.grid(row=6, column=0, columnspan=4, pady=10)
    txt = tk.Text(frm, width=100, height=20)
    txt.grid(row=7, column=0, columnspan=4)

    def log_func(msg):
        txt.insert(tk.END, msg)
        txt.see(tk.END)

    def update_prog():
        progress['value'] += 1

    # --- Run button ---
    btn_run = ttk.Button(frm, text=tr("en_US", "Run"))
    btn_run.grid(row=8, column=1, pady=10)
    translatable_widgets.append((btn_run, "Run"))

    def on_run():
        txt.delete('1.0', 'end')
        progress['value'] = 0
        r = root_folder.get()

        # Gather all PDFs
        pdfs = []
        for plant in os.listdir(r):
            for fld in ['CAL', 'WO CAL', 'FCPA']:
                d = os.path.join(r, plant, fld)
                if os.path.isdir(d):
                    for f in os.listdir(d):
                        if f.lower().endswith('.pdf'):
                            pdfs.append(os.path.join(d, f))

        # Identify broken PDFs (>5 pages)
        broken = []
        for pdf in pdfs:
            with pdfplumber.open(pdf) as doc:
                if len(doc.pages) > 5:
                    broken.append(os.path.basename(pdf))

        insufficient = []
        masters = [
            f for f in os.listdir(r)
            if f.lower().startswith("study_nl_") and f.lower().endswith(".xlsx")
        ]
        progress['maximum'] = len(pdfs) + len(masters)

        def worker():
            out_root = os.path.join(r, "EXCEL_OUTPUT")
            os.makedirs(out_root, exist_ok=True)

            # Convert PDFs in parallel
            def wrapper(pdf_path):
                ok = process_pdf(pdf_path, r, out_root, log_func, update_prog)
                if not ok:
                    insufficient.append(os.path.basename(pdf_path))
                return ok

            with ThreadPoolExecutor(max_workers=2) as exe:
                list(exe.map(wrapper, pdfs))

            # Update master workbooks
            update_masters(
                mapping_file_v.get(), out_root, r,
                last_monthly.get(), last_weekly.get(), last_daily.get(),
                prev_monthly.get(), prev_weekly.get(), prev_daily.get(),
                shift_monthly.get(), shift_weekly.get(), shift_daily.get(),
                log_func, update_prog
            )

            # Report skipped PDFs
            if insufficient:
                log_func(f"\n{len(insufficient)} skipped (insufficient data):\n")
                for fn in insufficient:
                    log_func(f"  • {fn}\n")

            # Report broken PDFs for manual review
            if broken:
                log_func("\nPDFs with >5 pages (please review manually):\n")
                for fn in broken:
                    log_func(f"  • {fn}\n")

            # Check locked masters
            locked = []
            for f in os.listdir(r):
                if f.lower().startswith("study_nl_") and f.lower().endswith(".xlsx"):
                    path = os.path.join(r, f)
                    try:
                        os.rename(path, path)
                    except PermissionError:
                        locked.append(f)
            if locked:
                log_func("\n[WARN] The following master files are still locked; please close them before opening:\n")
                for mf in locked:
                    log_func(f"  • {mf}\n")

            root.after(0, lambda: messagebox.showinfo(tr(current_lang_code.get(), "Done"), f"{tr(current_lang_code.get(), 'Complete')}. {len(insufficient)} {tr(current_lang_code.get(), 'skipped')}."))

        threading.Thread(target=worker, daemon=True).start()

    btn_run.config(command=on_run)

    # --- Version and contact ---
    lbl_version = ttk.Label(frm, text="Version 1.6.3")
    lbl_version.grid(row=9, column=0, sticky="w", pady=(10,0))
    lbl_email = ttk.Label(frm, text="grece@ford.com.tr")
    lbl_email.grid(row=9, column=3, sticky="e", pady=(10,0))

    # --- Language logic ---
    def set_language(lang_code):
        current_lang_code.set(lang_code)
        for widget, key in translatable_widgets:
            widget.config(text=tr(lang_code, key))
        lang_button.config(text=tr(lang_code, "Language"))

    lang_button.config(command=lambda: show_language_picker(root, set_language, current_lang_code.get()))

    root.mainloop()
