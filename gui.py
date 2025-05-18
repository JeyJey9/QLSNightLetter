import os
import sys
import threading
import pdfplumber
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor

from processing import process_pdf
from excel_updater import update_masters
from languages import LANG_CODES, TRANSLATIONS

def tr(lang, key):
    return TRANSLATIONS.get(lang, TRANSLATIONS["en_US"]).get(key, key)


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def get_flag_path(lang_name):
    fname = lang_name.replace(" ", "_") + ".png"
    return os.path.join("flags", fname)

def show_language_picker(parent, set_lang_callback, current_lang):
    import sys
    try:
        from PIL import Image, ImageTk
        PIL_OK = True
    except ImportError:
        PIL_OK = False

    popup = tk.Toplevel(parent)
    popup.title("Select Language")
    popup.resizable(False, False)
    popup.configure(bg="#f5f5f5")

    btn_cancel = tk.Button(popup, text="Cancel", font=("Segoe UI", 13),
                           width=9, command=popup.destroy, relief="raised")
    btn_cancel.place(x=15, y=15)

    flag_w, flag_h = 95, 65
    gap_x, gap_y = 22, 22
    start_x, start_y = 20, 65
    flags_per_row = 6

    # Store images on the popup itself (so they stay in memory!)
    popup.flag_images = {}

    langs = list(LANG_CODES.keys())
    max_label_width = flag_w + 12

    for i, lang_key in enumerate(langs):
        lang_code = LANG_CODES[lang_key]
        flag_name = lang_key.replace(" ", "_") + ".png"
        flag_path = flag_path = resource_path(os.path.join("flags", flag_name))


        img = None
        if os.path.isfile(flag_path):
            if PIL_OK:
                try:
                    im = Image.open(flag_path).resize((flag_w, flag_h), Image.LANCZOS)
                    img = ImageTk.PhotoImage(im)
                except Exception as e:
                    print(f"Pillow could not open {flag_path}: {e}", file=sys.stderr)
                    img = None
            else:
                try:
                    img = tk.PhotoImage(file=flag_path)
                except Exception as e:
                    print(f"Tkinter could not open {flag_path}: {e}", file=sys.stderr)
                    img = None

        popup.flag_images[lang_key] = img  # <-- CRUCIAL

        row = i // flags_per_row
        col = i % flags_per_row

        x = start_x + col * (flag_w + gap_x)
        y = start_y + row * (flag_h + gap_y + 30)

        def make_lang_handler(lc=lang_code):
            def handler(event=None):
                set_lang_callback(lc)
                popup.destroy()
            return handler

        btn = tk.Button(popup, image=img, width=flag_w, height=flag_h,
                        relief="groove", command=make_lang_handler(), bg="white")
        if img is None:
            btn.config(bg="#e57373", text="?", fg="white", font=("Segoe UI", 14, "bold"))
        btn.place(x=x, y=y)

        lbl = tk.Label(
            popup,
            text=tr(current_lang, lang_key),
            bg="#f5f5f5",
            font=("Segoe UI", 9, "bold"),
            anchor="center",
            justify="center",
            wraplength=max_label_width
        )
        lbl.place(x=x + flag_w // 2, y=y + flag_h + 14, anchor="n")

    total_rows = (len(langs) + flags_per_row - 1) // flags_per_row
    total_height = start_y + total_rows * (flag_h + gap_y + 30) + 30
    total_width = flags_per_row * (flag_w + gap_x) + start_x
    popup.geometry(f"{total_width}x{total_height}")
    popup.grab_set()


class QLSApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QLS Night Letter")
        self.current_lang_code = "en_US"

        root.geometry("770x570")
        root.configure(background="#d6d6d6", highlightbackground="#dFFFFF", highlightcolor="#000000")

        # Variable fields
        self.var_root_folder = tk.StringVar(value=os.getcwd())
        self.var_mapping_file = tk.StringVar(value="Sticker_Mapping.xlsx")

        self.var_last_monthly = tk.BooleanVar(value=True)
        self.var_prev_monthly = tk.BooleanVar(value=False)
        self.var_last_weekly  = tk.BooleanVar(value=True)
        self.var_prev_weekly  = tk.BooleanVar(value=False)
        self.var_last_daily   = tk.BooleanVar(value=True)
        self.var_prev_daily   = tk.BooleanVar(value=False)

        self.var_shift_monthly = tk.BooleanVar(value=False)
        self.var_shift_weekly  = tk.BooleanVar(value=False)
        self.var_shift_daily   = tk.BooleanVar(value=False)

        

        # Language button
        self.lang_button = tk.Button(
            self.root,
            text=tr(self.current_lang_code, "Language"),
            font=("Segoe UI", 10),            # match Browse
            bg="#fff",                        # white background
            fg="#000",                        # black text
            bd=2,                             # border size (match Browse)
            relief="raised",                  # match Browse
            activebackground="#ececec",       # subtle highlight on press
            highlightthickness=0,             # no focus border
            command=self.open_language_picker
        )
        self.lang_button.place(relx=0.86, rely=0.045, height=35, width=96)  # match Browse button size

        self.help_button = tk.Button(
            self.root,
            text=tr(self.current_lang_code, "Help"),
            font=("Segoe UI", 10),
            bg="#fff",
            fg="#000",
            bd=2,
            relief="raised",
            activebackground="#ececec",
            highlightthickness=0,
            command=self.open_help_popup
        )
        # Place just below the Language button (adjust relx/rely as needed to fit layout)
        self.help_button.place(relx=0.86, rely=0.125, height=35, width=96)

        # Root Folder
        self.Label1 = tk.Label(
            self.root, text=tr(self.current_lang_code, "Root Folder:"),
            background="#d9d9d9", foreground="#000000", anchor="w", borderwidth=0
        )
        self.Label1.place(relx=0.025, rely=0.055, height=21, width=94)

        self.entry_root_folder = tk.Entry(self.root, textvariable=self.var_root_folder)
        self.entry_root_folder.place(relx=0.21, rely=0.055, height=20, relwidth=0.425)

        self.btn_browse_root = tk.Button(
            self.root, text=tr(self.current_lang_code, "Browse"),
            bg="#FFFFFF", bd=0, highlightthickness=0,
            activebackground="#d9d9d9",
            command=self.browse_root
        )
        self.btn_browse_root.place(relx=0.67, rely=0.05, height=26, width=77)

        # Sticker Mapping
        self.Label2 = tk.Label(
            self.root, text=tr(self.current_lang_code, "Mapping File:"),
            background="#d9d9d9", foreground="#000000", anchor="w", borderwidth=0
        )
        self.Label2.place(relx=0.025, rely=0.128, height=21, width=140)

        self.entry_mapping_file = tk.Entry(self.root, textvariable=self.var_mapping_file)
        self.entry_mapping_file.place(relx=0.21, rely=0.128, height=20, relwidth=0.425)

        # --- Column Update Options label ---
        self.Label3 = tk.Label(
            self.root, text=tr(self.current_lang_code, "Column\nUpdate\nOptions"),
            background="#d9d9d9", foreground="#000000", anchor="w", borderwidth=0,
            font=("Segoe UI", 10)
        )
        # Place it between Mapping File and checkboxes
        self.Label3.place(relx=0.025, rely=0.214, height=62, width=170)

        self.btn_browse_map = tk.Button(
            self.root, text=tr(self.current_lang_code, "Browse"),
            bg="#FFFFFF", bd=0, highlightthickness=0,
            activebackground="#d9d9d9",
            command=self.browse_map
        )
        self.btn_browse_map.place(relx=0.67, rely=0.125, height=26, width=77)

        # --- Dynamic Checkboxes with wrapping ---
        self.chk_last_monthly = tk.Checkbutton(
            self.root, text=tr(self.current_lang_code, "Last Monthly  (F)"),
            variable=self.var_last_monthly, anchor="w", justify="left", font=("Segoe UI", 10)
        )
        self.chk_last_monthly.place(relx=0.175, rely=0.212, width=210, height=36)

        self.chk_prev_monthly = tk.Checkbutton(
            self.root, text=tr(self.current_lang_code, "Prev Monthly  (E)"),
            variable=self.var_prev_monthly, anchor="w", justify="left", font=("Segoe UI", 10)
        )
        self.chk_prev_monthly.place(relx=0.175, rely=0.275, width=210, height=36)

        self.chk_last_weekly = tk.Checkbutton(
            self.root, text=tr(self.current_lang_code, "Last Weekly   (M)"),
            variable=self.var_last_weekly, anchor="w", justify="left", font=("Segoe UI", 10)
        )
        self.chk_last_weekly.place(relx=0.405, rely=0.212, width=210, height=36)

        self.chk_prev_weekly = tk.Checkbutton(
            self.root, text=tr(self.current_lang_code, "Prev Weekly   (L)"),
            variable=self.var_prev_weekly, anchor="w", justify="left", font=("Segoe UI", 10)
        )
        self.chk_prev_weekly.place(relx=0.405, rely=0.275, width=210, height=36)

        self.chk_last_daily = tk.Checkbutton(
            self.root, text=tr(self.current_lang_code, "Last Daily    (Q)"),
            variable=self.var_last_daily, anchor="w", justify="left", font=("Segoe UI", 10)
        )
        self.chk_last_daily.place(relx=0.635, rely=0.212, width=175, height=36)

        self.chk_prev_daily = tk.Checkbutton(
            self.root, text=tr(self.current_lang_code, "Prev Daily    (P)"),
            variable=self.var_prev_daily, anchor="w", justify="left", font=("Segoe UI", 10)
        )
        self.chk_prev_daily.place(relx=0.635, rely=0.275, width=175, height=36)

        # Shift checkboxes (wider for full text)
        self.TLabel4 = tk.Label(
            self.root, text=tr(self.current_lang_code, "Shift Values:"),
            background="#d9d9d9", foreground="#000000", anchor="w", borderwidth=0
        )
        self.TLabel4.place(relx=0.025, rely=0.385, height=17, width=125)

        self.chk_shift_monthly = tk.Checkbutton(
            self.root, text=tr(self.current_lang_code, "Monthly (D→C–F→E)"),
            variable=self.var_shift_monthly, anchor="w", justify="left", font=("Segoe UI", 10)
        )
        self.chk_shift_monthly.place(relx=0.175, rely=0.366, width=210, height=36)

        self.chk_shift_weekly = tk.Checkbutton(
            self.root, text=tr(self.current_lang_code, "Weekly  (I→H–M→L)"),
            variable=self.var_shift_weekly, anchor="w", justify="left", font=("Segoe UI", 10)
        )
        self.chk_shift_weekly.place(relx=0.405, rely=0.366, width=210, height=36)

        self.chk_shift_daily = tk.Checkbutton(
            self.root, text=tr(self.current_lang_code, "Daily   (P→O–Q→P)"),
            variable=self.var_shift_daily, anchor="w", justify="left", font=("Segoe UI", 10)
        )
        self.chk_shift_daily.place(relx=0.635, rely=0.366, width=175, height=36)

        # -- Store checkboxes and keys for dynamic updates --
        self.checkbox_widgets = [
            (self.chk_last_monthly, "Last Monthly  (F)"),
            (self.chk_prev_monthly, "Prev Monthly  (E)"),
            (self.chk_last_weekly, "Last Weekly   (M)"),
            (self.chk_prev_weekly, "Prev Weekly   (L)"),
            (self.chk_last_daily, "Last Daily    (Q)"),
            (self.chk_prev_daily, "Prev Daily    (P)"),
            (self.chk_shift_monthly, "Monthly (D→C–F→E)"),
            (self.chk_shift_weekly, "Weekly  (I→H–M→L)"),
            (self.chk_shift_daily, "Daily   (P→O–Q→P)"),
        ]

        # Progress Bar
        self.progress_bar = ttk.Progressbar(self.root)
        self.progress_bar.place(relx=0.081, rely=0.44, relwidth=0.853, height=19)

        # Log text box
        self.text_log = tk.Text(self.root, wrap="word", background="white", foreground="black", font="TkTextFont")
        self.text_log.place(relx=0.054, rely=0.513, relheight=0.319, relwidth=0.899)
        self.text_log.configure(state='disabled')

        # Run button
        self.btn_run = ttk.Button(self.root, text=tr(self.current_lang_code, "Run"), command=self.on_run)
        self.btn_run.place(relx=0.419, rely=0.879, height=36, width=75)

        # -- Translation widgets mapping for language switching --
        self.translatable_widgets = [
            (self.lang_button, "Language"),
            (self.Label1, "Root Folder:"),
            (self.Label2, "Mapping File:"),
            (self.btn_browse_root, "Browse"),
            (self.btn_browse_map, "Browse"),
            (self.TLabel4, "Shift Values:"),
            (self.btn_run, "Run"),
            (self.Label3, "Column\nUpdate\nOptions")
        ]
        # Add checkboxes to translation as well
        for cb, key in self.checkbox_widgets:
            self.translatable_widgets.append((cb, key))

        self.update_checkbox_texts(self.current_lang_code)  # Initial sizing

    def open_help_popup(self):
            help_text = tr(self.current_lang_code, "Help_Message")
            tk.messagebox.showinfo(tr(self.current_lang_code, "Help"), help_text)

    def update_checkbox_texts(self, lang_code):
        labels = [tr(lang_code, key) for _, key in self.checkbox_widgets]
        max_chars = max(len(label) for label in labels)
        width_px = min(max(max_chars * 11, 160), 240)+30
        wrap_px = width_px - 25
        for cb, key in self.checkbox_widgets:
            cb.config(text=tr(lang_code, key), wraplength=wrap_px)

    def browse_root(self):
        path = filedialog.askdirectory()
        if path:
            self.var_root_folder.set(path)

    def browse_map(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.var_mapping_file.set(path)

    def log_func(self, msg):
        self.text_log.configure(state='normal')
        self.text_log.insert("end", msg)
        self.text_log.see("end")
        self.text_log.configure(state='disabled')

    def update_prog(self):
        self.progress_bar['value'] += 1
        self.root.update_idletasks()

    def on_run(self):
        self.text_log.configure(state='normal')
        self.text_log.delete('1.0', 'end')
        self.progress_bar['value'] = 0
        self.text_log.configure(state='disabled')
        r = self.var_root_folder.get()

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
        self.progress_bar['maximum'] = len(pdfs) + len(masters)

        def worker():
            out_root = os.path.join(r, "EXCEL_OUTPUT")
            os.makedirs(out_root, exist_ok=True)

            def wrapper(pdf_path):
                ok = process_pdf(pdf_path, r, out_root, self.log_func, self.update_prog)
                if not ok:
                    insufficient.append(os.path.basename(pdf_path))
                return ok

            with ThreadPoolExecutor(max_workers=2) as exe:
                list(exe.map(wrapper, pdfs))

            update_masters(
                self.var_mapping_file.get(), out_root, r,
                self.var_last_monthly.get(), self.var_last_weekly.get(), self.var_last_daily.get(),
                self.var_prev_monthly.get(), self.var_prev_weekly.get(), self.var_prev_daily.get(),
                self.var_shift_monthly.get(), self.var_shift_weekly.get(), self.var_shift_daily.get(),
                self.log_func, self.update_prog
            )

            if insufficient:
                self.log_func(f"\n{len(insufficient)} skipped (insufficient data):\n")
                for fn in insufficient:
                    self.log_func(f"  • {fn}\n")
            if broken:
                self.log_func("\nPDFs with >5 pages (please review manually):\n")
                for fn in broken:
                    self.log_func(f"  • {fn}\n")
            locked = []
            for f in os.listdir(r):
                if f.lower().startswith("study_nl_") and f.lower().endswith(".xlsx"):
                    path = os.path.join(r, f)
                    try:
                        os.rename(path, path)
                    except PermissionError:
                        locked.append(f)
            if locked:
                self.log_func("\n[WARN] The following master files are still locked; please close them before opening:\n")
                for mf in locked:
                    self.log_func(f"  • {mf}\n")

            self.root.after(0, lambda: messagebox.showinfo(tr(self.current_lang_code, "Done"), f"{tr(self.current_lang_code, 'Complete')}. {len(insufficient)} {tr(self.current_lang_code, 'skipped')}."))

        threading.Thread(target=worker, daemon=True).start()

    def open_language_picker(self):
        def set_language(lang_code):
            self.current_lang_code = lang_code
            for widget, key in self.translatable_widgets:
                widget.config(text=tr(lang_code, key))
            self.lang_button.config(text=tr(lang_code, "Language"))
            self.update_checkbox_texts(lang_code)
        show_language_picker(self.root, set_language, self.current_lang_code)

def launch_gui():
    root = tk.Tk()
    app = QLSApp(root)
    root.mainloop()
