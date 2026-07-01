import os
import re
import sqlite3
import json
import shutil
import sys
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from PIL import Image, ImageTk, ImageEnhance

try:
    from pdf2image import convert_from_path
except ImportError:
    convert_from_path = None

try:
    import pytesseract
    from pytesseract import Output
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

try:
    import cv2
    import numpy as np
    HAS_CV = True
except ImportError:
    HAS_CV = False

try:
    from licence_check import get_machine_id, activate, validate, is_activated
    HAS_LICENCE = True
except ImportError:
    HAS_LICENCE = False

EXCEL_FILE = "Processed_Vendor_Data.xlsx"
# Database in AppData folder for proper permissions in Program Files
_APP_DATA_DIR = os.path.join(os.getenv('APPDATA'), 'Receipts_OCR_Tool')
if not os.path.exists(_APP_DATA_DIR):
    os.makedirs(_APP_DATA_DIR)
DATABASE_FILE = os.path.join(_APP_DATA_DIR, 'doc_data.db')
BACKUP_DIR = "backups"
CONFIG_FILE = "app_config.json"

GRID_COLS = ("code", "desc", "qty", "unit", "total", "vat_amount")
GRID_HEAD = {"code": "Item Code", "desc": "Description", "qty": "Qty",
             "unit": "Unit Price", "total": "Line Total", "vat_amount": "VAT"}
FIELD_KEYS = ("vendor", "vendor_num", "date", "vat", "phone", "cell",
              "account_num", "ref", "address", "email", "contact")


class AccountingScannerApp:
    def __init__(self, root):
        self.root = root
        root.title("Receipts OCR Tool - Purple Cow Accounting")
        root.geometry("1800x1000")
        root.withdraw()  # Hide main window until licence + TOS done
        
        self.pages = []
        self.page_index = 0
        self.original_img = None
        self.display_img = None
        self.zoom_level = 1.0
        self.current_doc_index = 0
        self.document_pipeline = {0: self.blank_doc()}
        self.current_doc_index = 0
        self.page_data = {}  # Stores document_pipeline per page index
        self.active_targeting_field = None
        self.active_editor = None
        self.rect_start_x = None
        self.rect_start_y = None
        self.crop_rect = None
        self.fbtns = {}
        self.entries = {}
        self.tree = None
        self.totals_lbl = None
        self.orient_var = None
        self.canvas = None
        self.controls = None
        self.db = None
        self.edit_popup = None
        
        self.init_db()
        self.check_licence()
        self.check_tos()
        self.root.deiconify()  # Show main window now that checks are done
        self.setup_ui()

    def init_db(self):
        try:
            self.db = sqlite3.connect(DATABASE_FILE)
            self.db.execute("""CREATE TABLE IF NOT EXISTS vendors(
                vendor TEXT PRIMARY KEY, vendor_num TEXT, date TEXT, vat TEXT,
                phone TEXT, cell TEXT, account_num TEXT, ref TEXT, address TEXT,
                email TEXT, contact TEXT)""")
            cursor = self.db.execute("PRAGMA table_info(vendors)")
            columns = [row[1] for row in cursor.fetchall()]
            for col in ["email", "contact"]:
                if col not in columns:
                    self.db.execute(f"ALTER TABLE vendors ADD COLUMN {col} TEXT")
            self.db.commit()
        except Exception as e:
            messagebox.showerror("DB Error", str(e))
            self.db = None

    def check_licence(self):
        """Block app startup if no valid licence is present."""
        if not HAS_LICENCE:
            messagebox.showerror(
                "Application Error",
                "A required application file is missing or damaged.\n\n"
                "Please reinstall the application or contact Purple Cow Accounting.\n\n"
                "OCR@purplecow.site  |  +27 608 888 812"
            )
            self.root.destroy()
            sys.exit(1)
        if is_activated():
            return  # Already licenced - carry on
        self._show_activation()

    def _show_activation(self):
        """Show activation screen. App cannot proceed without a valid key."""
        machine_id = get_machine_id()

        win = tk.Toplevel(self.root)
        win.title("Activate Receipts OCR Tool")
        win.resizable(False, False)
        
        # Graceful close handler - exit app if user closes without activating
        def on_close():
            win.destroy()
            self.root.quit()
        
        win.protocol("WM_DELETE_WINDOW", on_close)

        # Centre on screen
        win.update_idletasks()
        w, h = 500, 440
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")

        # Force to front
        win.lift()
        win.attributes("-topmost", True)
        win.after(100, lambda: win.attributes("-topmost", False))
        win.grab_set()
        win.focus_force()

        # Header with branding
        header_frame = tk.Frame(win, bg="#6B2FA0", height=80)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="Receipts OCR Tool", font=("Arial", 16, "bold"),
                 fg="white", bg="#6B2FA0").pack(pady=(10, 2))
        tk.Label(header_frame, text="by Purple Cow Accounting",
                 font=("Arial", 10), fg="#D4AF37", bg="#6B2FA0").pack(pady=(0, 10))

        tk.Label(win, text="Activate your licence",
                 font=("Arial", 11, "bold"), fg="#2D3142").pack(pady=(20, 2))
        tk.Label(win,
                 text="Purchase or activate your licence from:",
                 font=("Arial", 10), fg="#555").pack(pady=(6, 2))
        tk.Label(win, text="OCR@purplecow.site  |  +27 608 888 812",
                 font=("Arial", 10, "bold"), fg="#6B2FA0").pack()

        tk.Frame(win, height=1, bg="#ddd").pack(fill=tk.X, padx=20, pady=12)

        # Machine ID display (customer copies this to send to Clare)
        tk.Label(win, text="Your Machine ID (send this to us):",
                 font=("Arial", 9, "bold")).pack(anchor=tk.W, padx=30)
        mid_frame = tk.Frame(win)
        mid_frame.pack(fill=tk.X, padx=30, pady=(2, 10))
        mid_var = tk.StringVar(value=machine_id)
        mid_entry = tk.Entry(mid_frame, textvariable=mid_var, state="readonly",
                             font=("Courier", 9), fg="#333", bg="#f0f0f0")
        mid_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        def copy_mid():
            win.clipboard_clear()
            win.clipboard_append(machine_id)
            copy_btn.config(text="Copied!", bg="#4CAF50")
            win.after(2000, lambda: copy_btn.config(text="Copy", bg="#e0e0e0"))
        copy_btn = tk.Button(mid_frame, text="Copy", command=copy_mid,
                             bg="#e0e0e0", font=("Arial", 8), width=6)
        copy_btn.pack(side=tk.LEFT, padx=(4, 0))

        # Customer name entry
        tk.Label(win, text="Customer Name (as provided):",
                 font=("Arial", 9, "bold")).pack(anchor=tk.W, padx=30)
        name_entry = tk.Entry(win, width=40, font=("Arial", 10))
        name_entry.pack(padx=30, pady=(2, 8), fill=tk.X)

        # Licence key entry
        tk.Label(win, text="Licence Key:",
                 font=("Arial", 9, "bold")).pack(anchor=tk.W, padx=30)
        key_entry = tk.Entry(win, width=40, font=("Courier", 10))
        key_entry.pack(padx=30, pady=(2, 12), fill=tk.X)

        status_lbl = tk.Label(win, text="", font=("Arial", 9), fg="red")
        status_lbl.pack()

        def try_activate():
            customer = name_entry.get().strip()
            key = key_entry.get().strip().upper()
            if not customer or not key:
                status_lbl.config(text="Please enter both your name and licence key.")
                return
            try:
                status_lbl.config(text="Validating...", fg="blue")
                win.update()
                ok, msg = activate(key, customer)
                if ok:
                    status_lbl.config(text="Activated for " + msg + "!", fg="green")
                    win.after(1200, win.destroy)
                else:
                    status_lbl.config(
                        text="Invalid key. Check your name and key match exactly.",
                        fg="red")
            except Exception as e:
                status_lbl.config(text="Error: " + str(e), fg="red")
        tk.Button(win, text="Activate", command=try_activate,
                  bg="#6B2FA0", fg="white", font=("Arial", 10, "bold"),
                  width=20).pack(pady=(4, 0))

        self.root.wait_window(win)

        # After window closes, check if activated - if not, quit
        if not is_activated():
            self.root.destroy()
            sys.exit(0)

    def check_tos(self):
        try:
            with open(CONFIG_FILE, 'r') as f:
                if json.load(f).get("tos_accepted"):
                    return
        except:
            pass
        self._show_tos()

    def _show_tos(self):
        tos = tk.Toplevel(self.root)
        tos.title("Terms of Service")
        tos.resizable(False, False)
        tos.protocol("WM_DELETE_WINDOW", lambda: self.root.destroy())

        # Centre on screen
        tos.update_idletasks()
        w, h = 700, 600
        sw = tos.winfo_screenwidth()
        sh = tos.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        tos.geometry(f"{w}x{h}+{x}+{y}")

        tos.lift()
        tos.attributes("-topmost", True)
        tos.after(100, lambda: tos.attributes("-topmost", False))
        tos.grab_set()
        tos.focus_force()
        txt = tk.Text(tos, wrap=tk.WORD, bg="#f9f9f9", font=("Arial", 9))
        txt.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        txt.insert("1.0", """RECEIPTS OCR PROGRAM - TERMS OF SERVICE

1. ACCURACY DISCLAIMER
OCR is not 100% accurate. You are responsible for verifying all extracted data.

2. NO LIABILITY
Not liable for data loss, extraction errors, or financial damages.

3. LOCAL PROCESSING
All data stays on your computer.

4. YOUR RESPONSIBILITY
Verify all data before use in financial records.""")
        txt.config(state=tk.DISABLED)
        def accept():
            with open(CONFIG_FILE, 'w') as f:
                json.dump({"tos_accepted": True}, f)
            tos.destroy()
        tk.Button(tos, text="I Agree", command=accept, bg="#4CAF50", fg="white", width=20).pack(pady=10)

    def vendor_lookup(self, name):
        if not self.db or not name.strip():
            return None
        try:
            return self.db.execute(
                "SELECT vendor_num, date, vat, phone, cell, account_num, ref, address, email, contact FROM vendors WHERE vendor=?",
                (name,)).fetchone()
        except:
            return None

    def vendor_save(self, vendor, vendor_num, date, vat, phone, cell, account_num, ref, address, email, contact):
        if not self.db or not vendor.strip():
            return
        try:
            self.db.execute(
                "INSERT OR REPLACE INTO vendors VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                (vendor, vendor_num, date, vat, phone, cell, account_num, ref, address, email, contact))
            self.db.commit()
        except Exception as e:
            messagebox.showerror("DB Error", str(e))

    def blank_doc(self):
        return {"orientation": "South", "fields": {"vendor": "", "vendor_num": "", "date": "", "vat": "",
                                    "phone": "", "cell": "", "account_num": "", "ref": "",
                                    "address": "", "po_box": "", "email": "", "contact": ""},
                "boxes": {k: False for k in ["vendor", "vendor_num", "date", "vat", "phone", "cell",
                                             "account_num", "ref", "address", "po_box", "email", "contact"]},
                "lines": []}

    def setup_ui(self):
        # LEFT PANEL
        # MENU BAR
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Licence Status", command=self.show_licence_status)
        help_menu.add_separator()
        help_menu.add_command(label="Check Licence", command=self.check_licence_menu)
        
        self.left = tk.Frame(self.root, width=500, bg="#f5f5f5")
        self.left.pack(side=tk.LEFT, fill=tk.BOTH, padx=6, pady=6)
        self.left.pack_propagate(False)  # Keep at 500px width
        
        # RIGHT PANEL
        self.right = tk.Frame(self.root, bg="#fff")
        self.right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=6, pady=6)

        # TOP SECTION
        topf = tk.Frame(self.left, bg="#f5f5f5")
        topf.pack(fill=tk.X, pady=4)
        tk.Button(topf, text="- Open PDF/Image", command=self.load_file,
                  bg="#D4AF37", fg="#2D3142", font=("Arial", 9, "bold")).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        tk.Button(topf, text="- Reset", command=self.reset,
                  bg="#6B2FA0", fg="white", font=("Arial", 9, "bold")).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        # PAGE NAV
        pf = tk.Frame(self.left, bg="#f5f5f5")
        pf.pack(fill=tk.X, pady=3)
        tk.Button(pf, text="- Page", command=self.prev_page, width=7).pack(side=tk.LEFT, padx=2)
        self.page_lbl = tk.Label(pf, text="1/1", font=("Arial", 9), bg="#f5f5f5")
        self.page_lbl.pack(side=tk.LEFT, expand=True)
        tk.Button(pf, text="Page -", command=self.next_page, width=7).pack(side=tk.LEFT, padx=2)

        # ZOOM
        zf = tk.LabelFrame(self.left, text="Zoom", bg="#f5f5f5", font=("Arial", 8))
        zf.pack(fill=tk.X, pady=2)
        zb = tk.Frame(zf, bg="#f5f5f5")
        zb.pack(pady=3)
        tk.Button(zb, text="-", width=4, command=lambda: self.adjust_zoom(0.2), font=("Arial", 10)).pack(side=tk.LEFT, padx=2)
        tk.Button(zb, text="-", width=4, command=lambda: self.adjust_zoom(-0.2), font=("Arial", 10)).pack(side=tk.LEFT, padx=2)
        tk.Button(zb, text="Fit", width=4, command=self.zoom_fit, font=("Arial", 10)).pack(side=tk.LEFT, padx=2)

        # TABS
        nb = ttk.Notebook(self.left)
        nb.pack(fill=tk.X, pady=2)
        t1 = tk.Frame(nb)
        t2 = tk.Frame(nb)
        nb.add(t1, text="Single")
        nb.add(t2, text="Multiple")
        t1.config(bg="#6B2FA0")
        t2.config(bg="#6B2FA0")
        tk.Label(t1, text="Default: 1 doc per page", font=("Arial", 9), bg="#6B2FA0", fg="white").pack(pady=4)
        tk.Label(t2, text="Docs per page:", font=("Arial", 9), bg="#6B2FA0", fg="white").pack(anchor=tk.W, pady=2, padx=4)
        self.count_entry = tk.Entry(t2, width=6, font=("Arial", 9), bg="white", fg="#2D3142")
        self.count_entry.insert(0, "2")
        self.count_entry.pack(anchor=tk.W, padx=4)
        tk.Button(t2, text="Initialize Slots", command=self.build_slots, bg="#6B2FA0", fg="white", font=("Arial", 9)).pack(fill=tk.X, pady=2, padx=2)

        # DOC LABEL
        self.doc_lbl = tk.Label(self.left, text="Doc [1/1]", font=("Arial", 10, "bold"), fg="#E91E63", bg="#f5f5f5")
        self.doc_lbl.pack(anchor=tk.W, pady=3)

        # SCROLLABLE AREA
        scroll_f = tk.Frame(self.left, bg="#f5f5f5")
        scroll_f.pack(fill=tk.BOTH, expand=True, pady=3)
        
        scr = tk.Canvas(scroll_f, bg="#f5f5f5", highlightthickness=0)
        sb = tk.Scrollbar(scroll_f, orient=tk.VERTICAL, command=scr.yview)
        self.controls = tk.Frame(scr, bg="#f5f5f5")
        self.controls.bind("<Configure>", lambda e: scr.configure(scrollregion=scr.bbox("all")))
        scr.create_window((0, 0), window=self.controls, anchor=tk.NW)
        scr.config(yscrollcommand=sb.set)
        scr.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.render_form()

        # BOTTOM BUTTONS
        exp_f = tk.Frame(self.left, bg="#f5f5f5")
        exp_f.pack(fill=tk.X, side=tk.BOTTOM, pady=3)
        tk.Button(exp_f, text="- Create Backup", command=self.backup, bg="#6B2FA0", fg="white", font=("Arial", 9, "bold")).pack(fill=tk.X, pady=2)
        tk.Button(exp_f, text="-- Export to Excel", command=self.export_menu,
                  bg="#D4AF37", fg="#2D3142", font=("Arial", 10, "bold")).pack(fill=tk.X, pady=2)

        # CANVAS
        self.canvas_title = tk.Label(self.right, text="Canvas - Select regions by dragging", font=("Arial", 10, "bold"))
        self.canvas_title.pack(pady=4)
        cw = tk.Frame(self.right)
        cw.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(cw, bg="grey20", cursor="crosshair")
        vb = tk.Scrollbar(cw, orient=tk.VERTICAL, command=self.canvas.yview)
        hb = tk.Scrollbar(cw, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.config(yscrollcommand=vb.set, xscrollcommand=hb.set)
        vb.pack(side=tk.RIGHT, fill=tk.Y)
        hb.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.canvas.bind("<ButtonPress-1>", self.m_down)
        self.canvas.bind("<B1-Motion>", self.m_drag)
        self.canvas.bind("<ButtonRelease-1>", self.m_up)
        self.canvas.bind("<MouseWheel>", self.wheel)
        
        self.root.bind("<Control-s>", lambda e: self.export_menu())
        self.root.bind("<Control-n>", lambda e: self.reset())

    def wheel(self, e):
        if e.state & 0x0004:
            self.adjust_zoom(0.2 if e.delta > 0 else -0.2)
        else:
            self.canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")

    def render_form(self):
        for w in self.controls.winfo_children():
            w.destroy()
        if not self.document_pipeline:
            self.document_pipeline = {0: self.blank_doc()}
            self.current_doc_index = 0
        st = self.document_pipeline[self.current_doc_index]

        # ORIENTATION
        of = tk.LabelFrame(self.controls, text="Orientation", bg="#f5f5f5", font=("Arial", 9, "bold"))
        of.pack(fill=tk.X, pady=4, padx=4)
        self.orient_var = tk.StringVar(value=st["orientation"])
        ob = tk.Frame(of, bg="#f5f5f5")
        ob.pack(fill=tk.X, padx=4)
        for t, v in [("- South", "South"), ("- North", "North"), ("- West", "West"), ("- East", "East")]:
            tk.Radiobutton(ob, text=t, variable=self.orient_var, value=v, bg="#f5f5f5", font=("Arial", 9),
                           command=lambda s=st: s.update({"orientation": self.orient_var.get()})).pack(side=tk.LEFT, expand=True)

        # FIELDS
        ff = tk.LabelFrame(self.controls, text="Metadata (- = Click to OCR)", bg="#f5f5f5", font=("Arial", 9, "bold"))
        ff.pack(fill=tk.X, pady=4, padx=4)
        self.entries = {}
        self.fbtns = {}
        field_list = [("Vendor Name", "vendor"), ("Vendor #", "vendor_num"), ("Date", "date"),
                      ("VAT Number", "vat"), ("Telephone", "phone"), ("Cell Phone", "cell"),
                      ("Account #", "account_num"), ("Reference", "ref"), ("Physical Address", "address"),
                      ("P.O. Box", "po_box"), ("Email", "email"), ("Contact", "contact")]
        for label, key in field_list:
            r = tk.Frame(ff, bg="#f5f5f5")
            r.pack(fill=tk.X, pady=3, padx=4)
            tk.Label(r, text=label, width=12, anchor=tk.W, font=("Arial", 9)).pack(side=tk.LEFT)
            if key in ["address", "po_box"]:
                w = tk.Text(r, height=3, width=40, font=("Arial", 9), wrap=tk.WORD)
                w.insert("1.0", st["fields"][key])
            else:
                w = tk.Entry(r, width=35, font=("Arial", 9))
                w.insert(0, st["fields"][key])
            w.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
            self.entries[key] = w
            btn = tk.Button(r, text="-", width=2, font=("Arial", 9),
                            command=lambda k=key: self.target(k), bg="#e0e0e0")
            btn.pack(side=tk.LEFT, padx=2)
            self.fbtns[key] = btn
            if st["boxes"][key]:
                btn.config(bg="#a5d6a7")

        tk.Button(ff, text="Auto-Fill from Vendor Memory", command=self.vendor_autofill,
                  bg="#4CAF50", fg="white", font=("Arial", 9)).pack(fill=tk.X, pady=3, padx=4)

        # TRANSACTIONS
        tf = tk.LabelFrame(self.controls, text="Transaction Lines (Double-click cell to edit | - for OCR)", bg="#f5f5f5", font=("Arial", 9, "bold"))
        tf.pack(fill=tk.X, pady=4, padx=4)
        
        self.tree = ttk.Treeview(tf, columns=GRID_COLS, show="headings", height=5)
        for c in GRID_COLS:
            self.tree.heading(c, text=GRID_HEAD[c])
            w = 60 if c == "vat_amount" else 75
            self.tree.column(c, width=w, anchor=tk.W)
        self.tree.pack(fill=tk.BOTH, padx=4, pady=3)
        self.tree.bind("<Double-1>", self.tree_double_click)
        
        # LOAD ROWS
        for row in st.get("lines", []):
            vals = tuple(row.get(c, "") for c in GRID_COLS)
            self.tree.insert("", tk.END, values=vals)
        
        # ADD FIRST BLANK ROW IF GRID IS EMPTY
        if not self.tree.get_children():
            self.tree.insert("", tk.END, values=tuple("" for _ in GRID_COLS))

        # TOTALS
        tot_f = tk.Frame(tf, bg="#e8f5e9", relief=tk.SUNKEN, borderwidth=2)
        tot_f.pack(fill=tk.X, padx=4, pady=3)
        self.totals_lbl = tk.Label(tot_f, text="", font=("Arial", 10, "bold"), bg="#e8f5e9", fg="#1565C0")
        self.totals_lbl.pack(anchor=tk.E, padx=8, pady=4)
        self.update_totals()

        # ROW BUTTONS
        rb = tk.Frame(tf, bg="#f5f5f5")
        rb.pack(fill=tk.X, padx=4, pady=3)
        tk.Button(rb, text="- Add Row", command=self.add_row, font=("Arial", 9)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        tk.Button(rb, text="- Delete Row", command=self.del_row, font=("Arial", 9)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        # NAV
        nf = tk.Frame(self.controls, bg="#f5f5f5")
        nf.pack(fill=tk.X, padx=4, pady=4)
        tk.Button(nf, text="- Previous Document", command=lambda: self.shift("prev"), font=("Arial", 9)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        tk.Button(nf, text="Next Document -", command=lambda: self.shift("next"), font=("Arial", 9)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self._setup_field_navigation()  # Enable keyboard navigation


    def _setup_field_navigation(self):
        """Bind Enter key to move between fields smoothly."""
        field_order = ["vendor", "vendor_num", "date", "vat", "phone", "cell", 
                       "account_num", "ref", "address", "po_box", "email", "contact"]
        
        def make_next_handler(current_key):
            def on_enter(event):
                current_idx = field_order.index(current_key)
                next_idx = (current_idx + 1) % len(field_order)  # Loop to first field
                next_key = field_order[next_idx]
                if next_key in self.entries:
                    self.entries[next_key].focus()
                return "break"  # Prevent default Enter behavior
            return on_enter
        
        # Bind Enter to all Entry fields
        for key in field_order:
            if key in self.entries:
                widget = self.entries[key]
                # For Entry widgets, bind Return
                if isinstance(widget, tk.Entry):
                    widget.bind("<Return>", make_next_handler(key))
                    widget.bind("<Tab>", lambda e, k=key: self._focus_next_field(k))
                # For Text widget (address), also bind Return to move next
                elif isinstance(widget, tk.Text):
                    widget.bind("<Control-Return>", make_next_handler(key))
                    # Ctrl+Return moves to next field; regular Return adds newline
    
    def _focus_next_field(self, current_key):
        """Focus the next field when Tab is pressed."""
        field_order = ["vendor", "vendor_num", "date", "vat", "phone", "cell", 
                       "account_num", "ref", "address", "po_box", "email", "contact"]
        current_idx = field_order.index(current_key)
        next_idx = (current_idx + 1) % len(field_order)
        next_key = field_order[next_idx]
        if next_key in self.entries:
            self.entries[next_key].focus()
        return "break"


    def tree_double_click(self, event):
        """Handle double-click on tree cell"""
        item = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not item or not col_id:
            return
        
        col_idx = int(col_id[1:]) - 1
        if col_idx < 0 or col_idx >= len(GRID_COLS):
            return
        col_key = GRID_COLS[col_idx]
        
        # Get current value
        current_value = self.tree.set(item, col_key)
        
        # Create popup editor
        popup = tk.Toplevel(self.root)
        popup.title(f"Edit {GRID_HEAD[col_key]}")
        popup.geometry("400x150")
        popup.grab_set()
        
        tk.Label(popup, text=f"Enter value for {GRID_HEAD[col_key]}:", font=("Arial", 9)).pack(pady=10)
        
        entry = tk.Entry(popup, width=40, font=("Arial", 10))
        entry.insert(0, current_value)
        entry.pack(pady=5)
        entry.focus_set()
        entry.select_range(0, tk.END)
        entry.bind("<Return>", lambda e: save_and_move_right())
        
        btn_frame = tk.Frame(popup)
        btn_frame.pack(pady=10)
        
        def save_value():
            self.tree.set(item, col_key, entry.get())
            self.update_totals()
            self.save_form()
            popup.destroy()
        
        def save_and_move_right():
            """Save current cell and move to next column (right)."""
            save_value()
            # Move to next column
            col_idx = GRID_COLS.index(col_key)
            next_col_idx = (col_idx + 1) % len(GRID_COLS)  # Loop back to first
            next_col = GRID_COLS[next_col_idx]
            # Open the next cell for editing
            self.root.after(100, lambda: self._edit_tree_cell(item, next_col))
        
        def ocr_value():
            popup.destroy()
            self.active_targeting_field = ("grid_ocr", item, col_key, entry)
            self.canvas_title.config(text=f">>> Draw box for [{col_key}] - Release to OCR <<<")
        
        tk.Button(btn_frame, text="Save", command=save_value, bg="#4CAF50", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="OCR -", command=ocr_value, bg="#2196F3", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Cancel", command=popup.destroy, bg="#f44336", fg="white", width=10).pack(side=tk.LEFT, padx=5)

    def update_totals(self):
        ta, tv = 0.0, 0.0
        for i in self.tree.get_children():
            v = self.tree.item(i, "values")
            try:
                ta += float(str(v[4] or 0).replace(",", "."))
                tv += float(str(v[5] or 0).replace(",", "."))
            except:
                pass
        self.totals_lbl.config(text=f"Excl VAT: {ta:.2f}  |  VAT: {tv:.2f}  |  Grand Total: {ta+tv:.2f}")

    def add_row(self):
        self.tree.insert("", tk.END, values=["", "", "", "", "", ""])
        self.update_totals()

    def del_row(self):
        for s in self.tree.selection():
            self.tree.delete(s)
        self.update_totals()

    def target(self, key):
        self.canvas_title.config(text=f">>> Draw box for [{key}] - Release to OCR <<<")
        self.active_targeting_field = key
        st = self.document_pipeline[self.current_doc_index]
        for k, b in self.fbtns.items():
            b.config(bg="#a5d6a7" if st["boxes"].get(k) else "#e0e0e0")
        if key in self.fbtns:
            self.fbtns[key].config(bg="#ff8a80")

    def prev_page(self):
        if not self.pages or self.page_index <= 0:
            return
        if not messagebox.askyesno("Prev", "Switch page?"):
            return
        self.save_form()
        self.page_data[self.page_index] = {
            "pipeline": self.document_pipeline,
            "doc_index": self.current_doc_index
        }
        self.page_index -= 1
        self.load_page()

    def next_page(self):
        if not self.pages or self.page_index >= len(self.pages) - 1:
            return
        if not messagebox.askyesno("Next", "Switch page?"):
            return
        self.save_form()
        self.page_data[self.page_index] = {
            "pipeline": self.document_pipeline,
            "doc_index": self.current_doc_index
        }
        self.page_index += 1
        self.load_page()

    def load_page(self):
        self.original_img = self.pages[self.page_index].convert("RGB")
        self.zoom_level = 1.0
        self.page_lbl.config(text=f"{self.page_index+1}/{len(self.pages)}")
        # Restore previously saved data for this page, or start fresh
        if self.page_index in self.page_data:
            saved = self.page_data[self.page_index]
            self.document_pipeline = saved["pipeline"]
            self.current_doc_index = saved["doc_index"]
        else:
            self.document_pipeline = {0: self.blank_doc()}
            self.current_doc_index = 0
        self.refresh()

    def reset(self):
        if not messagebox.askyesno("Reset", "Clear ALL data and canvas?"):
            return
        self.original_img = None
        self.pages = []
        self.page_data = {}
        self.canvas.delete("all")
        self.page_lbl.config(text="1/1")
        self.document_pipeline = {0: self.blank_doc()}
        self.current_doc_index = 0
        self.refresh()

    def build_slots(self):
        try:
            n = max(2, int(self.count_entry.get().strip() or 2))
        except:
            n = 2
        self.document_pipeline = {i: self.blank_doc() for i in range(n)}
        self.current_doc_index = 0
        self.refresh()

    def refresh(self):
        self.doc_lbl.config(text=f"Document [{self.current_doc_index+1} of {len(self.document_pipeline)}]")
        self.render_form()
        self.render_canvas()

    def shift(self, d):
        if not self.document_pipeline:
            return
        self.save_form()
        if d == "next" and self.current_doc_index < len(self.document_pipeline) - 1:
            self.current_doc_index += 1
        elif d == "prev" and self.current_doc_index > 0:
            self.current_doc_index -= 1
        self.refresh()

    def field_get(self, key):
        w = self.entries[key]
        return w.get("1.0", "end").strip() if isinstance(w, tk.Text) else w.get().strip()

    def field_set(self, key, val):
        w = self.entries[key]
        if isinstance(w, tk.Text):
            w.delete("1.0", "end")
            w.insert("1.0", val)
        else:
            w.delete(0, tk.END)
            w.insert(0, val)

    def load_file(self):
        fp = filedialog.askopenfilename(filetypes=[("PDF/Image", "*.pdf *.png *.jpg *.jpeg")])
        if not fp:
            return
        try:
            if fp.lower().endswith(".pdf"):
                self.pages = convert_from_path(fp, dpi=300)
            else:
                self.pages = [Image.open(fp).convert("RGB")]
            self.page_index = 0
            self.load_page()
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    def adjust_zoom(self, d):
        if not self.original_img:
            return
        self.zoom_level = round(max(0.3, min(5.0, self.zoom_level + d)), 2)
        self.render_canvas()

    def zoom_fit(self):
        if not self.original_img:
            return
        cw = self.canvas.winfo_width() or 800
        ch = self.canvas.winfo_height() or 700
        w, h = self.original_img.size
        self.zoom_level = round(max(0.1, min(cw / w, ch / h)), 2)
        self.render_canvas()

    def render_canvas(self):
        if not self.original_img:
            return
        w, h = self.original_img.size
        z = self.zoom_level
        disp = self.original_img.resize((int(w * z), int(h * z)), Image.Resampling.LANCZOS)
        disp = ImageEnhance.Contrast(disp).enhance(1.3)
        self.display_img = ImageTk.PhotoImage(disp)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.display_img)
        self.canvas.config(scrollregion=(0, 0, disp.width, disp.height))
        st = self.document_pipeline.get(self.current_doc_index)
        if st:
            for k, box in st["boxes"].items():
                if box:
                    c = "red"
                    self.canvas.create_rectangle(box[0]*z, box[1]*z, box[2]*z, box[3]*z, outline=c, width=2)

    def m_down(self, e):
        self.rect_start_x = self.canvas.canvasx(e.x)
        self.rect_start_y = self.canvas.canvasy(e.y)
        if self.crop_rect:
            self.canvas.delete(self.crop_rect)
        c = "orange" if isinstance(self.active_targeting_field, tuple) else "red"
        self.crop_rect = self.canvas.create_rectangle(self.rect_start_x, self.rect_start_y,
                                                      self.rect_start_x, self.rect_start_y, outline=c, width=2)

    def m_drag(self, e):
        if self.crop_rect:
            self.canvas.coords(self.crop_rect, self.rect_start_x, self.rect_start_y,
                               self.canvas.canvasx(e.x), self.canvas.canvasy(e.y))

    def m_up(self, e):
        if not self.active_targeting_field or not self.crop_rect:
            if self.crop_rect:
                self.canvas.delete(self.crop_rect)
            return
        z = self.zoom_level
        box = (self.rect_start_x / z, self.rect_start_y / z, self.canvas.canvasx(e.x) / z, self.canvas.canvasy(e.y) / z)
        
        # GRID CELL OCR
        if isinstance(self.active_targeting_field, tuple) and self.active_targeting_field[0] == "grid_ocr":
            _, item, col_key, entry = self.active_targeting_field
            st = self.document_pipeline[self.current_doc_index]
            txt = self.ocr_text(box, st["orientation"])
            self.tree.set(item, col_key, txt)
            self.update_totals()
            self.active_targeting_field = None
            self.crop_rect = None
            self.render_canvas()
            return

        # VENDOR METADATA OCR
        st = self.document_pipeline[self.current_doc_index]
        key = self.active_targeting_field
        st["boxes"][key] = box
        if key in self.fbtns:
            txt = self.ocr_text(box, st["orientation"])
            self.field_set(key, txt)
            self.fbtns[key].config(bg="#a5d6a7")

        self.active_targeting_field = None
        self.crop_rect = None
        self.render_canvas()

    def crop_rotate(self, box, orient):
        x0, y0, x1, y1 = [int(c) for c in box]
        x0, x1 = sorted((x0, x1))
        y0, y1 = sorted((y0, y1))
        w, h = self.original_img.size
        x0, y0, x1, y1 = max(0, x0), max(0, y0), min(w, x1), min(h, y1)
        if x1 <= x0 or y1 <= y0:
            return None
        crop = self.original_img.crop((x0, y0, x1, y1))
        if orient == "North":
            crop = crop.rotate(180, expand=True)
        elif orient == "West":
            crop = crop.rotate(90, expand=True)
        elif orient == "East":
            crop = crop.rotate(270, expand=True)
        return crop

    def preprocess(self, pil):
        """Simple preprocessing for better OCR"""
        img = pil.convert("L")
        img = img.resize((img.width * 2, img.height * 2), Image.Resampling.LANCZOS)
        if HAS_CV:
            arr = np.array(img)
            arr = cv2.bilateralFilter(arr, 9, 75, 75)
            img = Image.fromarray(arr)
        return img

    def ocr_text(self, box, orient):
        """Extract text from box via OCR"""
        if not HAS_OCR or not self.original_img:
            return ""
        crop = self.crop_rotate(box, orient)
        if crop is None:
            return ""
        proc = self.preprocess(crop)
        try:
            txt = pytesseract.image_to_string(proc, config="--psm 6")
            txt = txt.strip()
            txt = re.sub(r'\s+', ' ', txt)
            return txt
        except:
            return ""

    def vendor_autofill(self):
        v = self.field_get("vendor").strip()
        if not v:
            return messagebox.showwarning("Vendor", "Type vendor name first.")
        row = self.vendor_lookup(v)
        if not row:
            messagebox.showinfo("New", "Vendor not in DB; will save on export.")
            return
        self.field_set("vendor_num", row[0] or "")
        self.field_set("date", row[1] or "")
        self.field_set("vat", row[2] or "")
        self.field_set("phone", row[3] or "")
        self.field_set("cell", row[4] or "")
        self.field_set("account_num", row[5] or "")
        self.field_set("ref", row[6] or "")
        self.field_set("address", row[7] or "")
        self.field_set("email", row[8] or "")
        self.field_set("contact", row[9] or "")
        messagebox.showinfo("Loaded", f"Auto-filled from DB for '{v}'.")

    def backup(self):
        if not os.path.exists(EXCEL_FILE):
            return messagebox.showinfo("Backup", "No data file yet.")
        try:
            os.makedirs(BACKUP_DIR, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy(EXCEL_FILE, f"{BACKUP_DIR}/backup_{ts}.xlsx")
            messagebox.showinfo("Backup", "Saved successfully.")
        except Exception as e:
            messagebox.showerror("Backup", str(e))

    def save_form(self):
        if self.current_doc_index not in self.document_pipeline:
            return
        st = self.document_pipeline[self.current_doc_index]
        st["orientation"] = self.orient_var.get()
        field_list = ["vendor", "vendor_num", "date", "vat", "phone", "cell", 
                      "account_num", "ref", "address", "po_box", "email", "contact"]
        for k in field_list:
            if k in self.entries:
                st["fields"][k] = self.field_get(k)
        st["lines"] = [dict(zip(GRID_COLS, self.tree.item(i, "values"))) for i in self.tree.get_children()]


    def _edit_tree_cell(self, item, col_key):
        """Helper method to programmatically edit a tree cell."""
        if item not in self.tree.get_children():
            return
        col_idx = GRID_COLS.index(col_key) + 1
        current_value = self.tree.set(item, col_key)
        
        # Create popup editor
        popup = tk.Toplevel(self.root)
        popup.title(f"Edit {GRID_HEAD[col_key]}")
        popup.geometry("400x150")
        popup.grab_set()
        
        tk.Label(popup, text=f"Enter value for {GRID_HEAD[col_key]}:", font=("Arial", 9)).pack(pady=10)
        
        entry = tk.Entry(popup, width=40, font=("Arial", 10))
        entry.insert(0, current_value)
        entry.pack(pady=5)
        entry.focus_set()
        entry.select_range(0, tk.END)
        
        btn_frame = tk.Frame(popup)
        btn_frame.pack(pady=10)
        
        def save_value():
            self.tree.set(item, col_key, entry.get())
            self.update_totals()
            self.save_form()
            popup.destroy()
        
        def save_and_move_right():
            """Save current cell and move to next column (right)."""
            save_value()
            # Move to next column
            col_idx = GRID_COLS.index(col_key)
            next_col_idx = (col_idx + 1) % len(GRID_COLS)
            next_col = GRID_COLS[next_col_idx]
            self.root.after(100, lambda: self._edit_tree_cell(item, next_col))
        
        entry.bind("<Return>", lambda e: save_and_move_right())
        
        tk.Button(btn_frame, text="Save", command=save_value, bg="#4CAF50", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="OCR -", command=lambda: self._ocr_from_popup(item, col_key, entry, popup), bg="#2196F3", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Cancel", command=popup.destroy, bg="#f44336", fg="white", width=10).pack(side=tk.LEFT, padx=5)
    
    def _ocr_from_popup(self, item, col_key, entry, popup):
        """Handle OCR from grid popup."""
        popup.destroy()
        self.active_targeting_field = ("grid_ocr", item, col_key, entry)
        self.canvas_title.config(text=f">>> Draw box for [{col_key}] - Release to OCR <<<")


    def export_menu(self):
        """Ask user whether to export current page or all pages, then export."""
        self.save_form()
        # Save current page into page_data so it's included if exporting all
        self.page_data[self.page_index] = {
            "pipeline": self.document_pipeline,
            "doc_index": self.current_doc_index
        }

        total_pages = len(self.pages) if self.pages else 1
        has_multiple = total_pages > 1

        if has_multiple:
            choice_win = tk.Toplevel(self.root)
            choice_win.title("Export to Excel")
            choice_win.resizable(False, False)
            choice_win.grab_set()
            choice_win.focus_force()

            w, h = 380, 210
            sw = choice_win.winfo_screenwidth()
            sh = choice_win.winfo_screenheight()
            choice_win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
            choice_win.lift()
            choice_win.attributes("-topmost", True)
            choice_win.after(100, lambda: choice_win.attributes("-topmost", False))

            tk.Label(choice_win, text="What would you like to export?",
                     font=("Arial", 10, "bold")).pack(pady=(20, 6))
            tk.Label(choice_win,
                     text=f"This document has {total_pages} pages.",
                     font=("Arial", 9), fg="#555").pack(pady=(0, 16))

            btn_frame = tk.Frame(choice_win)
            btn_frame.pack()

            export_scope = {"value": None}

            def pick(val):
                export_scope["value"] = val
                choice_win.destroy()

            tk.Button(btn_frame,
                      text=f"This page only  (page {self.page_index + 1})",
                      command=lambda: pick("current"),
                      bg="#008CBA", fg="white",
                      font=("Arial", 9, "bold"), width=30).pack(pady=4)
            tk.Button(btn_frame,
                      text=f"All pages  ({total_pages} pages)",
                      command=lambda: pick("all"),
                      bg="#4CAF50", fg="white",
                      font=("Arial", 9, "bold"), width=30).pack(pady=4)
            tk.Button(btn_frame, text="Cancel",
                      command=lambda: pick(None),
                      bg="#aaa", fg="white",
                      font=("Arial", 9), width=30).pack(pady=4)

            self.root.wait_window(choice_win)

            if export_scope["value"] is None:
                return
            export_all = export_scope["value"] == "all"
        else:
            export_all = False

        if export_all:
            pipelines = [
                self.page_data[i]["pipeline"]
                for i in sorted(self.page_data.keys())
            ]
            scope_label = f"all {total_pages} pages"
        else:
            pipelines = [self.document_pipeline]
            scope_label = f"page {self.page_index + 1}"

        self._do_export(pipelines, scope_label)

    def _collect_rows(self, pipelines):
        """Build list of row dicts from a list of document pipelines."""
        out = []
        for pipeline in pipelines:
            for doc in pipeline.values():
                f = doc["fields"]
                if not f["vendor"].strip():
                    continue
                lines = doc.get("lines") or [{}]
                for ln in lines:
                    out.append({
                        "Vendor Name": f["vendor"], "Vendor#": f["vendor_num"],
                        "Account#": f["account_num"], "VAT#": f["vat"],
                        "Email": f["email"], "Contact": f["contact"],
                        "Date": f["date"], "Address": f["address"], "PO Box": f["po_box"],
                        "Telephone": f["phone"], "Cell": f["cell"],
                        "DocRef": f["ref"], "Code": ln.get("code", ""),
                        "Description": ln.get("desc", ""), "Qty": ln.get("qty", ""),
                        "UnitPrice": ln.get("unit", ""), "LineTotal": ln.get("total", ""),
                        "VATAmount": ln.get("vat_amount", "")
                    })
        return out

    def _do_export(self, pipelines, scope_label):
        """Clean, format and write rows to Excel."""
        out = self._collect_rows(pipelines)
        if not out:
            return messagebox.showerror(
                "Empty",
                f"No vendor data found in {scope_label}.\n"
                "Add at least one Vendor + transaction line.")

        df = pd.DataFrame(out)
        if os.path.exists(EXCEL_FILE):
            try:
                old = pd.read_excel(EXCEL_FILE)
                df = pd.concat([old, df], ignore_index=True)
            except:
                pass

        df["Date"] = pd.to_datetime(
            df["Date"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")

        def clean_phone(x):
            if not x or str(x).strip() == "":
                return ""
            s = str(x).replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
            if s.startswith("+"):
                return s
            if s.startswith("0"):
                return s
            if len(s) in (9, 10):
                return "0" + s
            return s

        for c in ["Telephone", "Cell"]:
            df[c] = df[c].apply(clean_phone)

        for c in ["Account#", "Vendor#"]:
            df[c] = df[c].apply(
                lambda x: str(x).replace(",", "").replace("-", "") if x else "")

        df["Qty"] = df["Qty"].apply(lambda x: int(x) if str(x).isdigit() else x)

        for c in ["UnitPrice", "LineTotal", "VATAmount"]:
            df[c] = df[c].apply(
                lambda x: float(str(x).replace(",", ".").strip()) if x else "")

        for pipeline in pipelines:
            for doc in pipeline.values():
                f = doc["fields"]
                if f["vendor"].strip():
                    self.vendor_save(
                        f["vendor"], f["vendor_num"], f["date"], f["vat"],
                        f["phone"], f["cell"], f["account_num"], f["ref"],
                        f["address"], f["email"], f["contact"])

        try:
            df.to_excel(EXCEL_FILE, index=False)
            row_count = len(out)
            messagebox.showinfo(
                "- Export Complete",
                f"Exported {row_count} row{'s' if row_count != 1 else ''} "
                f"from {scope_label}.\n\nSaved to:\n{os.path.abspath(EXCEL_FILE)}")
        except PermissionError:
            messagebox.showerror("Locked", "Close the Excel file first.")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))




    def show_licence_status(self):
        """Display current licence status and information."""
        if not HAS_LICENCE:
            messagebox.showerror(
                "Application Error",
                "A required application file is missing or damaged.\n\n"
                "Please reinstall the application or contact Purple Cow Accounting.\n\n"
                "OCR@purplecow.site  |  +27 608 888 812"
            )
            return
        
        status_window = tk.Toplevel(self.root)
        status_window.title("Licence Status")
        status_window.geometry("450x350")
        status_window.resizable(False, False)
        
        # Center window
        status_window.update_idletasks()
        w, h = 450, 350
        sw = status_window.winfo_screenwidth()
        sh = status_window.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        status_window.geometry(f"{w}x{h}+{x}+{y}")
        
        # Header
        header = tk.Frame(status_window, bg="#6B2FA0", height=60)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)
        
        tk.Label(header, text="Licence Information", font=("Arial", 14, "bold"),
                 fg="white", bg="#6B2FA0").pack(pady=(12, 8))
        
        # Content
        content = tk.Frame(status_window)
        content.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
        
        # Check if activated
        if is_activated():
            status_label = tk.Label(content, text="- ACTIVATED", font=("Arial", 12, "bold"),
                                   fg="#4CAF50")
            status_label.pack(pady=10)
            
            try:
                from licence_check import _licence_path
                import json
                with open(_licence_path(), 'r') as f:
                    lic_data = json.load(f)
                customer = lic_data.get('customer', 'Unknown')
                key = lic_data.get('key', '')
                
                tk.Label(content, text="Customer:", font=("Arial", 10, "bold"),
                        fg="#2D3142").pack(anchor=tk.W, pady=(10, 2))
                tk.Label(content, text=customer, font=("Arial", 10),
                        fg="#666").pack(anchor=tk.W, pady=(0, 10))
                
                tk.Label(content, text="Licence Key:", font=("Arial", 10, "bold"),
                        fg="#2D3142").pack(anchor=tk.W)
                key_frame = tk.Frame(content)
                key_frame.pack(anchor=tk.W, fill=tk.X, pady=(2, 10))
                tk.Label(key_frame, text=key, font=("Courier", 9),
                        fg="#6B2FA0", bg="#f5f5f5").pack(fill=tk.X, padx=8, pady=4)
                
                tk.Label(content, text="Validity: 12 months from activation",
                        font=("Arial", 9), fg="#666").pack(anchor=tk.W, pady=5)
                
            except Exception as e:
                tk.Label(content, text=f"Error reading licence: {str(e)}",
                        font=("Arial", 9), fg="#f44336").pack()
        else:
            status_label = tk.Label(content, text="- NOT ACTIVATED", font=("Arial", 12, "bold"),
                                   fg="#f44336")
            status_label.pack(pady=10)
            
            tk.Label(content, text="This copy of the Receipts OCR Tool is not activated.",
                    font=("Arial", 10), fg="#666", wraplength=350).pack(pady=10)
            tk.Label(content, text="Contact Purple Cow Accounting to obtain a licence key.",
                    font=("Arial", 9), fg="#666").pack()
            tk.Label(content, text="OCR@purplecow.site | +27 608 888 812",
                    font=("Arial", 9, "bold"), fg="#6B2FA0").pack(pady=(10, 0))
        
        tk.Button(content, text="Close", command=status_window.destroy,
                 bg="#6B2FA0", fg="white", font=("Arial", 9)).pack(pady=(20, 0))

    def check_licence_menu(self):
        """Safe licence check from menu - shows status with option to change key."""
        if not HAS_LICENCE:
            messagebox.showerror(
                "Application Error",
                "A required application file is missing or damaged.\n\n"
                "Please reinstall the application or contact Purple Cow Accounting.\n\n"
                "OCR@purplecow.site  |  +27 608 888 812"
            )
            return
        
        if not is_activated():
            # Not activated - show activation screen
            self._show_activation()
            return
        
        # Already activated - show status with option to change
        check_window = tk.Toplevel(self.root)
        check_window.title("Licence Management")
        check_window.geometry("450x300")
        check_window.resizable(False, False)
        
        # Center window
        check_window.update_idletasks()
        w, h = 450, 300
        sw = check_window.winfo_screenwidth()
        sh = check_window.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        check_window.geometry(f"{w}x{h}+{x}+{y}")
        
        # Header
        header = tk.Frame(check_window, bg="#6B2FA0", height=60)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)
        
        tk.Label(header, text="Current Licence", font=("Arial", 14, "bold"),
                 fg="white", bg="#6B2FA0").pack(pady=(12, 8))
        
        # Content
        content = tk.Frame(check_window)
        content.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
        
        try:
            from licence_check import _licence_path
            import json
            with open(_licence_path(), 'r') as f:
                lic_data = json.load(f)
            customer = lic_data.get('customer', 'Unknown')
            
            tk.Label(content, text="- Activated", font=("Arial", 12, "bold"),
                    fg="#4CAF50").pack(pady=10)
            tk.Label(content, text=f"Customer: {customer}", font=("Arial", 10),
                    fg="#2D3142").pack(pady=5)
            tk.Label(content, text="Machine ID: " + get_machine_id()[:16] + "...",
                    font=("Arial", 9), fg="#666").pack(pady=5)
            tk.Label(content, text="Validity: 12 months from activation",
                    font=("Arial", 9), fg="#666").pack(pady=10)
        except Exception as e:
            tk.Label(content, text="Licence Info", font=("Arial", 10, "bold"),
                    fg="#2D3142").pack(pady=5)
            tk.Label(content, text="- Activated and running",
                    font=("Arial", 10), fg="#4CAF50").pack(pady=10)
        
        # Buttons
        btn_frame = tk.Frame(content)
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="Close", command=check_window.destroy,
                 bg="#6B2FA0", fg="white", font=("Arial", 9), width=15).pack(pady=5)
        tk.Button(btn_frame, text="Enter Different Key", command=lambda: self._enter_different_key(check_window),
                 bg="#D4AF37", fg="#2D3142", font=("Arial", 9), width=15).pack(pady=5)

    def _enter_different_key(self, parent_window):
        """Open activation screen to enter a new licence key."""
        parent_window.destroy()
        self._show_activation()

    def show_about(self):
        """Show About dialog with branding and AI training services."""
        about_window = tk.Toplevel(self.root)
        about_window.title("About Receipts OCR Tool")
        about_window.geometry("450x450")
        about_window.resizable(False, False)
        
        # Center window
        about_window.update_idletasks()
        w, h = 450, 450
        sw = about_window.winfo_screenwidth()
        sh = about_window.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        about_window.geometry(f"{w}x{h}+{x}+{y}")
        
        # Header with branding
        header = tk.Frame(about_window, bg="#6B2FA0", height=60)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)
        
        tk.Label(header, text="Receipts OCR Tool", font=("Arial", 14, "bold"),
                 fg="white", bg="#6B2FA0").pack(pady=(8, 2))
        tk.Label(header, text="v1.0 - by Purple Cow Accounting",
                 font=("Arial", 9), fg="#D4AF37", bg="#6B2FA0").pack(pady=(0, 8))
        
        # Content
        content = tk.Frame(about_window)
        content.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
        
        tk.Label(content, text="Fast OCR processing for receipts & invoices",
                 font=("Arial", 10), fg="#555", wraplength=350).pack(pady=5)
        
        # Divider
        tk.Frame(content, height=1, bg="#ddd").pack(fill=tk.X, pady=10)
        
        # AI Training Services section
        tk.Label(content, text="AI Evaluation & Training", font=("Arial", 11, "bold"),
                 fg="#6B2FA0").pack(anchor=tk.W)
        tk.Label(content, text="We help your business implement AI automation and agentic workflows.",
                 font=("Arial", 9), fg="#666", wraplength=350, justify=tk.LEFT).pack(anchor=tk.W, pady=(3, 8))
        tk.Label(content, text="Services include:",
                 font=("Arial", 9, "bold"), fg="#2D3142").pack(anchor=tk.W)
        
        services = [
            "- AI Integration Training & Setup",
            "- Workflow Automation Consulting",
            "- Custom Agentic Solutions"
        ]
        for service in services:
            tk.Label(content, text=service, font=("Arial", 9), fg="#555",
                    justify=tk.LEFT).pack(anchor=tk.W, pady=1)
        
        # Divider
        tk.Frame(content, height=1, bg="#ddd").pack(fill=tk.X, pady=10)
        
        # Contact
        tk.Label(content, text="Contact & Support", font=("Arial", 11, "bold"),
                 fg="#6B2FA0").pack(anchor=tk.W)
        tk.Label(content, text="Email: OCR@purplecow.site",
                 font=("Arial", 9, "bold"), fg="#6B2FA0").pack(anchor=tk.W, pady=2)
        tk.Label(content, text="Phone: +27 608 888 812",
                 font=("Arial", 9, "bold"), fg="#6B2FA0").pack(anchor=tk.W, pady=2)
        tk.Label(content, text="Address: 40 Claremont Road, Mbombela",
                 font=("Arial", 8), fg="#999").pack(anchor=tk.W, pady=2)
        
        # Close button
        tk.Button(about_window, text="Close", command=about_window.destroy,
                 bg="#6B2FA0", fg="white", font=("Arial", 9)).pack(pady=(10, 0))

if __name__ == "__main__":
    root = tk.Tk()
    app = AccountingScannerApp(root)
    root.mainloop()