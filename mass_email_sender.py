import os
import time
import smtplib
import ssl
import sys
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from tkinter import *
from tkinter import ttk, filedialog, messagebox, scrolledtext
from threading import Thread
from datetime import datetime
import tkinter.font as tkfont

# ─── Color Palette (Light Professional Theme) ─────────────────────────────────
BG = "#f5f7fa"
PANEL = "#ffffff"
CARD = "#ffffff"
ACCENT = "#2563eb"
ACCENT2 = "#7c3aed"
SUCCESS = "#16a34a"
WARNING = "#d97706"
DANGER = "#dc2626"
TEXT = "#334155"
MUTED = "#64748b"
BORDER = "#e2e8f0"
ENTRY_BG = "#ffffff"

FONT_BODY = ("Segoe UI", 11)
FONT_BOLD = ("Segoe UI", 11, "bold")
FONT_TITLE = ("Segoe UI", 20, "bold")
FONT_SMALL = ("Segoe UI", 10)
FONT_MONO = ("Consolas", 10)


# ─── Helpers ───────────────────────────────────────────────────────────────────
def styled_entry(parent, width=30, **kw):
    e = Entry(
        parent,
        width=width,
        bg=ENTRY_BG,
        fg=TEXT,
        insertbackground=TEXT,
        relief=FLAT,
        font=FONT_BODY,
        highlightthickness=1,
        highlightbackground=BORDER,
        highlightcolor=ACCENT,
        **kw,
    )
    return e


def styled_button(parent, text, command=None, kind="primary", width=None, **kw):
    colors = {
        "primary": (ACCENT, "#ffffff"),
        "success": (SUCCESS, "#ffffff"),
        "danger": (DANGER, "#ffffff"),
        "ghost": (CARD, TEXT),
        "warning": (WARNING, "#000000"),
    }
    bg, fg = colors.get(kind, (ACCENT, "#ffffff"))
    b = Button(
        parent,
        text=text,
        command=command,
        bg=bg,
        fg=fg,
        activebackground=bg,
        activeforeground=fg,
        relief=FLAT,
        font=FONT_BOLD,
        cursor="hand2",
        padx=14,
        pady=8,
        bd=0,
        **kw,
    )
    if width:
        b.config(width=width)

    def on_enter(e):
        if kind == "ghost":
            b.config(bg="#f1f5f9", fg=ACCENT)
        else:
            b.config(bg=_lighten(bg))

    def on_leave(e):
        if kind == "ghost":
            b.config(bg=CARD, fg=TEXT)
        else:
            b.config(bg=bg)

    b.bind("<Enter>", on_enter)
    b.bind("<Leave>", on_leave)
    return b


def _lighten(hex_color):
    try:
        r = min(255, int(hex_color[1:3], 16) + 20)
        g = min(255, int(hex_color[3:5], 16) + 20)
        b = min(255, int(hex_color[5:7], 16) + 20)
        return f"#{r:02x}{g:02x}{b:02x}"
    except:
        return hex_color


def section_card(parent, title, **kw):
    outer = Frame(
        parent, bg=CARD, bd=0, highlightthickness=1, highlightbackground=BORDER
    )
    outer.pack(fill=X, padx=18, pady=(0, 12))
    hdr = Frame(outer, bg=ACCENT2, height=2)
    hdr.pack(fill=X)
    Label(
        outer,
        text=title.upper(),
        font=("Segoe UI", 8, "bold"),
        fg=MUTED,
        bg=CARD,
        pady=6,
        padx=14,
        anchor=W,
    ).pack(fill=X)
    inner = Frame(outer, bg=CARD, padx=14, pady=10)
    inner.pack(fill=X)
    return inner


# ─── Main App ──────────────────────────────────────────────────────────────────
class MassEmailSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Bulk Sender")
        self.root.geometry("1150x800")
        self.root.minsize(960, 680)
        self.root.configure(bg=BG)
        self.root.resizable(True, True)

        self._init_state()
        self._build_ui()

    def _init_state(self):
        self.stop_flag = False
        self.pause_flag = False
        self.sent_count = 0
        self.failed_count = 0
        self.total_emails = 0
        self.attachment_files = []
        self.available_sheets = []
        self.selected_sheets = []
        self.is_sending = False
        self.data_df = None  # In-memory DataFrame

    # ── UI BUILDING ────────────────────────────────────────────────────────────
    def _build_ui(self):
        self._build_titlebar()
        self._build_body()

    def _build_titlebar(self):
        bar = Frame(self.root, bg=PANEL, height=64)
        bar.pack(fill=X, side=TOP)
        bar.pack_propagate(False)

        logo_frame = Frame(bar, bg=PANEL)
        logo_frame.pack(side=LEFT, padx=20)
        Label(
            logo_frame,
            text="Email Bulk Sender",
            font=("Segoe UI", 16, "bold"),
            fg=ACCENT,
            bg=PANEL,
        ).pack(side=LEFT)
        Label(
            logo_frame,
            text="  |  Mass Email Sender",
            font=FONT_SMALL,
            fg=MUTED,
            bg=PANEL,
        ).pack(side=LEFT)

        self.stat_sent = self._stat_chip(bar, "Sent", "0", SUCCESS)
        self.stat_fail = self._stat_chip(bar, "Failed", "0", DANGER)
        self.stat_total = self._stat_chip(bar, "Total", "0", ACCENT)

    def _stat_chip(self, parent, label_text, value, color):
        f = Frame(
            parent,
            bg="#f8fafc",
            padx=8,
            pady=2,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        f.pack(side=RIGHT, padx=4, pady=6)
        Label(f, text=label_text, font=FONT_SMALL, fg=MUTED, bg="#f8fafc").pack()
        v = Label(f, text=value, font=("Segoe UI", 13, "bold"), fg=color, bg="#f8fafc")
        v.pack()
        return v

    def _build_body(self):
        body = Frame(self.root, bg=BG)
        body.pack(fill=BOTH, expand=True)

        sidebar = Frame(body, bg=PANEL, width=180)
        sidebar.pack(side=LEFT, fill=Y)
        sidebar.pack_propagate(False)

        Label(
            sidebar,
            text="NAVIGATION",
            font=("Segoe UI", 8, "bold"),
            fg=MUTED,
            bg=PANEL,
            pady=14,
        ).pack()

        self.tabs = {}
        self.active_tab = StringVar(value="smtp")
        tab_defs = [
            ("smtp", "SMTP Setup"),
            ("recipients", "Recipients"),
            ("compose", "Compose"),
            ("attachments", "Attachments"),
            ("send", "Send"),
            ("logs", "Logs"),
        ]
        for key, text in tab_defs:
            self._nav_btn(sidebar, key, text)

        self.content = Frame(body, bg=BG)
        self.content.pack(side=LEFT, fill=BOTH, expand=True)

        self.pages = {}
        self.pages["smtp"] = self._page_smtp()
        self.pages["recipients"] = self._page_recipients()
        self.pages["compose"] = self._page_compose()
        self.pages["attachments"] = self._page_attachments()
        self.pages["send"] = self._page_send()
        self.pages["logs"] = self._page_logs()

        self._show_tab("smtp")

    def _nav_btn(self, parent, key, text):
        def click():
            self._show_tab(key)

        b = Button(
            parent,
            text=text,
            command=click,
            bg=PANEL,
            fg=TEXT,
            activebackground="#f1f5f9",
            activeforeground=ACCENT,
            relief=FLAT,
            font=("Segoe UI", 10),
            anchor=W,
            padx=18,
            pady=12,
            bd=0,
            cursor="hand2",
            width=16,
        )
        b.pack(fill=X)
        self.tabs[key] = b

    def _show_tab(self, key):
        for k, page in self.pages.items():
            page.pack_forget()
        self.pages[key].pack(fill=BOTH, expand=True)
        for k, b in self.tabs.items():
            b.config(
                bg="#f1f5f9" if k == key else PANEL, fg=ACCENT if k == key else TEXT
            )
        self.active_tab.set(key)

    def _scrollable_page(self):
        outer = Frame(self.content, bg=BG)
        canvas = Canvas(outer, bg=BG, bd=0, highlightthickness=0)
        vsb = Scrollbar(outer, orient=VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=RIGHT, fill=Y)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)

        inner = Frame(canvas, bg=BG)
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _resize(e):
            canvas.itemconfig(win_id, width=e.width)

        canvas.bind("<Configure>", _resize)
        inner.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        # Mouse wheel / Linux scroll support
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<MouseWheel>", _on_mousewheel)
        inner.bind("<MouseWheel>", _on_mousewheel)

        def _on_linux_up(e):
            canvas.yview_scroll(-1, "units")

        def _on_linux_down(e):
            canvas.yview_scroll(1, "units")

        canvas.bind("<Button-4>", _on_linux_up)
        canvas.bind("<Button-5>", _on_linux_down)
        inner.bind("<Button-4>", _on_linux_up)
        inner.bind("<Button-5>", _on_linux_down)

        def _focus_canvas(e):
            canvas.focus_set()

        canvas.bind("<Enter>", _focus_canvas)
        inner.bind("<Enter>", _focus_canvas)

        return outer, inner

    # ────────────────────────── PAGE 1: SMTP ──────────────────────────────────
    def _page_smtp(self):
        outer, pg = self._scrollable_page()
        Label(
            pg,
            text="SMTP Configuration",
            font=FONT_TITLE,
            fg=TEXT,
            bg=BG,
            pady=18,
            padx=18,
            anchor=W,
        ).pack(fill=X)

        s = section_card(pg, "Quick Provider Presets")
        presets = [
            ("Gmail", "smtp.gmail.com", "587", "tls"),
            ("Outlook", "smtp-mail.outlook.com", "587", "tls"),
            ("Yahoo", "smtp.mail.yahoo.com", "465", "ssl"),
            ("Office365", "smtp.office365.com", "587", "tls"),
            ("Custom", "", "", "tls"),
        ]
        pf = Frame(s, bg=CARD)
        pf.pack(fill=X, pady=(0, 4))
        for name, srv, port, enc in presets:
            styled_button(
                pf,
                name,
                command=lambda sv=srv, po=port, en=enc: self._apply_preset(sv, po, en),
                kind="ghost",
            ).pack(side=LEFT, padx=(0, 6))

        s2 = section_card(pg, "Server & Credentials")
        rows = Frame(s2, bg=CARD)
        rows.pack(fill=X)
        Label(
            rows, text="SMTP Server", font=FONT_SMALL, fg=MUTED, bg=CARD, anchor=W
        ).grid(row=0, column=0, sticky=W, pady=(0, 2))
        self.smtp_server = styled_entry(rows, width=40)
        self.smtp_server.insert(0, "smtp.gmail.com")
        self.smtp_server.grid(row=1, column=0, sticky=EW, padx=(0, 14), pady=(0, 10))
        Label(rows, text="Port", font=FONT_SMALL, fg=MUTED, bg=CARD, anchor=W).grid(
            row=0, column=1, sticky=W
        )
        self.smtp_port = styled_entry(rows, width=8)
        self.smtp_port.insert(0, "587")
        self.smtp_port.grid(row=1, column=1, sticky=W, padx=(0, 14), pady=(0, 10))
        Label(
            rows, text="Encryption", font=FONT_SMALL, fg=MUTED, bg=CARD, anchor=W
        ).grid(row=0, column=2, sticky=W)
        self.ssl_var = StringVar(value="tls")
        ef = Frame(rows, bg=CARD)
        ef.grid(row=1, column=2, sticky=W, pady=(0, 10))
        for val, lbl in [("tls", "TLS"), ("ssl", "SSL")]:
            Radiobutton(
                ef,
                text=lbl,
                variable=self.ssl_var,
                value=val,
                bg=CARD,
                fg=TEXT,
                selectcolor=ENTRY_BG,
                activebackground=CARD,
                font=FONT_SMALL,
            ).pack(side=LEFT, padx=4)
        rows.columnconfigure(0, weight=1)

        Label(
            rows, text="Email Address", font=FONT_SMALL, fg=MUTED, bg=CARD, anchor=W
        ).grid(row=2, column=0, sticky=W, pady=(0, 2))
        self.smtp_email = styled_entry(rows, width=40)
        self.smtp_email.grid(row=3, column=0, sticky=EW, padx=(0, 14), pady=(0, 10))
        Label(
            rows,
            text="Password / App Password",
            font=FONT_SMALL,
            fg=MUTED,
            bg=CARD,
            anchor=W,
        ).grid(row=2, column=1, columnspan=2, sticky=W)
        self.smtp_pass = styled_entry(rows, width=32, show="•")
        self.smtp_pass.grid(row=3, column=1, columnspan=2, sticky=EW, pady=(0, 10))

        tip = Frame(
            s2,
            bg="#f0f9ff",
            padx=10,
            pady=8,
            highlightthickness=1,
            highlightbackground="#bae6fd",
        )
        tip.pack(fill=X, pady=(6, 0))
        Label(
            tip,
            text="Tip: Gmail users should enable 2-Step Verification → generate an App Password at myaccount.google.com → use it here.",
            font=FONT_SMALL,
            fg="#0369a1",
            bg="#f0f9ff",
            wraplength=700,
            justify=LEFT,
        ).pack(anchor=W)
        Label(
            tip,
            text="   Also enable: Allow less secure apps OR use App Password. Bulk sending works best with Google Workspace.",
            font=FONT_SMALL,
            fg=MUTED,
            bg="#f0f9ff",
            wraplength=700,
            justify=LEFT,
        ).pack(anchor=W, pady=(2, 0))

        s3 = section_card(pg, "Sender Identity")
        r3 = Frame(s3, bg=CARD)
        r3.pack(fill=X)
        Label(
            r3, text="Sender Name", font=FONT_SMALL, fg=MUTED, bg=CARD, anchor=W
        ).grid(row=0, column=0, sticky=W, pady=(0, 2))
        self.sender_name = styled_entry(r3, width=30)
        self.sender_name.insert(0, "Your Name / Organization")
        self.sender_name.grid(row=1, column=0, sticky=EW, padx=(0, 14))
        Label(
            r3, text="Reply-To (optional)", font=FONT_SMALL, fg=MUTED, bg=CARD, anchor=W
        ).grid(row=0, column=1, sticky=W)
        self.reply_to = styled_entry(r3, width=35)
        self.reply_to.grid(row=1, column=1, sticky=EW)
        r3.columnconfigure(0, weight=1)
        r3.columnconfigure(1, weight=1)

        bf = Frame(pg, bg=BG, pady=12, padx=18)
        bf.pack(fill=X)
        styled_button(
            bf, "Test Connection", command=self.test_connection, kind="primary"
        ).pack(side=LEFT)
        styled_button(
            bf,
            "Next →  Recipients",
            command=lambda: self._show_tab("recipients"),
            kind="ghost",
        ).pack(side=RIGHT)
        return outer

    def _apply_preset(self, srv, port, enc):
        self.smtp_server.delete(0, END)
        self.smtp_server.insert(0, srv)
        self.smtp_port.delete(0, END)
        self.smtp_port.insert(0, port)
        self.ssl_var.set(enc)

    # ─────────────────────── PAGE 2: RECIPIENTS (EDITABLE) ─────────────────────
    def _page_recipients(self):
        outer, pg = self._scrollable_page()
        Label(
            pg,
            text="Recipients",
            font=FONT_TITLE,
            fg=TEXT,
            bg=BG,
            pady=18,
            padx=18,
            anchor=W,
        ).pack(fill=X)

        s = section_card(pg, "Excel / CSV File")
        ff = Frame(s, bg=CARD)
        ff.pack(fill=X)
        self.file_path_var = StringVar()
        fe = Entry(
            ff,
            textvariable=self.file_path_var,
            font=FONT_SMALL,
            bg=ENTRY_BG,
            fg=TEXT,
            relief=FLAT,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        fe.pack(side=LEFT, fill=X, expand=True, padx=(0, 8), ipady=6)
        styled_button(ff, "Browse", command=self.browse_file, kind="primary").pack(
            side=LEFT
        )
        self.file_status_lbl = Label(
            s,
            text="No file selected",
            font=FONT_SMALL,
            fg=MUTED,
            bg=CARD,
            anchor=W,
            pady=4,
        )
        self.file_status_lbl.pack(fill=X)

        s2 = section_card(pg, "Column Mapping")
        Label(
            s2,
            text="Map your Excel columns to email fields:",
            font=FONT_SMALL,
            fg=MUTED,
            bg=CARD,
            anchor=W,
        ).pack(fill=X, pady=(0, 8))
        cm = Frame(s2, bg=CARD)
        cm.pack(fill=X)
        fields = [
            ("Email Column *", "Email"),
            ("Name Column", "Name"),
            ("Attachment Path Col", "Path"),
            ("CC Column", "CC_col"),
            ("BCC Column", "BCC_col"),
        ]
        self._col_vars = {}
        for i, (lbl_text, key) in enumerate(fields):
            r, c = i // 3, i % 3
            Label(cm, text=lbl_text, font=FONT_SMALL, fg=MUTED, bg=CARD).grid(
                row=r * 2, column=c, sticky=W, padx=(0, 14), pady=(4, 2)
            )
            v = StringVar(value=key)
            e = styled_entry(cm, width=16)
            e.insert(0, key)
            e.grid(row=r * 2 + 1, column=c, sticky=W, padx=(0, 14), pady=(0, 8))
            self._col_vars[key] = e

        s3 = section_card(pg, "Sheet Selection")
        sh_ctrl = Frame(s3, bg=CARD)
        sh_ctrl.pack(fill=X)
        lb_frame = Frame(sh_ctrl, bg=CARD)
        lb_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))
        self.sheets_lb = Listbox(
            lb_frame,
            selectmode=MULTIPLE,
            height=6,
            bg=ENTRY_BG,
            fg=TEXT,
            selectbackground=ACCENT,
            font=FONT_BODY,
            relief=FLAT,
            highlightthickness=1,
            highlightbackground=BORDER,
            activestyle="none",
        )
        sb2 = Scrollbar(lb_frame, orient=VERTICAL, command=self.sheets_lb.yview)
        self.sheets_lb.config(yscrollcommand=sb2.set)
        sb2.pack(side=RIGHT, fill=Y)
        self.sheets_lb.pack(fill=BOTH, expand=True)
        sf_btn = Frame(sh_ctrl, bg=CARD)
        sf_btn.pack(side=LEFT)
        styled_button(
            sf_btn, "Select All", command=self.select_all_sheets, kind="ghost", width=14
        ).pack(fill=X, pady=3)
        styled_button(
            sf_btn, "Clear All", command=self.clear_sheets, kind="ghost", width=14
        ).pack(fill=X, pady=3)
        styled_button(
            sf_btn, "Load & Preview", command=self.load_sheets, kind="primary", width=14
        ).pack(fill=X, pady=3)
        self.sheet_info_lbl = Label(
            s3,
            text="No sheets loaded",
            font=FONT_SMALL,
            fg=MUTED,
            bg=CARD,
            anchor=W,
            pady=4,
        )
        self.sheet_info_lbl.pack(fill=X)

        # Editable table
        s4 = section_card(pg, "Data Table (double-click to edit)")
        tool = Frame(s4, bg=CARD, pady=4)
        tool.pack(fill=X)
        styled_button(tool, "+ Add Row", command=self.add_row, kind="primary").pack(
            side=LEFT, padx=(0, 6)
        )
        styled_button(
            tool, "X Delete Selected", command=self.delete_row, kind="danger"
        ).pack(side=LEFT, padx=(0, 6))
        styled_button(
            tool, "Refresh from File", command=self.refresh_data, kind="ghost"
        ).pack(side=LEFT)
        tree_frame = Frame(s4, bg=CARD)
        tree_frame.pack(fill=BOTH, expand=True, pady=(4, 0))
        self.tree = ttk.Treeview(
            tree_frame, selectmode="browse", show="headings", height=12
        )
        vsb_tree = Scrollbar(tree_frame, orient=VERTICAL, command=self.tree.yview)
        hsb_tree = Scrollbar(tree_frame, orient=HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb_tree.set, xscrollcommand=hsb_tree.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb_tree.grid(row=0, column=1, sticky="ns")
        hsb_tree.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree_columns = []
        self.tree_editable = False
        self.row_count_lbl = Label(
            s4, text="0 rows", font=FONT_SMALL, fg=MUTED, bg=CARD, anchor=W
        )
        self.row_count_lbl.pack(fill=X, pady=(6, 0))

        bf = Frame(pg, bg=BG, pady=12, padx=18)
        bf.pack(fill=X)
        styled_button(
            bf, "← SMTP", command=lambda: self._show_tab("smtp"), kind="ghost"
        ).pack(side=LEFT)
        styled_button(
            bf,
            "Next →  Compose",
            command=lambda: self._show_tab("compose"),
            kind="primary",
        ).pack(side=RIGHT)
        return outer

    def _refresh_treeview(self):
        for col in self.tree_columns:
            self.tree.heading(col, text="")
            self.tree.column(col, width=0)
        self.tree.delete(*self.tree.get_children())
        self.tree_columns = []
        if self.data_df is None or self.data_df.empty:
            self.row_count_lbl.config(text="0 rows")
            self.tree_editable = False
            return
        cols = list(self.data_df.columns)
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="w")
        self.tree_columns = cols
        for idx, row in self.data_df.iterrows():
            values = [str(row[col]) if pd.notna(row[col]) else "" for col in cols]
            self.tree.insert("", "end", iid=str(idx), values=values)
        self.row_count_lbl.config(text=f"{len(self.data_df)} rows")
        self.tree_editable = True

    def on_tree_double_click(self, event):
        if not self.tree_editable:
            return
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not row_id or not column:
            return
        col_index = int(column[1:]) - 1
        col_name = self.tree_columns[col_index]
        current_val = self.tree.item(row_id)["values"][col_index]
        x, y, width, height = self.tree.bbox(row_id, column)
        entry = Entry(
            self.tree,
            font=FONT_SMALL,
            bg=ENTRY_BG,
            fg=TEXT,
            relief=SOLID,
            borderwidth=1,
        )
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_val)
        entry.focus_set()

        def save_edit(event=None):
            new_val = entry.get()
            entry.destroy()
            values = list(self.tree.item(row_id)["values"])
            values[col_index] = new_val
            self.tree.item(row_id, values=values)
            self.data_df.at[int(row_id), col_name] = new_val
            self.root.update_idletasks()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", lambda e: entry.destroy())

    def add_row(self):
        if self.data_df is None:
            return
        new_idx = self.data_df.index.max() + 1 if not self.data_df.empty else 0
        new_row = {col: "" for col in self.data_df.columns}
        self.data_df.loc[new_idx] = new_row
        self._refresh_treeview()
        # Refresh variable buttons in compose if needed
        if hasattr(self, "var_buttons_container"):
            self._refresh_var_buttons(self.data_df.columns.tolist())

    def delete_row(self):
        if self.data_df is None or self.tree_editable is False:
            return
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Info", "Select a row to delete.")
            return
        for iid in selected:
            idx = int(iid)
            self.data_df.drop(idx, inplace=True)
        self.data_df.reset_index(drop=True, inplace=True)
        self._refresh_treeview()
        if hasattr(self, "var_buttons_container"):
            self._refresh_var_buttons(self.data_df.columns.tolist())

    def refresh_data(self):
        """Reload data from original file, discarding edits."""
        fp = self.file_path_var.get()
        if not fp:
            messagebox.showwarning("Warning", "No file selected.")
            return
        self.update_selected_sheets()
        if not self.selected_sheets:
            messagebox.showwarning("Warning", "Select at least one sheet.")
            return
        try:
            combined = None
            for sh in self.selected_sheets:
                df = (
                    pd.read_excel(fp, sheet_name=sh)
                    if not fp.endswith(".csv")
                    else pd.read_csv(fp)
                )
                if combined is None:
                    combined = df
                else:
                    combined = pd.concat([combined, df], ignore_index=True)
            self.data_df = combined.reset_index(drop=True)
            self.total_emails = len(self.data_df)
            self.stat_total.config(text=str(self.total_emails))
            self.sheet_info_lbl.config(
                text=f"[OK] {len(self.selected_sheets)} sheet(s) — {self.total_emails} total recipients",
                fg=SUCCESS,
            )
            self._refresh_treeview()
            if hasattr(self, "var_buttons_container"):
                self._refresh_var_buttons(self.data_df.columns.tolist())
        except Exception as e:
            self.log(f"Error refreshing data: {e}", "error")

    # ─────────────────────── PAGE 3: COMPOSE (dynamic variables) ──────────────
    def _page_compose(self):
        outer, pg = self._scrollable_page()
        Label(
            pg,
            text="Compose Email",
            font=FONT_TITLE,
            fg=TEXT,
            bg=BG,
            pady=18,
            padx=18,
            anchor=W,
        ).pack(fill=X)

        s = section_card(pg, "Email Headers")
        hf = Frame(s, bg=CARD)
        hf.pack(fill=X)
        Label(hf, text="Subject *", font=FONT_SMALL, fg=MUTED, bg=CARD, anchor=W).grid(
            row=0, column=0, sticky=W, pady=(0, 2)
        )
        self.subject_entry = styled_entry(hf, width=80)
        self.subject_entry.insert(0, "Certificate of Appreciation")
        self.subject_entry.grid(row=1, column=0, columnspan=4, sticky=EW, pady=(0, 10))
        Label(
            hf,
            text="CC (comma separated)",
            font=FONT_SMALL,
            fg=MUTED,
            bg=CARD,
            anchor=W,
        ).grid(row=2, column=0, sticky=W, pady=(0, 2))
        self.cc_entry = styled_entry(hf, width=38)
        self.cc_entry.grid(
            row=3, column=0, columnspan=2, sticky=EW, padx=(0, 14), pady=(0, 10)
        )
        Label(
            hf,
            text="BCC (comma separated)",
            font=FONT_SMALL,
            fg=MUTED,
            bg=CARD,
            anchor=W,
        ).grid(row=2, column=2, sticky=W)
        self.bcc_entry = styled_entry(hf, width=38)
        self.bcc_entry.grid(row=3, column=2, columnspan=2, sticky=EW, pady=(0, 10))
        hf.columnconfigure(0, weight=1)
        hf.columnconfigure(2, weight=1)

        s2 = section_card(pg, "Email Body")
        ff = Frame(s2, bg=CARD)
        ff.pack(fill=X, pady=(0, 8))
        self.format_var = StringVar(value="plain")
        for val, lbl in [("plain", "Plain Text"), ("html", "HTML")]:
            Radiobutton(
                ff,
                text=lbl,
                variable=self.format_var,
                value=val,
                bg=CARD,
                fg=TEXT,
                selectcolor=ENTRY_BG,
                activebackground=CARD,
                font=FONT_SMALL,
            ).pack(side=LEFT, padx=(0, 12))

        # Dynamic variable chips
        self.var_chips_frame = Frame(s2, bg=CARD)
        self.var_chips_frame.pack(fill=X, pady=(0, 8))
        Label(
            self.var_chips_frame,
            text="Insert variable:",
            font=FONT_SMALL,
            fg=MUTED,
            bg=CARD,
        ).pack(side=LEFT, padx=(0, 6))
        self.var_buttons_container = Frame(self.var_chips_frame, bg=CARD)
        self.var_buttons_container.pack(side=LEFT)
        self._refresh_var_buttons([])  # start empty

        self.body_text = Text(
            s2,
            height=16,
            bg=ENTRY_BG,
            fg=TEXT,
            insertbackground=TEXT,
            font=FONT_MONO,
            relief=FLAT,
            padx=10,
            pady=10,
            wrap=WORD,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        self.body_text.pack(fill=X)
        self.body_text.insert(
            END,
            """Dear {Name},

Thank you for participating in "{Event}" organized by {Society} on {Date}.
Your role as {Role} was greatly appreciated.

Please find your certificate attached.

Best regards,
{Society}""",
        )

        bf = Frame(pg, bg=BG, pady=12, padx=18)
        bf.pack(fill=X)
        styled_button(
            bf,
            "← Recipients",
            command=lambda: self._show_tab("recipients"),
            kind="ghost",
        ).pack(side=LEFT)
        styled_button(
            bf,
            "Next →  Attachments",
            command=lambda: self._show_tab("attachments"),
            kind="primary",
        ).pack(side=RIGHT)
        return outer

    def _refresh_var_buttons(self, columns):
        """Rebuild insert-variable buttons based on actual DataFrame columns."""
        for w in self.var_buttons_container.winfo_children():
            w.destroy()
        if not columns:
            columns = ["Name", "Email", "Event", "Society", "Date", "Role", "Custom1"]
        for col in columns:
            b = Button(
                self.var_buttons_container,
                text=f"{{{col}}}",
                font=FONT_SMALL,
                bg=ENTRY_BG,
                fg=ACCENT,
                relief=FLAT,
                cursor="hand2",
                padx=6,
                pady=2,
                bd=0,
                command=lambda c=col: self._insert_var(f"{{{c}}}"),
            )
            b.pack(side=LEFT, padx=2)

    def _insert_var(self, var):
        try:
            self.body_text.insert(INSERT, var)
        except:
            pass

    # ─────────────────────── PAGE 4: ATTACHMENTS ──────────────────────────────
    def _page_attachments(self):
        outer, pg = self._scrollable_page()
        Label(
            pg,
            text="Attachments",
            font=FONT_TITLE,
            fg=TEXT,
            bg=BG,
            pady=18,
            padx=18,
            anchor=W,
        ).pack(fill=X)
        s = section_card(pg, "Attachment Mode")
        self.att_mode = IntVar(value=0)
        modes = [
            (0, "Per-recipient (path from Excel column 'Path')"),
            (1, "Same file(s) for all recipients"),
            (2, "No attachments"),
        ]
        for val, txt in modes:
            Radiobutton(
                s,
                text=txt,
                variable=self.att_mode,
                value=val,
                bg=CARD,
                fg=TEXT,
                selectcolor=ENTRY_BG,
                activebackground=CARD,
                font=FONT_BODY,
                command=self._toggle_att_mode,
            ).pack(anchor=W, pady=4)

        s2 = section_card(pg, "Select Files (for 'Same for all' mode)")
        self.att_list_frame = Frame(s2, bg=CARD)
        self.att_list_frame.pack(fill=X)
        self.att_listbox = Listbox(
            self.att_list_frame,
            bg=ENTRY_BG,
            fg=TEXT,
            selectbackground=ACCENT,
            font=FONT_SMALL,
            height=8,
            relief=FLAT,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        att_sb = Scrollbar(
            self.att_list_frame, orient=VERTICAL, command=self.att_listbox.yview
        )
        self.att_listbox.config(yscrollcommand=att_sb.set)
        att_sb.pack(side=RIGHT, fill=Y)
        self.att_listbox.pack(fill=X, expand=True)
        btn_row = Frame(s2, bg=CARD, pady=8)
        btn_row.pack(fill=X)
        styled_button(
            btn_row, "+ Add Files", command=self.add_attachments, kind="primary"
        ).pack(side=LEFT, padx=(0, 6))
        styled_button(
            btn_row, "X Remove Selected", command=self.remove_attachment, kind="danger"
        ).pack(side=LEFT, padx=(0, 6))
        styled_button(
            btn_row, "Clear All", command=self.clear_attachments, kind="ghost"
        ).pack(side=LEFT)
        self.att_count_lbl = Label(
            s2, text="0 file(s) selected", font=FONT_SMALL, fg=MUTED, bg=CARD, anchor=W
        )
        self.att_count_lbl.pack(fill=X, pady=(4, 0))
        bf = Frame(pg, bg=BG, pady=12, padx=18)
        bf.pack(fill=X)
        styled_button(
            bf, "← Compose", command=lambda: self._show_tab("compose"), kind="ghost"
        ).pack(side=LEFT)
        styled_button(
            bf, "Next →  Send", command=lambda: self._show_tab("send"), kind="primary"
        ).pack(side=RIGHT)
        return outer

    def _toggle_att_mode(self):
        pass

    # ─────────────────────── PAGE 5: SEND ─────────────────────────────────────
    def _page_send(self):
        outer, pg = self._scrollable_page()
        Label(
            pg,
            text="Send Emails",
            font=FONT_TITLE,
            fg=TEXT,
            bg=BG,
            pady=18,
            padx=18,
            anchor=W,
        ).pack(fill=X)
        s = section_card(pg, "Sending Options")
        of = Frame(s, bg=CARD)
        of.pack(fill=X)
        Label(
            of,
            text="Delay between emails (seconds):",
            font=FONT_SMALL,
            fg=MUTED,
            bg=CARD,
        ).grid(row=0, column=0, sticky=W, pady=4)
        self.delay_entry = styled_entry(of, width=6)
        self.delay_entry.insert(0, "3")
        self.delay_entry.grid(row=0, column=1, sticky=W, padx=8)
        self.draft_var = IntVar()
        Checkbutton(
            of,
            text="Simulate only (save .eml draft, don't send)",
            variable=self.draft_var,
            bg=CARD,
            fg=TEXT,
            selectcolor=ENTRY_BG,
            activebackground=CARD,
            font=FONT_SMALL,
        ).grid(row=1, column=0, columnspan=3, sticky=W, pady=4)
        self.retry_var = IntVar(value=1)
        Checkbutton(
            of,
            text="Retry failed emails once",
            variable=self.retry_var,
            bg=CARD,
            fg=TEXT,
            selectcolor=ENTRY_BG,
            activebackground=CARD,
            font=FONT_SMALL,
        ).grid(row=2, column=0, columnspan=3, sticky=W, pady=4)

        s2 = section_card(pg, "Pre-send Summary")
        self.summary_lbl = Label(
            s2,
            text="Click 'Refresh Summary' to see details",
            font=FONT_SMALL,
            fg=MUTED,
            bg=CARD,
            justify=LEFT,
            anchor=W,
        )
        self.summary_lbl.pack(fill=X)
        styled_button(
            s2, "Refresh Summary", command=self.refresh_summary, kind="ghost"
        ).pack(anchor=W, pady=(8, 0))

        s3 = section_card(pg, "Progress")
        self.prog_bar = ttk.Progressbar(
            s3,
            orient=HORIZONTAL,
            mode="determinate",
            style="custom.Horizontal.TProgressbar",
        )
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "custom.Horizontal.TProgressbar",
            troughcolor="#e2e8f0",
            background=ACCENT,
            darkcolor=ACCENT,
            lightcolor=ACCENT,
            thickness=16,
        )
        self.prog_bar.pack(fill=X, ipady=2)
        self.prog_lbl = Label(
            s3, text="Ready", font=FONT_SMALL, fg=MUTED, bg=CARD, anchor=W, pady=4
        )
        self.prog_lbl.pack(fill=X)

        bf = Frame(pg, bg=BG, pady=16, padx=18)
        bf.pack(fill=X)
        self.send_btn = styled_button(
            bf, "SEND NOW", command=self.start_sending, kind="success"
        )
        self.send_btn.config(font=("Segoe UI", 13, "bold"), pady=12, padx=28)
        self.send_btn.pack(side=LEFT, padx=(0, 10))
        self.pause_btn = styled_button(
            bf, "Pause", command=self.toggle_pause, kind="warning"
        )
        self.pause_btn.config(state=DISABLED)
        self.pause_btn.pack(side=LEFT, padx=(0, 6))
        self.stop_btn = styled_button(
            bf, "Stop", command=self.stop_sending, kind="danger"
        )
        self.stop_btn.config(state=DISABLED)
        self.stop_btn.pack(side=LEFT)
        styled_button(
            bf, "View Logs →", command=lambda: self._show_tab("logs"), kind="ghost"
        ).pack(side=RIGHT)
        return outer

    # ─────────────────────── PAGE 6: LOGS ─────────────────────────────────────
    def _page_logs(self):
        outer = Frame(self.content, bg=BG)
        Label(
            outer,
            text="Activity Logs",
            font=FONT_TITLE,
            fg=TEXT,
            bg=BG,
            pady=18,
            padx=18,
            anchor=W,
        ).pack(fill=X)
        toolbar = Frame(outer, bg=BG, padx=18, pady=0)
        toolbar.pack(fill=X)
        styled_button(
            toolbar, "Clear Logs", command=self.clear_logs, kind="ghost"
        ).pack(side=LEFT)
        styled_button(
            toolbar, "Export Logs", command=self.export_logs, kind="ghost"
        ).pack(side=LEFT, padx=6)
        log_frame = Frame(
            outer,
            bg=CARD,
            padx=18,
            pady=12,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        log_frame.pack(fill=BOTH, expand=True, padx=18, pady=12)
        self.log_text = Text(
            log_frame,
            bg=ENTRY_BG,
            fg=TEXT,
            font=FONT_MONO,
            relief=FLAT,
            wrap=WORD,
            state=DISABLED,
            highlightthickness=0,
            padx=8,
            pady=8,
        )
        log_sb = Scrollbar(log_frame, orient=VERTICAL, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=log_sb.set)
        log_sb.pack(side=RIGHT, fill=Y)
        self.log_text.pack(fill=BOTH, expand=True)
        self.log_text.tag_configure("error", foreground=DANGER)
        self.log_text.tag_configure("success", foreground=SUCCESS)
        self.log_text.tag_configure("warning", foreground=WARNING)
        self.log_text.tag_configure("info", foreground=ACCENT)
        return outer

    # ─────────────────────── ACTIONS / LOGGING ────────────────────────────────
    def log(self, msg, kind="normal"):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        tag = {
            "error": "error",
            "success": "success",
            "warning": "warning",
            "info": "info",
        }.get(kind, None)
        self.log_text.config(state=NORMAL)
        self.log_text.insert(END, line, tag)
        self.log_text.config(state=DISABLED)
        self.log_text.see(END)

    def clear_logs(self):
        self.log_text.config(state=NORMAL)
        self.log_text.delete(1.0, END)
        self.log_text.config(state=DISABLED)

    def export_logs(self):
        fp = filedialog.asksaveasfilename(
            defaultextension=".txt", filetypes=[("Text", "*.txt")]
        )
        if fp:
            with open(fp, "w") as f:
                f.write(self.log_text.get(1.0, END))
            self.log(f"Logs exported to {fp}", "success")

    def test_connection(self):
        def _test():
            srv = self.smtp_server.get()
            port_str = self.smtp_port.get()
            email = self.smtp_email.get()
            pwd = self.smtp_pass.get()
            enc = self.ssl_var.get()
            if not all([srv, port_str, email, pwd]):
                messagebox.showerror("Missing", "Fill all SMTP fields first")
                return
            try:
                port = int(port_str)
                ctx = ssl.create_default_context()
                self.log("Testing SMTP connection...", "info")
                if enc == "ssl":
                    with smtplib.SMTP_SSL(srv, port, context=ctx) as s:
                        s.login(email, pwd)
                else:
                    with smtplib.SMTP(srv, port) as s:
                        s.starttls(context=ctx)
                        s.login(email, pwd)
                self.log(f"[OK] Connection to {srv}:{port} successful!", "success")
                messagebox.showinfo("Success", "SMTP connection test passed!")
            except smtplib.SMTPAuthenticationError:
                self.log(
                    "[FAIL] Auth failed — check email/password or use App Password",
                    "error",
                )
                messagebox.showerror(
                    "Auth Failed",
                    "Authentication failed.\n\nFor Gmail:\n1. Enable 2-Step Verification\n"
                    "2. Create App Password at myaccount.google.com\n3. Use that as password here.",
                )
            except Exception as e:
                self.log(f"[FAIL] Connection failed: {e}", "error")
                messagebox.showerror("Failed", str(e))

        Thread(target=_test, daemon=True).start()

    def browse_file(self):
        fp = filedialog.askopenfilename(
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")]
        )
        if not fp:
            return
        self.file_path_var.set(fp)
        self.file_status_lbl.config(text=f"File: {os.path.basename(fp)}", fg=SUCCESS)
        self.sheets_lb.delete(0, END)
        self.available_sheets = []
        try:
            if fp.endswith(".csv"):
                self.sheets_lb.insert(END, "Sheet1")
                self.available_sheets = ["Sheet1"]
            else:
                xl = pd.ExcelFile(fp)
                self.available_sheets = xl.sheet_names
                for sh in self.available_sheets:
                    self.sheets_lb.insert(END, sh)
            if self.available_sheets:
                self.sheets_lb.selection_set(0)
                self.update_selected_sheets()
            self.log(
                f"Loaded file: {os.path.basename(fp)} — {len(self.available_sheets)} sheet(s)",
                "info",
            )
        except Exception as e:
            self.log(f"Error reading file: {e}", "error")

    def update_selected_sheets(self):
        self.selected_sheets = [
            self.sheets_lb.get(i) for i in self.sheets_lb.curselection()
        ]

    def select_all_sheets(self):
        self.sheets_lb.selection_set(0, END)
        self.update_selected_sheets()

    def clear_sheets(self):
        self.sheets_lb.selection_clear(0, END)
        self.selected_sheets = []

    def load_sheets(self):
        self.update_selected_sheets()
        fp = self.file_path_var.get()
        if not fp or not self.selected_sheets:
            messagebox.showwarning("Warning", "Select a file and at least one sheet")
            return
        try:
            combined = None
            for sh in self.selected_sheets:
                df = (
                    pd.read_excel(fp, sheet_name=sh)
                    if not fp.endswith(".csv")
                    else pd.read_csv(fp)
                )
                if combined is None:
                    combined = df
                else:
                    combined = pd.concat([combined, df], ignore_index=True)
            self.data_df = combined.reset_index(drop=True)
            self.total_emails = len(self.data_df)
            self.stat_total.config(text=str(self.total_emails))
            self.sheet_info_lbl.config(
                text=f"[OK] {len(self.selected_sheets)} sheet(s) — {self.total_emails} total recipients",
                fg=SUCCESS,
            )
            self._refresh_treeview()
            if hasattr(self, "var_buttons_container"):
                self._refresh_var_buttons(self.data_df.columns.tolist())
            self.log(f"Loaded {self.total_emails} recipients.", "info")
        except Exception as e:
            self.log(f"Error loading sheets: {e}", "error")

    def add_attachments(self):
        files = filedialog.askopenfilenames(title="Select attachments")
        if files:
            for f in files:
                if f not in self.attachment_files:
                    self.attachment_files.append(f)
                    self.att_listbox.insert(END, os.path.basename(f))
            self.att_count_lbl.config(
                text=f"{len(self.attachment_files)} file(s) selected", fg=SUCCESS
            )
            self.log(f"Added {len(files)} attachment(s)", "info")

    def remove_attachment(self):
        sel = self.att_listbox.curselection()
        for i in reversed(sel):
            self.att_listbox.delete(i)
            self.attachment_files.pop(i)
        self.att_count_lbl.config(text=f"{len(self.attachment_files)} file(s) selected")

    def clear_attachments(self):
        self.att_listbox.delete(0, END)
        self.attachment_files = []
        self.att_count_lbl.config(text="0 file(s) selected")

    def refresh_summary(self):
        lines = [
            f"SMTP Server : {self.smtp_server.get() or '—'}:{self.smtp_port.get() or '—'} ({self.ssl_var.get().upper()})",
            f"From        : {self.smtp_email.get() or '—'}",
            f"Subject     : {self.subject_entry.get() or '—'}",
            f"Recipients  : {self.total_emails}  (across {len(self.selected_sheets)} sheet(s))",
            f"Attachments : {len(self.attachment_files)} file(s) — mode: {['Per-recipient','Same for all','None'][self.att_mode.get()]}",
            f"Delay       : {self.delay_entry.get()}s between emails",
            f"Draft mode  : {'YES (not sending)' if self.draft_var.get() else 'NO (will send)'}",
        ]
        self.summary_lbl.config(text="\n".join(lines), fg=TEXT)

    def start_sending(self):
        if self.is_sending:
            return
        if not all(
            [
                self.smtp_server.get(),
                self.smtp_port.get(),
                self.smtp_email.get(),
                self.smtp_pass.get(),
            ]
        ):
            messagebox.showerror("Error", "Configure SMTP settings first (SMTP tab)")
            return
        if self.data_df is None or self.data_df.empty:
            messagebox.showerror(
                "Error", "No recipient data loaded. Load a file in Recipients tab."
            )
            return
        self.stop_flag = False
        self.pause_flag = False
        self.sent_count = 0
        self.failed_count = 0
        self.total_emails = len(self.data_df)
        self.prog_bar["maximum"] = self.total_emails
        self.prog_bar["value"] = 0
        self.is_sending = True
        self.send_btn.config(state=DISABLED)
        self.pause_btn.config(state=NORMAL)
        self.stop_btn.config(state=NORMAL)
        Thread(target=self._send_thread, daemon=True).start()

    def _send_thread(self):
        srv = self.smtp_server.get()
        port = int(self.smtp_port.get())
        email_addr = self.smtp_email.get()
        pwd = self.smtp_pass.get()
        enc = self.ssl_var.get()
        subj = self.subject_entry.get()
        cc = self.cc_entry.get()
        bcc = self.bcc_entry.get()
        template = self.body_text.get(1.0, END)
        fmt = self.format_var.get()
        delay = int(self.delay_entry.get() or "0")
        draft = bool(self.draft_var.get())
        att_mode = self.att_mode.get()
        sender_name = self.sender_name.get()
        reply_to = self.reply_to.get()
        email_col_name = (
            self._col_vars["Email"].get() if "Email" in self._col_vars else "Email"
        )
        path_col_name = (
            self._col_vars["Path"].get() if "Path" in self._col_vars else "Path"
        )

        df = self.data_df.copy()  # work with a snapshot
        self.total_emails = len(df)
        self.prog_bar["maximum"] = self.total_emails

        try:
            ctx = ssl.create_default_context()
            if enc == "ssl":
                smtp = smtplib.SMTP_SSL(srv, port, context=ctx)
            else:
                smtp = smtplib.SMTP(srv, port)
                smtp.starttls(context=ctx)
            smtp.login(email_addr, pwd)
            self.log("[OK] SMTP logged in", "success")

            for idx, row in df.iterrows():
                if self.stop_flag:
                    break
                while self.pause_flag and not self.stop_flag:
                    time.sleep(0.3)

                recip = str(row.get(email_col_name, "")).strip()
                if not recip or "@" not in recip:
                    self.log(
                        f"Row {idx+1}: skipped (invalid email '{recip}')", "warning"
                    )
                    self.failed_count += 1
                    continue

                try:
                    row_dict = row.to_dict()
                    try:
                        body = template.format_map(
                            {k: str(v) for k, v in row_dict.items()}
                        )
                    except Exception:
                        body = template

                    msg = MIMEMultipart()
                    from_str = (
                        f"{sender_name} <{email_addr}>" if sender_name else email_addr
                    )
                    msg["From"] = from_str
                    msg["To"] = recip
                    msg["Subject"] = subj
                    if cc:
                        msg["Cc"] = cc
                    if bcc:
                        msg["Bcc"] = bcc
                    if reply_to:
                        msg["Reply-To"] = reply_to
                    msg.attach(MIMEText(body, fmt))

                    files_to_attach = []
                    if att_mode == 0:
                        row_path = str(row.get(path_col_name, "")).strip()
                        if row_path and os.path.exists(row_path):
                            files_to_attach = [row_path]
                        elif row_path:
                            self.log(f"  Attachment not found: {row_path}", "warning")
                    elif att_mode == 1:
                        files_to_attach = self.attachment_files

                    for af in files_to_attach:
                        with open(af, "rb") as fh:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(fh.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            "Content-Disposition",
                            f'attachment; filename="{os.path.basename(af)}"',
                        )
                        msg.attach(part)

                    if draft:
                        dd = os.path.join(
                            os.path.dirname(self.file_path_var.get()), "drafts"
                        )
                        os.makedirs(dd, exist_ok=True)
                        dpath = os.path.join(dd, f"draft_{idx}.eml")
                        with open(dpath, "w") as f:
                            f.write(msg.as_string())
                        self.log(f"  Draft saved → {os.path.basename(dpath)}", "info")
                    else:
                        all_recip = [recip]
                        if cc:
                            all_recip += [x.strip() for x in cc.split(",") if x.strip()]
                        if bcc:
                            all_recip += [
                                x.strip() for x in bcc.split(",") if x.strip()
                            ]
                        smtp.sendmail(email_addr, all_recip, msg.as_string())

                    self.sent_count += 1
                    self.log(f"  [SENT] → {recip}", "success")
                except Exception as e:
                    self.failed_count += 1
                    self.log(f"  [FAIL] → {recip}: {e}", "error")

                self.prog_bar["value"] = idx + 1
                self.prog_lbl.config(
                    text=f"Processed {idx+1}/{self.total_emails} — Sent: {self.sent_count} | Failed: {self.failed_count}"
                )
                self.stat_sent.config(text=str(self.sent_count))
                self.stat_fail.config(text=str(self.failed_count))
                self.root.update_idletasks()
                if delay and idx < len(df) - 1 and not self.stop_flag:
                    for t in range(delay, 0, -1):
                        if self.stop_flag:
                            break
                        self.prog_lbl.config(text=f"Waiting {t}s…")
                        time.sleep(1)

            smtp.quit()
            status = "stopped" if self.stop_flag else "complete"
            self.log(
                f"\nDone ({status}) — Sent: {self.sent_count} | Failed: {self.failed_count}",
                "success",
            )
            if not self.stop_flag:
                messagebox.showinfo(
                    "Done",
                    f"Sending complete!\n\nSent: {self.sent_count}\nFailed: {self.failed_count}",
                )
        except Exception as e:
            self.log(f"Fatal error: {e}", "error")
            messagebox.showerror("Error", str(e))
        finally:
            self.is_sending = False
            self.send_btn.config(state=NORMAL)
            self.pause_btn.config(state=DISABLED)
            self.stop_btn.config(state=DISABLED)
            self.prog_lbl.config(text="Ready")

    def toggle_pause(self):
        self.pause_flag = not self.pause_flag
        self.pause_btn.config(text="Resume" if self.pause_flag else "Pause")
        self.log("Paused" if self.pause_flag else "Resumed", "warning")

    def stop_sending(self):
        self.stop_flag = True
        self.pause_flag = False
        self.log("Stop requested…", "warning")


# ─── Entry Point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    root = Tk()
    app = MassEmailSender(root)
    root.mainloop()
