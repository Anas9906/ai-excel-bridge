"""
gui_app.py — Tkinter Batch & Single Certificate GUI for GIECO Insurance Sync V2

Features:
    - Multi-Image Batch Processing (Select multiple files or load whole folders)
    - Asynchronous background extraction queue with progress bar
    - Interactive Batch Queue table with status badges (Pending, Extracting, Matched, Saved)
    - Side-by-side zoomable image viewer (mouse wheel + buttons)
    - Editable fields with instant live diff preview
    - Single & Batch Excel commit with automatic soft-audit highlighting and backups
    - Dual OCR Engine: High-precision Gemini Vision AI + Windows WinRT OCR fallback

Usage:
    python gui_app.py
"""

import os
os.environ["OPENBLAS_NUM_THREADS"] = "1"
os.environ["MKL_NUM_THREADS"] = "1"
os.environ["OMP_NUM_THREADS"] = "1"
os.environ["NUMEXPR_NUM_THREADS"] = "1"

import sys
import threading
import queue


if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import tkinter as tk

from tkinter import ttk, filedialog, messagebox, scrolledtext
import re
from pathlib import Path
from datetime import date, datetime, timedelta
from typing import Optional, List, Dict, Any

from PIL import Image, ImageTk

from stamp_remover import remove_stamp
from ocr_engine import extract_certificates_batch, extract_certificate_data
from ocr_provider import DailyQuotaExhaustedError, GEMINI_BATCH_SIZE, GEMINI_RPM
from excel_handler import ExcelHandler


class GIECOInsuranceSyncApp:
    """Main application window for GIECO Insurance Sync V2 with Batch Support."""

    # Default Excel path
    DEFAULT_EXCEL = str(
        Path(__file__).parent / "test_data" / "كشف 4 سيارات 2026 - رامي.xlsx"
    )

    # ── Single Unified Palette & Material Tokens ──
    # 1. Base Layer (Dark Navy)
    BASE_BG = "#0f1420"

    # 2. Surface / Panel Layer (1 step lighter)
    PANEL_BG = "#161c2c"

    # 3. Elevated Surface (Buttons, inputs, cards, table elements)
    ELEVATED_BG = "#1e2536"
    ELEVATED_ACTIVE = "#262f44"
    BORDER_COLOR = "#2a3449"
    CANVAS_BG = "#0b0f19"

    # 4. Primary Brand Accent (Single unified accent: Electric Sky-Blue)
    ACCENT = "#38bdf8"
    ACCENT_HOVER = "#7dd3fc"
    ACCENT_DARK = "#0f1420"

    # 5. Status Indicators (Strictly 3, used only for status signals)
    STATUS_SUCCESS = "#22c55e"  # Green: Matched / Connected
    STATUS_WARNING = "#f59e0b"  # Amber: Review Required / Partial Match
    STATUS_ERROR = "#ef4444"    # Red: Error / Not Found / Disconnected

    # 6. Typography & Text Hierarchy
    TEXT_PRIMARY = "#f8fafc"    # Bright near-white for headers, values, active text
    TEXT_SECONDARY = "#94a3b8"  # Slate grey for labels, column headers, status text
    TEXT_MUTED = "#64748b"      # Muted grey for placeholders, disabled copy

    # Backward-compatibility aliases
    BG_COLOR = BASE_BG
    TEXT_COLOR = TEXT_PRIMARY
    GREEN = STATUS_SUCCESS
    YELLOW = STATUS_WARNING
    RED = STATUS_ERROR
    BTN_BG = ELEVATED_BG
    BTN_ACTIVE = ELEVATED_ACTIVE
    ENTRY_BG = ELEVATED_BG
    TABLE_BG = PANEL_BG
    TABLE_FG = TEXT_PRIMARY

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("🏗️ GIECO Insurance Sync — Batch & OCR Studio V2")
        self.root.geometry("1350x850")
        self.root.configure(bg=self.BASE_BG)
        self.root.minsize(1000, 650)

        # UI State Guard
        self._is_updating_ui = False

        # Batch Queue State
        self.batch_items: List[Dict[str, Any]] = []
        self.active_index: Optional[int] = None
        self.is_processing_batch = False
        self.work_queue = queue.Queue()

        # Excel State
        self.handler: Optional[ExcelHandler] = None
        self.excel_path = self.DEFAULT_EXCEL

        # Active Image Zoom State
        self.original_image = None
        self.photo_image = None
        self.zoom_level = 1.0
        self.base_scale = 1.0

        # Editable Form Fields (StringVar) - Formatted in Day-Month-Year (DD-MM-YYYY) for UI
        self.insurance_no_var = tk.StringVar()
        self.date_from_var = tk.StringVar()
        self.date_to_var = tk.StringVar()
        self.print_date_var = tk.StringVar()
        self.receipt_date_var = tk.StringVar(value=self._to_ui_date(date.today()))

        # Bind variable changes to auto-update active batch item
        for var in (self.insurance_no_var, self.date_from_var, self.date_to_var, self.print_date_var, self.receipt_date_var):
            var.trace_add("write", self._on_field_edited)

        # Initialize
        self._setup_styles()
        self._load_excel()
        self._build_ui()

        # Start background polling for queue results
        self.root.after(100, self._check_work_queue)

    @staticmethod
    def _to_ui_date(d_val: Optional[Any]) -> str:
        """Convert ISO date string (YYYY-MM-DD), datetime/date object to UI display format DD-MM-YYYY."""
        if not d_val:
            return ""
        if isinstance(d_val, (date, datetime)):
            return d_val.strftime("%d-%m-%Y")
        s = str(d_val).strip()
        if not s:
            return ""
        # Check if already DD-MM-YYYY or DD/MM/YYYY
        m_dd = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$", s)
        if m_dd:
            d, m, y = m_dd.groups()
            return f"{int(d):02d}-{int(m):02d}-{y}"
        # Check if YYYY-MM-DD or YYYY/MM/DD
        m_iso = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
        if m_iso:
            y, m, d = m_iso.groups()
            return f"{int(d):02d}-{int(m):02d}-{y}"
        return s

    @staticmethod
    def _to_iso_date(d_val: Optional[Any]) -> Optional[str]:
        """Convert UI date string (DD-MM-YYYY or DD/MM/YYYY or YYYY-MM-DD) to ISO format YYYY-MM-DD for backend/Excel."""
        if not d_val:
            return None
        if isinstance(d_val, (date, datetime)):
            return d_val.strftime("%Y-%m-%d")
        s = str(d_val).strip()
        if not s:
            return None
        # Check DD-MM-YYYY or DD/MM/YYYY
        m_dd = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$", s)
        if m_dd:
            d, m, y = m_dd.groups()
            return f"{y}-{int(m):02d}-{int(d):02d}"
        # Check YYYY-MM-DD or YYYY/MM/DD
        m_iso = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", s)
        if m_iso:
            y, m, d = m_iso.groups()
            return f"{y}-{int(m):02d}-{int(d):02d}"
        return s

    def _setup_styles(self):
        """Configure ttk styles for a cohesive dark palette."""
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Treeview (Batch Queue)
        self.style.configure(
            "Batch.Treeview",
            background=self.PANEL_BG,
            foreground=self.TEXT_PRIMARY,
            fieldbackground=self.PANEL_BG,
            font=("Helvetica", 10),
            rowheight=28,
            borderwidth=0,
            relief=tk.FLAT,
        )
        self.style.configure(
            "Batch.Treeview.Heading",
            background=self.ELEVATED_BG,
            foreground=self.TEXT_SECONDARY,
            font=("Helvetica", 9, "bold"),
            relief=tk.FLAT,
            borderwidth=0,
            padding=(6, 4),
        )
        self.style.map(
            "Batch.Treeview",
            background=[("selected", self.ELEVATED_BG)],
            foreground=[("selected", self.ACCENT)],
        )
        self.style.map(
            "Batch.Treeview.Heading",
            background=[("active", self.ELEVATED_ACTIVE)],
            foreground=[("active", self.TEXT_PRIMARY)],
        )

        # Progressbar
        self.style.configure(
            "Batch.Horizontal.TProgressbar",
            troughcolor=self.ELEVATED_BG,
            background=self.ACCENT,
            bordercolor=self.PANEL_BG,
            lightcolor=self.ACCENT,
            darkcolor=self.ACCENT,
        )

    def _load_excel(self):
        """Initialize the Excel handler."""
        if os.path.exists(self.excel_path):
            try:
                self.handler = ExcelHandler(self.excel_path)
            except Exception as e:
                print(f"[WARN] Cannot load Excel: {e}")
                self.handler = None
        else:
            self.handler = None

    def _build_ui(self):
        """Build the main layout with 3 core panels: Batch Queue, Image Viewer, Fields & Diffs."""
        # ── 1. Top Title & Global Controls ──
        top_frame = tk.Frame(self.root, bg=self.BASE_BG, pady=8, padx=14)
        top_frame.pack(fill=tk.X)

        title_label = tk.Label(
            top_frame,
            text="🏗️ GIECO Insurance Sync — Batch & OCR Studio",
            font=("Helvetica", 15, "bold"),
            fg=self.ACCENT,
            bg=self.BASE_BG,
        )
        title_label.pack(side=tk.LEFT)

        self.excel_status_lbl = tk.Label(
            top_frame,
            text=f"Excel: {'✅ ' + os.path.basename(self.excel_path) if self.handler else '❌ Not Found'}",
            font=("Helvetica", 10, "bold"),
            fg=self.STATUS_SUCCESS if self.handler else self.STATUS_ERROR,
            bg=self.BASE_BG,
        )
        self.excel_status_lbl.pack(side=tk.RIGHT, padx=10)

        change_excel_btn = tk.Button(
            top_frame,
            text="📊 Change Excel",
            font=("Helvetica", 9, "bold"),
            fg=self.TEXT_PRIMARY,
            bg=self.ELEVATED_BG,
            activebackground=self.ELEVATED_ACTIVE,
            activeforeground=self.ACCENT,
            highlightbackground=self.BASE_BG,
            highlightthickness=0,
            bd=0,
            relief=tk.FLAT,
            padx=10,
            pady=4,
            command=self._change_excel,
        )
        change_excel_btn.pack(side=tk.RIGHT, padx=5)

        # ── 2. Main Content Split Panels ──
        main_paned = tk.PanedWindow(
            self.root, orient=tk.HORIZONTAL, bg=self.BASE_BG, bd=0, sashwidth=6, sashrelief=tk.FLAT
        )
        main_paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        # Left: Batch Queue Panel
        left_batch_frame = self._build_batch_panel(main_paned)
        main_paned.add(left_batch_frame, minsize=320, width=360)

        # Middle: Image Canvas Panel
        middle_img_frame = self._build_image_panel(main_paned)
        main_paned.add(middle_img_frame, minsize=380, width=480)

        # Right: Fields & Diff Controls
        right_fields_frame = self._build_fields_panel(main_paned)
        main_paned.add(right_fields_frame, minsize=420, width=540)

        # ── 3. Bottom Status Bar & Global Progress ──
        bottom_bar = tk.Frame(self.root, bg=self.PANEL_BG, padx=12, pady=4, highlightthickness=1, highlightbackground=self.BORDER_COLOR)
        bottom_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_var = tk.StringVar(value="Ready. Load images to begin.")
        status_lbl = tk.Label(
            bottom_bar,
            textvariable=self.status_var,
            font=("Helvetica", 9),
            fg=self.TEXT_SECONDARY,
            bg=self.PANEL_BG,
            anchor=tk.W,
        )
        status_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.batch_count_var = tk.StringVar(value="Queue: 0 items")
        count_lbl = tk.Label(
            bottom_bar,
            textvariable=self.batch_count_var,
            font=("Helvetica", 9, "bold"),
            fg=self.ACCENT,
            bg=self.PANEL_BG,
        )
        count_lbl.pack(side=tk.RIGHT, padx=10)

    # ─────────────────────────────────────────────────────────────
    # Panel 1: Batch Queue Panel
    # ─────────────────────────────────────────────────────────────
    def _build_batch_panel(self, parent):
        frame = tk.Frame(parent, bg=self.PANEL_BG, padx=8, pady=8)

        # Header & Actions
        hdr_frame = tk.Frame(frame, bg=self.PANEL_BG)
        hdr_frame.pack(fill=tk.X, pady=(0, 6))

        tk.Label(
            hdr_frame,
            text="📑 Batch Queue",
            font=("Helvetica", 12, "bold"),
            fg=self.TEXT_PRIMARY,
            bg=self.PANEL_BG,
        ).pack(side=tk.LEFT)

        # Load buttons row
        btn_row = tk.Frame(frame, bg=self.PANEL_BG)
        btn_row.pack(fill=tk.X, pady=(0, 6))

        tk.Button(
            btn_row,
            text="📁 Load Images",
            font=("Helvetica", 10, "bold"),
            fg=self.TEXT_PRIMARY,
            bg=self.ELEVATED_BG,
            activebackground=self.ELEVATED_ACTIVE,
            activeforeground=self.ACCENT,
            highlightbackground=self.PANEL_BG,
            highlightthickness=0,
            bd=0,
            relief=tk.FLAT,
            padx=8,
            pady=5,
            command=self._load_multiple_images,
        ).pack(side=tk.LEFT, padx=(0, 4), expand=True, fill=tk.X)

        tk.Button(
            btn_row,
            text="📂 Load Folder",
            font=("Helvetica", 10, "bold"),
            fg=self.TEXT_PRIMARY,
            bg=self.ELEVATED_BG,
            activebackground=self.ELEVATED_ACTIVE,
            activeforeground=self.ACCENT,
            highlightbackground=self.PANEL_BG,
            highlightthickness=0,
            bd=0,
            relief=tk.FLAT,
            padx=8,
            pady=5,
            command=self._load_folder,
        ).pack(side=tk.LEFT, expand=True, fill=tk.X)

        # Batch Run Button Row
        run_row = tk.Frame(frame, bg=self.PANEL_BG)
        run_row.pack(fill=tk.X, pady=(0, 6))

        self.process_batch_btn = tk.Button(
            run_row,
            text="⚡ Extract All Batch",
            font=("Helvetica", 11, "bold"),
            fg=self.ACCENT_DARK,
            bg=self.ACCENT,
            activebackground=self.ACCENT_HOVER,
            activeforeground=self.ACCENT_DARK,
            highlightbackground=self.PANEL_BG,
            highlightthickness=0,
            bd=0,
            relief=tk.FLAT,
            padx=10,
            pady=6,
            command=self._start_batch_processing,
        )
        self.process_batch_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 4))

        tk.Button(
            run_row,
            text="🗑️ Clear",
            font=("Helvetica", 10, "bold"),
            fg=self.STATUS_ERROR,
            bg=self.ELEVATED_BG,
            activebackground=self.ELEVATED_ACTIVE,
            activeforeground=self.STATUS_ERROR,
            highlightbackground=self.PANEL_BG,
            highlightthickness=0,
            bd=0,
            relief=tk.FLAT,
            padx=8,
            pady=6,
            command=self._clear_queue,
        ).pack(side=tk.LEFT)

        # Progress bar
        self.progress_bar = ttk.Progressbar(
            frame, orient=tk.HORIZONTAL, mode="determinate", style="Batch.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 6))

        # Treeview Batch List
        tree_frame = tk.Frame(frame, bg=self.PANEL_BG, highlightthickness=1, highlightbackground=self.BORDER_COLOR)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("id", "filename", "status", "match")
        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            style="Batch.Treeview",
            selectmode="browse",
        )
        self.tree.heading("id", text="#")
        self.tree.heading("filename", text="File Name")
        self.tree.heading("status", text="Status")
        self.tree.heading("match", text="Excel Row")

        self.tree.column("id", width=30, anchor=tk.CENTER)
        self.tree.column("filename", width=140, anchor=tk.W)
        self.tree.column("status", width=90, anchor=tk.CENTER)
        self.tree.column("match", width=80, anchor=tk.CENTER)

        v_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scroll.set)

        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_item_selected)

        return frame

    # ─────────────────────────────────────────────────────────────
    # Panel 2: Image Canvas & Zoom Viewer
    # ─────────────────────────────────────────────────────────────
    def _build_image_panel(self, parent):
        frame = tk.Frame(parent, bg=self.PANEL_BG, padx=8, pady=8)

        # Header & Zoom Controls
        hdr_frame = tk.Frame(frame, bg=self.PANEL_BG)
        hdr_frame.pack(fill=tk.X, pady=(0, 4))

        self.image_title_lbl = tk.Label(
            hdr_frame,
            text="📄 Certificate View",
            font=("Helvetica", 12, "bold"),
            fg=self.TEXT_PRIMARY,
            bg=self.PANEL_BG,
        )
        self.image_title_lbl.pack(side=tk.LEFT)

        # Zoom Controls
        zoom_frame = tk.Frame(hdr_frame, bg=self.PANEL_BG)
        zoom_frame.pack(side=tk.RIGHT)

        for text, cmd, pad, font_w in [("🔍 +", self._zoom_in, 8, "bold"), ("🔍 -", self._zoom_out, 8, "bold"), ("Reset", self._zoom_reset, 6, "normal")]:
            tk.Button(
                zoom_frame,
                text=text,
                font=("Helvetica", 9, font_w),
                fg=self.TEXT_PRIMARY,
                bg=self.ELEVATED_BG,
                activebackground=self.ELEVATED_ACTIVE,
                activeforeground=self.ACCENT,
                highlightbackground=self.PANEL_BG,
                highlightthickness=0,
                bd=0,
                relief=tk.FLAT,
                padx=pad,
                pady=2,
                command=cmd,
            ).pack(side=tk.LEFT, padx=2)

        # Canvas with dual scrollbars
        canvas_frame = tk.Frame(frame, bg=self.PANEL_BG)
        canvas_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        self.canvas = tk.Canvas(
            canvas_frame,
            bg=self.CANVAS_BG,
            highlightthickness=1,
            highlightbackground=self.BORDER_COLOR,
        )
        v_scroll = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        h_scroll = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas.bind("<MouseWheel>", self._on_mouse_wheel)

        # Placeholder text
        self._show_canvas_placeholder()

        return frame

    def _show_canvas_placeholder(self):
        """Draw empty state placeholder on canvas."""
        self.canvas.delete("all")
        self.canvas.create_text(
            250, 250,
            text="📥\n\nLoad images or a folder to begin ↑",
            fill=self.TEXT_MUTED,
            font=("Helvetica", 12),
            justify=tk.CENTER,
            tags="placeholder"
        )

    # ─────────────────────────────────────────────────────────────
    # Panel 3: Fields Inspector, Diffs, and Single/Batch Commit
    # ─────────────────────────────────────────────────────────────
    def _build_fields_panel(self, parent):
        frame = tk.Frame(parent, bg=self.PANEL_BG, padx=8, pady=8)

        # ── Bottom Docked Action Buttons ──
        bottom_actions = tk.Frame(frame, bg=self.PANEL_BG)
        bottom_actions.pack(side=tk.BOTTOM, fill=tk.X)

        # Batch Write All Button
        self.batch_write_btn = tk.Button(
            bottom_actions,
            text="🚀 WRITE 0 VALIDATED ROW(S) TO EXCEL",
            font=("Helvetica", 12, "bold"),
            fg=self.TEXT_MUTED,
            bg=self.ELEVATED_BG,
            activebackground=self.ACCENT_HOVER,
            activeforeground=self.ACCENT_DARK,
            highlightbackground=self.PANEL_BG,
            highlightthickness=0,
            bd=0,
            relief=tk.FLAT,
            padx=15,
            pady=10,
            state=tk.DISABLED,
            command=self._write_all_batch_to_excel,
        )
        self.batch_write_btn.pack(fill=tk.X, pady=(4, 2))

        # Single Write & Refresh Row
        single_row = tk.Frame(bottom_actions, bg=self.PANEL_BG)
        single_row.pack(fill=tk.X, pady=(2, 4))

        self.single_write_btn = tk.Button(
            single_row,
            text="💾 Save This Item",
            font=("Helvetica", 10, "bold"),
            fg=self.TEXT_PRIMARY,
            bg=self.ELEVATED_BG,
            activebackground=self.ELEVATED_ACTIVE,
            activeforeground=self.ACCENT,
            highlightbackground=self.PANEL_BG,
            highlightthickness=0,
            bd=0,
            relief=tk.FLAT,
            padx=10,
            pady=6,
            state=tk.DISABLED,
            command=self._write_single_to_excel,
        )
        self.single_write_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 3))

        self.refresh_btn = tk.Button(
            single_row,
            text="🔄 Refresh Diff",
            font=("Helvetica", 10),
            fg=self.TEXT_PRIMARY,
            bg=self.ELEVATED_BG,
            activebackground=self.ELEVATED_ACTIVE,
            activeforeground=self.ACCENT,
            highlightbackground=self.PANEL_BG,
            highlightthickness=0,
            bd=0,
            relief=tk.FLAT,
            padx=10,
            pady=6,
            state=tk.DISABLED,
            command=self._refresh_diff,
        )
        self.refresh_btn.pack(side=tk.LEFT, padx=(3, 0))

        # ── Editable Date Fields Box (Elevated Surface Card) ──
        date_box = tk.Frame(
            bottom_actions,
            bg=self.ELEVATED_BG,
            highlightthickness=1,
            highlightbackground=self.BORDER_COLOR,
            bd=0,
            relief=tk.FLAT,
        )
        date_box.pack(fill=tk.X, pady=(4, 4))

        tk.Label(
            date_box,
            text="📅 Editable Dates & Insurance No (Live Auto-Sync)",
            font=("Helvetica", 10, "bold"),
            fg=self.ACCENT,
            bg=self.ELEVATED_BG,
            anchor=tk.W,
        ).pack(fill=tk.X, padx=8, pady=(6, 2))

        # Divider
        tk.Frame(date_box, bg=self.BORDER_COLOR, height=1).pack(fill=tk.X, padx=8, pady=2)

        for label, var in [
            ("Insurance No:", self.insurance_no_var),
            ("Date From (من):", self.date_from_var),
            ("Date To (الى):", self.date_to_var),
            ("Print Date:", self.print_date_var),
            ("Receipt Date:", self.receipt_date_var),
        ]:
            row = tk.Frame(date_box, bg=self.ELEVATED_BG)
            row.pack(fill=tk.X, padx=8, pady=2)

            tk.Label(
                row,
                text=label,
                font=("Helvetica", 9),
                fg=self.TEXT_SECONDARY,
                bg=self.ELEVATED_BG,
                width=17,
                anchor=tk.W,
            ).pack(side=tk.LEFT)

            entry = tk.Entry(
                row,
                textvariable=var,
                font=("Courier", 11, "bold"),
                bg=self.PANEL_BG,
                fg=self.TEXT_PRIMARY,
                insertbackground=self.ACCENT,
                highlightbackground=self.BORDER_COLOR,
                highlightcolor=self.ACCENT,
                highlightthickness=1,
                relief=tk.FLAT,
                bd=0,
            )
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 2))
            
            # Use ghost text placeholder when empty
            def _on_focus_out(e, v=var, l=label):
                if not v.get():
                    e.widget.configure(fg=self.TEXT_MUTED)
            def _on_focus_in(e):
                e.widget.configure(fg=self.TEXT_PRIMARY)
            
            entry.bind("<FocusIn>", _on_focus_in)
            entry.bind("<FocusOut>", _on_focus_out)
            if not var.get():
                entry.configure(fg=self.TEXT_MUTED)

        # ── Scrollable Inspector Text (Console Card) ──
        self.fields_text = scrolledtext.ScrolledText(
            frame,
            wrap=tk.WORD,
            bg=self.PANEL_BG,
            fg=self.TEXT_PRIMARY,
            font=("Courier", 11),
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=self.BORDER_COLOR,
            insertbackground=self.ACCENT,
            state=tk.DISABLED,
        )
        self.fields_text.pack(fill=tk.BOTH, expand=True, pady=(0, 4))

        # Configure text tags
        self.fields_text.tag_configure("header", foreground=self.ACCENT, font=("Courier", 11, "bold"))
        self.fields_text.tag_configure("success", foreground=self.STATUS_SUCCESS)
        self.fields_text.tag_configure("warning", foreground=self.STATUS_WARNING)
        self.fields_text.tag_configure("error", foreground=self.STATUS_ERROR)
        self.fields_text.tag_configure("label", foreground=self.TEXT_SECONDARY)
        self.fields_text.tag_configure("value", foreground=self.TEXT_PRIMARY, font=("Courier", 11, "bold"))
        self.fields_text.tag_configure("diff_old", foreground=self.STATUS_ERROR)
        self.fields_text.tag_configure("diff_new", foreground=self.STATUS_SUCCESS)

        self._display_text("BATCH & OCR INSPECTOR\n", "header")
        self._display_text("─" * 40 + "\n", "label")
        self._display_text("Load images or a folder to start batch processing.\n", "label")

        return frame

    # ─────────────────────────────────────────────────────────────
    # Batch Loading & File Handling
    # ─────────────────────────────────────────────────────────────
    def _load_multiple_images(self):
        """Open file dialog allowing multiple certificate image selection."""
        try:
            paths = filedialog.askopenfilenames(
                title="Select Certificate Images",
                filetypes=[
                    ("Image files", "*.jpeg;*.jpg;*.png;*.bmp;*.tiff"),
                    ("All files", "*.*"),
                ],
            )
            if not paths:
                return

            if isinstance(paths, str):
                paths = self.root.tk.splitlist(paths)

            self._add_paths_to_queue(list(paths))
        except Exception as e:
            messagebox.showerror("Error Loading Images", f"Failed to open files: {e}")

    def _load_folder(self):
        """Load all images inside a selected directory."""
        folder = filedialog.askdirectory(title="Select Folder with Certificate Images")
        if not folder:
            return

        valid_exts = {".jpeg", ".jpg", ".png", ".bmp", ".tiff"}
        paths = [
            str(p) for p in Path(folder).glob("*") if p.suffix.lower() in valid_exts
        ]

        if not paths:
            messagebox.showinfo("No Images", "No image files found in selected directory.")
            return

        self._add_paths_to_queue(paths)

    def _add_paths_to_queue(self, paths: List[str]):
        """Add image paths to the batch queue."""
        existing_paths = {item["path"] for item in self.batch_items}
        added_count = 0

        for p in paths:
            if p in existing_paths:
                continue

            item_id = len(self.batch_items) + 1
            filename = os.path.basename(p)

            item = {
                "id": item_id,
                "path": p,
                "filename": filename,
                "status": "⏳ Pending",
                "ocr_data": None,
                "matches": None,
                "diffs": None,
                "date_from": "",
                "date_to": "",
                "print_date": "",
                "receipt_date": date.today().isoformat(),
                "insurance_no": "",
                "cert_code": "",
                "error": None,
            }
            self.batch_items.append(item)
            self.tree.insert(
                "",
                tk.END,
                iid=str(item_id - 1),
                values=(item_id, filename, "⏳ Pending", "-"),
            )
            added_count += 1

        self.batch_count_var.set(f"Queue: {len(self.batch_items)} items")
        self.status_var.set(f"Added {added_count} image(s) to batch queue.")

        if self.active_index is None and self.batch_items:
            self._select_item(0)

        self._update_action_buttons()

    def _clear_queue(self):
        """Clear the entire batch queue."""
        if self.is_processing_batch:
            messagebox.showwarning("Busy", "Cannot clear while batch processing is running.")
            return

        # Clear data
        self.batch_items.clear()
        
        # Clear Treeview visually
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Reset labels & progress
        self.batch_count_var.set("Queue: 0 items")
        self.status_var.set("Queue cleared.")
        self.progress_bar["value"] = 0
        self.active_index = None
        self.image_title_lbl.configure(text="📄 Certificate View")
        
        # Reset panels
        self._clear_fields()
        self._show_canvas_placeholder()
        self._update_action_buttons()

    # ─────────────────────────────────────────────────────────────
    # Asynchronous Batch Processing Worker
    # ─────────────────────────────────────────────────────────────
    def _start_batch_processing(self):
        """Start background thread to extract OCR and match Excel for all pending items."""
        if not self.batch_items:
            messagebox.showinfo("Queue Empty", "Please load images first.")
            return

        if self.is_processing_batch:
            return

        if not self.handler:
            messagebox.showerror("Error", "Excel file is not loaded. Please change/load Excel.")
            return

        self.is_processing_batch = True
        self.process_batch_btn.configure(state=tk.DISABLED, text="🔄 Processing...")
        self.progress_bar["maximum"] = len(self.batch_items)
        self.progress_bar["value"] = 0

        # Start worker thread
        threading.Thread(target=self._batch_worker, daemon=True).start()

    def _batch_worker(self):
        """Rate-limit-aware batch worker. Uses extract_certificates_batch for 3-6x API efficiency."""
        import json
        import time

        # Collect pending indices
        pending_indices = [
            idx for idx, item in enumerate(self.batch_items)
            if not (item["status"] in ("✅ Matched", "💾 Saved") and item["ocr_data"])
        ]

        if not pending_indices:
            self.work_queue.put(("batch_complete", None, None))
            return

        # Compute inter-batch delay to stay within GEMINI_RPM
        inter_batch_delay = 60.0 / max(GEMINI_RPM, 1)
        pending_paths = [self.batch_items[i]["path"] for i in pending_indices]

        # Signal all pending items as extracting
        for idx in pending_indices:
            self.work_queue.put(("status_update", idx, "🔄 Extracting..."))

        # Phase 1: Batch OCR Extraction with rate-limit awareness
        ocr_results: dict = {}  # idx -> dict or Exception
        chunk_size = GEMINI_BATCH_SIZE
        total_chunks = (len(pending_paths) + chunk_size - 1) // chunk_size

        for chunk_num in range(total_chunks):
            chunk_start = chunk_num * chunk_size
            chunk_end = min(chunk_start + chunk_size, len(pending_paths))
            chunk_paths = pending_paths[chunk_start:chunk_end]
            chunk_indices = pending_indices[chunk_start:chunk_end]

            self.work_queue.put((
                "status_bar",
                None,
                f"🧠 Extracting batch {chunk_num + 1}/{total_chunks} ({len(chunk_paths)} images)...",
            ))

            try:
                results = extract_certificates_batch(chunk_paths)
                for local_i, global_idx in enumerate(chunk_indices):
                    ocr_results[global_idx] = results[local_i]

            except DailyQuotaExhaustedError as e:
                # Save uncompleted items to resumable state file
                remaining_paths = pending_paths[chunk_start:]
                remaining_indices = pending_indices[chunk_start:]
                state = {
                    "remaining_paths": remaining_paths,
                    "remaining_indices": remaining_indices,
                }
                state_path = ".ocr_state.json"
                try:
                    with open(state_path, "w", encoding="utf-8") as f:
                        json.dump(state, f, ensure_ascii=False, indent=2)
                except Exception as write_err:
                    print(f"[WARN] Could not save state file: {write_err}")

                done_count = chunk_start
                remaining_count = len(pending_paths) - chunk_start
                self.work_queue.put((
                    "quota_exhausted",
                    None,
                    f"Daily Gemini quota exhausted after {done_count} images. "
                    f"{remaining_count} images saved to {state_path} for tomorrow.",
                ))
                # Mark remaining items as paused
                for idx in remaining_indices:
                    self.work_queue.put(("status_update", idx, "⏸ Paused (quota)"))
                # Process only what we have so far and stop
                break

            except Exception as e:
                for global_idx in chunk_indices:
                    ocr_results[global_idx] = e
                    self.work_queue.put(("item_error", global_idx, str(e)))

            # RPM pacing: sleep between chunks (not after the last one)
            if chunk_num < total_chunks - 1:
                time.sleep(inter_batch_delay)

        # Phase 2: Sequential Excel matching (safe for COM objects, not thread-safe)
        for idx, item in enumerate(self.batch_items):
            if idx not in ocr_results:
                self.work_queue.put(("progress", idx + 1, None))
                continue

            res = ocr_results[idx]
            if isinstance(res, Exception):
                self.work_queue.put(("progress", idx + 1, None))
                continue

            ocr_data = res

            try:
                chassis = ocr_data.get("chassis_no")
                plate = ocr_data.get("plate_digits")

                matches = []
                if self.handler:
                    if chassis:
                        matches = self.handler.find_vehicle(chassis)
                    if not matches and plate:
                        matches = self.handler.find_vehicle(plate)

                # Fallback date calculation if OCR missed dates
                date_from = ocr_data.get("date_from") or ""
                date_to = ocr_data.get("date_to") or ""

                if (not date_from or not date_to) and matches and matches[0].current_date_to:
                    prev_to = matches[0].current_date_to
                    calc_from = prev_to + timedelta(days=1)
                    try:
                        calc_to = calc_from.replace(year=calc_from.year + 1) - timedelta(days=1)
                    except ValueError:
                        calc_to = calc_from + (date(calc_from.year + 1, 1, 1) - date(calc_from.year, 1, 1)) - timedelta(days=1)

                    if not date_from:
                        date_from = calc_from.strftime("%Y-%m-%d")
                    if not date_to:
                        date_to = calc_to.strftime("%Y-%m-%d")

                # Flag partial chassis matches for supervisor review
                has_partial = any(
                    "partial" in (m.match_type or "") for m in matches
                ) if matches else False

                status = (
                    "⚠️ Review Required" if has_partial
                    else ("✅ Matched" if matches else "❌ Not Found")
                )

                result_payload = {
                    "ocr_data": ocr_data,
                    "matches": matches,
                    "status": status,
                    "date_from": date_from,
                    "date_to": date_to,
                    "print_date": ocr_data.get("print_date") or "",
                    "insurance_no": ocr_data.get("insurance_no") or "",
                    "cert_code": str(matches[0].current_cert_code) if matches else "",
                    "error": None,
                }
                self.work_queue.put(("item_done", idx, result_payload))

            except Exception as e:
                self.work_queue.put(("item_error", idx, str(e)))

            self.work_queue.put(("progress", idx + 1, None))

        self.work_queue.put(("batch_complete", None, None))

    def _check_work_queue(self):
        """Poll the work queue for updates from the background thread."""
        try:
            while True:
                msg_type, idx, payload = self.work_queue.get_nowait()

                if msg_type == "status_update":
                    if idx < len(self.batch_items):
                        self.batch_items[idx]["status"] = payload
                        self.tree.item(str(idx), values=(idx + 1, self.batch_items[idx]["filename"], payload, "-"))

                elif msg_type == "item_done":
                    if idx < len(self.batch_items):
                        item = self.batch_items[idx]
                        item.update(payload)
                        if item.get("matches"):
                            m0 = item["matches"][0]
                            code_tag = f" (كود: {m0.current_cert_code})" if m0.current_cert_code is not None else ""
                            match_str = f"Row {m0.row_number}{code_tag}"
                        else:
                            match_str = "❌ Not Found"


                        self.tree.item(
                            str(idx),
                            values=(idx + 1, item["filename"], item["status"], match_str),
                        )


                        # If active item finished, refresh display
                        if self.active_index == idx:
                            self._load_item_into_view(idx)

                elif msg_type == "status_bar":
                    if payload:
                        self.status_var.set(payload)

                elif msg_type == "quota_exhausted":
                    self.is_processing_batch = False
                    self.process_batch_btn.configure(state=tk.NORMAL, text="\u26a1 Extract All Batch")
                    self.status_var.set(f"\u26a0\ufe0f Quota exhausted — partial progress saved.")
                    messagebox.showwarning(
                        "Daily Quota Exhausted",
                        payload or "Gemini daily quota reached. Remaining items saved to .ocr_state.json."
                    )
                    self._update_action_buttons()

                elif msg_type == "item_error":
                    if idx < len(self.batch_items):
                        item = self.batch_items[idx]
                        item["status"] = "❌ Error"
                        item["error"] = payload
                        self.tree.item(str(idx), values=(idx + 1, item["filename"], "❌ Error", "Error"))

                elif msg_type == "progress":
                    self.progress_bar["value"] = idx

                elif msg_type == "batch_complete":
                    self.is_processing_batch = False
                    self.process_batch_btn.configure(state=tk.NORMAL, text="⚡ Extract All Batch")
                    self.status_var.set("✅ Batch processing completed!")
                    self._update_action_buttons()

        except queue.Empty:
            pass

        self.root.after(100, self._check_work_queue)

    # ─────────────────────────────────────────────────────────────
    # Treeview Navigation & Item Selection
    # ─────────────────────────────────────────────────────────────
    def _on_tree_item_selected(self, event):
        """Handle user clicking a certificate in the batch table."""
        if self._is_updating_ui:
            return
        selection = self.tree.selection()
        if not selection:
            return
        try:
            idx = int(selection[0])
            if idx == self.active_index:
                return
            self._select_item(idx)
        except Exception:
            pass

    def _select_item(self, idx: int):
        """Select and display an item from the batch items list."""
        if idx < 0 or idx >= len(self.batch_items):
            return

        self._is_updating_ui = True
        try:
            self.active_index = idx
            current_sel = self.tree.selection()
            if not current_sel or current_sel[0] != str(idx):
                self.tree.selection_set(str(idx))
                self.tree.see(str(idx))

            self._load_item_into_view(idx)
        finally:
            self._is_updating_ui = False

    def _load_item_into_view(self, idx: int):
        """Populate canvas, fields, and diffs for the specified item."""
        if idx < 0 or idx >= len(self.batch_items):
            return
        item = self.batch_items[idx]

        # 1. Load image on canvas
        self._show_image(item["path"])
        self.image_title_lbl.configure(text=f"📄 {item['filename']} (#{item['id']})")

        # 2. Populate editable date fields (formatted as DD-MM-YYYY for UI display)
        self._is_updating_ui = True
        try:
            self.insurance_no_var.set(item.get("insurance_no") or "")
            self.date_from_var.set(self._to_ui_date(item.get("date_from")))
            self.date_to_var.set(self._to_ui_date(item.get("date_to")))
            self.print_date_var.set(self._to_ui_date(item.get("print_date")))
            self.receipt_date_var.set(self._to_ui_date(item.get("receipt_date")) or self._to_ui_date(date.today()))
        finally:
            self._is_updating_ui = False

        # 3. Render Fields Inspector
        self._render_inspector_text(item)

        # 4. Update single write button state
        self._update_action_buttons()

    def _render_inspector_text(self, item: Dict[str, Any]):
        """Render details of the active item in the fields inspector panel."""
        self._clear_fields()

        self._display_text(f"CERTIFICATE: {item['filename']}\n", "header")
        self._display_text("─" * 40 + "\n", "label")

        ocr_data = item.get("ocr_data")
        if not ocr_data:
            self._display_text(f"  Status: {item['status']}\n", "warning")
            self._display_text("  Click 'Extract All Batch' to process.\n", "label")
            return

        # Show engine
        engine = ocr_data.get("engine_used", "WinRT OCR")
        if "Gemini" in engine:
            self._display_text("  ✨ AI VISION: 100% Extracted with Gemini\n\n", "success")
        else:
            self._display_text("  ⚡ ENGINE: Windows WinRT OCR\n\n", "label")

        fields = [
            ("Service No", "service_no"),
            ("Office Code", "office_code"),
            ("Plate Digits", "plate_digits"),
            ("Insurance No", "insurance_no"),
            ("Chassis No", "chassis_no"),
        ]

        for label, key in fields:
            val = ocr_data.get(key)
            conf = ocr_data.get("confidence", {}).get(key, 0)
            if val:
                icon = "✅" if conf >= 0.7 else "⚠️"
                tag = "success" if conf >= 0.7 else "warning"
            else:
                icon = "❌"
                tag = "error"
                val = "NOT FOUND"
            self._display_text(f"  {icon} {label:15s}: ", "label")
            self._display_text(f"{val}\n", tag)

        # Dates (Formatted as Day-Month-Year for UI clarity)
        self._display_text("\n  📅 Certificate Dates (DD-MM-YYYY):\n", "header")
        for label, key in [("Date From", "date_from"), ("Date To", "date_to"), ("Print Date", "print_date")]:
            val = item.get(key)
            if val:
                ui_val = self._to_ui_date(val)
                self._display_text(f"  ✅ {label:15s}: ", "label")
                self._display_text(f"{ui_val}\n", "success")
            else:
                self._display_text(f"  ⚠️  {label:15s}: ", "label")
                self._display_text("Empty (Enter below)\n", "warning")

        # Excel Match info
        self._display_text("\nEXCEL MATCH\n", "header")
        self._display_text("─" * 40 + "\n", "label")

        matches = item.get("matches")
        if not matches:
            self._display_text("  ❌ Vehicle NOT FOUND in master sheet.\n", "error")
            return

        if item.get("needs_review"):
            self._display_text(
                "  ⚠️  PARTIAL CHASSIS MATCH — Manual review required before writing!\n", "warning"
            )

        for m in matches:
            icon = "⚠️" if "partial" in (m.match_type or "") else "✅"
            self._display_text(f"  {icon} Row {m.row_number} ({m.match_type})\n", "warning" if "partial" in (m.match_type or "") else "success")
            self._display_text(f"     Driver:    ", "label")
            self._display_text(f"{m.driver_name}\n", "value")
            self._display_text(f"     Office:    ", "label")
            self._display_text(f"{m.office}\n", "value")
            self._display_text(f"     Plate:     ", "label")
            self._display_text(f"{m.plate}\n", "value")
            self._display_text(f"     Type:      ", "label")
            self._display_text(f"{m.vehicle_type}\n", "value")
            
            code_display = f"{m.current_cert_code}" if m.current_cert_code is not None else "None"
            self._display_text(f"\n  🏷️  كود الشهادة (عمود W): ", "label")
            self._display_text(f"[{code_display}]", "header")
            self._display_text(f"  ✍️ (اكتب هذا الرقم على أصل الشهادة الورقية)\n\n", "success")


        # Show Diff preview
        self._show_active_diff(item)

    def _show_active_diff(self, item: Dict[str, Any]):
        """Calculate and display diff for the active item."""
        if not self.handler or not item.get("matches"):
            return

        self._display_text("\nCHANGES PREVIEW\n", "header")
        self._display_text("─" * 40 + "\n", "label")

        ocr_with_dates = dict(item.get("ocr_data", {}))
        ocr_with_dates["date_from"] = item.get("date_from")
        ocr_with_dates["date_to"] = item.get("date_to")
        ocr_with_dates["print_date"] = item.get("print_date")
        ocr_with_dates["receipt_date"] = item.get("receipt_date")
        ocr_with_dates["insurance_no"] = item.get("insurance_no")

        diffs = self.handler.prepare_update(item["matches"], ocr_with_dates)
        item["diffs"] = diffs

        for diff in diffs:
            self._display_text(f"  Row {diff.row_number}:\n", "value")
            for col_name, change in diff.changes.items():
                old = change["old"]
                new = change["new"]
                col = change["col"]

                self._display_text(f"    Col{col:2d} {col_name:20s}\n", "label")
                if new is None:
                    self._display_text(f"      ⚠️  EMPTY — enter date\n", "warning")
                elif str(old) == str(new):
                    self._display_text(f"      = {self._to_ui_date(new)}\n", "label")
                else:
                    old_ui = self._to_ui_date(old)
                    new_ui = self._to_ui_date(new)
                    self._display_text(f"      Old: {old_ui}\n", "diff_old")
                    self._display_text(f"      New: {new_ui}\n", "diff_new")

    def _on_field_edited(self, *args):
        """Triggered whenever user edits an active date/insurance field."""
        if self._is_updating_ui or self.active_index is None or self.active_index >= len(self.batch_items):
            return

        item = self.batch_items[self.active_index]
        item["insurance_no"] = self.insurance_no_var.get().strip() or None
        item["date_from"] = self._to_iso_date(self.date_from_var.get())
        item["date_to"] = self._to_iso_date(self.date_to_var.get())
        item["print_date"] = self._to_iso_date(self.print_date_var.get())
        
        # Global sync for Receipt Date across the entire batch
        new_receipt = self._to_iso_date(self.receipt_date_var.get())
        for b_item in self.batch_items:
            b_item["receipt_date"] = new_receipt

    def _refresh_diff(self):
        """Manually refresh diff preview for active item."""
        if self.active_index is not None:
            self._load_item_into_view(self.active_index)
            self.status_var.set("✅ Diff refreshed.")

    # ─────────────────────────────────────────────────────────────
    # Excel Writing & Highlighting (Single & Batch)
    # ─────────────────────────────────────────────────────────────
    def _update_action_buttons(self):
        """Enable/disable action buttons based on batch state."""
        valid_items = [
            i for i in self.batch_items if i.get("matches") and i["status"] != "💾 Saved"
        ]
        has_active = (
            self.active_index is not None
            and self.active_index < len(self.batch_items)
            and self.batch_items[self.active_index].get("matches")
            and self.batch_items[self.active_index]["status"] != "💾 Saved"
        )

        self.batch_write_btn.configure(
            state=tk.NORMAL if valid_items else tk.DISABLED,
            text=f"🚀 WRITE {len(valid_items)} VALIDATED ROW(S) TO EXCEL",
            bg=self.ACCENT if valid_items else self.ELEVATED_BG,
            fg=self.ACCENT_DARK if valid_items else self.TEXT_MUTED,
            activebackground=self.ACCENT_HOVER if valid_items else self.ELEVATED_BG,
            activeforeground=self.ACCENT_DARK if valid_items else self.TEXT_MUTED,
        )
        self.single_write_btn.configure(
            state=tk.NORMAL if has_active else tk.DISABLED,
            fg=self.TEXT_PRIMARY if has_active else self.TEXT_MUTED,
        )
        self.refresh_btn.configure(
            state=tk.NORMAL if self.active_index is not None else tk.DISABLED,
            fg=self.TEXT_PRIMARY if self.active_index is not None else self.TEXT_MUTED,
        )

    def _write_single_to_excel(self):
        """Write the active item to Excel."""
        if self.active_index is None:
            return

        item = self.batch_items[self.active_index]
        if not item.get("matches") or not self.handler:
            messagebox.showwarning("Cannot Write", "Vehicle is not matched in Excel.")
            return

        # Prepare diff
        ocr_with_dates = dict(item.get("ocr_data", {}))
        ocr_with_dates.update({
            "date_from": item.get("date_from"),
            "date_to": item.get("date_to"),
            "print_date": item.get("print_date"),
            "receipt_date": item.get("receipt_date"),
            "insurance_no": item.get("insurance_no"),
        })
        diffs = self.handler.prepare_update(item["matches"], ocr_with_dates)

        confirm = messagebox.askyesno(
            "Confirm Single Write",
            f"Write update for Row {item['matches'][0].row_number} ({item['filename']}) to Excel?\n"
            f"Cell and row audit highlighting will be applied.",
        )
        if not confirm:
            return

        try:
            output = self.handler.apply_update(diffs, dry_run=False)
            item["status"] = "💾 Saved"
            self.tree.item(
                str(self.active_index),
                values=(self.active_index + 1, item["filename"], "💾 Saved", f"Row {item['matches'][0].row_number}"),
            )
            self._update_action_buttons()
            messagebox.showinfo("Success", f"Row saved and highlighted successfully!\n\nWorkbook: {output}")
        except Exception as e:
            messagebox.showerror("Error", f"Write failed: {e}")

    def _write_all_batch_to_excel(self):
        """Batch write all validated and matched items into Excel in a single transaction."""
        valid_items = [
            i for i in self.batch_items if i.get("matches") and i["status"] != "💾 Saved"
        ]
        if not valid_items or not self.handler:
            messagebox.showwarning("No Items", "No matched items available to write.")
            return

        # Collect all diffs
        all_diffs = []
        rows_to_update = []
        for item in valid_items:
            ocr_with_dates = dict(item.get("ocr_data", {}))
            ocr_with_dates.update({
                "date_from": item.get("date_from"),
                "date_to": item.get("date_to"),
                "print_date": item.get("print_date"),
                "receipt_date": item.get("receipt_date"),
                "insurance_no": item.get("insurance_no"),
            })
            item_diffs = self.handler.prepare_update(item["matches"], ocr_with_dates)
            all_diffs.extend(item_diffs)
            rows_to_update.append(f"Row {item['matches'][0].row_number} ({item['filename']})")

        confirm = messagebox.askyesno(
            "Confirm Batch Write",
            f"Commit {len(valid_items)} vehicle records to Excel in a single transaction?\n\n"
            f"Target Rows:\n  • " + "\n  • ".join(rows_to_update[:8]) +
            (f"\n  ... and {len(rows_to_update) - 8} more" if len(rows_to_update) > 8 else "") +
            f"\n\nAll updated rows will receive soft audit green highlighting.",
        )
        if not confirm:
            return

        try:
            output = self.handler.apply_update(all_diffs, dry_run=False)

            # Mark all as saved
            for item in valid_items:
                item["status"] = "💾 Saved"
                idx = item["id"] - 1
                self.tree.item(
                    str(idx),
                    values=(idx + 1, item["filename"], "💾 Saved", f"Row {item['matches'][0].row_number}"),
                )

            self._update_action_buttons()
            messagebox.showinfo(
                "Batch Write Complete",
                f"✅ Successfully wrote and highlighted {len(valid_items)} vehicle records in Excel!\n\n"
                f"Saved to:\n{output}",
            )
            self.status_var.set(f"✅ Batch committed: {len(valid_items)} rows highlighted and saved.")
        except Exception as e:
            messagebox.showerror("Batch Write Error", f"Failed to commit batch: {e}")

    # ─────────────────────────────────────────────────────────────
    # Canvas Image Rendering & Zoom
    # ─────────────────────────────────────────────────────────────
    def _show_image(self, path: str):
        """Load and display image on canvas."""
        try:
            if not os.path.exists(path):
                return
            self.original_image = Image.open(path)
            canvas_w = max(400, self.canvas.winfo_width())
            canvas_h = max(500, self.canvas.winfo_height())
            self.base_scale = min(
                canvas_w / max(1, self.original_image.width),
                canvas_h / max(1, self.original_image.height),
                1.0,
            )
            self.zoom_level = 1.0
            self._render_image()
        except Exception as e:
            print(f"[WARN] Failed to show image: {e}")

    def _render_image(self):
        """Render PIL image to Tkinter PhotoImage on canvas."""
        if not self.original_image:
            return

        try:
            scale = self.base_scale * self.zoom_level
            new_w = max(20, int(self.original_image.width * scale))
            new_h = max(20, int(self.original_image.height * scale))

            if new_w > 12000 or new_h > 12000:
                return

            resized = self.original_image.resize((new_w, new_h), Image.Resampling.LANCZOS)
            self.photo_image = ImageTk.PhotoImage(resized)

            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo_image)
            self.canvas.configure(scrollregion=(0, 0, new_w, new_h))
        except Exception as e:
            print(f"[WARN] Render image failed: {e}")

    def _zoom_in(self):
        if self.original_image:
            self.zoom_level *= 1.2
            self._render_image()

    def _zoom_out(self):
        if self.original_image:
            self.zoom_level /= 1.2
            self._render_image()

    def _zoom_reset(self):
        if self.original_image:
            self.zoom_level = 1.0
            self._render_image()

    def _on_mouse_wheel(self, event):
        if not self.original_image:
            return
        if event.delta > 0:
            self._zoom_in()
        elif event.delta < 0:
            self._zoom_out()

    # ─────────────────────────────────────────────────────────────
    # Utility Helpers
    # ─────────────────────────────────────────────────────────────
    def _display_text(self, text: str, tag: str = None):
        """Append formatted text to inspector pane."""
        self.fields_text.configure(state=tk.NORMAL)
        if tag:
            self.fields_text.insert(tk.END, text, tag)
        else:
            self.fields_text.insert(tk.END, text)
        self.fields_text.configure(state=tk.DISABLED)
        self.fields_text.see(tk.END)

    def _clear_fields(self):
        """Clear fields text area."""
        self.fields_text.configure(state=tk.NORMAL)
        self.fields_text.delete("1.0", tk.END)
        self.fields_text.configure(state=tk.DISABLED)

    def _change_excel(self):
        """Switch target Excel workbook."""
        path = filedialog.askopenfilename(
            title="Select Excel Workbook",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return

        try:
            self.handler = ExcelHandler(path)
            self.excel_path = path
            self.excel_status_lbl.configure(
                text=f"Excel: ✅ {os.path.basename(path)}", fg=self.STATUS_SUCCESS
            )
            self.status_var.set(f"Excel loaded: {os.path.basename(path)}")
            if self.active_index is not None:
                self._load_item_into_view(self.active_index)
            self._update_action_buttons()
        except Exception as e:
            self.handler = None
            self.excel_status_lbl.configure(text="Excel: ❌ Error", fg=self.STATUS_ERROR)
            messagebox.showerror("Error", f"Failed to load Excel:\n{e}")


def main():
    root = tk.Tk()
    app = GIECOInsuranceSyncApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
