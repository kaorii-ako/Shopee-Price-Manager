#!/usr/bin/env python3
"""
Shopee Mass Price Editor — Desktop GUI
Full Tkinter interface for bulk-adjusting prices in Shopee Thailand Excel files.
"""

import sys
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from datetime import datetime

# ── Core library ──
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from price_editor_core import (
    ProductRow, load_file, save_file, save_summary_report, make_output_path,
    calc_price_pct, calc_price_flat, calc_price_vat, calc_price_markup_vat, apply_adj, build_changes, merge_changes,
    parse_variation, unique_products, price_range, get_fill, make_fill,
    CAPPED_FILL, NO_FILL, SHOPEE_MIN, SHOPEE_MAX, HEADER_ROWS,
    THAI_COLORS, BUILD_LABEL,
    make_size_filter, make_color_filter, make_sku_filter, make_tag,
    make_search_price_filter,
)

# ── Color definitions ──
COLORS_GREEN_BAND = "#E8F5E9"
COLORS_RED_BAND   = "#FFEBEE"
COLORS_GREEN_DARK = "#2E7D32"
COLORS_RED_DARK   = "#C62828"
COLORS_GREEN_MOD  = "#66BB6A"
COLORS_RED_MOD    = "#EF5350"
COLORS_GREEN_LITE = "#A5D6A7"
COLORS_RED_LITE   = "#EF9A9A"
COLORS_GREEN_PALE = "#E8F5E9"
COLORS_RED_PALE   = "#FFEBEE"
COLORS_ORANGE     = "#FFE0B2"
COLOR_HEADER_BG   = "#4472C4"
COLOR_HEADER_FG   = "#FFFFFF"
COLOR_ROW_ALT     = "#F5F5F5"
COLOR_ROW_NORMAL  = "#FFFFFF"
COLOR_STATUS_BAR  = "#E0E0E0"
COLOR_ACCENT      = "#1976D2"

# ── Window title ──
APP_TITLE = f"Shopee Price Editor — {BUILD_LABEL}"

# ═══════════════════════════════════════════════════════════
#  PREVIEW DIALOG
# ═══════════════════════════════════════════════════════════
class PreviewDialog(tk.Toplevel):
    """Modal dialog showing before/after price changes."""

    def __init__(self, parent, rows, changes, original_prices, current_prices):
        super().__init__(parent)
        self.transient(parent)
        self.title("Preview Changes")
        self.geometry("700x500")
        self.resizable(True, True)
        self.grab_set()

        # Center
        x = parent.winfo_x() + (parent.winfo_width() - 700) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 500) // 2
        self.geometry(f"+{x}+{y}")

        # ── Summary frame ──
        sum_frame = ttk.Frame(self)
        sum_frame.pack(fill="x", padx=10, pady=(10, 0))

        changed_count = sum(1 for r in rows if r.row_idx in changes and changes[r.row_idx][0] != changes[r.row_idx][1])
        capped_count  = sum(1 for idx, (o, n, c) in changes.items() if o != n and c)

        # Summary labels
        sum_text = f"SKUs affected: {changed_count}  |  Capped: {capped_count}  |  "
        if changed_count > 0:
            diffs = [n - o for (o, n, _) in changes.values() if o != n]
            avg_pct = sum((n - o) / o * 100 for (o, n, _) in changes.values() if o != n and o) / changed_count
            sign = "+" if avg_pct >= 0 else ""
            sum_text += f"Avg change: {sign}{avg_pct:.1f}%"

        ttk.Label(sum_frame, text=sum_text, font=("Arial", 9)).pack(anchor="w")

        # ── Warning frame ──
        product_prices = {}
        for r in rows:
            p = changes.get(r.row_idx, (r.price, r.price, False))[1] if r.row_idx in changes else r.price
            if p is None or r.row_idx not in original_prices:
                continue
            pid = r.product_id
            if pid not in product_prices:
                product_prices[pid] = {"name": r.product_name, "prices": []}
            product_prices[pid]["prices"].append(p)

        warnings = []
        for pid, info in product_prices.items():
            ps = info["prices"]
            if len(ps) < 2:
                continue
            ratio = max(ps) / min(ps) if min(ps) > 0 else 0
            if ratio > 5:
                warnings.append(f"Product \"{info['name'][:40]}\": max ฿{max(ps):.0f} / min ฿{min(ps):.0f} ({ratio:.1f}x)")

        if warnings:
            warn_frame = ttk.Frame(self)
            warn_frame.pack(fill="x", padx=10, pady=(5, 0))
            for w in warnings:
                ttk.Label(warn_frame, text=f"  WARNING: {w}", foreground="#E65100", font=("Arial", 8)).pack(anchor="w")

        # ── Table ──
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=(5, 0))

        cols = ("Product", "SKU", "Original ฿", "New ฿", "Change ฿", "Change %")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=20)
        for col_name, w in [("Product", 180), ("SKU", 120), ("Original ฿", 80), ("New ฿", 80), ("Change ฿", 80), ("Change %", 70)]:
            self.tree.heading(col_name, text=col_name)
            self.tree.column(col_name, width=w, anchor="center" if col_name != "Product" else "w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Populate
        for r in rows:
            if r.row_idx not in changes:
                continue
            orig, new, capped = changes[r.row_idx]
            if orig == new:
                continue
            pct = (new - orig) / orig * 100 if orig else 0
            sign = "+" if new - orig >= 0 else ""
            fill, _ = get_fill(orig, new)
            tag = "green" if new > orig else "red" if new < orig else ""
            if capped:
                tag = "orange"
            self.tree.insert("", "end", values=(
                r.product_name[:25],
                r.sku[:12],
                f"{orig:.0f}",
                f"{new:.0f}",
                f"{sign}{new - orig:.0f}",
                f"{sign}{pct:.1f}%",
            ), tags=(tag,))

        self.tree.tag_configure("green", foreground=COLORS_GREEN_DARK, background=COLORS_GREEN_BAND)
        self.tree.tag_configure("red", foreground=COLORS_RED_DARK, background=COLORS_RED_PALE)
        self.tree.tag_configure("orange", foreground="#E65100", background=COLORS_ORANGE)

        # ── Buttons ──
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=10, pady=(5, 10))
        ttk.Label(btn_frame, text=f"Total price range: {min((changes.get(r.row_idx, (r.price, r.price, False))[1] for r in rows if r.price is not None), default=0):.0f} — {max((changes.get(r.row_idx, (r.price, r.price, False))[1] for r in rows if r.price is not None), default=0):.0f}").pack(anchor="w")

        btns = ttk.Frame(btn_frame)
        btns.pack(fill="x", pady=5)
        self.result = None

        ttk.Button(btns, text="Apply", width=12, command=self._apply).pack(side="left", padx=5)
        ttk.Button(btns, text="Discard", width=12, command=self._discard).pack(side="right", padx=5)

    def _apply(self):
        self.result = "apply"
        self.destroy()

    def _discard(self):
        self.result = "discard"
        self.destroy()


# ═══════════════════════════════════════════════════════════
#  SIZE SELECTION DIALOG
# ═══════════════════════════════════════════════════════════
class SizeSelectorDialog(tk.Toplevel):
    """Multi-select dialog for choosing size ranges."""

    def __init__(self, parent, size_map, sorted_sizes, unit_labels):
        super().__init__(parent)
        self.transient(parent)
        self.title("Select Sizes")
        self.geometry("400x500")
        self.grab_set()
        self.selected = set()
        self.result = None
        self.checkboxes = {}
        self.sorted_sizes = sorted_sizes

        x = parent.winfo_x() + (parent.winfo_width() - 400) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 450) // 2
        self.geometry(f"+{x}+{y}")

        # ── Title ──
        ttk.Label(self, text="Select sizes",
                  font=("Arial", 10, "bold")).pack(pady=(10, 5))

        # ── Canvas with scrollbar for checkboxes ──
        canvas_frame = ttk.Frame(self)
        canvas_frame.pack(fill="both", expand=True, padx=10, pady=5)

        canvas = tk.Canvas(canvas_frame, height=340)
        vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)

        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        cb_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=cb_frame, anchor="nw")

        for i, size in enumerate(sorted_sizes):
            num, unit = size
            count = size_map[size]
            label = f"{num} {unit_labels[unit]} ({count} SKUs)"
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(cb_frame, text=label, variable=var)
            cb.pack(anchor="w", pady=2)
            self.checkboxes[size] = (cb, var)

        cb_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

        # ── Range buttons ──
        range_frame = ttk.Frame(self)
        range_frame.pack(fill="x", padx=10, pady=(0, 5))
        ttk.Button(range_frame, text="Select All", width=10,
                   command=self._select_all).pack(side="left", padx=2)
        ttk.Button(range_frame, text="Deselect All", width=10,
                   command=self._deselect_all).pack(side="left", padx=2)

        # ── OK/Cancel ──
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=10, pady=(5, 10))
        ttk.Button(btn_frame, text="OK", width=8, command=self._ok).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancel", width=8, command=self._cancel).pack(side="right", padx=5)

        self.bind("<Escape>", lambda e: self._cancel())

    def _select_all(self):
        for _, var in self.checkboxes.values():
            var.set(True)

    def _deselect_all(self):
        for _, var in self.checkboxes.values():
            var.set(False)

    def _ok(self):
        selected = [size for size, (_, var) in self.checkboxes.items() if var.get()]
        if not selected:
            messagebox.showwarning("No Selection", "Please select at least one size.")
            return
        self.result = selected
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# ═══════════════════════════════════════════════════════════
#  MAIN WINDOW
# ═══════════════════════════════════════════════════════════
class PriceEditorApp(tk.Tk):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1000x700")

        # State
        self.file_path = None
        self.rows = []
        self.original_prices = {}
        self.current_prices = {}
        self.all_changes = {}
        self.wb = None
        self.ws = None
        self.product_prices = {}
        self._last_action = None  # tag for tracking
        self._is_processing = False
        self.search_var = tk.StringVar(value="")
        self.max_price_var = tk.StringVar(value="")
        self.filtered_rows = []

        self._setup_menu()
        self._setup_toolbar()
        self._setup_statusbar()
        self._setup_table()

    # ── Menu ──
    def _setup_menu(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open...", accelerator="Ctrl+O", command=self._open_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)

        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="All SKUs (+/- %)", command=lambda: self._start_edit("all", "pct"))
        edit_menu.add_command(label="By Size (+/- %)", command=lambda: self._start_edit("size", "pct"))
        edit_menu.add_command(label="By Color (+/- %)", command=lambda: self._start_edit("color", "pct"))
        edit_menu.add_command(label="By SKU (+/- %)", command=lambda: self._start_edit("sku", "pct"))
        edit_menu.add_separator()
        edit_menu.add_command(label="VAT +7%", command=self._apply_vat)
        edit_menu.add_command(label="Markup x2.5 + VAT", command=self._apply_markup_vat)
        edit_menu.add_separator()
        edit_menu.add_command(label="All SKUs (+/- ฿)", command=lambda: self._start_edit("all", "flat"))
        edit_menu.add_command(label="By Size (+/- ฿)", command=lambda: self._start_edit("size", "flat"))
        edit_menu.add_command(label="By Color (+/- ฿)", command=lambda: self._start_edit("color", "flat"))
        edit_menu.add_command(label="By SKU (+/- ฿)", command=lambda: self._start_edit("sku", "flat"))

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self._about)

    # ── Status bar ──
    def _setup_statusbar(self):
        self.status_var = tk.StringVar(value="Ready — Open a file to begin")
        self.statusbar = ttk.Label(self, textvariable=self.status_var,
                                   relief="sunken", anchor="w")
        self.statusbar.pack(side="bottom", fill="x")

    # ── Toolbar ──
    def _setup_toolbar(self):
        tb = ttk.Frame(self)
        tb.pack(fill="x", padx=5, pady=(5, 0))

        ttk.Button(tb, text="📂 Open", width=10, command=self._open_file).pack(side="left", padx=2)
        ttk.Separator(tb, orient="vertical").pack(side="left", fill="y", padx=5)
        ttk.Button(tb, text="+/- Price %", width=12, command=lambda: self._start_edit("all", "pct")).pack(side="left", padx=2)
        ttk.Button(tb, text="+/- Price ฿", width=12, command=lambda: self._start_edit("all", "flat")).pack(side="left", padx=2)
        ttk.Button(tb, text="VAT +7%", width=10, command=self._apply_vat).pack(side="left", padx=2)
        ttk.Button(tb, text="x2.5+VAT", width=10, command=self._apply_markup_vat).pack(side="left", padx=2)
        ttk.Button(tb, text="📐 By Size", width=10, command=lambda: self._start_edit("size", "pct")).pack(side="left", padx=2)
        ttk.Button(tb, text="🎨 By Color", width=10, command=lambda: self._start_edit("color", "pct")).pack(side="left", padx=2)
        ttk.Button(tb, text="🏷 By SKU", width=10, command=lambda: self._start_edit("sku", "pct")).pack(side="left", padx=2)

        ttk.Separator(tb, orient="vertical").pack(side="left", fill="y", padx=5)
        ttk.Button(tb, text="💾 Save", width=10, command=self._save).pack(side="left", padx=2)
        ttk.Button(tb, text="👁 Preview", width=10, command=self._preview_all).pack(side="left", padx=2)

        # Search and price filters
        ttk.Label(tb, text="Search").pack(side="left", padx=(6, 2))
        ttk.Entry(tb, textvariable=self.search_var, width=18).pack(side="left", padx=2)

        ttk.Label(tb, text="Min price").pack(side="left", padx=(8, 2))
        self.min_price_var = tk.DoubleVar(value=0)
        self.min_price_spin = ttk.Spinbox(tb, from_=0, to=500000, increment=100,
                                          textvariable=self.min_price_var, width=12)
        self.min_price_spin.pack(side="left", padx=2)
        ttk.Label(tb, text="Max price").pack(side="left", padx=(8, 2))
        self.max_price_entry = ttk.Entry(tb, textvariable=self.max_price_var, width=10)
        self.max_price_entry.pack(side="left", padx=2)
        ttk.Button(tb, text="Clear", width=7, command=self._clear_filters).pack(side="left", padx=2)
        ttk.Button(tb, text="Filtered %", width=11, command=lambda: self._start_edit("filtered", "pct")).pack(side="left", padx=2)
        ttk.Button(tb, text="Filtered ฿", width=12, command=lambda: self._start_edit("filtered", "flat")).pack(side="left", padx=2)

        self.search_var.trace_add("write", self._on_filter_changed)
        self.min_price_var.trace_add("write", self._on_filter_changed)
        self.max_price_var.trace_add("write", self._on_filter_changed)

    # ── Table ──
    def _setup_table(self):
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=(5, 0))

        self.tree = ttk.Treeview(tree_frame, show="headings", height=25)
        cols = ("#", "Product", "SKU", "Variation", "Price ฿")
        self.tree["columns"] = cols
        widths = (35, 220, 110, 260, 80)
        for col_name, w in zip(cols, widths):
            self.tree.heading(col_name, text=col_name)
            self.tree.column(col_name, width=w, anchor="center" if col_name != "Product" else "w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.tag_configure("green", foreground=COLORS_GREEN_DARK)
        self.tree.tag_configure("red", foreground=COLORS_RED_DARK)
        self.tree.tag_configure("orange", foreground="#E65100")
        self.tree.tag_configure("normal", foreground="black")

        self.tree.bind("<Double-1>", lambda e: self._on_tree_click(e))

    # ── Open file ──
    def _open_file(self):
        if self._is_processing:
            return
        path = filedialog.askopenfilename(
            title="Open Shopee Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return

        self.status_var.set("Loading...")
        self.update_idletasks()

        try:
            wb, ws, rows = load_file(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file:\n{e}")
            self.status_var.set("Error — Open a file to begin")
            return

        # Filter data rows (must have product_id, product_name, or sku)
        data_rows = [r for r in rows if any([r.product_id, r.product_name, r.sku])]
        if not data_rows:
            messagebox.showwarning("No Data", "No product rows found in this file.")
            self.status_var.set("No data rows — Open a file to begin")
            return

        self.file_path = path
        self.rows = data_rows
        self.original_prices = {r.row_idx: r.price for r in data_rows if r.price is not None}
        self.current_prices = dict(self.original_prices)
        self.all_changes = {}
        self.wb = wb
        self.ws = ws

        # Compute product price range for warnings
        self._compute_product_prices()

        # Update UI
        n_products = unique_products(data_rows)
        n_skus = sum(1 for r in data_rows if r.price is not None)
        mn, mx = price_range(data_rows)
        self.status_var.set(f"Loaded: {os.path.basename(path)}  |  Products: {n_products}  |  SKUs: {n_skus}  |  Price: ฿{mn:.0f}–฿{mx:.0f}" if mn else "Loaded")

        # Populate tree
        self._populate_tree()

    def _parse_optional_price(self, raw):
        raw = str(raw or "").strip()
        if not raw:
            return None
        return float(raw.replace(",", ""))

    def _get_filter_bounds(self, show_errors=False):
        try:
            min_price = self.min_price_var.get()
            min_price = min_price if min_price > 0 else None
        except (tk.TclError, ValueError):
            if show_errors:
                messagebox.showerror("Invalid Filter", "Min price must be a number.")
            return None, None, False

        try:
            max_price = self._parse_optional_price(self.max_price_var.get())
        except ValueError:
            if show_errors:
                messagebox.showerror("Invalid Filter", "Max price must be a number.")
            return None, None, False

        if min_price is not None and max_price is not None and min_price > max_price:
            if show_errors:
                messagebox.showerror("Invalid Filter", "Min price cannot be greater than max price.")
            return None, None, False

        return min_price, max_price, True

    def _get_filtered_rows(self, show_errors=False):
        min_price, max_price, ok = self._get_filter_bounds(show_errors)
        if not ok:
            return []
        row_filter = make_search_price_filter(
            self.search_var.get(),
            min_price,
            max_price,
            self.original_prices,
        )
        return [r for r in self.rows if row_filter(r)]

    def _on_filter_changed(self, *args):
        if hasattr(self, "tree"):
            self._populate_tree()

    def _clear_filters(self):
        self.search_var.set("")
        self.min_price_var.set(0)
        self.max_price_var.set("")
        self._populate_tree()

    def _populate_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.filtered_rows = self._get_filtered_rows()
        display_rows = self.filtered_rows if self.rows else []
        for r in display_rows:
            price_str = f"{r.price:.0f}" if r.price is not None else ""
            tag = "normal"
            if r.row_idx in self.all_changes:
                orig, new, capped = self.all_changes[r.row_idx]
                if new > orig:
                    tag = "green"
                elif new < orig:
                    tag = "red"
                if capped:
                    tag = "orange"
            elif r.row_idx in self.original_prices:
                if r.price is not None and r.price != self.original_prices[r.row_idx]:
                    tag = "green" if r.price > self.original_prices[r.row_idx] else "red"

            self.tree.insert("", "end", values=(
                r.row_idx - HEADER_ROWS,
                r.product_name[:25],
                r.sku[:15],
                r.variation_name[:35],
                price_str,
            ), tags=(tag,))

        has_active_filter = (
            bool(self.search_var.get().strip())
            or self.min_price_var.get() > 0
            or bool(self.max_price_var.get().strip())
        )
        if self.file_path and self.rows and has_active_filter:
            self.status_var.set(f"Showing {len(display_rows)} of {len(self.rows)} rows")

    # ── Product price range warnings ──
    def _compute_product_prices(self):
        self.product_prices = {}
        for r in self.rows:
            pid = r.product_id
            if pid not in self.product_prices:
                self.product_prices[pid] = {"name": r.product_name, "prices": []}
            self.product_prices[pid]["prices"].append(self.current_prices.get(r.row_idx, r.price))

    # ── Edit flow ──
    def _start_edit_legacy_unused(self, group, edit_type):
        if not self.file_path or not self.rows:
            messagebox.showinfo("No File", "Please open a file first.")
            return
        if self._is_processing:
            return

        if edit_type == "pct_neg":
            val_raw = self._prompt_adjustment("pct", group)
            if val_raw is None:
                return
            try:
                value, detected_pct = self._parse_adj(val_raw)
                value = max(value, -99.0)  # prevent -100%
            except ValueError:
                messagebox.showerror("Invalid", "Please enter a valid percentage (e.g., -5% or -10)")
                return
            tag = make_tag(group, value, True)
        elif edit_type == "flat_neg":
            val_raw = self._prompt_adjustment("flat", group)
            if val_raw is None:
                return
            try:
                value, _ = self._parse_adj(val_raw)
            except ValueError:
                messagebox.showerror("Invalid", "Please enter a valid amount (e.g., -50)")
                return
            tag = make_tag(group, value, False)
        elif edit_type == "pct":
            if group == "size":
                self._edit_by_size(True)
                return
            elif group == "color":
                self._edit_by_color(True)
                return
            elif group == "sku":
                self._edit_by_sku(True)
                return
            else:
                val_raw = self._prompt_adjustment("pct", group)
                if val_raw is None:
                    return
                try:
                    value, detected_pct = self._parse_adj(val_raw)
                except ValueError:
                    messagebox.showerror("Invalid", "Please enter a valid percentage (e.g., 7%)")
                    return
                tag = make_tag(group, value, True)
        elif edit_type == "flat":
            if group == "size":
                self._edit_by_size(False)
                return
            elif group == "color":
                self._edit_by_color(False)
                return
            elif group == "sku":
                self._edit_by_sku(False)
                return
            else:
                val_raw = self._prompt_adjustment("flat", group)
                if val_raw is None:
                    return
                try:
                    value, _ = self._parse_adj(val_raw)
                except ValueError:
                    messagebox.showerror("Invalid", "Please enter a valid amount (e.g., 50)")
                    return
                tag = make_tag(group, value, False)
        else:
            return

        # Build changes
        min_price = self.min_price_var.get() if self.min_price_var.get() > 0 else None

        if group == "all":
            changes = build_changes(self.rows, lambda r: True, value, True, min_price)
        elif group == "size":
            changes = build_changes(self.rows, lambda r: True, value, True, min_price)
        elif group == "color":
            changes = build_changes(self.rows, lambda r: True, value, True, min_price)
        elif group == "sku":
            changes = build_changes(self.rows, lambda r: True, value, True, min_price)

        if not changes:
            messagebox.showwarning("No Changes", "No SKUs matched the criteria (or all prices are unchanged).")
            return

        # Preview
        self._show_preview(changes, tag)

    def _prompt_adjustment(self, kind, group):
        kind_str = "%" if kind == "pct" else "฿"
        prompt = f"Enter {kind_str} adjustment for {group}:"
        default = "+7" if kind == "pct" else "+50"
        result = self._ask_user_input(prompt, default)
        return result

    def _ask_user_input(self, title, default):
        """Simple modal input dialog."""
        d = tk.Toplevel(self)
        d.transient(self)
        d.title(title)
        d.geometry("300x150")
        d.grab_set()

        ttk.Label(d, text=title).pack(pady=(10, 5))
        var = tk.StringVar(value=default)
        ttk.Entry(d, textvariable=var, width=20).pack(pady=5)

        def ok():
            d.result = var.get()
            d.destroy()

        ttk.Button(d, text="OK", command=ok).pack(pady=5)
        ttk.Button(d, text="Cancel", command=lambda: d.destroy()).pack()

        self.wait_window(d)
        return getattr(d, "result", None)

    def _parse_adj(self, raw):
        raw = raw.strip()
        is_pct = raw.endswith('%')
        num_str = raw.rstrip('%').strip()
        value = float(num_str)
        return value, is_pct

    # ── Interactive: By Size ──
    def _edit_by_size(self, is_pct):
        # Collect sizes
        size_map = {}
        for r in self.rows:
            if r.size_num is None:
                continue
            key = (r.size_num, r.unit)
            size_map[key] = size_map.get(key, 0) + 1

        def sort_key(k):
            num, unit = k
            return (0 if unit == "เดือน" else 1, num)

        sorted_sizes = sorted(size_map.keys(), key=sort_key)

        unit_labels = {"ปี": "years", "เดือน": "months"}
        dlg = SizeSelectorDialog(self, size_map, sorted_sizes, unit_labels)
        self.wait_window(dlg)
        selected = getattr(dlg, "result", None)
        if not selected:
            return

        # Get adjustment
        kind = "%" if is_pct else "฿"
        prompt = f"Enter adjustment for selected sizes ({kind}):"
        val_raw = self._ask_user_input(prompt, "+7%")
        if val_raw is None:
            return
        try:
            value, detected_pct = self._parse_adj(val_raw)
        except ValueError:
            messagebox.showerror("Invalid", "Please enter a valid amount or percentage.")
            return

        min_price = self.min_price_var.get() if self.min_price_var.get() > 0 else None
        changes = build_changes(
            self.rows,
            make_size_filter(selected),
            value, True, min_price,
        )
        tag = make_tag("size", value, True)

        if not changes:
            messagebox.showwarning("No Changes", "No SKUs matched the criteria.")
            return

        self._show_preview(changes, tag)

    # ── Interactive: By Color ──
    def _edit_by_color(self, is_pct):
        color_map = {}
        for r in self.rows:
            if r.color is None:
                continue
            color_map[r.color] = color_map.get(r.color, 0) + 1

        sorted_colors = sorted(color_map.keys(), key=lambda c: -color_map[c])

        if not sorted_colors:
            messagebox.showinfo("No Color", "No color data found.")
            return

        d = tk.Toplevel(self)
        d.transient(self)
        d.title("Select Colors")
        d.geometry("400x500")
        d.grab_set()
        checkboxes = {}

        ttk.Label(d, text="Select colors (English translation in brackets)",
                  font=("Arial", 9, "bold")).pack(pady=(10, 5))

        canvas_frame = ttk.Frame(d)
        canvas_frame.pack(fill="both", expand=True, padx=10, pady=5)

        canvas = tk.Canvas(canvas_frame, height=380)
        vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)

        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        cb_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=cb_frame, anchor="nw")

        for c in sorted_colors:
            eng = THAI_COLORS.get(c, "")
            cnt = color_map[c]
            if eng:
                label = f"{c} ({eng}) - {cnt} SKUs"
            else:
                label = f"{c} (no translation) - {cnt} SKUs"
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(cb_frame, text=label, variable=var)
            cb.pack(anchor="w", pady=2)
            checkboxes[c] = var

        cb_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

        range_frame = ttk.Frame(d)
        range_frame.pack(fill="x", padx=10, pady=(0, 5))
        ttk.Button(range_frame, text="Select All", width=12,
                   command=lambda: [v.set(True) for v in checkboxes.values()]).pack(side="left", padx=2)
        ttk.Button(range_frame, text="Deselect All", width=12,
                   command=lambda: [v.set(False) for v in checkboxes.values()]).pack(side="left", padx=2)

        btn_frame = ttk.Frame(d)
        btn_frame.pack(fill="x", padx=10, pady=(5, 10))
        ttk.Button(btn_frame, text="OK", width=8, command=lambda: _ok_color()).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancel", width=8, command=d.destroy).pack(side="right", padx=5)

        def _ok_color():
            selected = [c for c, var in checkboxes.items() if var.get()]
            if not selected:
                messagebox.showwarning("No Selection", "Please select at least one color.")
                return
            d.result = selected
            d.destroy()

        self.wait_window(d)
        selected_colors = getattr(d, "result", None)
        if not selected_colors:
            return

        kind = "%" if is_pct else "฿"
        prompt = f"Enter adjustment for selected colors ({kind}):"
        val_raw = self._ask_user_input(prompt, "+7%")
        if val_raw is None:
            return
        try:
            value, _ = self._parse_adj(val_raw)
        except ValueError:
            messagebox.showerror("Invalid", "Please enter a valid amount or percentage.")
            return

        min_price = self.min_price_var.get() if self.min_price_var.get() > 0 else None
        changes = build_changes(
            self.rows,
            make_color_filter(selected_colors),
            value, True, min_price,
        )
        tag = make_tag("color", value, True)

        if not changes:
            messagebox.showwarning("No Changes", "No SKUs matched the criteria.")
            return

        self._show_preview(changes, tag)

    # ── Interactive: By SKU ──
    def _edit_by_sku(self, is_pct):
        d = tk.Toplevel(self)
        d.transient(self)
        d.title("Search SKUs")
        d.geometry("350x350")
        d.grab_set()

        ttk.Label(d, text="Enter part of SKU code:", font=("Arial", 9, "bold")).pack(pady=(5, 0))
        ttk.Label(d, text="Changes will apply to all SKUs matching this text.",
                  foreground="#666", font=("Arial", 8)).pack()

        ttk.Separator(d, orient="horizontal").pack(fill="x", pady=5)

        var = tk.StringVar()
        ttk.Entry(d, textvariable=var, width=30).pack(pady=5)

        def search():
            token = var.get().strip()
            if not token:
                messagebox.showwarning("Empty", "Please enter SKU text.")
                return
            matched = [r.sku for r in self.rows if token in r.sku]
            if not matched:
                messagebox.showinfo("Not Found", f"No SKUs containing '{token}'.")
                return

            d2 = tk.Toplevel(d)
            d2.transient(d)
            d2.title(f"Found {len(matched)} SKUs")
            d2.geometry("400x250")
            d2.grab_set()

            ttk.Label(d2, text=f"Found {len(matched)} SKUs matching '{token}':",
                      font=("Arial", 9, "bold")).pack(pady=(5, 0))

            lb = tk.Listbox(d2, height=8, font=("Courier New", 8))
            vsb = ttk.Scrollbar(d2, orient="vertical", command=lb.yview)
            lb.configure(yscrollcommand=vsb.set)
            lb.pack(side="left", fill="both", expand=True)
            vsb.pack(side="right", fill="y")
            for s in matched:
                lb.insert("end", s)

            ttk.Label(d2, text="Adjustment:", font=("Arial", 9)).pack(pady=(5, 0))
            adj_var = tk.StringVar(value="+7%")
            ttk.Entry(d2, textvariable=adj_var, width=20).pack()

            kind = "%" if is_pct else "฿"
            ttk.Label(d2, text=f"{kind}").pack(pady=(0, 5))

            def apply_sku():
                val_raw = adj_var.get()
                try:
                    value, _ = self._parse_adj(val_raw)
                except ValueError:
                    messagebox.showerror("Invalid", "Please enter a valid amount or percentage.")
                    return

                min_price = self.min_price_var.get() if self.min_price_var.get() > 0 else None
                changes = build_changes(
                    self.rows,
                    make_sku_filter(token),
                    value, True, min_price,
                )
                tag = make_tag("sku", value, True)

                d2.destroy()

                if not changes:
                    messagebox.showwarning("No Changes", "No SKUs matched the criteria.")
                    return

                self._show_preview(changes, tag)

            ttk.Button(d2, text="Apply", command=apply_sku).pack(pady=5)
            ttk.Button(d2, text="Cancel", command=d2.destroy).pack()

        ttk.Button(d, text="Search", command=search).pack(pady=5)
        ttk.Button(d, text="Cancel", command=d.destroy).pack()

        self.wait_window(d)

    # ── Preview ──
    def _show_preview(self, changes, tag):
        # Store for apply
        self._last_changes = changes
        self._last_tag = tag
        self._last_action = tag

        self.status_var.set(f"Previewing: {tag} — {len(changes)} SKUs affected")

        d = PreviewDialog(
            self, self.rows, changes,
            self.original_prices, self.current_prices,
        )
        self.wait_window(d)

        self._apply_changes(changes, tag)

    def _apply_changes(self, changes, tag):
        self._is_processing = True
        self.status_var.set(f"Applying {tag}...")
        self.update_idletasks()

        for idx, (_, new_p, capped) in changes.items():
            self.current_prices[idx] = new_p
            self.all_changes[idx] = (self.original_prices.get(idx, 0), new_p, capped)
            row = None
            for r in self.rows:
                if r.row_idx == idx:
                    row = r
                    break
            if row:
                row.price = new_p

        self._compute_product_prices()
        self._populate_tree()

        changed = sum(1 for (o, n, c) in changes.values() if o != n)
        self.status_var.set(f"Applied: {tag} — {changed} SKUs changed")
        self._is_processing = False

    # ── Preview all (current state) ──
    def _preview_all(self):
        if not self.all_changes:
            messagebox.showinfo("No Changes", "No changes to preview.")
            return
        d = PreviewDialog(
            self, self.rows, self.all_changes,
            self.original_prices, self.current_prices,
        )
        self.wait_window(d)
        if getattr(d, "result", None) == "apply":
            # Already applied in _show_preview
            pass

    # ── Save ──
    def _save(self):
        if not self.file_path:
            messagebox.showinfo("No File", "Please open a file first.")
            return
        if not self.all_changes:
            messagebox.showinfo("No Changes", "No changes to save.")
            return

        self._is_processing = True
        self.status_var.set("Saving...")
        self.update_idletasks()

        try:
            out_path = make_output_path(self.file_path, suffix="updated", edit_label=self._last_action or "")
            save_file(self.wb, self.ws, self.rows, self.all_changes, out_path)

            # Summary report
            try:
                summary_path = save_summary_report(
                    self.original_prices, self.current_prices,
                    self.rows, out_path,
                )
            except Exception:
                summary_path = None

            # Reset for fresh session
            self.all_changes = {}
            self._last_action = None

            self.status_var.set(f"Saved: {os.path.basename(out_path)}")
            messagebox.showinfo("Saved", f"File saved:\n{out_path}"
                                f"{chr(10)}Summary: {os.path.basename(summary_path) if summary_path else 'skipped'}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save:\n{e}")
            self.status_var.set("Save failed")
        self._is_processing = False

    def _about(self):
        messagebox.showinfo("About",
            f"{APP_TITLE}\n\n"
            "Bulk price editor for Shopee Thailand Excel files.\n"
            "Supports percentage and flat adjustments.\n"
            "Filter: cheap SKUs below min price are skipped.\n\n"
            f"Build: {BUILD_LABEL}")

    def _on_tree_click(self, event):
        """Double-click to edit that row's price."""
        item = self.tree.focus()
        if not item:
            return
        values = self.tree.item(item, "values")
        idx = int(values[0]) if values else None
        if idx is None:
            return
        for r in self.rows:
            if r.row_idx - HEADER_ROWS == idx:
                new_price = self._ask_user_input(f"Edit price for {r.product_name[:30]}", f"{r.price:.0f}")
                if new_price is None:
                    return
                try:
                    new_val = float(new_price)
                except ValueError:
                    messagebox.showerror("Invalid", "Must be a number.")
                    return

                # This is a single SKU edit — wrap as dict
                changes = {r.row_idx: (r.price, new_val, False)}
                self.current_prices[r.row_idx] = new_val
                self.all_changes[r.row_idx] = (r.price, new_val, False)
                self._last_action = "single"
                self._populate_tree()
                self.status_var.set(f"Updated: ฿{r.price:.0f} → ฿{new_val:.0f}")
                break

    def _edit_filtered(self, is_pct):
        filtered = self._get_filtered_rows(show_errors=True)
        if not filtered:
            messagebox.showwarning("No Matches", "No rows match the current search and price filters.")
            return

        kind = "pct" if is_pct else "flat"
        val_raw = self._prompt_adjustment(kind, "filtered rows")
        if val_raw is None:
            return

        try:
            value, _ = self._parse_adj(val_raw)
            if is_pct:
                value = max(value, -99.0)
        except ValueError:
            messagebox.showerror("Invalid", "Please enter a valid amount or percentage.")
            return

        row_ids = {r.row_idx for r in filtered}
        changes = build_changes(
            self.rows,
            lambda r: r.row_idx in row_ids,
            value,
            is_pct,
            None,
        )
        if not changes:
            messagebox.showwarning("No Changes", "No filtered SKUs can be changed.")
            return

        tag = make_tag("filtered", value, is_pct)
        self._show_preview(changes, tag)

    def _apply_vat(self):
        if not self.file_path or not self.rows:
            messagebox.showinfo("No File", "Please open a file first.")
            return
        if self._is_processing:
            return

        min_price = self.min_price_var.get() if self.min_price_var.get() > 0 else None
        changes = {}
        for r in self.rows:
            if r.price is None:
                continue
            if min_price is not None and r.price < min_price:
                continue
            new_p, capped = calc_price_vat(r.price)
            changes[r.row_idx] = (r.price, new_p, capped)

        if not changes:
            messagebox.showwarning("No Changes", "No SKUs can be changed.")
            return

        tag = "vat_p7pct"
        self._show_preview(changes, tag)

    def _apply_markup_vat(self):
        if not self.file_path or not self.rows:
            messagebox.showinfo("No File", "Please open a file first.")
            return
        if self._is_processing:
            return

        min_price = self.min_price_var.get() if self.min_price_var.get() > 0 else None
        changes = {}
        for r in self.rows:
            if r.price is None:
                continue
            if min_price is not None and r.price < min_price:
                continue
            new_p, capped = calc_price_markup_vat(r.price)
            changes[r.row_idx] = (r.price, new_p, capped)

        if not changes:
            messagebox.showwarning("No Changes", "No SKUs can be changed.")
            return

        tag = "markup_x2.5_vat_p7pct"
        self._show_preview(changes, tag)

    def _start_edit(self, group, edit_type):
        """Start a bulk edit flow."""
        if not self.file_path or not self.rows:
            messagebox.showinfo("No File", "Please open a file first.")
            return
        if self._is_processing:
            return

        if group == "filtered":
            self._edit_filtered(edit_type == "pct")
            return

        # Handle special negated types
        if edit_type == "pct_neg":
            val_raw = self._prompt_adjustment("pct", group)
            if val_raw is None:
                return
            try:
                value, detected_pct = self._parse_adj(val_raw)
                value = max(value, -99.0)
            except ValueError:
                messagebox.showerror("Invalid", "Please enter a valid percentage (e.g., -5%)")
                return
            tag = make_tag(group, value, True)
            min_price = self.min_price_var.get() if self.min_price_var.get() > 0 else None
            changes = build_changes(self.rows, lambda r: True, value, True, min_price)
        elif edit_type == "flat_neg":
            val_raw = self._prompt_adjustment("flat", group)
            if val_raw is None:
                return
            try:
                value, _ = self._parse_adj(val_raw)
            except ValueError:
                messagebox.showerror("Invalid", "Please enter a valid amount (e.g., -50)")
                return
            tag = make_tag(group, value, False)
            min_price = self.min_price_var.get() if self.min_price_var.get() > 0 else None
            changes = build_changes(self.rows, lambda r: True, value, False, min_price)
        else:
            if group == "size":
                self._edit_by_size(edit_type == "pct")
                return
            elif group == "color":
                self._edit_by_color(edit_type == "pct")
                return
            elif group == "sku":
                self._edit_by_sku(edit_type == "pct")
                return
            else:
                is_pct = edit_type == "pct"
                val_raw = self._prompt_adjustment(edit_type, group)
                if val_raw is None:
                    return
                try:
                    value = float(val_raw.strip().rstrip('%').strip())
                except ValueError:
                    messagebox.showerror("Invalid", "Please enter a valid amount or percentage.")
                    return
                tag = make_tag(group, value, is_pct)
                min_price = self.min_price_var.get() if self.min_price_var.get() > 0 else None
                changes = build_changes(self.rows, lambda r: True, value, is_pct, min_price)

        if not changes:
            messagebox.showwarning("No Changes", "No SKUs matched the criteria.")
            return

        self._show_preview(changes, tag)


# ═══════════════════════════════════════════════════════════
#  Entry point
# ═══════════════════════════════════════════════════════════
def main():
    app = PriceEditorApp()
    app.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print(f"ERROR: {e}")
        traceback.print_exc()
