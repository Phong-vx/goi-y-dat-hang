#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gợi Ý Đặt Hàng
- File bán hàng : data.warehouse export (giao dịch thô, cột Date + Quantity)
- File tồn kho  : stock.quant export   (nhiều dòng/địa điểm, cột Số lượng)
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, re
from datetime import datetime
import sys

# ─── Apple-style palette ─────────────────────────────────────────────────────
C = {
    'bg':         '#F5F5F7',   # Apple light gray
    'card':       '#FFFFFF',
    'border':     '#D2D2D7',
    'sep':        '#E5E5EA',
    'primary':    '#0071E3',   # Apple blue
    'primary_dk': '#005BBB',
    'text':       '#1D1D1F',
    'text2':      '#6E6E73',
    'green':      '#34C759',
    'green_dk':   '#248A3D',
    'red_light':  '#FFF0F0',
    'red_text':   '#C0392B',
    'blue_light': '#F0F7FF',
    'input_bg':   '#F5F5F7',
    'check_sel':  '#E8F3FF',
    'console_bg': '#1C1C1E',
    'console_fg': '#32D74B',
    'tag_bg':     '#E8F3FF',
    'tag_fg':     '#0058BD',
}

SALES = {
    'sku':      'SKU',
    'name':     'Product Item',
    'brand':    'Brand',
    'category': 'Category',
    'date':     'Date',
    'qty':      'Quantity',
}
INV = {
    'sku':      'Sản phẩm/Mã nội bộ',
    'qty':      'Số lượng',
    'brand':    'Sản phẩm/Brand/Display Name',
    'category': 'Sản phẩm/Nganh Hang/Name',
}


def strip_sku_prefix(t: str) -> str:
    return re.sub(r'^\[.*?\]\s*', '', str(t)).strip()


# ─── Reusable widgets ─────────────────────────────────────────────────────────

def make_card(parent, title, subtitle=''):
    """Apple-style card: white rounded-ish frame with title + hairline separator."""
    outer = tk.Frame(parent, bg=C['border'], padx=1, pady=1)
    inner = tk.Frame(outer, bg=C['card'])
    inner.pack(fill=tk.BOTH, expand=True)

    hdr = tk.Frame(inner, bg=C['card'], padx=18, pady=10)
    hdr.pack(fill=tk.X)
    tk.Label(hdr, text=title, font=('Helvetica Neue', 11, 'bold'),
             bg=C['card'], fg=C['text']).pack(side=tk.LEFT)
    if subtitle:
        tk.Label(hdr, text=subtitle, font=('Helvetica Neue', 9),
                 bg=C['card'], fg=C['text2']).pack(side=tk.LEFT, padx=(8, 0))

    tk.Frame(inner, bg=C['sep'], height=1).pack(fill=tk.X)

    body = tk.Frame(inner, bg=C['card'], padx=18, pady=12)
    body.pack(fill=tk.BOTH, expand=True)

    return outer, body


def make_btn(parent, text, command, style='primary', small=False):
    """Flat Apple-style button."""
    if style == 'primary':
        bg, fg, abg = C['primary'],   'white', C['primary_dk']
    elif style == 'green':
        bg, fg, abg = C['green'],     'white', C['green_dk']
    elif style == 'ghost':
        bg, fg, abg = C['blue_light'], C['primary'], '#DAEEFF'
    else:
        bg, fg, abg = C['sep'], C['text'], '#CACACE'

    font = ('Helvetica Neue', 9 if small else 11, 'bold')
    px, py = (10, 4) if small else (20, 9)
    return tk.Button(parent, text=text, command=command,
                     font=font, bg=bg, fg=fg, bd=0, relief='flat',
                     cursor='hand2', padx=px, pady=py,
                     activebackground=abg, activeforeground=fg)


# ─── Scrollable multi-select panel ───────────────────────────────────────────

class FilterPanel(tk.Frame):
    """
    Scrollable checkbox panel.
    Không chọn gì = không lọc (lấy tất cả).
    """
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=C['card'], **kw)
        self._vars: dict[str, tk.BooleanVar] = {}
        self._widgets: list = []
        self._build()

    def _build(self):
        # Toolbar
        tb = tk.Frame(self, bg=C['card'])
        tb.pack(fill=tk.X, pady=(0, 6))

        make_btn(tb, '✓ Tất Cả', lambda: self._set_all(True),  style='ghost', small=True).pack(side=tk.LEFT, padx=(0, 6))
        make_btn(tb, '✕ Bỏ Chọn', lambda: self._set_all(False), style='neutral', small=True).pack(side=tk.LEFT)

        self.lbl_count = tk.Label(tb, text='—  chưa tải', font=('Helvetica Neue', 9),
                                   bg=C['card'], fg=C['text2'])
        self.lbl_count.pack(side=tk.RIGHT)

        # Scroll area
        wrap = tk.Frame(self, bg=C['input_bg'], bd=1, relief='solid',
                        highlightbackground=C['border'], highlightthickness=1)
        wrap.pack(fill=tk.BOTH, expand=True)

        self._canvas = tk.Canvas(wrap, bg=C['card'], height=150,
                                  highlightthickness=0, bd=0)
        vsb = tk.Scrollbar(wrap, orient='vertical', command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._inner = tk.Frame(self._canvas, bg=C['card'])
        self._win   = self._canvas.create_window((0, 0), window=self._inner, anchor='nw')

        self._inner.bind('<Configure>', lambda e: self._canvas.configure(
            scrollregion=self._canvas.bbox('all')))
        self._canvas.bind('<Configure>', lambda e: self._canvas.itemconfig(
            self._win, width=e.width))
        self._canvas.bind('<MouseWheel>', lambda e: self._canvas.yview_scroll(
            int(-1 * (e.delta / 120)), 'units'))

        # Placeholder
        self._ph = tk.Label(self._inner,
                             text='Nhấn  "Đọc Files"  để tải danh sách',
                             font=('Helvetica Neue', 9, 'italic'),
                             bg=C['card'], fg=C['text2'], pady=24)
        self._ph.pack()

    def populate(self, items: list):
        for w in self._inner.winfo_children():
            w.destroy()
        self._vars.clear()
        self._widgets.clear()

        cols = 2
        for idx, item in enumerate(items):
            var = tk.BooleanVar(value=False)  # mặc định KHÔNG chọn = không lọc
            self._vars[item] = var
            cb = tk.Checkbutton(
                self._inner, text=item, variable=var,
                font=('Helvetica Neue', 10), bg=C['card'], fg=C['text'],
                selectcolor=C['check_sel'], activebackground=C['card'],
                anchor='w', command=self._update_count,
            )
            r, c = divmod(idx, cols)
            cb.grid(row=r, column=c, sticky='w', padx=10, pady=2)
            self._widgets.append(cb)

        for c in range(cols):
            self._inner.columnconfigure(c, weight=1)

        self._update_count()

    def _set_all(self, val: bool):
        for v in self._vars.values():
            v.set(val)
        self._update_count()

    def _update_count(self):
        total    = len(self._vars)
        selected = sum(1 for v in self._vars.values() if v.get())
        if total == 0:
            self.lbl_count.config(text='—  chưa tải', fg=C['text2'])
        elif selected == 0:
            self.lbl_count.config(text='Trống = lấy tất cả', fg=C['text2'])
        else:
            self.lbl_count.config(text=f'{selected} / {total} đã chọn',
                                   fg=C['primary'])

    def selected(self) -> list:
        """Trả về list item đã chọn. Rỗng = không lọc (lấy tất cả)."""
        return [k for k, v in self._vars.items() if v.get()]

    def all_items(self) -> list:
        return list(self._vars.keys())


# ─── Main App ─────────────────────────────────────────────────────────────────

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title('Gợi Ý Đặt Hàng')
        self.root.geometry('900x860')
        self.root.configure(bg=C['bg'])
        self.root.resizable(True, True)

        self.v_sales  = tk.StringVar()
        self.v_inv    = tk.StringVar()
        self.v_months = tk.StringVar(value='3')

        self._build_ui()
        self.log('✅ Ứng dụng khởi động thành công.')
        self.log('📌 Bước 1: Chọn 2 file  →  Bước 2: Đọc Files  →  Bước 3: Lọc (tuỳ chọn)  →  Bước 4: Tạo Gợi Ý.')

    # ── Build UI ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header ───────────────────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg=C['primary'])
        hdr.pack(fill=tk.X)

        hdr_inner = tk.Frame(hdr, bg=C['primary'], padx=32, pady=20)
        hdr_inner.pack(fill=tk.X)

        tk.Label(hdr_inner, text='Gợi Ý Đặt Hàng',
                 font=('Helvetica Neue', 22, 'bold'),
                 bg=C['primary'], fg='white').pack(anchor='w')
        tk.Label(hdr_inner,
                 text='Phân tích dữ liệu bán hàng & tồn kho · Xuất file gợi ý đặt hàng',
                 font=('Helvetica Neue', 11), bg=C['primary'], fg='#A8CAFF').pack(anchor='w', pady=(2, 0))

        # ── Scrollable body ───────────────────────────────────────────────────
        outer  = tk.Frame(self.root, bg=C['bg'])
        outer.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(outer, bg=C['bg'], highlightthickness=0)
        vsb    = tk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.body = tk.Frame(canvas, bg=C['bg'], padx=28, pady=20)
        win = canvas.create_window((0, 0), window=self.body, anchor='nw')

        self.body.bind('<Configure>', lambda e: canvas.configure(
            scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(win, width=e.width))
        canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(
            int(-1 * (e.delta / 120)), 'units'))

        self._build_body()

    def _build_body(self):
        b = self.body

        # ── 1. Import Files ───────────────────────────────────────────────────
        card, body = make_card(b, '① Import Files')
        self._file_row(body, 'File Bán Hàng', self.v_sales)
        sep = tk.Frame(body, bg=C['sep'], height=1)
        sep.pack(fill=tk.X, pady=6)
        self._file_row(body, 'File Tồn Kho', self.v_inv)

        btn_row = tk.Frame(body, bg=C['card'])
        btn_row.pack(fill=tk.X, pady=(10, 0))
        self.btn_read = make_btn(btn_row, '  Đọc Files & Tải Bộ Lọc  →',
                                  self._read_files, style='green')
        self.btn_read.pack(side=tk.RIGHT)
        card.pack(fill=tk.X, pady=(0, 12))

        # ── 2. Bộ lọc Brand + Category (2 cột) ───────────────────────────────
        filter_row = tk.Frame(b, bg=C['bg'])
        filter_row.pack(fill=tk.X, pady=(0, 12))
        filter_row.columnconfigure(0, weight=1)
        filter_row.columnconfigure(1, weight=1)

        # Brand
        bc, bbody = make_card(filter_row, '② Thương Hiệu', '(tuỳ chọn · trống = tất cả)')
        self.brand_panel = FilterPanel(bbody)
        self.brand_panel.pack(fill=tk.BOTH, expand=True)
        bc.grid(row=0, column=0, sticky='nsew', padx=(0, 6))

        # Category
        cc, cbody = make_card(filter_row, '③ Danh Mục', '(tuỳ chọn · trống = tất cả)')
        self.cat_panel = FilterPanel(cbody)
        self.cat_panel.pack(fill=tk.BOTH, expand=True)
        cc.grid(row=0, column=1, sticky='nsew', padx=(6, 0))

        # ── 3. Cài đặt số tháng ───────────────────────────────────────────────
        sc, sbody = make_card(b, '④ Tồn Kho Tối Thiểu')
        row = tk.Frame(sbody, bg=C['card'])
        row.pack(fill=tk.X)
        tk.Label(row, text='Số tháng bán hàng cần tồn kho :',
                 font=('Helvetica Neue', 11), bg=C['card'], fg=C['text']).pack(side=tk.LEFT)

        rb_frame = tk.Frame(row, bg=C['card'])
        rb_frame.pack(side=tk.LEFT, padx=16)
        for val, lbl in [('3','3 tháng'), ('6','6 tháng'), ('9','9 tháng')]:
            tk.Radiobutton(rb_frame, text=lbl, variable=self.v_months, value=val,
                           font=('Helvetica Neue', 11), bg=C['card'], fg=C['text'],
                           selectcolor=C['check_sel'], activebackground=C['card']
                           ).pack(side=tk.LEFT, padx=10)
        sc.pack(fill=tk.X, pady=(0, 12))

        # ── CTA button ────────────────────────────────────────────────────────
        self.btn_run = make_btn(b, '       🚀   Tạo Gợi Ý Đặt Hàng       ',
                                 self.run, style='primary')
        self.btn_run.config(font=('Helvetica Neue', 13, 'bold'), pady=13)
        self.btn_run.pack(fill=tk.X, pady=(0, 12))

        # ── Log ───────────────────────────────────────────────────────────────
        lc, lbody = make_card(b, 'Nhật Ký')
        lbody.config(padx=0, pady=0)
        self.log_box = tk.Text(lbody, height=8, font=('Menlo', 9),
                                bg=C['console_bg'], fg=C['console_fg'],
                                padx=14, pady=8, wrap=tk.WORD, bd=0,
                                insertbackground=C['console_fg'])
        self.log_box.pack(fill=tk.BOTH, expand=True)
        lc.pack(fill=tk.X)

    # ── Widgets helpers ───────────────────────────────────────────────────────

    def _file_row(self, parent, label, var):
        f = tk.Frame(parent, bg=C['card'])
        f.pack(fill=tk.X)

        tk.Label(f, text=label, font=('Helvetica Neue', 11),
                 bg=C['card'], fg=C['text'], width=16, anchor='w').pack(side=tk.LEFT)

        entry = tk.Entry(f, textvariable=var, font=('Helvetica Neue', 10),
                         bg=C['input_bg'], fg=C['text2'], relief='flat',
                         bd=0, state='readonly', readonlybackground=C['input_bg'])
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 10), ipady=5)

        make_btn(f, 'Chọn File', lambda v=var: self._browse(v),
                 style='ghost', small=True).pack(side=tk.RIGHT)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _browse(self, var):
        path = filedialog.askopenfilename(
            filetypes=[('Excel files', '*.xlsx *.xls'), ('All files', '*.*')])
        if path:
            var.set(path)
            self.log(f'📂  {os.path.basename(path)}')

    def log(self, msg: str):
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_box.insert(tk.END, f'[{ts}]  {msg}\n')
        self.log_box.see(tk.END)
        self.root.update()

    # ── Đọc files & populate filters ─────────────────────────────────────────

    def _read_files(self):
        s = self.v_sales.get()
        i = self.v_inv.get()
        if not s:
            messagebox.showerror('Lỗi', 'Vui lòng chọn file bán hàng!'); return
        if not i:
            messagebox.showerror('Lỗi', 'Vui lòng chọn file tồn kho!'); return

        try:
            self.btn_read.config(state='disabled', text='⏳  Đang đọc…')
            self.log('─── Đọc files ───')

            sales = pd.read_excel(s)
            sales.columns = [str(c).strip() for c in sales.columns]
            self._check_cols(sales, [SALES['brand'], SALES['category']], 'file bán hàng')

            # Lọc SERVICE/COUPON
            sales = sales[
                sales[SALES['sku']].notna() &
                (~sales[SALES['sku']].astype(str).str.upper().str.contains('COUPON|DISCOUNT', na=False)) &
                (sales[SALES['category']].astype(str).str.upper() != 'SERVICE')
            ]

            inv = pd.read_excel(i)
            inv.columns = [str(c).strip() for c in inv.columns]

            def clean(series):
                return set(series.dropna().astype(str).str.strip()
                           .replace('', pd.NA).dropna().unique())

            # Brand
            brands = sorted(
                clean(sales[SALES['brand']]) |
                (clean(inv[INV['brand']]) if INV['brand'] in inv.columns else set()),
                key=str.upper)

            # Category
            cats = sorted(
                clean(sales[SALES['category']]) |
                (clean(inv[INV['category']]) if INV['category'] in inv.columns else set()),
                key=str.upper)

            self.brand_panel.populate(brands)
            self.cat_panel.populate(cats)

            self.log(f'✅  {len(brands)} thương hiệu  ·  {len(cats)} danh mục  →  sẵn sàng lọc')

        except Exception as e:
            self.log(f'❌  {e}')
            messagebox.showerror('Lỗi đọc file', str(e))
        finally:
            self.btn_read.config(state='normal', text='  Đọc Files & Tải Bộ Lọc  →')

    # ── Run ───────────────────────────────────────────────────────────────────

    def run(self):
        s = self.v_sales.get()
        i = self.v_inv.get()
        m = int(self.v_months.get())

        if not s: messagebox.showerror('Lỗi', 'Vui lòng chọn file bán hàng!'); return
        if not i: messagebox.showerror('Lỗi', 'Vui lòng chọn file tồn kho!'); return

        sel_brands = self.brand_panel.selected()   # [] = tất cả
        sel_cats   = self.cat_panel.selected()     # [] = tất cả

        brand_desc = ', '.join(sel_brands) if sel_brands else 'Tất cả'
        cat_desc   = ', '.join(sel_cats)   if sel_cats   else 'Tất cả'

        try:
            self.btn_run.config(state='disabled', text='⏳   Đang xử lý…')
            self.log(f'─── Xử lý: {m} tháng · Brand: {brand_desc} · Danh mục: {cat_desc} ───')
            df  = self._process(s, i, m, sel_brands, sel_cats)
            out = self._export(df, m, sel_brands, sel_cats)
            self.log(f'✅  Hoàn thành → {out}')
            messagebox.showinfo('Thành công', f'Đã tạo file gợi ý đặt hàng!\n\n📄 {out}')
            if sys.platform == 'win32':
                os.startfile(os.path.dirname(out))
        except Exception as e:
            self.log(f'❌  {e}')
            messagebox.showerror('Lỗi xử lý', str(e))
        finally:
            self.btn_run.config(state='normal', text='       🚀   Tạo Gợi Ý Đặt Hàng       ')

    # ── Core logic ────────────────────────────────────────────────────────────

    def _process(self, sales_path, inv_path, months, sel_brands, sel_cats) -> pd.DataFrame:

        # 1. Đọc & làm sạch bán hàng
        self.log('📊  Đọc file bán hàng…')
        sales = pd.read_excel(sales_path)
        sales.columns = [str(c).strip() for c in sales.columns]
        self._check_cols(sales, list(SALES.values()), 'file bán hàng')

        sales[SALES['date']] = pd.to_datetime(sales[SALES['date']], errors='coerce')
        sales = sales.dropna(subset=[SALES['date']])
        sales['_month'] = sales[SALES['date']].dt.to_period('M')
        sales[SALES['qty']] = pd.to_numeric(sales[SALES['qty']], errors='coerce').fillna(0)

        before = len(sales)
        sales = sales[
            sales[SALES['sku']].notna() &
            (sales[SALES['sku']].astype(str).str.strip() != '') &
            (~sales[SALES['sku']].astype(str).str.upper().str.contains('COUPON|DISCOUNT', na=False)) &
            (sales[SALES['category']].astype(str).str.upper() != 'SERVICE')
        ]
        if len(sales) < before:
            self.log(f'   ⚠️  Lọc {before - len(sales)} dòng không hợp lệ')

        # 2. Áp dụng bộ lọc Brand + Category (trống = lấy tất cả)
        if sel_brands:
            sales = sales[sales[SALES['brand']].astype(str).str.strip().isin(sel_brands)]
        if sel_cats:
            sales = sales[sales[SALES['category']].astype(str).str.strip().isin(sel_cats)]

        if sales.empty:
            raise ValueError('Không có dữ liệu sau khi lọc. Hãy thử chọn lại Brand / Danh mục.')

        self.log(f'   {len(sales):,} giao dịch  ·  {sales["_month"].nunique()} tháng  ·  {sales[SALES["sku"]].nunique():,} SKU')

        # 3. Thông tin sản phẩm
        prod_info = (
            sales[[SALES['sku'], SALES['name'], SALES['brand'], SALES['category']]]
            .drop_duplicates(subset=[SALES['sku']], keep='first').copy()
        )
        prod_info.columns = ['SKU', 'Tên Sản Phẩm', 'Brand', 'Category']
        prod_info['Tên Sản Phẩm'] = prod_info['Tên Sản Phẩm'].apply(strip_sku_prefix)
        prod_info['SKU'] = prod_info['SKU'].astype(str).str.strip()

        # 4. Pivot theo tháng (max 12 tháng gần nhất)
        monthly = (sales.groupby([SALES['sku'], '_month'])[SALES['qty']]
                   .sum().reset_index())
        monthly.columns = ['SKU', 'Month', 'Qty']
        monthly['SKU'] = monthly['SKU'].astype(str).str.strip()

        all_months = sorted(monthly['Month'].unique())
        use_months = all_months[-12:] if len(all_months) > 12 else all_months
        monthly    = monthly[monthly['Month'].isin(use_months)]

        pivot = monthly.pivot_table(index='SKU', columns='Month',
                                    values='Qty', fill_value=0).reset_index()
        mcmap = {p: f'Th.{p.month}/{p.year}' for p in use_months}
        pivot.rename(columns=mcmap, inplace=True)
        mlabels = [mcmap[p] for p in use_months]
        self.log(f'   Tháng: {mlabels[0]}  →  {mlabels[-1]}')

        # 5. Ghép + tính toán
        result = prod_info.merge(pivot, on='SKU', how='left')
        for ml in mlabels:
            if ml not in result.columns:
                result[ml] = 0

        n  = len(mlabels)
        md = result[mlabels].fillna(0)
        result['Tổng Năm']         = md.sum(axis=1)
        result['TB Quý']           = (result['Tổng Năm'] / 4).round(0)
        result['TB Tháng (Năm)']   = (result['Tổng Năm'] / n).round(0)
        last6 = mlabels[-6:] if n >= 6 else mlabels
        result['Tổng 6T Gần Nhất'] = result[last6].sum(axis=1)
        result['TB Tháng (6T GN)'] = (result['Tổng 6T Gần Nhất'] / len(last6)).round(0)

        # 6. Tồn kho
        self.log('📦  Đọc file tồn kho…')
        inv = pd.read_excel(inv_path)
        inv.columns = [str(c).strip() for c in inv.columns]
        self._check_cols(inv, [INV['sku'], INV['qty']], 'file tồn kho')
        inv[INV['qty']] = pd.to_numeric(inv[INV['qty']], errors='coerce').fillna(0)

        inv_f = inv.copy()
        if sel_brands and INV['brand'] in inv_f.columns:
            inv_f = inv_f[inv_f[INV['brand']].astype(str).str.strip().isin(sel_brands)]
        if sel_cats and INV['category'] in inv_f.columns:
            inv_f = inv_f[inv_f[INV['category']].astype(str).str.strip().isin(sel_cats)]

        inv_agg = (inv_f.groupby(INV['sku'])[INV['qty']].sum().reset_index())
        inv_agg.columns = ['SKU', 'Tồn Kho Hiện Tại']
        inv_agg['SKU']  = inv_agg['SKU'].astype(str).str.strip()
        self.log(f'   {inv_agg["SKU"].nunique():,} SKU  ·  Tổng tồn: {inv_agg["Tồn Kho Hiện Tại"].sum():,.0f}')

        result = result.merge(inv_agg, on='SKU', how='left')
        result['Tồn Kho Hiện Tại'] = result['Tồn Kho Hiện Tại'].fillna(0)

        # 7. Gợi ý đặt hàng
        needed = (result['TB Tháng (6T GN)'] * months).round(0)
        scol   = f'Gợi Ý Đặt Hàng ({months} Tháng)'
        result[scol] = (needed - result['Tồn Kho Hiện Tại']).clip(lower=0).round(0)

        result = result.fillna(0)
        self.log(f'✅  {len(result):,} SKU')
        return result

    def _check_cols(self, df, required, label):
        miss = [c for c in required if c not in df.columns]
        if miss:
            raise ValueError(f'Thiếu cột trong {label}: {miss}\nCột hiện có: {list(df.columns)}')

    # ── Export Excel ──────────────────────────────────────────────────────────

    def _export(self, df: pd.DataFrame, months: int,
                sel_brands: list, sel_cats: list) -> str:
        self.log('📝  Xuất Excel…')

        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        ts      = datetime.now().strftime('%Y%m%d_%H%M%S')

        tag_parts = []
        if sel_brands:
            tag_parts.append('_'.join(sel_brands[:2]) + (f'+{len(sel_brands)-2}' if len(sel_brands) > 2 else ''))
        if sel_cats:
            tag_parts.append('_'.join(sel_cats[:2])   + (f'+{len(sel_cats)-2}'   if len(sel_cats)   > 2 else ''))
        tag = ('_' + '_'.join(tag_parts)) if tag_parts else '_TatCa'

        out = os.path.join(desktop, f'GoiYDatHang_{months}T{tag}_{ts}.xlsx')

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f'Gợi Ý {months}T'

        cols       = list(df.columns)
        scol       = f'Gợi Ý Đặt Hàng ({months} Tháng)'
        fixed      = ['SKU', 'Tên Sản Phẩm', 'Brand', 'Category']
        month_cols = [c for c in cols if re.match(r'^Th\.\d{1,2}/\d{4}$', c)]
        calc_cols  = ['Tổng Năm', 'TB Quý', 'TB Tháng (Năm)', 'Tổng 6T Gần Nhất', 'TB Tháng (6T GN)']

        HDR  = {'fixed':('1A56DB','FFF'), 'month':('2563EB','FFF'),
                'calc': ('7C3AED','FFF'), 'inv':  ('D97706','FFF'), 'suggest':('059669','FFF')}
        FILL = {'month':'EFF6FF', 'calc':'F5F3FF', 'inv':'FFFBEB', 'suggest':'ECFDF5'}

        def grp(c):
            if c in fixed:           return 'fixed'
            if c in month_cols:      return 'month'
            if c in calc_cols:       return 'calc'
            if c == 'Tồn Kho Hiện Tại': return 'inv'
            if c == scol:            return 'suggest'
            return 'fixed'

        thin = Border(*(Side(style='thin', color='CBD5E0'),)*4)
        thin = Border(left=Side(style='thin',color='CBD5E0'),right=Side(style='thin',color='CBD5E0'),
                      top=Side(style='thin',color='CBD5E0'), bottom=Side(style='thin',color='CBD5E0'))
        ctr = Alignment(horizontal='center', vertical='center', wrap_text=True)
        lft = Alignment(horizontal='left',   vertical='center')
        rgt = Alignment(horizontal='right',  vertical='center')

        for ci, col in enumerate(cols, 1):
            bg, fg = HDR[grp(col)]
            cell = ws.cell(row=1, column=ci, value=col)
            cell.font = Font(name='Arial', bold=True, color=fg, size=10)
            cell.fill = PatternFill('solid', fgColor=bg)
            cell.alignment = ctr; cell.border = thin
        ws.row_dimensions[1].height = 40

        for ri, (_, row) in enumerate(df.iterrows(), start=2):
            for ci, col in enumerate(cols, 1):
                val = row[col]
                if pd.isna(val): val = 0
                elif isinstance(val, float) and val.is_integer(): val = int(val)
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.border = thin
                g = grp(col)
                if g == 'month':
                    cell.fill = PatternFill('solid', fgColor=FILL['month'])
                    cell.font = Font(name='Arial', size=10); cell.alignment = ctr
                elif g == 'calc':
                    cell.fill = PatternFill('solid', fgColor=FILL['calc'])
                    cell.font = Font(name='Arial', bold=True, size=10); cell.alignment = rgt
                elif g == 'inv':
                    cell.fill = PatternFill('solid', fgColor=FILL['inv'])
                    cell.font = Font(name='Arial', bold=True, size=10); cell.alignment = ctr
                elif g == 'suggest':
                    cell.fill = PatternFill('solid', fgColor=FILL['suggest'])
                    cell.font = Font(name='Arial', bold=True, color='065F46', size=10)
                    cell.alignment = ctr
                elif col == 'SKU':
                    cell.font = Font(name='Arial', size=10); cell.alignment = ctr
                else:
                    cell.font = Font(name='Arial', size=10); cell.alignment = lft
            ws.row_dimensions[ri].height = 18

        W = {'SKU':18,'Tên Sản Phẩm':40,'Brand':16,'Category':20,
             'Tổng Năm':13,'TB Quý':11,'TB Tháng (Năm)':15,
             'Tổng 6T Gần Nhất':17,'TB Tháng (6T GN)':17,'Tồn Kho Hiện Tại':17,scol:24}
        for ci, col in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(ci)].width = W.get(col, 11 if col in month_cols else 14)

        ws.freeze_panes = 'C2'
        ws.auto_filter.ref = f'A1:{get_column_letter(len(cols))}1'
        wb.save(out)
        self.log(f'   → {os.path.basename(out)}')
        return out


# ─── Entry ───────────────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    App(root)
    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    x = (root.winfo_screenwidth()  - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry(f'+{x}+{y}')
    root.mainloop()


if __name__ == '__main__':
    main()
