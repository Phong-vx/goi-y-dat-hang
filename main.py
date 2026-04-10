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
import os, re, platform
from datetime import datetime
import sys
from PIL import Image, ImageTk

# Đường dẫn logo — dùng sys._MEIPASS khi chạy từ PyInstaller exe
if getattr(sys, 'frozen', False):
    _DIR = sys._MEIPASS
else:
    _DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(_DIR, 'File_template', 'Bluecircle.png')

# ─── Font detection (SF Pro trên macOS, Helvetica Neue fallback) ──────────────
IS_MAC = platform.system() == 'Darwin'
SANS   = '.AppleSystemUIFont' if IS_MAC else 'Helvetica Neue'
MONO   = 'SF Mono'            if IS_MAC else 'Menlo'

# ─── Apple-style palette ─────────────────────────────────────────────────────
C = {
    'bg':         '#F2F2F7',   # iOS system background
    'card':       '#FFFFFF',
    'border':     '#C6C6C8',
    'sep':        '#E5E5EA',
    'primary':    '#007AFF',   # iOS blue
    'primary_dk': '#0062CC',
    'text':       '#000000',
    'text2':      '#8E8E93',
    'green':      '#34C759',
    'green_dk':   '#248A3D',
    'red_light':  '#FFF0F0',
    'red_text':   '#FF3B30',
    'blue_light': '#EAF3FF',
    'input_bg':   '#F2F2F7',
    'check_sel':  '#EAF3FF',
    'console_bg': '#1C1C1E',
    'console_fg': '#32D74B',
    'tag_bg':     '#EAF3FF',
    'tag_fg':     '#007AFF',
    'hdr_top':    '#1C1C1E',
}

SALES = {
    'sku':       'SKU',
    'name':      'Product Item',
    'brand':     'Brand',
    'category':  'Category',
    'date':      'Date',
    'qty':       'Quantity',
    'sale_team': 'Sale Team',
    'revenue':   'Revenue',
    'model':     'Model',
}
INV = {
    'sku':          'Sản phẩm/Mã nội bộ',
    'qty':          'Số lượng',
    'qty_reserved': 'Số lượng bảo lưu',
    'brand':        'Sản phẩm/Brand/Display Name',
    'category':     'Sản phẩm/Nganh Hang/Name',
}


def strip_sku_prefix(t: str) -> str:
    return re.sub(r'^\[.*?\]\s*', '', str(t)).strip()


# ─── Reusable widgets ─────────────────────────────────────────────────────────

def make_card(parent, title, subtitle=''):
    """Apple-style card: white frame with hairline border + separator."""
    outer = tk.Frame(parent, bg=C['border'], padx=1, pady=1)
    inner = tk.Frame(outer, bg=C['card'])
    inner.pack(fill=tk.BOTH, expand=True)

    hdr = tk.Frame(inner, bg=C['card'], padx=20, pady=12)
    hdr.pack(fill=tk.X)
    tk.Label(hdr, text=title, font=(SANS, 13, 'bold'),
             bg=C['card'], fg=C['text']).pack(side=tk.LEFT)
    if subtitle:
        tk.Label(hdr, text=subtitle, font=(SANS, 11),
                 bg=C['card'], fg=C['text2']).pack(side=tk.LEFT, padx=(8, 0))

    tk.Frame(inner, bg=C['sep'], height=1).pack(fill=tk.X)

    body = tk.Frame(inner, bg=C['card'], padx=20, pady=14)
    body.pack(fill=tk.BOTH, expand=True)

    return outer, body


def make_btn(parent, text, command, style='primary', small=False):
    """Flat Apple-style button."""
    if style == 'primary':
        bg, fg, abg = C['primary'],    'white', C['primary_dk']
    elif style == 'green':
        bg, fg, abg = C['green'],      'white', C['green_dk']
    elif style == 'ghost':
        bg, fg, abg = C['blue_light'], C['primary'], '#D0E8FF'
    else:
        bg, fg, abg = C['sep'],        C['text'], '#D0D0D5'

    font = (SANS, 10 if small else 13, 'bold')
    px, py = (12, 5) if small else (22, 10)
    return tk.Button(parent, text=text, command=command,
                     font=font, bg=bg, fg=fg, bd=0, relief='flat',
                     cursor='hand2', padx=px, pady=py,
                     activebackground=abg, activeforeground=fg)


# ─── Loading popup với spinner ───────────────────────────────────────────────

class LoadingPopup:
    """
    Popup nhỏ giữa màn hình, có arc xoay tròn và message tuỳ chỉnh.
    Dùng:  popup = LoadingPopup(root, 'message')  →  popup.close()
    """
    _SIZE   = 56   # đường kính vòng spinner
    _THICK  = 5
    _GAP    = 270  # phần trống của arc (°)
    _SPEED  = 18   # ms giữa mỗi frame

    def __init__(self, parent: tk.Tk, message: str):
        self._running = False
        self._angle   = 0

        # ── Cửa sổ không có title bar ─────────────────────────────────────
        self.top = tk.Toplevel(parent)
        self.top.overrideredirect(True)
        self.top.configure(bg=C['card'])
        self.top.attributes('-topmost', True)
        self.top.resizable(False, False)

        # viền mỏng quanh popup
        border = tk.Frame(self.top, bg=C['border'], padx=1, pady=1)
        border.pack(fill=tk.BOTH, expand=True)
        inner  = tk.Frame(border, bg=C['card'], padx=36, pady=28)
        inner.pack(fill=tk.BOTH, expand=True)

        # spinner canvas
        s = self._SIZE
        self._cv = tk.Canvas(inner, width=s, height=s,
                             bg=C['card'], highlightthickness=0)
        self._cv.pack()

        # vòng nền mờ
        m = self._THICK // 2 + 1
        self._cv.create_oval(m, m, s - m, s - m,
                             outline=C['sep'], width=self._THICK)
        # arc xoay
        self._arc_id = self._cv.create_arc(
            m, m, s - m, s - m,
            start=0, extent=360 - self._GAP,
            outline=C['primary'], width=self._THICK, style='arc'
        )

        # dòng message (hỗ trợ xuống dòng)
        tk.Label(inner, text=message,
                 font=(SANS, 12), bg=C['card'], fg=C['text'],
                 justify='center', wraplength=280,
                 pady=(14)).pack(pady=(14, 0))

        # căn giữa so với parent
        self.top.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        ww = self.top.winfo_reqwidth()
        wh = self.top.winfo_reqheight()
        x  = px + (pw - ww) // 2
        y  = py + (ph - wh) // 2
        self.top.geometry(f'+{x}+{y}')

        # chặn tương tác với cửa sổ chính
        self.top.grab_set()

        # bắt đầu animate
        self._running = True
        self._animate()

    def _animate(self):
        if not self._running:
            return
        self._angle = (self._angle - 8) % 360   # quay ngược chiều kim đồng hồ
        self._cv.itemconfigure(self._arc_id, start=self._angle)
        self._cv.after(self._SPEED, self._animate)

    def close(self):
        self._running = False
        try:
            self.top.grab_release()
            self.top.destroy()
        except tk.TclError:
            pass


# ─── Scrollable multi-select panel ───────────────────────────────────────────

class FilterPanel(tk.Frame):
    """
    Scrollable checkbox panel với ô tìm kiếm.
    Không chọn gì = không lọc (lấy tất cả).
    """
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=C['card'], **kw)
        self._vars: dict[str, tk.BooleanVar] = {}
        self._all_items: list = []
        self._widgets: list = []
        self._placeholder_active = True
        self._build()

    # ── scroll helper ─────────────────────────────────────────────────────────
    def _on_scroll(self, event):
        """Route MouseWheel to this panel's canvas, stop propagation."""
        if event.delta:
            # Windows / macOS: delta is ±120 per notch (or smaller for trackpad)
            units = int(-1 * (event.delta / 120)) or (-1 if event.delta > 0 else 1)
        elif event.num == 4:
            units = -1
        else:
            units = 1
        self._canvas.yview_scroll(units, 'units')
        return 'break'   # ← stops the event from reaching the outer canvas

    def _build(self):
        # ── Ô tìm kiếm ───────────────────────────────────────────────────────
        search_wrap = tk.Frame(self, bg=C['border'], padx=1, pady=1)
        search_wrap.pack(fill=tk.X, pady=(0, 8))
        search_inner = tk.Frame(search_wrap, bg=C['input_bg'])
        search_inner.pack(fill=tk.X)
        tk.Label(search_inner, text='⌕', font=(SANS, 14),
                 bg=C['input_bg'], fg=C['text2']).pack(side=tk.LEFT, padx=(10, 4))
        self._search_entry = tk.Entry(
            search_inner,
            font=(SANS, 12), bg=C['input_bg'], fg=C['text2'],
            relief='flat', bd=0, insertbackground=C['primary'])
        self._search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=7, padx=(0, 10))
        self._search_entry.insert(0, 'Tìm kiếm...')
        self._search_entry.bind('<FocusIn>',   self._on_search_focus_in)
        self._search_entry.bind('<FocusOut>',  self._on_search_focus_out)
        self._search_entry.bind('<KeyRelease>', lambda _e: self._apply_search())

        # ── Toolbar ───────────────────────────────────────────────────────────
        tb = tk.Frame(self, bg=C['card'])
        tb.pack(fill=tk.X, pady=(0, 8))
        make_btn(tb, '✓ Tất Cả',  lambda: self._set_all(True),  style='ghost',   small=True).pack(side=tk.LEFT, padx=(0, 6))
        make_btn(tb, '✕ Bỏ Chọn', lambda: self._set_all(False), style='neutral', small=True).pack(side=tk.LEFT)
        self.lbl_count = tk.Label(tb, text='—  chưa tải', font=(SANS, 10),
                                   bg=C['card'], fg=C['text2'])
        self.lbl_count.pack(side=tk.RIGHT)

        # ── Scroll area ───────────────────────────────────────────────────────
        wrap = tk.Frame(self, bg=C['border'], padx=1, pady=1)
        wrap.pack(fill=tk.BOTH, expand=True)
        inner_bg = tk.Frame(wrap, bg=C['card'])
        inner_bg.pack(fill=tk.BOTH, expand=True)

        self._canvas = tk.Canvas(inner_bg, bg=C['card'], height=160,
                                  highlightthickness=0, bd=0)
        vsb = tk.Scrollbar(inner_bg, orient='vertical', command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._inner = tk.Frame(self._canvas, bg=C['card'])
        self._win   = self._canvas.create_window((0, 0), window=self._inner, anchor='nw')

        self._inner.bind('<Configure>', lambda e: self._canvas.configure(
            scrollregion=self._canvas.bbox('all')))
        self._canvas.bind('<Configure>', lambda e: self._canvas.itemconfig(
            self._win, width=e.width))

        # Scroll bindings – both canvas AND inner frame, return 'break' to stop bubbling
        self._canvas.bind('<MouseWheel>', self._on_scroll)
        self._inner.bind('<MouseWheel>',  self._on_scroll)
        self._canvas.bind('<Button-4>',   self._on_scroll)   # Linux
        self._canvas.bind('<Button-5>',   self._on_scroll)

        tk.Label(self._inner, text='Nhấn  "Đọc Files"  để tải danh sách',
                 font=(SANS, 10, 'italic'),
                 bg=C['card'], fg=C['text2'], pady=28).pack()

    # ── Search ────────────────────────────────────────────────────────────────

    def _on_search_focus_in(self, e):
        if self._placeholder_active:
            self._search_entry.delete(0, tk.END)
            self._search_entry.config(fg=C['text'])
            self._placeholder_active = False

    def _on_search_focus_out(self, e):
        if not self._search_entry.get().strip():
            self._search_entry.delete(0, tk.END)
            self._search_entry.insert(0, 'Tìm kiếm...')
            self._search_entry.config(fg=C['text2'])
            self._placeholder_active = True

    def _apply_search(self):
        if self._placeholder_active:
            self._render(self._all_items)
            return
        keyword = self._search_entry.get().lower().strip()
        filtered = [i for i in self._all_items if keyword in i.lower()] if keyword else self._all_items
        self._render(filtered)

    def _render(self, items: list):
        """Vẽ lại checkbox list theo danh sách items (giữ nguyên trạng thái tick)."""
        for w in self._inner.winfo_children():
            w.destroy()
        self._widgets.clear()

        if not items:
            tk.Label(self._inner, text='Không tìm thấy kết quả',
                     font=(SANS, 10, 'italic'),
                     bg=C['card'], fg=C['text2'], pady=20).pack()
            return

        cols = 2
        for idx, item in enumerate(items):
            if item not in self._vars:
                self._vars[item] = tk.BooleanVar(value=False)
            cb = tk.Checkbutton(
                self._inner, text=item, variable=self._vars[item],
                font=(SANS, 11), bg=C['card'], fg=C['text'],
                selectcolor=C['check_sel'], activebackground=C['card'],
                anchor='w', command=self._update_count,
            )
            r, c = divmod(idx, cols)
            cb.grid(row=r, column=c, sticky='w', padx=10, pady=3)
            # Scroll trên checkbox cũng cuộn panel, không bắn lên canvas ngoài
            cb.bind('<MouseWheel>', self._on_scroll)
            cb.bind('<Button-4>',   self._on_scroll)
            cb.bind('<Button-5>',   self._on_scroll)
            self._widgets.append(cb)

        for c in range(cols):
            self._inner.columnconfigure(c, weight=1)

    def populate(self, items: list):
        self._all_items = items
        self._vars.clear()
        self._widgets.clear()
        # Reset ô tìm kiếm về placeholder
        self._search_entry.delete(0, tk.END)
        self._search_entry.insert(0, 'Tìm kiếm...')
        self._search_entry.config(fg=C['text2'])
        self._placeholder_active = True

        self._render(items)
        self._update_count()

    def _set_all(self, val: bool):
        # Chỉ áp dụng cho items đang hiển thị (sau filter tìm kiếm)
        if self._placeholder_active:
            visible = self._all_items
        else:
            keyword = self._search_entry.get().lower().strip()
            visible = [i for i in self._all_items if keyword in i.lower()] if keyword else self._all_items
        for item in visible:
            if item in self._vars:
                self._vars[item].set(val)
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

    def _scroll_main(self, event):
        """Scroll handler cho outer body canvas."""
        if event.delta:
            units = int(-1 * (event.delta / 120)) or (-1 if event.delta > 0 else 1)
        elif event.num == 4:
            units = -1
        else:
            units = 1
        self._body_canvas.yview_scroll(units, 'units')

    def _build_ui(self):
        # ── Header ───────────────────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg=C['hdr_top'])
        hdr.pack(fill=tk.X)

        hdr_inner = tk.Frame(hdr, bg=C['hdr_top'], padx=32, pady=22)
        hdr_inner.pack(fill=tk.X)

        # Logo bên phải
        try:
            pil_img = Image.open(LOGO_PATH).convert('RGBA')
            logo_h  = 52
            logo_w  = int(pil_img.width * logo_h / pil_img.height)
            pil_img = pil_img.resize((logo_w, logo_h), Image.LANCZOS)
            # Ghép lên nền tối để alpha không bị vỡ
            bg_img  = Image.new('RGBA', (logo_w, logo_h), C['hdr_top'])
            bg_img.paste(pil_img, mask=pil_img.split()[3])
            self._logo_img = ImageTk.PhotoImage(bg_img.convert('RGB'))
            tk.Label(hdr_inner, image=self._logo_img,
                     bg=C['hdr_top']).pack(side=tk.RIGHT, anchor='center')
        except Exception:
            pass   # Nếu không tìm thấy ảnh thì bỏ qua

        title_row = tk.Frame(hdr_inner, bg=C['hdr_top'])
        title_row.pack(anchor='w')
        tk.Label(title_row, text='Gợi Ý Đặt Hàng',
                 font=(SANS, 24, 'bold'),
                 bg=C['hdr_top'], fg='white').pack(side=tk.LEFT)
        tk.Label(hdr_inner,
                 text='Phân tích dữ liệu bán hàng & tồn kho  ·  Xuất file Excel gợi ý đặt hàng',
                 font=(SANS, 12), bg=C['hdr_top'], fg='#98989F').pack(anchor='w', pady=(4, 0))

        # ── Scrollable body ───────────────────────────────────────────────────
        outer  = tk.Frame(self.root, bg=C['bg'])
        outer.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(outer, bg=C['bg'], highlightthickness=0)
        self._body_canvas = canvas
        vsb    = tk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.body = tk.Frame(canvas, bg=C['bg'], padx=30, pady=22)
        win = canvas.create_window((0, 0), window=self.body, anchor='nw')

        self.body.bind('<Configure>', lambda e: canvas.configure(
            scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(win, width=e.width))

        # Chỉ bind scroll vào body frame (KHÔNG dùng bind_all để tránh conflict với FilterPanel)
        canvas.bind('<MouseWheel>', self._scroll_main)
        self.body.bind('<MouseWheel>', self._scroll_main)
        canvas.bind('<Button-4>', self._scroll_main)   # Linux
        canvas.bind('<Button-5>', self._scroll_main)

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
        btn_row.pack(fill=tk.X, pady=(12, 0))
        self.btn_read = make_btn(btn_row, '  Đọc Files & Tải Bộ Lọc  →',
                                  self._read_files, style='green')
        self.btn_read.pack(side=tk.RIGHT)
        card.pack(fill=tk.X, pady=(0, 14))

        # ── 2. Bộ lọc Brand + Category (2 cột) ───────────────────────────────
        filter_row = tk.Frame(b, bg=C['bg'])
        filter_row.pack(fill=tk.X, pady=(0, 14))
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
                 font=(SANS, 13), bg=C['card'], fg=C['text']).pack(side=tk.LEFT)

        rb_frame = tk.Frame(row, bg=C['card'])
        rb_frame.pack(side=tk.LEFT, padx=18)
        for val, lbl in [('3','3 tháng'), ('6','6 tháng'), ('9','9 tháng'),
                         ('12','12 tháng'), ('18','18 tháng'), ('24','24 tháng')]:
            tk.Radiobutton(rb_frame, text=lbl, variable=self.v_months, value=val,
                           font=(SANS, 13), bg=C['card'], fg=C['text'],
                           selectcolor=C['check_sel'], activebackground=C['card']
                           ).pack(side=tk.LEFT, padx=12)
        sc.pack(fill=tk.X, pady=(0, 14))

        # ── CTA button ────────────────────────────────────────────────────────
        self.btn_run = make_btn(b, '  🚀   Tạo Gợi Ý Đặt Hàng  ',
                                 self.run, style='primary')
        self.btn_run.config(font=(SANS, 15, 'bold'), pady=15)
        self.btn_run.pack(fill=tk.X, pady=(0, 14))

        # ── Log ───────────────────────────────────────────────────────────────
        lc, lbody = make_card(b, 'Nhật Ký')
        lbody.config(padx=0, pady=0)
        self.log_box = tk.Text(lbody, height=8, font=(MONO, 10),
                                bg=C['console_bg'], fg=C['console_fg'],
                                padx=16, pady=10, wrap=tk.WORD, bd=0,
                                insertbackground=C['console_fg'])
        self.log_box.pack(fill=tk.BOTH, expand=True)
        # Scroll trong log box cũng cần không bắn ra ngoài
        self.log_box.bind('<MouseWheel>', lambda e: 'break')
        lc.pack(fill=tk.X)

    # ── Widgets helpers ───────────────────────────────────────────────────────

    def _file_row(self, parent, label, var):
        f = tk.Frame(parent, bg=C['card'])
        f.pack(fill=tk.X)

        tk.Label(f, text=label, font=(SANS, 13),
                 bg=C['card'], fg=C['text'], width=16, anchor='w').pack(side=tk.LEFT)

        entry = tk.Entry(f, textvariable=var, font=(SANS, 11),
                         bg=C['input_bg'], fg=C['text2'], relief='flat',
                         bd=0, state='readonly', readonlybackground=C['input_bg'])
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 10), ipady=6)

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

        popup = LoadingPopup(
            self.root,
            'Sếp đợi chút\nem đang tính toán,\nem biết sếp tính lóng như kem 🍦'
        )
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
            popup.close()
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

        popup = LoadingPopup(
            self.root,
            'Sếp đợi em chút nha,\nem đang xử lý đây 🚀'
        )
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
            popup.close()
            self.btn_run.config(state='normal', text='  🚀   Tạo Gợi Ý Đặt Hàng  ')

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
        has_model   = SALES['model'] in sales.columns
        extra_cols  = [SALES['model']] if has_model else []
        prod_info   = (
            sales[[SALES['sku'], SALES['name'], SALES['brand'], SALES['category']] + extra_cols]
            .drop_duplicates(subset=[SALES['sku']], keep='first').copy()
        )
        prod_info.columns = ['SKU', 'Tên Sản Phẩm', 'Brand', 'Category'] + (['Model'] if has_model else [])
        prod_info['Tên Sản Phẩm'] = prod_info['Tên Sản Phẩm'].apply(strip_sku_prefix)
        prod_info['SKU'] = prod_info['SKU'].astype(str).str.strip()

        # 4. Pivot theo tháng (max 12 tháng gần nhất)
        monthly = (sales.groupby([SALES['sku'], '_month'])[SALES['qty']]
                   .sum().reset_index())
        monthly.columns = ['SKU', 'Month', 'Qty']
        monthly['SKU'] = monthly['SKU'].astype(str).str.strip()

        all_months = sorted(monthly['Month'].unique())
        use_months = all_months   # lấy tất cả tháng có trong file
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

        md = result[mlabels].fillna(0)

        # Nhóm tháng theo năm → Tổng YYYY + TB YYYY
        from collections import defaultdict
        year_months_map = defaultdict(list)
        for ml in mlabels:
            year = ml.split('/')[1]
            year_months_map[year].append(ml)

        sorted_years  = sorted(year_months_map.keys())
        year_stat_cols = []
        for year in sorted_years:
            ycols = year_months_map[year]
            result[f'Tổng {year}']  = result[ycols].sum(axis=1)
            result[f'TB {year}']    = (result[f'Tổng {year}'] / len(ycols)).round(0)
            year_stat_cols.extend([f'Tổng {year}', f'TB {year}'])

        result['Tổng Toàn TG'] = md.sum(axis=1)

        last6 = mlabels[-6:] if len(mlabels) >= 6 else mlabels
        result['Tổng 6T Gần Nhất'] = result[last6].sum(axis=1)
        result['TB Tháng (6T GN)'] = (result['Tổng 6T Gần Nhất'] / len(last6)).round(0)

        # Doanh thu tổng
        if SALES['revenue'] in sales.columns:
            sales[SALES['revenue']] = pd.to_numeric(sales[SALES['revenue']], errors='coerce').fillna(0)
            rev_agg = sales.groupby(SALES['sku'])[SALES['revenue']].sum().reset_index()
            rev_agg.columns = ['SKU', 'Doanh Thu Tổng']
            rev_agg['SKU'] = rev_agg['SKU'].astype(str).str.strip()
            result = result.merge(rev_agg, on='SKU', how='left')
            result['Doanh Thu Tổng'] = result['Doanh Thu Tổng'].fillna(0)
            has_revenue = True
        else:
            has_revenue = False

        # Tổng bán theo Sale Team
        team_labels = []
        if SALES['sale_team'] in sales.columns:
            team_agg = (
                sales.groupby([SALES['sku'], SALES['sale_team']])[SALES['qty']]
                .sum().reset_index()
            )
            team_agg['SKU'] = team_agg[SALES['sku']].astype(str).str.strip()
            teams = sorted(team_agg[SALES['sale_team']].dropna().unique())

            team_wide = team_agg.pivot_table(
                index='SKU', columns=SALES['sale_team'],
                values=SALES['qty'], fill_value=0
            ).reset_index()
            team_wide.columns.name = None
            team_labels = [c for c in team_wide.columns if c != 'SKU']

            result = result.merge(team_wide, on='SKU', how='left')
            for tl in team_labels:
                result[tl] = result[tl].fillna(0)

            self.log(f'   Sale Team: {", ".join(str(t) for t in teams)}')

        result.attrs['year_stat_cols'] = year_stat_cols
        result.attrs['team_labels']    = team_labels
        result.attrs['has_revenue']    = has_revenue

        # 6. Tồn kho
        self.log('📦  Đọc file tồn kho…')
        inv = pd.read_excel(inv_path)
        inv.columns = [str(c).strip() for c in inv.columns]
        self._check_cols(inv, [INV['sku'], INV['qty']], 'file tồn kho')
        inv[INV['qty']] = pd.to_numeric(inv[INV['qty']], errors='coerce').fillna(0)

        has_reserved = INV['qty_reserved'] in inv.columns
        if has_reserved:
            inv[INV['qty_reserved']] = pd.to_numeric(inv[INV['qty_reserved']], errors='coerce').fillna(0)

        inv_f = inv.copy()
        if sel_brands and INV['brand'] in inv_f.columns:
            inv_f = inv_f[inv_f[INV['brand']].astype(str).str.strip().isin(sel_brands)]

        # Aggregate: Số lượng + Số lượng bảo lưu
        agg_cols = {INV['qty']: 'sum'}
        if has_reserved:
            agg_cols[INV['qty_reserved']] = 'sum'

        inv_agg = inv_f.groupby(INV['sku']).agg(agg_cols).reset_index()
        inv_agg.columns = ['SKU', 'Tồn Kho (Số Lượng)'] + (['Tồn Bảo Lưu'] if has_reserved else [])
        inv_agg['SKU'] = inv_agg['SKU'].astype(str).str.strip()

        if has_reserved:
            inv_agg['Tồn Khả Dụng'] = (inv_agg['Tồn Kho (Số Lượng)'] - inv_agg['Tồn Bảo Lưu']).clip(lower=0)
        else:
            inv_agg['Tồn Khả Dụng'] = inv_agg['Tồn Kho (Số Lượng)']

        self.log(f'   {inv_agg["SKU"].nunique():,} SKU  ·  Tổng tồn: {inv_agg["Tồn Kho (Số Lượng)"].sum():,.0f}  ·  Bảo lưu: {inv_agg["Tồn Bảo Lưu"].sum():,.0f}' if has_reserved else
                 f'   {inv_agg["SKU"].nunique():,} SKU  ·  Tổng tồn: {inv_agg["Tồn Kho (Số Lượng)"].sum():,.0f}')

        result = result.merge(inv_agg, on='SKU', how='left')
        for col in ['Tồn Kho (Số Lượng)', 'Tồn Bảo Lưu', 'Tồn Khả Dụng']:
            if col in result.columns:
                result[col] = result[col].fillna(0)

        # 7. Gợi ý đặt hàng — dựa trên Tồn Khả Dụng
        needed = (result['TB Tháng (6T GN)'] * months).round(0)
        scol   = f'Gợi Ý Đặt Hàng ({months} Tháng)'
        result[scol] = (needed - result['Tồn Khả Dụng']).clip(lower=0).round(0)

        # Sắp xếp lại thứ tự cột
        inv_out = [c for c in ['Tồn Kho (Số Lượng)', 'Tồn Bảo Lưu', 'Tồn Khả Dụng'] if c in result.columns]
        rev_out    = ['Doanh Thu Tổng'] if has_revenue else []
        model_out  = ['Model'] if has_model else []
        ordered = (['SKU', 'Tên Sản Phẩm', 'Brand', 'Category'] + model_out
                   + mlabels
                   + year_stat_cols
                   + ['Tổng Toàn TG']
                   + rev_out
                   + team_labels
                   + ['Tổng 6T Gần Nhất', 'TB Tháng (6T GN)']
                   + inv_out
                   + [scol])
        result = result[[c for c in ordered if c in result.columns]].fillna(0)
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
            tag_parts.append('_'.join(sel_cats[:2]) + (f'+{len(sel_cats)-2}' if len(sel_cats) > 2 else ''))
        tag = ('_' + '_'.join(tag_parts)) if tag_parts else '_TatCa'
        out = os.path.join(desktop, f'GoiYDatHang_{months}T{tag}_{ts}.xlsx')

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f'Gợi Ý {months}T'

        cols       = list(df.columns)
        scol       = f'Gợi Ý Đặt Hàng ({months} Tháng)'
        fixed      = ['SKU', 'Tên Sản Phẩm', 'Brand', 'Category', 'Model']
        month_cols = [c for c in cols if re.match(r'^Th\.\d{1,2}/\d{4}$', c)]
        inv_cols   = ['Tồn Kho (Số Lượng)', 'Tồn Bảo Lưu', 'Tồn Khả Dụng']
        stat6_cols = ['Tổng 6T Gần Nhất', 'TB Tháng (6T GN)']
        # Cột không nên SUBTOTAL (text hoặc trung bình)
        NOT_SUM    = set(fixed) | {c for c in cols if c.startswith('TB ')}

        # Detect year stat columns  (Tổng YYYY / TB YYYY / Tổng Toàn TG)
        year_total_cols = [c for c in cols if re.match(r'^Tổng \d{4}$', c)]
        year_avg_cols   = [c for c in cols if re.match(r'^TB \d{4}$', c)]
        grand_col       = 'Tổng Toàn TG'
        team_cols       = df.attrs.get('team_labels', [])

        # ── Màu header theo nhóm ──────────────────────────────────────────────
        # (bg_header, fg_header, fill_odd, fill_even)
        GRP_STYLE = {
            'fixed':    ('1A56DB', 'FFFFFF', 'FFFFFF',  'F0F4FF'),
            'month':    ('2563EB', 'FFFFFF', 'EFF6FF',  'DBEAFE'),
            'yr_total': ('6D28D9', 'FFFFFF', 'F5F3FF',  'EDE9FE'),
            'yr_avg':   ('7C3AED', 'FFFFFF', 'FAF5FF',  'F3E8FF'),
            'grand':    ('4C1D95', 'FFFFFF', 'EDE9FE',  'DDD6FE'),
            'revenue':  ('B45309', 'FFFFFF', 'FFFBEB',  'FDE68A'),
            'team':     ('0F766E', 'FFFFFF', 'F0FDFA',  'CCFBF1'),
            'stat6':    ('0369A1', 'FFFFFF', 'E0F2FE',  'BAE6FD'),
            'inv':      ('B45309', 'FFFFFF', 'FFFBEB',  'FEF3C7'),
            'suggest':  ('065F46', 'FFFFFF', 'ECFDF5',  'D1FAE5'),
        }

        def grp(c):
            if c in fixed:          return 'fixed'
            if c in month_cols:     return 'month'
            if c in year_total_cols:return 'yr_total'
            if c in year_avg_cols:  return 'yr_avg'
            if c == grand_col:      return 'grand'
            if c == 'Doanh Thu Tổng': return 'revenue'
            if c in team_cols:      return 'team'
            if c in stat6_cols:     return 'stat6'
            if c in inv_cols:       return 'inv'
            if c == scol:           return 'suggest'
            return 'fixed'

        thin = Border(
            left=Side(style='thin',  color='D1D5DB'),
            right=Side(style='thin', color='D1D5DB'),
            top=Side(style='thin',   color='D1D5DB'),
            bottom=Side(style='thin',color='D1D5DB'),
        )
        ctr = Alignment(horizontal='center', vertical='center', wrap_text=True)
        lft = Alignment(horizontal='left',   vertical='center')
        rgt = Alignment(horizontal='right',  vertical='center')

        # ── Row 1: Subtotal ───────────────────────────────────────────────────
        for ci, col in enumerate(cols, 1):
            st        = GRP_STYLE[grp(col)]
            cell      = ws.cell(row=1, column=ci)
            col_ltr   = get_column_letter(ci)
            cell.fill   = PatternFill('solid', fgColor=st[0])
            cell.border = thin

            if col in NOT_SUM:
                cell.value     = 'SUBTOTAL' if ci == 1 else None
                cell.font      = Font(name='Calibri', bold=True, color=st[1], size=10)
                cell.alignment = ctr
            else:
                cell.value     = f'=SUBTOTAL(9,{col_ltr}3:{col_ltr}1048576)'
                cell.font      = Font(name='Calibri', bold=True, color=st[1], size=10)
                cell.alignment = rgt
                if col == 'Doanh Thu Tổng':
                    cell.number_format = '#,##0'

        ws.row_dimensions[1].height = 26

        # ── Row 2: Header ─────────────────────────────────────────────────────
        for ci, col in enumerate(cols, 1):
            st = GRP_STYLE[grp(col)]
            cell = ws.cell(row=2, column=ci, value=col)
            cell.font      = Font(name='Calibri', bold=True, color=st[1], size=10)
            cell.fill      = PatternFill('solid', fgColor=st[0])
            cell.alignment = ctr
            cell.border    = thin
        ws.row_dimensions[2].height = 38

        # ── Data rows với màu xen kẽ (bắt đầu từ row 3) ─────────────────────
        for ri, (_, row) in enumerate(df.iterrows(), start=3):
            is_odd = (ri % 2 == 0)   # hàng chẵn (dữ liệu lẻ) = nền nhạt hơn
            for ci, col in enumerate(cols, 1):
                val = row[col]
                if pd.isna(val):     val = 0
                elif isinstance(val, float) and val.is_integer(): val = int(val)

                cell        = ws.cell(row=ri, column=ci, value=val)
                cell.border = thin
                g  = grp(col)
                st = GRP_STYLE[g]
                fill_color  = st[2] if is_odd else st[3]
                cell.fill   = PatternFill('solid', fgColor=fill_color)

                if g == 'fixed':
                    cell.font      = Font(name='Calibri', size=10)
                    cell.alignment = ctr if col == 'SKU' else lft
                elif g == 'month':
                    cell.font      = Font(name='Calibri', size=10)
                    cell.alignment = ctr
                elif g in ('yr_total', 'grand', 'team'):
                    cell.font      = Font(name='Calibri', bold=True, size=10)
                    cell.alignment = rgt
                elif g == 'revenue':
                    cell.font         = Font(name='Calibri', bold=True, size=10)
                    cell.alignment    = rgt
                    cell.number_format = '#,##0'
                elif g == 'yr_avg':
                    cell.font      = Font(name='Calibri', size=10, italic=True)
                    cell.alignment = rgt
                elif g == 'stat6':
                    cell.font      = Font(name='Calibri', bold=True, size=10)
                    cell.alignment = rgt
                elif g == 'inv':
                    cell.font      = Font(name='Calibri', bold=True, size=10)
                    cell.alignment = ctr
                elif g == 'suggest':
                    cell.font      = Font(name='Calibri', bold=True, size=11)
                    cell.alignment = ctr
            ws.row_dimensions[ri].height = 18

        # ── Độ rộng cột ───────────────────────────────────────────────────────
        W = {
            'SKU': 18, 'Tên Sản Phẩm': 42, 'Brand': 14, 'Category': 18,
            grand_col: 16, 'Tổng 6T Gần Nhất': 16, 'TB Tháng (6T GN)': 16,
            'Tồn Kho (Số Lượng)': 17, 'Tồn Bảo Lưu': 13, 'Tồn Khả Dụng': 14,
            scol: 24,
        }
        for c in year_total_cols + year_avg_cols:
            W[c] = 13
        for c in team_cols:
            W[c] = 16
        W['Doanh Thu Tổng'] = 18

        for ci, col in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(ci)].width = W.get(
                col, 10 if col in month_cols else 14)

        # ── Freeze + filter ───────────────────────────────────────────────────
        ws.freeze_panes = 'C3'   # cố định 2 dòng đầu + cột A, B
        ws.auto_filter.ref = f'A2:{get_column_letter(len(cols))}2'

        # ── Tab màu ───────────────────────────────────────────────────────────
        ws.sheet_properties.tabColor = '0071E3'

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
