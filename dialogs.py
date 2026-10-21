"""Dialog boxes for PyExcel: Format Cells, conditional-formatting rules, sorting, series,
insert/delete cells, paste special, the filter dropdown, insert function, format as table,
remove duplicates, rule management, data validation (with its in-cell list) and notes."""

import tkinter as tk
import tkinter.font as tkfont
from tkinter import colorchooser, ttk

import formula_engine as fe
from function_info import CATEGORIES, FUNCTION_INFO
from ribbon import GROUP_LINE, HOVER_BG, STANDARD_COLORS, TEXT_FG, Popup, color_popup, theme_palette

SYMBOLS = ["₫", "$", "€", "¥", "£", "Không"]
LONG_DATE = 'dddd, "ngày" dd "tháng" mm "năm" yyyy'
DATE_TYPES = ["dd/mm/yyyy", "d/m/yyyy", "dd-mm-yyyy", "yyyy-mm-dd", "dd/mm/yy", "dd/mm", LONG_DATE,
              "mmmm yyyy", "dd/mm/yyyy hh:mm"]
TIME_TYPES = ["hh:mm", "hh:mm:ss", "dd/mm/yyyy hh:mm"]
FRACTION_TYPES = ["# ?/?", "# ??/??", "# ???/???"]
NUM_CATEGORIES = ["Chung", "Số", "Tiền tệ", "Kế toán", "Ngày", "Thời gian", "Phần trăm", "Phân số",
                  "Khoa học", "Văn bản", "Tùy chỉnh"]
BORDER_NAMES = [("thin", "Mảnh"), ("medium", "Vừa"), ("thick", "Dày"), ("double", "Kép"),
                ("dashed", "Gạch"), ("dotted", "Chấm")]
CF_PRESETS = [("Tô đỏ nhạt, chữ đỏ sẫm", {"bg": "#ffc7ce", "fg": "#9c0006"}),
              ("Tô vàng, chữ vàng sẫm", {"bg": "#ffeb9c", "fg": "#9c5700"}),
              ("Tô xanh lá, chữ xanh sẫm", {"bg": "#c6efce", "fg": "#006100"}),
              ("Tô đỏ nhạt", {"bg": "#ffc7ce"}),
              ("Chữ đỏ", {"fg": "#ff0000"}),
              ("Chữ đậm", {"bold": True})]


def _decimals(n):
    return "." + "0" * n if n else ""


def currency_format(symbol, decimals):
    core = "#,##0" + _decimals(decimals)
    return {"₫": f"{core} ₫", "$": f"${core}", "€": f"{core} €", "¥": f"¥{core}", "£": f"£{core}"}.get(symbol, core)


def accounting_format(symbol, decimals):
    core = "#,##0" + _decimals(decimals)
    if symbol in ("₫", "€"):
        return f'{core} {symbol};-{core} {symbol};"-" {symbol}'
    if symbol in ("$", "¥", "£"):
        return f'{symbol} {core};{symbol} -{core};{symbol} "-"'
    return f'{core};-{core};"-"'


def canonical_format(fmt):
    """A number format with literal-text quotes removed, so '#,##0 "₫"' matches '#,##0 ₫'."""
    return (fmt or "").replace('"', "")


def classify_format(fmt):
    """(category, options) describing a number format for the Format Cells dialog."""
    if not fmt:
        return "Chung", {}
    fmt = canonical_format(fmt)
    if fmt == "@":
        return "Văn bản", {}
    if fe.is_date_format(fmt):
        return ("Thời gian" if fmt in TIME_TYPES[:2] else "Ngày"), {"type": fmt}
    if "?/" in fmt:
        return "Phân số", {"type": fmt}
    sections = fe.split_sections(fmt)
    m = fe._NUM_CORE_RE.search(sections[0])
    decimals = len(m.group(0).partition(".")[2]) if m else 0
    if "E+" in sections[0].upper():
        return "Khoa học", {"decimals": decimals}
    symbol = next((s for s in SYMBOLS[:-1] if s in fmt), None)
    if symbol and len(sections) >= 3 and fmt == canonical_format(accounting_format(symbol, decimals)):
        return "Kế toán", {"decimals": decimals, "symbol": symbol}
    if symbol and fmt == currency_format(symbol, decimals):
        return "Tiền tệ", {"decimals": decimals, "symbol": symbol}
    if len(sections) == 1 and sections[0] == "0" + _decimals(decimals) + "%":
        return "Phần trăm", {"decimals": decimals}
    if len(sections) == 1 and sections[0] in ("0" + _decimals(decimals), "#,##0" + _decimals(decimals)):
        return "Số", {"decimals": decimals, "thousands": sections[0].startswith("#")}
    return "Tùy chỉnh", {}


class Dialog(tk.Toplevel):
    """Modal dialog: subclasses fill self.body, call add_buttons() and implement collect()."""

    def __init__(self, master, title):
        super().__init__(master)
        self.withdraw()
        self.title(title)
        self.transient(master.winfo_toplevel())
        self.resizable(False, False)
        self.result = None
        self.body = ttk.Frame(self, padding=12)
        self.body.pack(fill="both", expand=True)
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.bind("<Escape>", lambda e: self.cancel())

    def add_buttons(self, ok_text="OK"):
        bar = ttk.Frame(self, padding=(12, 0, 12, 12))
        bar.pack(fill="x")
        ttk.Button(bar, text="Hủy", command=self.cancel, width=10).pack(side="right")
        ttk.Button(bar, text=ok_text, command=self.ok, width=10).pack(side="right", padx=(0, 6))
        self.bind("<Return>", lambda e: self.ok())
        return bar

    def show(self, focus=None):
        self.update_idletasks()
        top = self.master.winfo_toplevel()
        x = top.winfo_rootx() + (top.winfo_width() - self.winfo_reqwidth()) // 2
        y = top.winfo_rooty() + (top.winfo_height() - self.winfo_reqheight()) // 3
        self.geometry(f"+{max(0, x)}+{max(0, y)}")
        self.deiconify()
        self.grab_set()
        (focus or self).focus_set()
        self.wait_window()
        return self.result

    def collect(self):
        return True

    def ok(self):
        result = self.collect()
        if result is None:
            return
        self.result = result
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()


class ColorButton(tk.Frame):
    """A swatch that opens the Excel colour palette; None means 'automatic/no colour'."""

    def __init__(self, master, ui, color, none_label, on_change=None):
        super().__init__(master, bg="#a0a0a0")
        self.ui, self.color, self.none_label, self.on_change = ui, color, none_label, on_change
        self.label = tk.Label(self, width=14, relief="flat", bd=0, cursor="hand2", font=ui.small_font)
        self.label.pack(padx=1, pady=1)
        self.label.bind("<ButtonRelease-1>", lambda e: self._open())
        self._paint()

    def _paint(self):
        if self.color:
            dark = sum(int(self.color[i:i + 2], 16) for i in (1, 3, 5)) < 380
            self.label.config(bg=self.color, text=self.color.upper(), fg="#ffffff" if dark else "#000000")
        else:
            self.label.config(bg="#ffffff", text=self.none_label, fg="#333333")

    def _open(self):
        color_popup(self.label, self.ui, self._set, self.none_label, self._more)

    def _more(self):
        _, color = colorchooser.askcolor(color=self.color, parent=self, title="Chọn màu")
        if color:
            self._set(color)

    def _set(self, color):
        self.color = color
        self._paint()
        if self.on_change:
            self.on_change(color)


class FormatCellsDialog(Dialog):
    """Excel's Format Cells (Ctrl+1). collect() returns only what the user changed."""

    TABS = ("number", "align", "font", "border", "fill")

    def __init__(self, master, ui, fmt, sample, families, default_family, tab="number", merged=False,
                 multi_rows=False, multi_cols=False):
        super().__init__(master, "Định dạng ô")
        self.ui, self.fmt, self.sample = ui, dict(fmt or {}), sample
        self.families, self.default_family = families, default_family
        self.merged, self.multi_rows, self.multi_cols = merged, multi_rows, multi_cols
        self.nb = ttk.Notebook(self.body)
        self.nb.pack(fill="both", expand=True)
        self._number_tab()
        self._align_tab()
        self._font_tab()
        self._border_tab()
        self._fill_tab()
        self.nb.select(self.TABS.index(tab) if tab in self.TABS else 0)
        self.add_buttons()

    # number
    def _number_tab(self):
        page = ttk.Frame(self.nb, padding=10)
        self.nb.add(page, text="Số")
        ttk.Label(page, text="Thể loại:").grid(row=0, column=0, sticky="w")
        self.cat_list = tk.Listbox(page, height=12, width=16, exportselection=False, activestyle="none")
        for name in NUM_CATEGORIES:
            self.cat_list.insert("end", name)
        self.cat_list.grid(row=1, column=0, sticky="ns", padx=(0, 12))
        right = ttk.Frame(page, width=380)
        right.grid(row=1, column=1, sticky="nsew")
        right.grid_propagate(False)
        page.rowconfigure(1, minsize=250)
        sample_box = ttk.LabelFrame(right, text="Mẫu", padding=6)
        sample_box.pack(fill="x")
        self.sample_label = ttk.Label(sample_box, text="", width=44)
        self.sample_label.pack(anchor="w")
        self.opts = ttk.Frame(right)
        self.opts.pack(fill="both", expand=True, pady=(8, 0))
        cat, opts = classify_format(self.fmt.get("num"))
        self.dec_var = tk.IntVar(value=opts.get("decimals", 2))
        self.thou_var = tk.BooleanVar(value=opts.get("thousands", True))
        self.sym_var = tk.StringVar(value=opts.get("symbol", "₫"))
        self.type_var = tk.StringVar(value=opts.get("type", ""))
        self.custom_var = tk.StringVar(value=self.fmt.get("num") or "General")
        for var in (self.dec_var, self.thou_var, self.sym_var, self.custom_var):
            var.trace_add("write", lambda *a: self._update_sample())
        self.cat_list.bind("<<ListboxSelect>>", lambda e: self._show_category())
        self.cat_list.selection_set(NUM_CATEGORIES.index(cat))
        self._initial_num = self.fmt.get("num")
        self._show_category()

    def _category(self):
        sel = self.cat_list.curselection()
        return NUM_CATEGORIES[sel[0]] if sel else "Chung"

    def _show_category(self):
        for w in self.opts.winfo_children():
            w.destroy()
        cat = self._category()
        f = self.opts
        if cat in ("Số", "Tiền tệ", "Kế toán", "Phần trăm", "Khoa học"):
            row = ttk.Frame(f)
            row.pack(anchor="w", pady=2)
            ttk.Label(row, text="Số chữ số thập phân:").pack(side="left")
            ttk.Spinbox(row, from_=0, to=15, width=5, textvariable=self.dec_var).pack(side="left", padx=6)
        if cat == "Số":
            ttk.Checkbutton(f, text="Dùng dấu phân cách hàng nghìn (,)", variable=self.thou_var).pack(anchor="w", pady=2)
        if cat in ("Tiền tệ", "Kế toán"):
            row = ttk.Frame(f)
            row.pack(anchor="w", pady=2)
            ttk.Label(row, text="Ký hiệu:").pack(side="left")
            ttk.Combobox(row, values=SYMBOLS, textvariable=self.sym_var, state="readonly", width=8).pack(side="left", padx=6)
        if cat in ("Ngày", "Thời gian", "Phân số"):
            types = {"Ngày": DATE_TYPES, "Thời gian": TIME_TYPES, "Phân số": FRACTION_TYPES}[cat]
            ttk.Label(f, text="Kiểu:").pack(anchor="w")
            lb = tk.Listbox(f, height=7, width=44, exportselection=False, activestyle="none")
            serial = self.sample if fe.is_number(self.sample) else fe.parse_literal("21/10/2026")[0]
            for t in types:
                lb.insert("end", fe.format_number(serial if cat != "Phân số" else 1.25, t))
            lb.pack(anchor="w")
            if self.type_var.get() not in types:
                self.type_var.set(types[0])
            lb.selection_set(types.index(self.type_var.get()))
            lb.bind("<<ListboxSelect>>", lambda e: (self.type_var.set(types[lb.curselection()[0]]),
                                                    self._update_sample()) if lb.curselection() else None)
        if cat == "Tùy chỉnh":
            ttk.Label(f, text="Mã định dạng:").pack(anchor="w")
            if self.custom_var.get() == "General" and self.fmt.get("num"):
                self.custom_var.set(self.fmt["num"])
            ttk.Entry(f, textvariable=self.custom_var, width=46).pack(anchor="w", pady=2)
            ttk.Label(f, text='Ví dụ: #,##0.00   0%   dd/mm/yyyy   #,##0 "đ";[âm] -#,##0', foreground="#666").pack(anchor="w")
        notes = {"Chung": "Ô định dạng Chung không có định dạng số cụ thể.",
                 "Văn bản": "Nội dung được giữ nguyên như khi nhập, kể cả số.",
                 "Kế toán": "Số 0 hiển thị là dấu gạch ngang.",
                 "Phần trăm": "Giá trị được nhân với 100 và thêm ký hiệu %."}
        if cat in notes:
            ttk.Label(f, text=notes[cat], foreground="#555", wraplength=360).pack(anchor="w", pady=(8, 0))
        self._update_sample()

    def number_format(self):
        cat = self._category()
        d = max(0, min(15, int(self.dec_var.get() or 0)))
        if cat == "Chung":
            return None
        if cat == "Số":
            return ("#,##0" if self.thou_var.get() else "0") + _decimals(d)
        if cat == "Tiền tệ":
            return currency_format(self.sym_var.get(), d)
        if cat == "Kế toán":
            return accounting_format(self.sym_var.get(), d)
        if cat == "Phần trăm":
            return "0" + _decimals(d) + "%"
        if cat == "Khoa học":
            return "0" + _decimals(d) + "E+00"
        if cat == "Văn bản":
            return "@"
        if cat in ("Ngày", "Thời gian", "Phân số"):
            return self.type_var.get()
        text = self.custom_var.get().strip()
        return None if text.lower() in ("", "general") else text

    def _update_sample(self):
        try:
            fmt = self.number_format()
        except (tk.TclError, ValueError):
            return
        value = self.sample
        if fe.is_number(value) and fmt and fmt != "@":
            text = fe.format_number(value, fmt)
        else:
            text = fe.format_value(value)
        self.sample_label.config(text=text or " ")

    # alignment
    def _align_tab(self):
        page = ttk.Frame(self.nb, padding=10)
        self.nb.add(page, text="Căn chỉnh")
        box = ttk.LabelFrame(page, text="Căn chỉnh văn bản", padding=8)
        box.grid(row=0, column=0, sticky="nsew")
        self.h_names = {"Chung": None, "Trái (thụt lề)": "left", "Giữa": "center", "Phải (thụt lề)": "right"}
        self.v_names = {"Trên": "top", "Giữa": "center", "Dưới": None}
        ttk.Label(box, text="Ngang:").grid(row=0, column=0, sticky="w")
        self.h_var = tk.StringVar(value=next(k for k, v in self.h_names.items() if v == self.fmt.get("align")))
        ttk.Combobox(box, values=list(self.h_names), textvariable=self.h_var, state="readonly", width=18).grid(row=1, column=0, sticky="w")
        ttk.Label(box, text="Thụt lề:").grid(row=0, column=1, sticky="w", padx=(12, 0))
        self.indent_var = tk.IntVar(value=self.fmt.get("indent") or 0)
        ttk.Spinbox(box, from_=0, to=15, width=5, textvariable=self.indent_var).grid(row=1, column=1, sticky="w", padx=(12, 0))
        ttk.Label(box, text="Dọc:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.v_var = tk.StringVar(value=next(k for k, v in self.v_names.items() if v == self.fmt.get("valign")))
        ttk.Combobox(box, values=list(self.v_names), textvariable=self.v_var, state="readonly", width=18).grid(row=3, column=0, sticky="w")
        ctrl = ttk.LabelFrame(page, text="Điều khiển văn bản", padding=8)
        ctrl.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        self.wrap_var = tk.BooleanVar(value=bool(self.fmt.get("wrap")))
        self.merge_var = tk.BooleanVar(value=self.merged)
        ttk.Checkbutton(ctrl, text="Xuống dòng tự động", variable=self.wrap_var).pack(anchor="w")
        ttk.Checkbutton(ctrl, text="Hợp nhất ô", variable=self.merge_var).pack(anchor="w")
        orient = ttk.LabelFrame(page, text="Hướng", padding=8)
        orient.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=(12, 0))
        rot = self.fmt.get("rotate") or 0
        self.vertical_var = tk.BooleanVar(value=rot == 255)
        self.rot_var = tk.IntVar(value=0 if rot == 255 else rot)
        ttk.Checkbutton(orient, text="Văn bản dọc", variable=self.vertical_var,
                        command=self._draw_orientation).pack(anchor="w")
        self.orient_cv = tk.Canvas(orient, width=150, height=120, bg="#ffffff", highlightthickness=1,
                                   highlightbackground="#c8c8c8")
        self.orient_cv.pack(pady=6)
        row = ttk.Frame(orient)
        row.pack()
        ttk.Spinbox(row, from_=-90, to=90, width=5, textvariable=self.rot_var,
                    command=self._draw_orientation).pack(side="left")
        ttk.Label(row, text="độ").pack(side="left", padx=4)
        self.rot_var.trace_add("write", lambda *a: self._draw_orientation())
        self._draw_orientation()

    def _draw_orientation(self):
        cv = self.orient_cv
        cv.delete("all")
        if self.vertical_var.get():
            cv.create_text(75, 60, text="V\nă\nn\nb\nả\nn", justify="center")
            return
        try:
            angle = int(self.rot_var.get())
        except (tk.TclError, ValueError):
            angle = 0
        cv.create_arc(20, 10, 130, 110, start=-90, extent=180, style="arc", outline="#c8c8c8")
        cv.create_text(75, 60, text="Văn bản ——", angle=max(-90, min(90, angle)), anchor="center")

    # font
    def _font_tab(self):
        page = ttk.Frame(self.nb, padding=10)
        self.nb.add(page, text="Phông chữ")
        self.family_var = tk.StringVar(value=self.fmt.get("font") or self.default_family)
        style = ("Đậm nghiêng" if self.fmt.get("bold") and self.fmt.get("italic") else
                 "Đậm" if self.fmt.get("bold") else "Nghiêng" if self.fmt.get("italic") else "Thường")
        self.style_var = tk.StringVar(value=style)
        self.size_var = tk.StringVar(value=str(self.fmt.get("size") or 11))
        ttk.Label(page, text="Phông:").grid(row=0, column=0, sticky="w")
        ttk.Label(page, text="Kiểu phông:").grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(page, text="Cỡ:").grid(row=0, column=2, sticky="w")
        ttk.Entry(page, textvariable=self.family_var, width=28).grid(row=1, column=0, sticky="w")
        ttk.Entry(page, textvariable=self.style_var, width=14, state="readonly").grid(row=1, column=1, sticky="w", padx=8)
        ttk.Entry(page, textvariable=self.size_var, width=6).grid(row=1, column=2, sticky="w")
        fam_list = tk.Listbox(page, height=7, width=30, exportselection=False, activestyle="none")
        for fam in self.families:
            fam_list.insert("end", fam)
        fam_list.grid(row=2, column=0, sticky="w")
        if self.family_var.get() in self.families:
            i = self.families.index(self.family_var.get())
            fam_list.selection_set(i)
            fam_list.see(i)
        fam_list.bind("<<ListboxSelect>>", lambda e: self.family_var.set(self.families[fam_list.curselection()[0]])
                      if fam_list.curselection() else None)
        styles = ["Thường", "Nghiêng", "Đậm", "Đậm nghiêng"]
        st_list = tk.Listbox(page, height=7, width=14, exportselection=False, activestyle="none")
        for s in styles:
            st_list.insert("end", s)
        st_list.selection_set(styles.index(style))
        st_list.grid(row=2, column=1, sticky="w", padx=8)
        st_list.bind("<<ListboxSelect>>", lambda e: self.style_var.set(styles[st_list.curselection()[0]])
                     if st_list.curselection() else None)
        sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
        sz_list = tk.Listbox(page, height=7, width=6, exportselection=False, activestyle="none")
        for s in sizes:
            sz_list.insert("end", s)
        sz_list.grid(row=2, column=2, sticky="w")
        sz_list.bind("<<ListboxSelect>>", lambda e: self.size_var.set(str(sizes[sz_list.curselection()[0]]))
                     if sz_list.curselection() else None)
        row = ttk.Frame(page)
        row.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Label(row, text="Gạch chân:").pack(side="left")
        u = self.fmt.get("underline")
        self.underline_var = tk.StringVar(value="Kép" if u == "double" else "Đơn" if u else "Không")
        ttk.Combobox(row, values=["Không", "Đơn", "Kép"], textvariable=self.underline_var, state="readonly",
                     width=8).pack(side="left", padx=(6, 18))
        ttk.Label(row, text="Màu:").pack(side="left")
        self.fg_btn = ColorButton(row, self.ui, self.fmt.get("fg"), "Tự động", lambda c: self._update_preview())
        self.fg_btn.pack(side="left", padx=6)
        self.strike_var = tk.BooleanVar(value=bool(self.fmt.get("strike")))
        ttk.Checkbutton(page, text="Gạch ngang", variable=self.strike_var).grid(row=4, column=0, sticky="w", pady=6)
        prev = ttk.LabelFrame(page, text="Xem trước", padding=6)
        prev.grid(row=5, column=0, columnspan=3, sticky="ew")
        self.preview = tk.Label(prev, text="AaBbCcYyZz  Tiếng Việt", bg="#ffffff", width=40, height=2)
        self.preview.pack(fill="x")
        for var in (self.family_var, self.style_var, self.size_var, self.underline_var, self.strike_var):
            var.trace_add("write", lambda *a: self._update_preview())
        self._update_preview()

    def _font_values(self):
        style = self.style_var.get().lower()
        try:
            size = float(self.size_var.get().replace(",", "."))
        except ValueError:
            size = 11
        size = max(1.0, min(409.0, size))
        return {"font": self.family_var.get().strip() or self.default_family,
                "size": int(size) if size.is_integer() else size,
                "bold": "đậm" in style, "italic": "nghiêng" in style,
                "underline": {"Không": None, "Đơn": "single", "Kép": "double"}[self.underline_var.get()],
                "strike": self.strike_var.get(), "fg": self.fg_btn.color}

    def _update_preview(self):
        v = self._font_values()
        font = tkfont.Font(family=v["font"], size=-round(min(v["size"], 28) * 4 / 3),
                           weight="bold" if v["bold"] else "normal", slant="italic" if v["italic"] else "roman",
                           underline=bool(v["underline"]), overstrike=v["strike"])
        self._preview_font = font
        self.preview.config(font=font, fg=v["fg"] or "#000000")

    # borders
    def _border_tab(self):
        page = ttk.Frame(self.nb, padding=10)
        self.nb.add(page, text="Viền")
        left = ttk.LabelFrame(page, text="Đường kẻ", padding=8)
        left.grid(row=0, column=0, sticky="ns")
        ttk.Label(left, text="Kiểu:").pack(anchor="w")
        self.line_list = tk.Listbox(left, height=6, width=12, exportselection=False, activestyle="none")
        for _, name in BORDER_NAMES:
            self.line_list.insert("end", name)
        self.line_list.selection_set(0)
        self.line_list.pack(anchor="w")
        self.line_list.bind("<<ListboxSelect>>", lambda e: self._touch_border())
        ttk.Label(left, text="Màu:").pack(anchor="w", pady=(8, 0))
        self.line_color = ColorButton(left, self.ui, "#000000", "Tự động", lambda c: self._touch_border())
        self.line_color.pack(anchor="w")
        right = ttk.Frame(page)
        right.grid(row=0, column=1, sticky="nsew", padx=(14, 0))
        presets = ttk.Frame(right)
        presets.pack(anchor="w")
        ttk.Label(presets, text="Thiết lập sẵn:").pack(side="left")
        for text, command in (("Không", self._preset_none), ("Viền ngoài", self._preset_outline),
                              ("Bên trong", self._preset_inside)):
            ttk.Button(presets, text=text, command=command).pack(side="left", padx=3)
        self.edge_vars = {}
        edges = ttk.Frame(right)
        edges.pack(anchor="w", pady=8)
        names = [("top", "Trên"), ("inside_h", "Ngang giữa"), ("bottom", "Dưới"),
                 ("left", "Trái"), ("inside_v", "Dọc giữa"), ("right", "Phải")]
        initial = {"top": "bt", "bottom": "bb", "left": "bl", "right": "br"}
        for i, (key, label) in enumerate(names):
            var = tk.BooleanVar(value=bool(self.fmt.get(initial.get(key, ""))))
            self.edge_vars[key] = var
            state = "normal"
            if (key == "inside_h" and not self.multi_rows) or (key == "inside_v" and not self.multi_cols):
                state = "disabled"
            ttk.Checkbutton(edges, text=label, variable=var, state=state,
                            command=self._touch_border).grid(row=i % 3, column=i // 3, sticky="w", padx=(0, 18))
        self.border_cv = tk.Canvas(right, width=220, height=120, bg="#ffffff", highlightthickness=1,
                                   highlightbackground="#c8c8c8")
        self.border_cv.pack(anchor="w")
        existing = next((self.fmt[k] for k in fe.BORDER_KEYS if self.fmt.get(k)), None)
        if existing:
            names_only = [n for n, _ in BORDER_NAMES]
            if existing[0] in names_only:
                self.line_list.selection_clear(0, "end")
                self.line_list.selection_set(names_only.index(existing[0]))
            self.line_color._set(existing[1])
        self.border_touched = False
        self._draw_borders()

    def _line(self):
        sel = self.line_list.curselection()
        return [BORDER_NAMES[sel[0] if sel else 0][0], self.line_color.color or "#000000"]

    def _touch_border(self):
        self.border_touched = True
        self._draw_borders()

    def _preset_none(self):
        for var in self.edge_vars.values():
            var.set(False)
        self._touch_border()

    def _preset_outline(self):
        for key in ("top", "bottom", "left", "right"):
            self.edge_vars[key].set(True)
        self._touch_border()

    def _preset_inside(self):
        if self.multi_rows:
            self.edge_vars["inside_h"].set(True)
        if self.multi_cols:
            self.edge_vars["inside_v"].set(True)
        self._touch_border()

    def _draw_borders(self):
        if not hasattr(self, "border_cv"):
            return
        cv = self.border_cv
        cv.delete("all")
        x0, y0, x1, y1 = 30, 15, 190, 105
        cv.create_text(70, 45, text="Văn bản", fill="#999")
        cv.create_text(150, 75, text="Văn bản", fill="#999")
        style, color = self._line()
        width = {"medium": 2, "thick": 3}.get(style, 1)
        dash = {"dashed": "-", "dotted": "."}.get(style, "")
        lines = {"top": (x0, y0, x1, y0), "bottom": (x0, y1, x1, y1), "left": (x0, y0, x0, y1),
                 "right": (x1, y0, x1, y1), "inside_h": (x0, 60, x1, 60), "inside_v": (110, y0, 110, y1)}
        for key, pts in lines.items():
            if self.edge_vars[key].get():
                cv.create_line(*pts, fill=color, width=width, dash=dash)
            else:
                cv.create_line(*pts, fill="#e6e6e6", dash=".")

    # fill
    def _fill_tab(self):
        page = ttk.Frame(self.nb, padding=10)
        self.nb.add(page, text="Tô màu")
        self.fill_color = self.fmt.get("bg")
        ttk.Button(page, text="Không màu", command=lambda: self._set_fill(None)).pack(anchor="w")
        grid = tk.Frame(page)
        grid.pack(anchor="w", pady=8)
        for i, row in enumerate(theme_palette() + [STANDARD_COLORS]):
            for j, color in enumerate(row):
                sw = tk.Frame(grid, width=18, height=18, bg=color, highlightthickness=1, highlightbackground="#c8c8c8")
                sw.grid(row=i, column=j, padx=2, pady=(6 if i in (1, 6) else 1, 1))
                sw.bind("<ButtonRelease-1>", lambda e, c=color: self._set_fill(c))
        ttk.Button(page, text="Màu khác...", command=self._more_fill).pack(anchor="w")
        ttk.Label(page, text="Mẫu:").pack(anchor="w", pady=(10, 2))
        self.fill_preview = tk.Label(page, width=40, height=2, relief="solid", bd=1)
        self.fill_preview.pack(anchor="w")
        self._set_fill(self.fill_color)

    def _set_fill(self, color):
        self.fill_color = color
        self.fill_preview.config(bg=color or "#ffffff", text="" if color else "Không màu")

    def _more_fill(self):
        _, color = colorchooser.askcolor(color=self.fill_color, parent=self, title="Chọn màu")
        if color:
            self._set_fill(color)

    def collect(self):
        out = {}
        try:
            num = self.number_format()
            indent = max(0, min(15, int(self.indent_var.get())))
            rot = 255 if self.vertical_var.get() else max(-90, min(90, int(self.rot_var.get())))
        except (tk.TclError, ValueError):
            return None
        if num != self._initial_num:
            out["num"] = num
        align, valign = self.h_names[self.h_var.get()], self.v_names[self.v_var.get()]
        for key, value in (("align", align), ("valign", valign), ("wrap", self.wrap_var.get() or None),
                           ("indent", indent or None), ("rotate", rot or None)):
            if value != (self.fmt.get(key) or None):
                out[key] = value
        if self.merge_var.get() != self.merged:
            out["merge"] = self.merge_var.get()
        fv = self._font_values()
        if fv["font"] == self.default_family:
            fv["font"] = None
        if fv["size"] == 11:
            fv["size"] = None
        for key, value in fv.items():
            old = self.fmt.get(key)
            if key == "underline" and old is True:
                old = "single"
            if (value or None) != (old or None):
                out[key] = value or None
        if self.fill_color != self.fmt.get("bg"):
            out["bg"] = self.fill_color
        if self.border_touched:
            line = self._line()
            out["borders"] = {k: (line if v.get() else None) for k, v in self.edge_vars.items()}
        return out


class RuleDialog(Dialog):
    """Value(s) and highlight style for a conditional-formatting rule."""

    def __init__(self, master, title, prompt, fields=(), choice=None):
        super().__init__(master, title)
        ttk.Label(self.body, text=prompt).pack(anchor="w")
        row = ttk.Frame(self.body)
        row.pack(anchor="w", pady=8)
        self.vars = []
        first = None
        for field in fields:
            if field.get("label"):
                ttk.Label(row, text=field["label"]).pack(side="left", padx=4)
            var = tk.StringVar(value=field.get("default", ""))
            widget = (ttk.Spinbox(row, from_=1, to=1000, width=6, textvariable=var) if field.get("spin")
                      else ttk.Entry(row, textvariable=var, width=18))
            widget.pack(side="left")
            if field.get("after"):
                ttk.Label(row, text=field["after"]).pack(side="left", padx=4)
            first = first or widget
            self.vars.append(var)
        self.choice_var = None
        if choice:
            self.choice_var = tk.StringVar(value=choice[0])
            ttk.Combobox(row, values=choice, textvariable=self.choice_var, state="readonly", width=12).pack(side="left", padx=4)
        ttk.Label(row, text="với").pack(side="left", padx=6)
        self.style_var = tk.StringVar(value=CF_PRESETS[0][0])
        ttk.Combobox(row, values=[n for n, _ in CF_PRESETS], textvariable=self.style_var, state="readonly",
                     width=26).pack(side="left")
        self.first = first
        self.add_buttons()

    def show(self):
        return super().show(self.first)

    def collect(self):
        values = [v.get().strip() for v in self.vars]
        if any(v == "" for v in values):
            self.bell()
            return None
        return {"values": values, "style": dict(dict(CF_PRESETS)[self.style_var.get()]),
                "choice": self.choice_var.get() if self.choice_var else None}


class SortDialog(Dialog):
    def __init__(self, master, columns, header):
        """columns: [(col, header_text)]."""
        super().__init__(master, "Sắp xếp")
        self.columns = columns
        self.header_var = tk.BooleanVar(value=header)
        top = ttk.Frame(self.body)
        top.pack(fill="x")
        ttk.Button(top, text="+ Thêm mức", command=self._add_level).pack(side="left")
        ttk.Button(top, text="− Xóa mức", command=self._remove_level).pack(side="left", padx=6)
        ttk.Checkbutton(top, text="Dữ liệu có tiêu đề", variable=self.header_var,
                        command=self._relabel).pack(side="right")
        self.levels_frame = ttk.Frame(self.body)
        self.levels_frame.pack(fill="x", pady=8)
        self.levels = []
        self._add_level()
        self.add_buttons()

    def _labels(self):
        return [(text if self.header_var.get() and text else f"Cột {fe.col_name(c)}") for c, text in self.columns]

    def _add_level(self):
        if len(self.levels) >= 3:
            return
        row = ttk.Frame(self.levels_frame)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text="Sắp xếp theo" if not self.levels else "Sau đó theo", width=13).pack(side="left")
        col_var = tk.StringVar(value=self._labels()[min(len(self.levels), len(self.columns) - 1)])
        box = ttk.Combobox(row, values=self._labels(), textvariable=col_var, state="readonly", width=22)
        box.pack(side="left")
        order_var = tk.StringVar(value="A → Z (nhỏ → lớn)")
        ttk.Combobox(row, values=["A → Z (nhỏ → lớn)", "Z → A (lớn → nhỏ)"], textvariable=order_var,
                     state="readonly", width=18).pack(side="left", padx=6)
        self.levels.append((row, box, col_var, order_var))

    def _remove_level(self):
        if len(self.levels) > 1:
            self.levels.pop()[0].destroy()

    def _relabel(self):
        labels = self._labels()
        for _, box, var, _ in self.levels:
            index = box.current()
            box.config(values=labels)
            var.set(labels[max(0, index)])

    def collect(self):
        labels = self._labels()
        keys = [(self.columns[labels.index(var.get())][0], order.get().startswith("Z")) for _, _, var, order in self.levels]
        return {"keys": keys, "header": self.header_var.get()}


class SeriesDialog(Dialog):
    def __init__(self, master, axis):
        super().__init__(master, "Chuỗi")
        self.axis_var = tk.StringVar(value=axis)
        self.kind_var = tk.StringVar(value="linear")
        self.step_var, self.stop_var = tk.StringVar(value="1"), tk.StringVar()
        box = ttk.LabelFrame(self.body, text="Chuỗi trong", padding=6)
        box.grid(row=0, column=0, sticky="nsew")
        ttk.Radiobutton(box, text="Hàng", value="row", variable=self.axis_var).pack(anchor="w")
        ttk.Radiobutton(box, text="Cột", value="col", variable=self.axis_var).pack(anchor="w")
        box = ttk.LabelFrame(self.body, text="Loại", padding=6)
        box.grid(row=0, column=1, sticky="nsew", padx=8)
        for text, value in (("Tuyến tính", "linear"), ("Tăng trưởng (nhân)", "growth"), ("Ngày", "date")):
            ttk.Radiobutton(box, text=text, value=value, variable=self.kind_var).pack(anchor="w")
        row = ttk.Frame(self.body)
        row.grid(row=1, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Label(row, text="Giá trị bước:").pack(side="left")
        ttk.Entry(row, textvariable=self.step_var, width=10).pack(side="left", padx=(4, 14))
        ttk.Label(row, text="Giá trị dừng:").pack(side="left")
        ttk.Entry(row, textvariable=self.stop_var, width=10).pack(side="left", padx=4)
        self.add_buttons()

    def collect(self):
        try:
            step = float(self.step_var.get().replace(",", "."))
            stop = float(self.stop_var.get().replace(",", ".")) if self.stop_var.get().strip() else None
        except ValueError:
            self.bell()
            return None
        return {"axis": self.axis_var.get(), "kind": self.kind_var.get(), "step": step, "stop": stop}


class ChoiceDialog(Dialog):
    """A list of radio options (Insert/Delete cells, ...)."""

    def __init__(self, master, title, options, default=None):
        super().__init__(master, title)
        self.var = tk.StringVar(value=default or options[0][1])
        for text, value in options:
            ttk.Radiobutton(self.body, text=text, value=value, variable=self.var).pack(anchor="w", pady=1)
        self.add_buttons()

    def collect(self):
        return self.var.get()


def cells_dialog(master, mode):
    if mode == "insert":
        options = [("Dịch chuyển ô sang phải", "right"), ("Dịch chuyển ô xuống dưới", "down"),
                   ("Toàn bộ hàng", "rows"), ("Toàn bộ cột", "cols")]
        return ChoiceDialog(master, "Chèn", options, "down").show()
    options = [("Dịch chuyển ô sang trái", "left"), ("Dịch chuyển ô lên trên", "up"),
               ("Toàn bộ hàng", "rows"), ("Toàn bộ cột", "cols")]
    return ChoiceDialog(master, "Xóa", options, "up").show()


class PasteSpecialDialog(Dialog):
    def __init__(self, master):
        super().__init__(master, "Dán đặc biệt")
        box = ttk.LabelFrame(self.body, text="Dán", padding=6)
        box.pack(fill="x")
        self.mode = tk.StringVar(value="all")
        for text, value in (("Tất cả", "all"), ("Công thức", "formulas"), ("Giá trị", "values"),
                            ("Định dạng", "formats")):
            ttk.Radiobutton(box, text=text, value=value, variable=self.mode).pack(anchor="w")
        self.skip = tk.BooleanVar()
        self.transpose = tk.BooleanVar()
        ttk.Checkbutton(self.body, text="Bỏ qua ô trống", variable=self.skip).pack(anchor="w", pady=(8, 0))
        ttk.Checkbutton(self.body, text="Chuyển vị", variable=self.transpose).pack(anchor="w")
        self.add_buttons()

    def collect(self):
        return {"mode": self.mode.get(), "skip_blanks": self.skip.get(), "transpose": self.transpose.get()}


class FunctionDialog(Dialog):
    def __init__(self, master):
        super().__init__(master, "Chèn hàm")
        ttk.Label(self.body, text="Tìm hàm:").grid(row=0, column=0, sticky="w")
        self.search = tk.StringVar()
        entry = ttk.Entry(self.body, textvariable=self.search, width=40)
        entry.grid(row=1, column=0, sticky="w")
        ttk.Label(self.body, text="Thể loại:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.cat = tk.StringVar(value="Tất cả")
        ttk.Combobox(self.body, values=["Tất cả"] + CATEGORIES, textvariable=self.cat, state="readonly",
                     width=30).grid(row=3, column=0, sticky="w")
        ttk.Label(self.body, text="Chọn hàm:").grid(row=4, column=0, sticky="w", pady=(8, 0))
        self.listbox = tk.Listbox(self.body, height=10, width=44, exportselection=False, activestyle="none")
        self.listbox.grid(row=5, column=0, sticky="w")
        self.syntax = ttk.Label(self.body, text="", font=("Segoe UI", 9, "bold"))
        self.syntax.grid(row=6, column=0, sticky="w", pady=(8, 0))
        self.desc = ttk.Label(self.body, text="", wraplength=330)
        self.desc.grid(row=7, column=0, sticky="w")
        self.search.trace_add("write", lambda *a: self._fill())
        self.cat.trace_add("write", lambda *a: self._fill())
        self.listbox.bind("<<ListboxSelect>>", lambda e: self._describe())
        self.listbox.bind("<Double-Button-1>", lambda e: self.ok())
        self.names = []
        self._fill()
        self.add_buttons()
        self.entry = entry

    def show(self):
        return super().show(self.entry)

    def _fill(self):
        q, cat = self.search.get().strip().upper(), self.cat.get()
        self.names = [n for n in sorted(FUNCTION_INFO)
                      if (cat == "Tất cả" or FUNCTION_INFO[n][0] == cat)
                      and (not q or q in n or q.lower() in FUNCTION_INFO[n][2].lower())]
        self.listbox.delete(0, "end")
        for n in self.names:
            self.listbox.insert("end", n)
        if self.names:
            self.listbox.selection_set(0)
        self._describe()

    def _describe(self):
        sel = self.listbox.curselection()
        if not sel:
            self.syntax.config(text="")
            self.desc.config(text="")
            return
        _, syntax, desc = FUNCTION_INFO[self.names[sel[0]]]
        self.syntax.config(text=syntax)
        self.desc.config(text=desc)

    def collect(self):
        sel = self.listbox.curselection()
        return self.names[sel[0]] if sel else None


class TableDialog(Dialog):
    def __init__(self, master, range_text, header=True):
        super().__init__(master, "Tạo bảng")
        ttk.Label(self.body, text="Dữ liệu của bảng nằm ở đâu?").pack(anchor="w")
        self.range_var = tk.StringVar(value=range_text)
        self.entry = ttk.Entry(self.body, textvariable=self.range_var, width=30)
        self.entry.pack(anchor="w", pady=6)
        self.header_var = tk.BooleanVar(value=header)
        ttk.Checkbutton(self.body, text="Bảng có tiêu đề", variable=self.header_var).pack(anchor="w")
        self.add_buttons()

    def show(self):
        return super().show(self.entry)

    def collect(self):
        rect = fe.parse_range(self.range_var.get().replace("$", "").strip())
        if rect is None:
            self.bell()
            return None
        return rect, self.header_var.get()


class RemoveDuplicatesDialog(Dialog):
    def __init__(self, master, columns, header):
        super().__init__(master, "Loại bỏ trùng lặp")
        self.columns = columns
        ttk.Label(self.body, text="Chọn các cột dùng để so sánh trùng lặp:").pack(anchor="w")
        self.header_var = tk.BooleanVar(value=header)
        ttk.Checkbutton(self.body, text="Dữ liệu có tiêu đề", variable=self.header_var).pack(anchor="e")
        box = ttk.Frame(self.body)
        box.pack(fill="x", pady=6)
        self.vars = []
        for c, text in columns:
            var = tk.BooleanVar(value=True)
            ttk.Checkbutton(box, text=f"{fe.col_name(c)}  {text}" if text else f"Cột {fe.col_name(c)}",
                            variable=var).pack(anchor="w")
            self.vars.append(var)
        self.add_buttons()

    def collect(self):
        cols = [c for (c, _), var in zip(self.columns, self.vars) if var.get()]
        if not cols:
            self.bell()
            return None
        return {"cols": cols, "header": self.header_var.get()}


class ManageRulesDialog(Dialog):
    def __init__(self, master, rules, describe):
        super().__init__(master, "Quản lý quy tắc định dạng có điều kiện")
        self.rules = list(rules)
        self.describe = describe
        self.listbox = tk.Listbox(self.body, width=70, height=10, activestyle="none")
        self.listbox.pack(fill="both")
        row = ttk.Frame(self.body)
        row.pack(fill="x", pady=(6, 0))
        ttk.Button(row, text="Xóa quy tắc", command=self._delete).pack(side="left")
        ttk.Label(row, text="Quy tắc ở cuối danh sách được ưu tiên áp dụng sau cùng.", foreground="#666").pack(side="left", padx=10)
        self._fill()
        self.add_buttons()

    def _fill(self):
        self.listbox.delete(0, "end")
        for rule in self.rules:
            self.listbox.insert("end", self.describe(rule))

    def _delete(self):
        sel = self.listbox.curselection()
        if sel:
            del self.rules[sel[0]]
            self._fill()

    def collect(self):
        return self.rules


FILTER_OP_LABELS = [("=", "bằng"), ("<>", "không bằng"), (">", "lớn hơn"), (">=", "lớn hơn hoặc bằng"),
                    ("<", "nhỏ hơn"), ("<=", "nhỏ hơn hoặc bằng"), ("begins", "bắt đầu bằng"),
                    ("nbegins", "không bắt đầu bằng"), ("ends", "kết thúc bằng"), ("nends", "không kết thúc bằng"),
                    ("contains", "có chứa"), ("ncontains", "không chứa")]
DATE_OP_LABELS = {">": "sau", ">=": "sau hoặc bằng", "<": "trước", "<=": "trước hoặc bằng"}
DYNAMIC_LABELS = {"tomorrow": "Ngày mai", "today": "Hôm nay", "yesterday": "Hôm qua", "nextweek": "Tuần tới",
                  "thisweek": "Tuần này", "lastweek": "Tuần trước", "nextmonth": "Tháng tới", "thismonth": "Tháng này",
                  "lastmonth": "Tháng trước", "nextyear": "Năm tới", "thisyear": "Năm nay", "lastyear": "Năm trước"}


def describe_filter(crit, kind="text"):
    """A short Vietnamese description of a column's filter, e.g. 'lớn hơn 10 và nhỏ hơn 20'."""
    if crit is None:
        return ""
    if not isinstance(crit, dict):
        return f"{len(crit)} giá trị được chọn"
    labels = dict(FILTER_OP_LABELS, **(DATE_OP_LABELS if kind == "date" else {}))
    t = crit["type"]
    if t == "custom":
        return (" và " if crit.get("join", "and") == "and" else " hoặc ").join(
            f"{labels[op]} “{value}”" for op, value in crit["conds"])
    if t == "top":
        return (f"{'dưới' if crit.get('bottom') else 'trên'} {crit.get('n', 10)}"
                f"{'%' if crit.get('percent') else ' mục'}")
    if t == "average":
        return "trên trung bình" if crit.get("above", True) else "dưới trung bình"
    if t == "dynamic":
        return DYNAMIC_LABELS[crit["kind"]].lower()
    return ("màu ô " if crit.get("what", "bg") == "bg" else "màu phông ") + (crit.get("color") or "tự động")


class FilterPopup:
    """Excel's AutoFilter dropdown: sort (also by colour), clear, filter by colour, the Number / Text /
    Date Filters menu, a search box and the checklist of values.

    kind is "number", "date" or "text"; colors {"bg": [...], "fg": [...]} the colours to offer (None means
    no fill / automatic); active the column's current criterion. Callbacks: on_sort(descending),
    on_sort_color(what, color), on_clear(), on_apply(set of texts or None for all), on_criterion(criterion)
    and on_dialog(name, preset) for the "custom" and "top" dialogs."""

    def __init__(self, master, ui, title, values, selected, x, y, *, kind="text", colors=None, active=None,
                 on_sort, on_clear, on_apply, on_sort_color=None, on_criterion=None, on_dialog=None, icons=None):
        self.values, self.on_apply, self.ui, self.icons = values, on_apply, ui, icons
        self.active, self.kind = active, kind
        self.on_criterion, self.on_dialog = on_criterion, on_dialog
        colors = colors or {}
        self.pop = pop = Popup(master, x, y)
        f = pop.inner
        self.pad = pad = ui.px(6)
        asc, desc = {"number": ("Sắp xếp từ nhỏ nhất đến lớn nhất", "Sắp xếp từ lớn nhất đến nhỏ nhất"),
                     "date": ("Sắp xếp từ cũ nhất đến mới nhất", "Sắp xếp từ mới nhất đến cũ nhất")}.get(
            kind, ("Sắp xếp từ A đến Z", "Sắp xếp từ Z đến A"))
        has_colors = bool(colors.get("bg") or colors.get("fg"))
        self._entry(f, "sort_az", asc, lambda: on_sort(False))
        self._entry(f, "sort_za", desc, lambda: on_sort(True))
        self._entry(f, "sort_color", "Sắp xếp theo màu", lambda w: self._color_menu(w, colors, on_sort_color),
                    has_colors and on_sort_color is not None, submenu=True)
        tk.Frame(f, bg=GROUP_LINE, height=1).pack(fill="x", padx=2, pady=2)
        self._entry(f, "filter_clear", f"Xóa bộ lọc khỏi “{title}”", on_clear, active is not None)
        color_on = isinstance(active, dict) and active["type"] == "color"
        self._entry(f, "filter_color", "Lọc theo màu",
                    lambda w: self._color_menu(w, colors, lambda what, c: on_criterion(
                        {"type": "color", "what": what, "color": c}), active if color_on else None),
                    has_colors and on_criterion is not None, submenu=True, checked=color_on)
        rules_on = isinstance(active, dict) and active["type"] != "color"
        label = {"number": "Bộ lọc số", "date": "Bộ lọc ngày"}.get(kind, "Bộ lọc văn bản")
        self._entry(f, None, label, self._rules_menu, on_criterion is not None, submenu=True, checked=rules_on)
        if active is not None and not isinstance(active, set):
            tk.Label(f, text="Đang lọc: " + describe_filter(active, kind), bg="#ffffff", fg="#2b579a", anchor="w",
                     font=ui.small_font, padx=pad, wraplength=ui.px(250), justify="left").pack(fill="x", padx=2)
        tk.Frame(f, bg=GROUP_LINE, height=1).pack(fill="x", padx=2, pady=2)
        self.search = tk.StringVar()
        search = ttk.Entry(f, textvariable=self.search, width=30, font=ui.font)
        search.pack(fill="x", padx=pad, pady=(2, 4))
        holder = tk.Frame(f, bg="#ffffff", highlightthickness=1, highlightbackground="#c8c8c8")
        holder.pack(fill="both", padx=pad)
        self.canvas = tk.Canvas(holder, width=ui.px(230), height=ui.px(180), bg="#ffffff", highlightthickness=0)
        bar = tk.Scrollbar(holder, orient="vertical", command=self.canvas.yview)
        self.canvas.config(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both")
        self.list_frame = tk.Frame(self.canvas, bg="#ffffff")
        self.canvas.create_window(0, 0, window=self.list_frame, anchor="nw")
        self.list_frame.bind("<Configure>", lambda e: self.canvas.config(scrollregion=self.canvas.bbox("all")))
        pop.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(-round(e.delta / 120) or -1, "units"))
        self.vars = {v: tk.BooleanVar(value=selected is None or v in selected) for v in values}
        self.all_var = tk.BooleanVar(value=all(var.get() for var in self.vars.values()))
        self.changed = False  # OK without touching the list keeps a rule filter as it is
        self.font = ui.font
        self._fill()
        self.search.trace_add("write", lambda *a: (setattr(self, "changed", True), self._fill()))
        buttons = tk.Frame(f, bg="#ffffff")
        buttons.pack(fill="x", padx=pad, pady=pad)
        ttk.Button(buttons, text="Hủy", width=9, command=pop.close).pack(side="right")
        ttk.Button(buttons, text="OK", width=9, command=self._ok).pack(side="right", padx=4)
        pop.show()
        search.focus_set()
        pop.bind("<Return>", lambda e: self._ok())

    def _entry(self, parent, icon, text, command, enabled=True, submenu=False, checked=False, swatch=None,
               popup=None):
        """A menu-like row; `command` takes the row widget when it opens a submenu. Picking an entry of a
        submenu (`popup`) closes it and the dropdown before running the command."""
        # a picture (or a blank one, to keep labels aligned), unless a colour swatch stands in for it
        image = self.icons.get(icon) if self.icons is not None and swatch is None else None
        row = tk.Frame(parent, bg="#ffffff")
        row.pack(fill="x", padx=2)
        widgets, sw = [row], None
        if swatch is not None:
            sw = tk.Label(row, text="", bg=swatch[0], width=2, relief="solid", bd=1)
            sw.pack(side="left", padx=(self.pad, 0), pady=2)
            widgets.append(sw)
        lbl = tk.Label(row, text=("✓ " if checked else "") + text, bg="#ffffff",
                       fg=TEXT_FG if enabled else "#a0a0a0", anchor="w", font=self.ui.font, padx=self.pad,
                       pady=self.ui.px(3), **({"image": image, "compound": "left"} if image is not None else {}))
        lbl.pack(side="left", fill="x", expand=True)
        widgets.append(lbl)
        if submenu:
            arrow = tk.Label(row, text="›", bg="#ffffff", fg="#606060" if enabled else "#c0c0c0", font=self.ui.font,
                             padx=self.pad)
            arrow.pack(side="right")
            widgets.append(arrow)
        if not enabled:
            return row

        def hover(on):
            for w in widgets:
                if w is not sw:
                    w.config(bg=HOVER_BG if on else "#ffffff")

        def run(e):
            if submenu:
                command(row)
            elif popup is not None:
                popup.close()
                self.pop.close()
                command()
            else:
                self.pop.close()
                command()
        for w in widgets:
            w.bind("<Enter>", lambda e: hover(True))
            w.bind("<Leave>", lambda e: hover(False))
            w.bind("<ButtonRelease-1>", run)
        return row

    def _submenu(self, row, items):
        """A dropdown to the right of `row`: items are (icon, label, command[, checked, swatch]) or None."""
        sub = Popup(self.pop, row.winfo_rootx() + row.winfo_width(), row.winfo_rooty())
        for item in items:
            if item is None:
                tk.Frame(sub.inner, bg=GROUP_LINE, height=1).pack(fill="x", padx=2, pady=2)
            elif item[2] is None:  # a heading
                tk.Label(sub.inner, text=item[1], bg="#f3f3f3", fg="#555555", anchor="w", font=self.ui.small_font,
                         padx=self.pad, pady=self.ui.px(2)).pack(fill="x", padx=2)
            else:
                icon, label, command = item[:3]
                checked = len(item) > 3 and item[3]
                swatch = item[4] if len(item) > 4 else None
                self._entry(sub.inner, icon, label, command, checked=checked, swatch=swatch, popup=sub)
        sub.show()

    def _color_menu(self, row, colors, pick, active=None):
        items = []
        for what, heading, none_label in (("bg", "Theo màu ô", "Không tô"), ("fg", "Theo màu phông", "Tự động")):
            if not colors.get(what):
                continue
            items.append((None, heading, None))
            for color in colors[what]:
                checked = active is not None and active.get("what") == what and active.get("color") == color
                swatch = (color or ("#ffffff" if what == "bg" else "#000000"),)
                items.append((None, (color or none_label).upper() if color else none_label,
                              lambda w=what, c=color: pick(w, c), checked, swatch))
        self._submenu(row, items)

    def _rules_menu(self, row):
        active = self.active if isinstance(self.active, dict) else {}
        on, dialog = self.on_criterion, self.on_dialog

        def custom(*conds, join="and"):
            return lambda: dialog("custom", {"type": "custom", "join": join, "conds": [[op, ""] for op in conds]})

        def current_or(preset):  # "Custom Filter..." starts from the filter in place, like Excel
            return lambda: dialog("custom", active if active.get("type") == "custom" else preset)

        if self.kind == "date":
            items = [(None, "Bằng...", custom("=")), (None, "Trước...", custom("<")), (None, "Sau...", custom(">")),
                     (None, "Nằm giữa...", custom(">=", "<=")), None]
            for i, kind in enumerate(DYNAMIC_LABELS):
                items.append((None, DYNAMIC_LABELS[kind], lambda k=kind: on({"type": "dynamic", "kind": k}),
                              active.get("type") == "dynamic" and active.get("kind") == kind))
                if i % 3 == 2:
                    items.append(None)
        elif self.kind == "number":
            items = [(None, "Bằng...", custom("=")), (None, "Không bằng...", custom("<>")),
                     (None, "Lớn hơn...", custom(">")), (None, "Lớn hơn hoặc bằng...", custom(">=")),
                     (None, "Nhỏ hơn...", custom("<")), (None, "Nhỏ hơn hoặc bằng...", custom("<=")),
                     (None, "Nằm giữa...", custom(">=", "<=")), None,
                     ("rank_top", "10 mục đầu...", lambda: dialog("top", active if active.get("type") == "top" else
                                                                 {"type": "top", "n": 10}), active.get("type") == "top"),
                     ("rank_above", "Trên trung bình", lambda: on({"type": "average", "above": True}),
                      active.get("type") == "average" and active.get("above", True)),
                     ("rank_below", "Dưới trung bình", lambda: on({"type": "average", "above": False}),
                      active.get("type") == "average" and not active.get("above", True)), None]
        else:
            items = [(None, "Bằng...", custom("=")), (None, "Không bằng...", custom("<>")), None,
                     (None, "Bắt đầu bằng...", custom("begins")), (None, "Kết thúc bằng...", custom("ends")), None,
                     (None, "Có chứa...", custom("contains")), (None, "Không chứa...", custom("ncontains")), None]
        items.append(("filter", "Bộ lọc tùy chỉnh...", current_or({"type": "custom", "join": "and",
                                                                    "conds": [["=", ""]]}),
                      active.get("type") == "custom"))
        self._submenu(row, items)

    def _visible(self):
        q = self.search.get().strip().lower()
        return [v for v in self.values if q in (v or "(trống)").lower()]

    def _fill(self):
        for w in self.list_frame.winfo_children():
            w.destroy()
        tk.Checkbutton(self.list_frame, text="(Chọn tất cả)", variable=self.all_var, bg="#ffffff", anchor="w",
                       font=self.font, command=self._toggle_all, activebackground="#ffffff").pack(fill="x", anchor="w")
        for v in self._visible()[:1000]:
            tk.Checkbutton(self.list_frame, text=v if v != "" else "(Trống)", variable=self.vars[v], bg="#ffffff",
                           anchor="w", font=self.font, activebackground="#ffffff",
                           command=self._sync_all).pack(fill="x", anchor="w")
        self.canvas.yview_moveto(0)

    def _toggle_all(self):
        self.changed = True
        for v in self._visible():
            self.vars[v].set(self.all_var.get())

    def _sync_all(self):
        self.changed = True
        self.all_var.set(all(self.vars[v].get() for v in self._visible()))

    def _ok(self):
        visible = set(self._visible())
        allowed = {v for v, var in self.vars.items() if var.get() and v in visible}
        self.pop.close()
        if self.changed or not isinstance(self.active, dict):
            self.on_apply(None if allowed == set(self.values) else allowed)


class CustomFilterDialog(Dialog):
    """Excel's Custom AutoFilter: one or two conditions joined by And / Or. collect() gives the criterion."""

    def __init__(self, master, title, values, crit, kind="text"):
        super().__init__(master, "Bộ lọc tự động tùy chỉnh")
        self.labels = [(op, DATE_OP_LABELS.get(op, label) if kind == "date" else label)
                       for op, label in FILTER_OP_LABELS]
        conds = [list(c) for c in crit.get("conds", [])][:2] + [["", ""]] * 2
        ttk.Label(self.body, text="Hiện các hàng có:").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(self.body, text=title, font=("TkDefaultFont", 9, "bold")).grid(row=1, column=0, columnspan=2,
                                                                                sticky="w", pady=(2, 6))
        self.ops, self.vals = [], []
        for i in range(2):
            op = tk.StringVar(value=dict(self.labels).get(conds[i][0], ""))
            box = ttk.Combobox(self.body, values=([""] if i else []) + [label for _, label in self.labels],
                               textvariable=op, state="readonly", width=20)
            box.grid(row=2 + 2 * i, column=0, sticky="w", pady=2)
            val = tk.StringVar(value=conds[i][1])
            entry = ttk.Combobox(self.body, values=[v for v in values if v][:500], textvariable=val, width=26)
            entry.grid(row=2 + 2 * i, column=1, sticky="w", padx=(6, 0), pady=2)
            self.ops.append(op)
            self.vals.append(val)
            if i == 0:
                self.first = entry
                self.join = tk.StringVar(value=crit.get("join", "and"))
                row = ttk.Frame(self.body)
                row.grid(row=3, column=0, columnspan=2, sticky="w", padx=20)
                ttk.Radiobutton(row, text="Và", value="and", variable=self.join).pack(side="left")
                ttk.Radiobutton(row, text="Hoặc", value="or", variable=self.join).pack(side="left", padx=12)
        ttk.Label(self.body, text="Dùng ? thay cho một ký tự bất kỳ\nDùng * thay cho một chuỗi ký tự bất kỳ",
                  foreground="#666").grid(row=6, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.add_buttons()

    def show(self):
        return super().show(self.first)

    def collect(self):
        ops = {label: op for op, label in self.labels}
        conds = [[ops[o.get()], v.get().strip()] for o, v in zip(self.ops, self.vals) if o.get()]
        if not conds:
            self.bell()
            return None
        return {"type": "custom", "join": self.join.get(), "conds": conds}


class Top10Dialog(Dialog):
    """Excel's Top 10 AutoFilter: the top or bottom n items or percent."""

    def __init__(self, master, crit):
        super().__init__(master, "Bộ lọc 10 mục đầu")
        ttk.Label(self.body, text="Hiện").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        self.side = tk.StringVar(value="Dưới" if crit.get("bottom") else "Trên")
        ttk.Combobox(self.body, values=["Trên", "Dưới"], textvariable=self.side, state="readonly",
                     width=8).grid(row=1, column=0)
        self.n = tk.StringVar(value=str(crit.get("n", 10)))
        self.spin = ttk.Spinbox(self.body, from_=1, to=500, textvariable=self.n, width=6)
        self.spin.grid(row=1, column=1, padx=6)
        self.unit = tk.StringVar(value="Phần trăm" if crit.get("percent") else "Mục")
        ttk.Combobox(self.body, values=["Mục", "Phần trăm"], textvariable=self.unit, state="readonly",
                     width=10).grid(row=1, column=2)
        self.add_buttons()

    def show(self):
        return super().show(self.spin)

    def collect(self):
        try:
            n = int(self.n.get())
        except ValueError:
            n = 0
        if not 1 <= n <= (100 if self.unit.get() == "Phần trăm" else 500):
            self.bell()
            return None
        return {"type": "top", "n": n, "bottom": self.side.get() == "Dưới", "percent": self.unit.get() == "Phần trăm"}


VALIDATION_ALLOW = [("any", "Bất kỳ giá trị nào"), ("whole", "Số nguyên"), ("decimal", "Số thập phân"),
                    ("list", "Danh sách"), ("date", "Ngày"), ("length", "Độ dài văn bản"), ("custom", "Tùy chỉnh")]
VALIDATION_OPS = [("between", "nằm giữa"), ("notbetween", "không nằm giữa"), ("=", "bằng"), ("<>", "khác"),
                  (">", "lớn hơn"), ("<", "nhỏ hơn"), (">=", "lớn hơn hoặc bằng"), ("<=", "nhỏ hơn hoặc bằng")]
ERROR_STYLES = [("stop", "Dừng"), ("warning", "Cảnh báo"), ("info", "Thông tin")]


def _text_box(master, text, height):
    box = tk.Text(master, width=46, height=height, wrap="word", font="TkDefaultFont", relief="solid", bd=1,
                  highlightthickness=0)
    box.insert("1.0", text or "")
    return box


class ValidationDialog(Dialog):
    """Excel's Data Validation: what may be typed, the message shown while the cell is selected and
    the alert for a wrong entry. collect() returns the rule without its range, or {"clear": True}."""

    def __init__(self, master, rule):
        super().__init__(master, "Xác thực dữ liệu")
        rule = rule or {}
        kind = rule.get("type", "any")
        nb = ttk.Notebook(self.body)
        nb.pack(fill="both", expand=True)

        f = ttk.Frame(nb, padding=10)
        nb.add(f, text="Cài đặt")
        ttk.Label(f, text="Cho phép:").grid(row=0, column=0, sticky="w")
        self.allow = tk.StringVar(value=dict(VALIDATION_ALLOW)[kind])
        allow_box = ttk.Combobox(f, values=[label for _, label in VALIDATION_ALLOW], textvariable=self.allow,
                                 state="readonly", width=24)
        allow_box.grid(row=1, column=0, sticky="w")
        self.blank = tk.BooleanVar(value=rule.get("blank", True))
        ttk.Checkbutton(f, text="Bỏ qua ô trống", variable=self.blank).grid(row=1, column=1, sticky="w", padx=12)
        self.op_label = ttk.Label(f, text="Dữ liệu:")
        self.op_label.grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.op = tk.StringVar(value=dict(VALIDATION_OPS)[rule.get("op", "between")])
        self.op_box = ttk.Combobox(f, values=[label for _, label in VALIDATION_OPS], textvariable=self.op,
                                   state="readonly", width=24)
        self.op_box.grid(row=3, column=0, sticky="w")
        self.v1 = tk.StringVar(value=rule.get("source", "") if kind == "list" else rule.get("v1", ""))
        self.v2 = tk.StringVar(value=rule.get("v2", ""))
        self.label1 = ttk.Label(f)
        self.label1.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.entry1 = ttk.Entry(f, textvariable=self.v1, width=48)
        self.entry1.grid(row=5, column=0, columnspan=2, sticky="we")
        self.label2 = ttk.Label(f)
        self.label2.grid(row=6, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.entry2 = ttk.Entry(f, textvariable=self.v2, width=48)
        self.entry2.grid(row=7, column=0, columnspan=2, sticky="we")
        self.dropdown = tk.BooleanVar(value=rule.get("dropdown", True))
        self.dropdown_check = ttk.Checkbutton(f, text="Hiện mũi tên thả xuống trong ô", variable=self.dropdown)
        self.dropdown_check.grid(row=8, column=0, columnspan=2, sticky="w", pady=(8, 0))
        for box in (allow_box, self.op_box):
            box.bind("<<ComboboxSelected>>", lambda e: self._update())

        f = ttk.Frame(nb, padding=10)
        nb.add(f, text="Thông báo nhập")
        ttk.Label(f, text="Khi chọn ô, hiện thông báo này:").pack(anchor="w")
        ttk.Label(f, text="Tiêu đề:").pack(anchor="w", pady=(8, 0))
        self.prompt_title = tk.StringVar(value=rule.get("prompt_title", ""))
        ttk.Entry(f, textvariable=self.prompt_title, width=48).pack(anchor="w")
        ttk.Label(f, text="Thông báo:").pack(anchor="w", pady=(8, 0))
        self.prompt = _text_box(f, rule.get("prompt"), 5)
        self.prompt.pack(anchor="w")

        f = ttk.Frame(nb, padding=10)
        nb.add(f, text="Cảnh báo lỗi")
        ttk.Label(f, text="Khi nhập dữ liệu không hợp lệ, hiện cảnh báo:").pack(anchor="w")
        ttk.Label(f, text="Kiểu:").pack(anchor="w", pady=(8, 0))
        self.style = tk.StringVar(value=dict(ERROR_STYLES)[rule.get("style", "stop")])
        ttk.Combobox(f, values=[label for _, label in ERROR_STYLES], textvariable=self.style, state="readonly",
                     width=14).pack(anchor="w")
        ttk.Label(f, text="Tiêu đề:").pack(anchor="w", pady=(8, 0))
        self.error_title = tk.StringVar(value=rule.get("error_title", ""))
        ttk.Entry(f, textvariable=self.error_title, width=48).pack(anchor="w")
        ttk.Label(f, text="Thông báo lỗi:").pack(anchor="w", pady=(8, 0))
        self.error = _text_box(f, rule.get("error"), 4)
        self.error.pack(anchor="w")

        bar = self.add_buttons()
        ttk.Button(bar, text="Xóa tất cả", command=self._clear).pack(side="left")
        self.unbind("<Return>")  # Enter belongs to the multi-line message boxes
        self._update()

    def _kind(self):
        return next(k for k, label in VALIDATION_ALLOW if label == self.allow.get())

    def _op(self):
        return next(k for k, label in VALIDATION_OPS if label == self.op.get())

    def _update(self):
        kind, op = self._kind(), self._op()
        numeric = kind in ("whole", "decimal", "date", "length")
        two = numeric and op in ("between", "notbetween")
        word = {"date": "Ngày", "length": "Độ dài"}.get(kind, "Giá trị")
        for widget, shown in ((self.op_label, numeric), (self.op_box, numeric), (self.label1, kind != "any"),
                              (self.entry1, kind != "any"), (self.label2, two), (self.entry2, two),
                              (self.dropdown_check, kind == "list")):
            if shown:
                widget.grid()
            else:
                widget.grid_remove()
        if kind == "list":
            self.label1.config(text="Nguồn (ví dụ  Có,Không  hoặc  =$A$1:$A$10):")
        elif kind == "custom":
            self.label1.config(text="Công thức phải cho kết quả TRUE (ví dụ  =A1>0):")
        elif two:
            self.label1.config(text=f"{word} tối thiểu:")
            self.label2.config(text=f"{word} tối đa:")
        else:
            self.label1.config(text=f"{word}:")

    def _clear(self):
        self.result = {"clear": True}
        self.destroy()

    def collect(self):
        kind, op = self._kind(), self._op()
        rule = {"type": kind, "blank": self.blank.get(),
                "style": next(k for k, label in ERROR_STYLES if label == self.style.get()),
                "error_title": self.error_title.get().strip(), "error": self.error.get("1.0", "end").strip(),
                "prompt_title": self.prompt_title.get().strip(), "prompt": self.prompt.get("1.0", "end").strip()}
        v1, v2 = self.v1.get().strip(), self.v2.get().strip()
        if kind == "any":
            return rule
        if not v1:
            self.bell()
            return None
        if kind == "list":
            return dict(rule, source=v1, dropdown=self.dropdown.get())
        if kind == "custom":
            return dict(rule, v1=v1 if v1.startswith("=") else "=" + v1)
        values = [v1, v2] if op in ("between", "notbetween") else [v1]
        for v in values:
            if not (fe.is_formula(v) or fe.is_number(fe.parse_literal(v)[0])):
                self.bell()
                return None
        rule.update(op=op, v1=v1)
        if len(values) == 2:
            rule["v2"] = v2
        return rule


class NoteDialog(Dialog):
    """Text of a cell note; collect() returns it ("" removes the note)."""

    def __init__(self, master, cell, text):
        super().__init__(master, f"Ghi chú - {cell}")
        self.box = _text_box(self.body, text, 7)
        self.box.pack(fill="both", expand=True)
        ttk.Label(self.body, text="Ctrl+Enter để lưu.", foreground="#666").pack(anchor="w", pady=(4, 0))
        self.add_buttons()
        self.unbind("<Return>")
        self.box.bind("<Control-Return>", lambda e: (self.ok(), "break")[1])

    def show(self):
        self.box.mark_set("insert", "end")
        return super().show(self.box)

    def collect(self):
        return self.box.get("1.0", "end").strip()


class ListPopup:
    """The in-cell dropdown of a list validation: click or Enter picks an item."""

    def __init__(self, master, ui, items, current, x, y, width, on_pick):
        self.on_pick, self.items = on_pick, items
        self.pop = pop = Popup(master, x, y)
        holder = tk.Frame(pop.inner, bg="#ffffff")
        holder.pack(fill="both", expand=True)
        self.listbox = lb = tk.Listbox(holder, height=min(8, max(1, len(items))), activestyle="none",
                                       font=ui.font, relief="flat", bd=0, highlightthickness=0,
                                       selectbackground="#cce4f7", selectforeground=TEXT_FG,
                                       width=max(4, round(width / max(1, ui.font.measure("0")))))
        if len(items) > 8:
            bar = ttk.Scrollbar(holder, orient="vertical", command=lb.yview)
            lb.config(yscrollcommand=bar.set)
            bar.pack(side="right", fill="y")
        lb.pack(side="left", fill="both", expand=True)
        for item in items:
            lb.insert("end", item)
        index = next((i for i, item in enumerate(items) if item.lower() == (current or "").lower()), 0)
        lb.selection_set(index)
        lb.activate(index)
        lb.see(index)
        lb.bind("<ButtonRelease-1>", lambda e: self._pick())
        lb.bind("<Return>", lambda e: self._pick())
        lb.bind("<Motion>", lambda e: (lb.selection_clear(0, "end"), lb.selection_set(lb.nearest(e.y))))
        pop.show()
        lb.focus_set()

    def _pick(self):
        sel = self.listbox.curselection()
        self.pop.close()
        if sel:
            self.on_pick(self.items[sel[0]])


CHART_KINDS = [("column", "Cột"), ("bar", "Thanh"), ("line", "Đường"), ("pie", "Tròn")]


class ChartDialog(Dialog):
    """Type, data range and title of a chart; collect() returns them as a dict."""

    def __init__(self, master, chart):
        super().__init__(master, "Biểu đồ")
        f = self.body
        ttk.Label(f, text="Kiểu biểu đồ:").grid(row=0, column=0, sticky="w", pady=3)
        self.kind = tk.StringVar(value=dict(CHART_KINDS)[chart["type"]])
        ttk.Combobox(f, values=[label for _, label in CHART_KINDS], textvariable=self.kind, state="readonly",
                     width=12).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(f, text="Vùng dữ liệu:").grid(row=1, column=0, sticky="w", pady=3)
        self.range_var = tk.StringVar(value=fe.range_name(*chart["range"]))
        self.range_entry = ttk.Entry(f, textvariable=self.range_var, width=30)
        self.range_entry.grid(row=1, column=1, sticky="w", padx=6)
        ttk.Label(f, text="Tiêu đề:").grid(row=2, column=0, sticky="w", pady=3)
        self.title_var = tk.StringVar(value=chart.get("title", ""))
        ttk.Entry(f, textvariable=self.title_var, width=30).grid(row=2, column=1, sticky="w", padx=6)
        ttk.Label(f, text="Để trống tiêu đề: biểu đồ một chuỗi lấy tên chuỗi làm tiêu đề.",
                  foreground="#666").grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self.add_buttons()

    def show(self):
        return super().show(self.range_entry)

    def collect(self):
        rect = fe.parse_range(self.range_var.get().replace("$", "").strip())
        if rect is None:
            self.bell()
            return None
        return {"type": next(k for k, label in CHART_KINDS if label == self.kind.get()), "range": rect,
                "title": self.title_var.get().strip()}


class PivotDialog(Dialog):
    """Create PivotTable: the source range, and a new sheet or a place on an existing one. collect()
    gives {"source": (sheet, rect), "new_sheet": bool, "dest": (sheet, row, col)}."""

    def __init__(self, master, wb, sheet, rect):
        super().__init__(master, "Tạo PivotTable")
        self.wb, self.sheet = wb, sheet
        ttk.Label(self.body, text="Chọn dữ liệu cần phân tích (gồm hàng tiêu đề):").pack(anchor="w")
        self.source = tk.StringVar(value=f"{fe.quote_sheet(sheet)}!{fe.range_name(*rect)}")
        self.entry = ttk.Entry(self.body, textvariable=self.source, width=40)
        self.entry.pack(anchor="w", pady=(4, 10))
        ttk.Label(self.body, text="Chọn nơi đặt báo cáo PivotTable:").pack(anchor="w")
        self.new_sheet = tk.BooleanVar(value=True)
        ttk.Radiobutton(self.body, text="Trang tính mới", value=True, variable=self.new_sheet).pack(anchor="w")
        row = ttk.Frame(self.body)
        row.pack(anchor="w", pady=(2, 0))
        ttk.Radiobutton(row, text="Trang tính hiện có, tại ô:", value=False, variable=self.new_sheet).pack(side="left")
        self.dest = tk.StringVar(value="")
        dest = ttk.Entry(row, textvariable=self.dest, width=18)
        dest.pack(side="left", padx=6)
        dest.bind("<FocusIn>", lambda e: self.new_sheet.set(False))
        self.add_buttons()

    def show(self):
        return super().show(self.entry)

    def _place(self, text):
        """(sheet, a, b) from 'B3', 'Sheet2!B3' or 'Sheet2!B3:D9' (a, b: the cell or range tuple)."""
        sheet_text, _, addr = text.strip().rpartition("!")
        sheet = self.sheet
        if sheet_text:
            name = sheet_text[1:-1].replace("''", "'") if sheet_text.startswith("'") else sheet_text
            sheet = self.wb.resolve_sheet(name)
        return sheet, fe.parse_address(addr.replace("$", ""))

    def collect(self):
        sheet, rect = self._place(self.source.get())
        if sheet is None or rect is None or len(rect) != 4 or rect[2] <= rect[0]:
            self.bell()
            return None
        result = {"source": (sheet, rect), "new_sheet": self.new_sheet.get(), "dest": None}
        if not result["new_sheet"]:
            dest_sheet, cell = self._place(self.dest.get())
            if dest_sheet is None or cell is None:
                self.bell()
                return None
            result["dest"] = (dest_sheet, cell[0], cell[1])
        return result


PIVOT_FORMATS = [("Giống dữ liệu nguồn", None), ("Chung", "General"), ("Số: 1,234", "#,##0"),
                 ("Số: 1,234.56", "#,##0.00"), ("Phần trăm: 12%", "0%"), ("Phần trăm: 12.34%", "0.00%")]


class ValueFieldDialog(Dialog):
    """Value Field Settings: caption, summary function, 'show values as' and number format."""

    def __init__(self, master, value, funcs, show_as):
        super().__init__(master, "Cài đặt trường giá trị")
        self.funcs, self.show_as = funcs, show_as
        ttk.Label(self.body, text=f"Trường nguồn:  {value['field']}").pack(anchor="w")
        ttk.Label(self.body, text="Tên tùy chỉnh:").pack(anchor="w", pady=(8, 0))
        self.name = tk.StringVar(value=value.get("name") or "")
        self.entry = ttk.Entry(self.body, textvariable=self.name, width=40)
        self.entry.pack(anchor="w")
        ttk.Label(self.body, text="(để trống: tự đặt theo hàm, ví dụ “Tổng của …”)", foreground="#666").pack(anchor="w")
        ttk.Label(self.body, text="Tóm tắt giá trị theo:").pack(anchor="w", pady=(8, 0))
        self.listbox = tk.Listbox(self.body, height=7, exportselection=False, activestyle="none")
        for key, label in funcs.items():
            self.listbox.insert("end", label)
        index = list(funcs).index(value["func"])
        self.listbox.selection_set(index)
        self.listbox.see(index)
        self.listbox.pack(fill="x")
        ttk.Label(self.body, text="Hiển thị giá trị dưới dạng:").pack(anchor="w", pady=(8, 0))
        self.show_var = tk.StringVar(value=show_as[value.get("show", "normal")])
        ttk.Combobox(self.body, values=list(show_as.values()), textvariable=self.show_var, state="readonly",
                     width=30).pack(anchor="w")
        ttk.Label(self.body, text="Định dạng số:").pack(anchor="w", pady=(8, 0))
        self.num = tk.StringVar(value=next((label for label, f in PIVOT_FORMATS if f == value.get("num")),
                                           value.get("num") or PIVOT_FORMATS[0][0]))
        ttk.Combobox(self.body, values=[label for label, _ in PIVOT_FORMATS], textvariable=self.num,
                     width=30).pack(anchor="w")
        self.add_buttons()

    def show(self):
        return super().show(self.entry)

    def collect(self):
        sel = self.listbox.curselection()
        func = list(self.funcs)[sel[0]] if sel else "sum"
        num = dict(PIVOT_FORMATS).get(self.num.get(), self.num.get().strip() or None)
        return {"name": self.name.get().strip() or None, "func": func,
                "show": next(k for k, label in self.show_as.items() if label == self.show_var.get()), "num": num}


class PivotFieldDialog(Dialog):
    """Settings of a row, column or filter field: item order, date grouping and the items shown."""

    def __init__(self, master, field, items, allowed, descending, group, groups, is_date):
        super().__init__(master, f"Cài đặt trường - {field}")
        self.items, self.groups = items, groups
        self.order = tk.StringVar(value="desc" if descending else "asc")
        row = ttk.Frame(self.body)
        row.pack(anchor="w")
        ttk.Label(row, text="Sắp xếp mục:").pack(side="left")
        ttk.Radiobutton(row, text="A → Z", value="asc", variable=self.order).pack(side="left", padx=6)
        ttk.Radiobutton(row, text="Z → A", value="desc", variable=self.order).pack(side="left")
        self.group = tk.StringVar(value=groups[group])
        if is_date:
            row = ttk.Frame(self.body)
            row.pack(anchor="w", pady=(8, 0))
            ttk.Label(row, text="Nhóm ngày theo:").pack(side="left")
            ttk.Combobox(row, values=list(groups.values()), textvariable=self.group, state="readonly",
                         width=14).pack(side="left", padx=6)
            ttk.Label(self.body, text="(đổi cách nhóm sẽ chọn lại tất cả các mục)", foreground="#666").pack(anchor="w")
        ttk.Label(self.body, text="Hiện các mục:").pack(anchor="w", pady=(8, 2))
        holder = ttk.Frame(self.body)
        holder.pack(fill="both")
        canvas = tk.Canvas(holder, width=280, height=200, bg="#ffffff", highlightthickness=1,
                           highlightbackground="#c8c8c8")
        bar = ttk.Scrollbar(holder, orient="vertical", command=canvas.yview)
        canvas.config(yscrollcommand=bar.set)
        canvas.pack(side="left", fill="both")
        bar.pack(side="right", fill="y")
        inner = tk.Frame(canvas, bg="#ffffff")
        canvas.create_window(0, 0, window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.config(scrollregion=canvas.bbox("all")))
        self.vars = {item: tk.BooleanVar(value=allowed is None or item in allowed) for item in items}
        self.all_var = tk.BooleanVar(value=all(v.get() for v in self.vars.values()))
        tk.Checkbutton(inner, text="(Chọn tất cả)", variable=self.all_var, bg="#ffffff", anchor="w",
                       command=lambda: [v.set(self.all_var.get()) for v in self.vars.values()]).pack(fill="x")
        for item in items[:2000]:
            tk.Checkbutton(inner, text=item, variable=self.vars[item], bg="#ffffff", anchor="w").pack(fill="x")
        self.add_buttons()

    def collect(self):
        allowed = {item for item, var in self.vars.items() if var.get()}
        if not allowed:
            self.bell()
            return None
        group = next(k for k, label in self.groups.items() if label == self.group.get())
        return {"desc": self.order.get() == "desc", "group": group,
                "items": None if len(allowed) == len(self.items) else allowed}


class PivotOptionsDialog(Dialog):
    """PivotTable options: name, layout, subtotals and grand totals."""

    def __init__(self, master, pivot):
        super().__init__(master, "Tùy chọn PivotTable")
        ttk.Label(self.body, text="Tên PivotTable:").grid(row=0, column=0, sticky="w")
        self.name = tk.StringVar(value=pivot["name"])
        self.entry = ttk.Entry(self.body, textvariable=self.name, width=28)
        self.entry.grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(self.body, text="Bố cục báo cáo:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.layout = tk.StringVar(value=pivot["layout"])
        row = ttk.Frame(self.body)
        row.grid(row=1, column=1, sticky="w", pady=(8, 0), padx=6)
        ttk.Radiobutton(row, text="Thu gọn", value="compact", variable=self.layout).pack(side="left")
        ttk.Radiobutton(row, text="Dạng bảng", value="tabular", variable=self.layout).pack(side="left", padx=8)
        self.flags = {}
        for i, (key, text) in enumerate((("subtotals", "Hiện tổng phụ"), ("grand_rows", "Tổng Cuối cho các hàng"),
                                         ("grand_cols", "Tổng Cuối cho các cột"))):
            self.flags[key] = tk.BooleanVar(value=pivot[key])
            ttk.Checkbutton(self.body, text=text, variable=self.flags[key]).grid(row=2 + i, column=0, columnspan=2,
                                                                                 sticky="w", pady=(6 if i == 0 else 0, 0))
        self.add_buttons()

    def show(self):
        return super().show(self.entry)

    def collect(self):
        if not self.name.get().strip():
            self.bell()
            return None
        return {"name": self.name.get().strip(), "layout": self.layout.get(),
                **{key: var.get() for key, var in self.flags.items()}}
