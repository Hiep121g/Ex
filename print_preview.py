"""File > Print for PyExcel: Excel's print page - the settings on the left, the pages on the right.

Pages are drawn by printing.render_page, the code that writes the PDF, onto a Tk canvas through
CanvasSurface, so what the preview shows is what gets printed.
"""

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

import pdf
import printing
from dialogs import PAPER_LABELS
from ribbon import RibbonButton

SCOPES = [("sheet", "In trang tính hiện hành"), ("book", "In toàn bộ sổ làm việc"), ("selection", "In vùng chọn")]
ORIENTS = [("portrait", "Hướng dọc"), ("landscape", "Hướng ngang")]
MARGIN_PRESETS = [("normal", "Lề thường"), ("wide", "Lề rộng"), ("narrow", "Lề hẹp")]
FITS = [(None, "Không co giãn"), ("width", "Vừa tất cả các cột trên một trang"),
        ("page", "Vừa trang tính trên một trang")]
CUSTOM_MARGINS, CUSTOM_SCALE = "Lề tùy chỉnh...", "Tùy chọn co giãn tùy chỉnh..."
PREVIEW_BG = "#e6e6e6"


def _clip_segment(a, b, box):
    """The part of segment a-b inside box (x0, y0, x1, y1) as (x0, y0, x1, y1), or None (Liang-Barsky)."""
    (x0, y0), (x1, y1) = a, b
    dx, dy = x1 - x0, y1 - y0
    t0, t1 = 0.0, 1.0
    for p, q in ((-dx, x0 - box[0]), (dx, box[2] - x0), (-dy, y0 - box[1]), (dy, box[3] - y0)):
        if p == 0:
            if q < 0:
                return None
        elif p < 0:
            t0 = max(t0, q / p)
        else:
            t1 = min(t1, q / p)
    if t0 > t1:
        return None
    return x0 + t0 * dx, y0 + t0 * dy, x0 + t1 * dx, y0 + t1 * dy


class CanvasSurface:
    """printing's drawing surface over a Tk canvas: points scaled by `zoom` and moved to (ox, oy).

    Text is measured with the PDF's TrueType metrics, so lines and pages break exactly as in the PDF,
    and drawn in the matching Tk font centred on that measured width. A canvas cannot clip, so inside
    clip() text is cut to whole characters and lines and boxes to the clip box."""

    def __init__(self, canvas, library):
        self.cv, self.library = canvas, library
        self.doc = pdf.PdfDocument(library)  # only for measuring
        self.z, self.ox, self.oy = 1.0, 0.0, 0.0
        self._ttf, self._tk, self._box = {}, {}, None

    def place(self, zoom, ox, oy):
        self.z, self.ox, self.oy = zoom, ox, oy

    def _xy(self, *coords):
        return [self.ox + v * self.z if i % 2 == 0 else self.oy + v * self.z for i, v in enumerate(coords)]

    def _w(self, width):
        return max(1, round(width * self.z))

    def font(self, spec):
        font = self._ttf.get(spec)
        if font is None:
            font = self._ttf[spec] = self.library.find(*spec) or self.library.find("Arial")
        return font

    def measure(self, text, spec, size):
        return self.doc.measure(text, self.font(spec), size)

    def ascent(self, spec, size):
        return self.font(spec).ascent * size / 1000

    def descent(self, spec, size):
        return -self.font(spec).descent * size / 1000

    def _tk_font(self, spec, size):
        key = (spec, max(1, round(size * self.z)))
        font = self._tk.get(key)
        if font is None:
            ttf = self.font(spec)
            font = self._tk[key] = tkfont.Font(root=self.cv, family=ttf.family if ttf else spec[0], size=-key[1],
                                               weight="bold" if spec[1] else "normal",
                                               slant="italic" if spec[2] else "roman")
        return font

    def text(self, x, y, text, spec, size, color):
        if self._box:
            x0, y0, x1, y1 = self._box
            if y - self.ascent(spec, size) >= y1 or y + self.descent(spec, size) <= y0:
                return
            widths = [self.measure(ch, spec, size) for ch in text]
            start, end, left = 0, len(text), x
            while start < end and left < x0 - 0.5:
                left += widths[start]
                start += 1
            right = left + sum(widths[start:end])
            while end > start and right > x1 + 0.5:
                end -= 1
                right -= widths[end]
            text = text[start:end]
        else:
            left, right = x, x + self.measure(text, spec, size)
        if not text.strip():
            return
        font = self._tk_font(spec, size)
        cx = self.ox + (left + right) / 2 * self.z
        self.cv.create_text(cx - font.measure(text) / 2, self.oy + y * self.z - font.metrics("ascent"), text=text,
                            font=font, fill=color, anchor="nw")

    def rect(self, x0, y0, x1, y1, fill=None, stroke=None, width=1.0):
        if not (fill or stroke):
            return
        if self._box:
            bx0, by0, bx1, by1 = self._box
            x0, y0, x1, y1 = max(x0, bx0), max(y0, by0), min(x1, bx1), min(y1, by1)
            if x0 >= x1 or y0 >= y1:
                return
        self.cv.create_rectangle(*self._xy(x0, y0, x1, y1), fill=fill or "", outline=stroke or "",
                                 width=self._w(width) if stroke else 0)

    def polyline(self, points, color="#000000", width=1.0, dash=None, closed=False, fill=None):
        if len(points) < 2:
            return
        if closed or fill:
            self.polygon(points, fill=fill, stroke=color, width=width)
            return
        opts = {"fill": color, "width": self._w(width)}
        if dash:
            opts["dash"] = tuple(min(255, max(1, round(d * self.z))) for d in dash)
        if self._box:
            for a, b in zip(points, points[1:]):
                seg = _clip_segment(a, b, self._box)
                if seg:
                    self.cv.create_line(*self._xy(*seg), **opts)
        else:
            self.cv.create_line(*self._xy(*[v for p in points for v in p]), **opts)

    def line(self, x0, y0, x1, y1, color="#000000", width=1.0, dash=None):
        self.polyline([(x0, y0), (x1, y1)], color, width, dash)

    def _outside(self, x0, y0, x1, y1):
        b = self._box
        return b is not None and (max(x0, x1) < b[0] or min(x0, x1) > b[2] or max(y0, y1) < b[1] or min(y0, y1) > b[3])

    def polygon(self, points, fill=None, stroke=None, width=1.0):
        xs, ys = [p[0] for p in points], [p[1] for p in points]
        if len(points) < 3 or self._outside(min(xs), min(ys), max(xs), max(ys)):
            return
        self.cv.create_polygon(*self._xy(*[v for p in points for v in p]), fill=fill or "", outline=stroke or "",
                               width=self._w(width) if stroke else 0)

    def ellipse(self, x0, y0, x1, y1, fill=None, stroke=None, width=1.0):
        if not self._outside(x0, y0, x1, y1):
            self.cv.create_oval(*self._xy(x0, y0, x1, y1), fill=fill or "", outline=stroke or "",
                                width=self._w(width) if stroke else 0)

    def arc(self, x0, y0, x1, y1, start, extent, fill=None, stroke=None, width=1.0, style="pieslice"):
        if self._outside(x0, y0, x1, y1):
            return
        opts = {"start": start, "extent": max(-359.99, min(359.99, extent)), "style": style,
                "outline": stroke or "", "width": self._w(width) if stroke else 0}
        if style != "arc":
            opts["fill"] = fill or ""
        self.cv.create_arc(*self._xy(x0, y0, x1, y1), **opts)

    def clip(self, x0, y0, x1, y1):
        self._box = (x0, y0, x1, y1)

    def unclip(self):
        self._box = None


class PrintPreview(tk.Toplevel):
    """The print page: print / export buttons and page settings beside a preview of the pages.

    `app` is the SpreadsheetApp; settings changed here go to the sheet of the page on show, with undo."""

    def __init__(self, app, scope="sheet"):
        super().__init__(app.root)
        self.withdraw()
        self.app, self.ui = app, app.ui
        self.title("In")
        self.transient(app.root)
        self.configure(bg=PREVIEW_BG)
        self.pages, self.index = [], 0
        self.surface = CanvasSurface(None, app.font_library)
        self._build(scope)
        top = app.root
        w, h = max(900, round(top.winfo_width() * 0.92)), max(600, round(top.winfo_height() * 0.92))
        x = top.winfo_rootx() + (top.winfo_width() - w) // 2
        y = top.winfo_rooty() + (top.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")
        self.rebuild()
        self.deiconify()
        self.grab_set()
        self.focus_set()

    # ------------------------------------------------------------------ widgets

    def _build(self, scope):
        ui, px = self.ui, self.ui.px
        side = tk.Frame(self, bg="#ffffff", padx=px(16), pady=px(12))
        side.pack(side="left", fill="y")
        tk.Label(side, text="In", bg="#ffffff", fg="#262626",
                 font=tkfont.Font(root=self, family=ui.family, size=-px(26))).pack(anchor="w")
        row = tk.Frame(side, bg="#ffffff")
        row.pack(anchor="w", pady=(px(8), px(14)))
        for text, icon, command, tip in (
                ("In", "print", self.print_, "Gửi tới máy in (Ctrl+P)"),
                ("Xuất PDF", "pdf", self.export, "Lưu các trang này thành tệp PDF")):
            RibbonButton(row, ui, text=text, icon=icon, command=command, large=True, tooltip=tip, bg="#ffffff",
                         hover_bg="#e1e1e1").pack(side="left", padx=(0, px(6)))
        tk.Label(side, text="Cài đặt", bg="#ffffff", fg="#262626",
                 font=tkfont.Font(root=self, family=ui.family, size=-px(15), weight="bold")).pack(anchor="w",
                                                                                                pady=(0, px(6)))
        self.scope_var = tk.StringVar(self, value=dict(SCOPES)[scope])
        self.orient_var, self.paper_var = tk.StringVar(self), tk.StringVar(self)
        self.margin_var, self.fit_var = tk.StringVar(self), tk.StringVar(self)
        self.boxes = {}
        for key, var, values in (("scope", self.scope_var, [label for _, label in SCOPES]),
                                 ("orient", self.orient_var, [label for _, label in ORIENTS]),
                                 ("paper", self.paper_var, list(PAPER_LABELS.values())),
                                 ("margins", self.margin_var, [label for _, label in MARGIN_PRESETS] + [CUSTOM_MARGINS]),
                                 ("fit", self.fit_var, [label for _, label in FITS] + [CUSTOM_SCALE])):
            box = ttk.Combobox(side, textvariable=var, values=values, state="readonly", width=34, font=ui.font)
            box.pack(anchor="w", fill="x", pady=px(3))
            box.bind("<<ComboboxSelected>>", lambda e, k=key: self._setting(k))
            self.boxes[key] = box
        link = tk.Label(side, text="Thiết lập Trang...", bg="#ffffff", fg="#2b579a", cursor="hand2",
                        font=tkfont.Font(root=self, family=ui.family, size=-px(12), underline=True))
        link.pack(anchor="e", pady=(px(6), 0))
        link.bind("<Button-1>", lambda e: self.page_setup())
        self.info = tk.Label(side, bg="#ffffff", fg="#5f5f5f", font=ui.small_font, justify="left", anchor="w")
        self.info.pack(anchor="w", fill="x", pady=(px(18), 0))

        main = tk.Frame(self, bg=PREVIEW_BG)
        main.pack(side="left", fill="both", expand=True)
        nav = tk.Frame(main, bg=PREVIEW_BG)
        nav.pack(side="bottom", fill="x", padx=px(10), pady=px(6))
        ttk.Button(nav, text="◀", width=3, command=lambda: self.go(self.index - 1)).pack(side="left")
        self.page_var = tk.StringVar(self)
        entry = ttk.Entry(nav, textvariable=self.page_var, width=4, justify="center", font=ui.font)
        entry.pack(side="left", padx=px(4))
        entry.bind("<Return>", lambda e: self._go_typed())
        self.count_label = tk.Label(nav, bg=PREVIEW_BG, font=ui.font)
        self.count_label.pack(side="left")
        ttk.Button(nav, text="▶", width=3, command=lambda: self.go(self.index + 1)).pack(side="left", padx=px(4))
        self.actual = tk.BooleanVar(self, value=False)
        ttk.Checkbutton(nav, text="Phóng to trang", variable=self.actual, command=self.draw).pack(side="right")
        self.canvas = tk.Canvas(main, bg=PREVIEW_BG, highlightthickness=0)
        ys = ttk.Scrollbar(main, orient="vertical", command=self.canvas.yview)
        xs = ttk.Scrollbar(main, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        xs.pack(side="bottom", fill="x")
        ys.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.surface.cv = self.canvas
        self.canvas.bind("<Configure>", lambda e: self.draw())
        for seq, step in (("<MouseWheel>", None), ("<Button-4>", -1), ("<Button-5>", 1)):
            self.canvas.bind(seq, lambda e, s=step: self._wheel(e, s))
        for seq, func in (("<Prior>", lambda: self.go(self.index - 1)), ("<Next>", lambda: self.go(self.index + 1)),
                          ("<Control-Home>", lambda: self.go(0)),
                          ("<Control-End>", lambda: self.go(len(self.pages) - 1)),
                          ("<Escape>", self.destroy), ("<Control-p>", self.print_), ("<Control-P>", self.print_)):
            self.bind(seq, lambda e, f=func: (f(), "break")[1])
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    # ------------------------------------------------------------------ pages

    def scope(self):
        return next(key for key, label in SCOPES if label == self.scope_var.get())

    def rebuild(self):
        """Lay out the pages again (after a change of scope or settings)."""
        sheet_now = self.pages[self.index][0].sheet if self.pages else None
        self.pages = []
        for sheet, setup in self.app.print_jobs(self.scope()):
            layout = printing.Layout(self.app.wb, sheet, setup)
            self.pages += [(layout, i) for i in range(len(layout.pages))]
        self.index = next((i for i, (lay, _) in enumerate(self.pages) if lay.sheet == sheet_now),
                          min(self.index, max(0, len(self.pages) - 1)))
        self.draw()

    def go(self, index):
        if self.pages and 0 <= index < len(self.pages) and index != self.index:
            self.index = index
            self.canvas.yview_moveto(0)
            self.draw()

    def _go_typed(self):
        try:
            self.go(int(self.page_var.get()) - 1)
        except ValueError:
            pass
        self.draw()

    def _wheel(self, event, step):
        step = step or (-1 if event.delta > 0 else 1)
        first, last = self.canvas.yview()
        if self.actual.get() and (step < 0 < first or step > 0 and last < 1):
            self.canvas.yview_scroll(step * 3, "units")
        else:
            self.go(self.index + step)

    def draw(self):
        cv = self.canvas
        cv.delete("all")
        W, H = cv.winfo_width(), cv.winfo_height()
        self._sync_settings()
        total = len(self.pages)
        self.page_var.set(str(self.index + 1 if total else 0))
        self.count_label.configure(text=f"của {total}")
        if W < 50 or H < 50:
            return
        if not total:
            cv.configure(scrollregion=(0, 0, W, H))
            cv.create_text(W / 2, H / 2, text="Không tìm thấy gì để in.", font=self.ui.font, fill="#5f5f5f")
            return
        layout, index = self.pages[self.index]
        pw, ph = layout.paper
        pad = self.ui.px(18)
        if self.actual.get():
            zoom = self.app.scale * 96 / 72  # paper at its real size
        else:
            zoom = min((W - 2 * pad) / pw, (H - 2 * pad) / ph)
        w, h = pw * zoom, ph * zoom
        ox, oy = max(pad, (W - w) / 2), max(pad, (H - h) / 2)
        cv.configure(scrollregion=(0, 0, max(W, w + 2 * pad), max(H, h + 2 * pad)))
        cv.create_rectangle(ox + 3, oy + 3, ox + w + 3, oy + h + 3, fill="#b4b4b4", outline="")
        cv.create_rectangle(ox, oy, ox + w, oy + h, fill="#ffffff", outline="#a0a0a0")
        self.surface.place(zoom, ox, oy)
        printing.render_page(self.surface, layout, index, self.app.file_label())

    # ------------------------------------------------------------------ settings

    def _sheet(self):
        return self.pages[self.index][0].sheet if self.pages else self.app.sheet

    def _sync_settings(self):
        setup = printing.setup_of(self.app.wb.sheets[self._sheet()])
        self.orient_var.set(dict(ORIENTS)[setup["orientation"]])
        self.paper_var.set(PAPER_LABELS.get(setup["paper"], PAPER_LABELS["A4"]))
        preset = next((label for key, label in MARGIN_PRESETS if printing.MARGINS[key] == list(setup["margins"])),
                      "Lề tùy chỉnh")
        self.margin_var.set(preset)
        if setup["fit"] or setup["scale"] == 100:
            self.fit_var.set(dict(FITS)[setup["fit"]])
        else:
            self.fit_var.set(f"Tỷ lệ {setup['scale']}%")
        sheets = {lay.sheet for lay, _ in self.pages}
        self.info.configure(text=f"Trang tính: {self._sheet()}" + (f"\n({len(sheets)} trang tính)" if len(sheets) > 1 else "")
                            + "\nCài đặt bên trên áp dụng cho trang tính\nđang xem.")

    def _setting(self, key):
        sheet = self._sheet()
        if key == "scope":
            self.rebuild()
            return
        if key == "orient":
            changes = {"orientation": next(k for k, label in ORIENTS if label == self.orient_var.get())}
        elif key == "paper":
            changes = {"paper": next(k for k, label in PAPER_LABELS.items() if label == self.paper_var.get())}
        elif key == "margins":
            if self.margin_var.get() == CUSTOM_MARGINS:
                self.page_setup("margins")
                return
            changes = {"margins": printing.MARGINS[next(k for k, label in MARGIN_PRESETS
                                                         if label == self.margin_var.get())]}
        else:
            if self.fit_var.get() == CUSTOM_SCALE:
                self.page_setup("page")
                return
            changes = {"fit": next(k for k, label in FITS if label == self.fit_var.get()), "scale": 100}
        self.app.update_page_setup(changes, sheet)
        self.rebuild()

    def page_setup(self, tab="page"):
        self.app.page_setup_dialog(tab, self._sheet(), parent=self)
        self.rebuild()

    # ------------------------------------------------------------------ output

    def print_(self):
        if self.pages and self.app.print_now(self.scope(), parent=self):
            self.destroy()

    def export(self):
        if self.pages:
            self.app.export_pdf(self.scope(), parent=self)
