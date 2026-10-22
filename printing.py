"""Printing for PyExcel: page setup, splitting a sheet into pages, and drawing a page.

The same drawing code serves the PDF export (PdfSurface over pdf.Page) and the print preview
(CanvasSurface over a Tk canvas); both measure text with the TrueType metrics of ttf.py, so the preview
breaks lines and pages exactly like the PDF. Units are points (1/72 inch); sheet sizes stored in pixels
at 96 DPI are converted with PT (0.75).
"""

import datetime as dt
import os
from operator import itemgetter

import base64

import charts
import formula_engine as fe
import images
import pdf

PT = 0.75  # points per pixel at 96 DPI
PAPERS = {"A4": (595.28, 841.89), "A3": (841.89, 1190.55), "A5": (419.53, 595.28), "Letter": (612.0, 792.0),
          "Legal": (612.0, 1008.0)}
XL_PAPERS = {"A4": 9, "A3": 8, "A5": 11, "Letter": 1, "Legal": 5}
# top, bottom, left, right, header, footer in inches (Excel's presets)
MARGINS = {"normal": [0.75, 0.75, 0.7, 0.7, 0.3, 0.3], "wide": [1.0, 1.0, 1.0, 1.0, 0.5, 0.5],
           "narrow": [0.75, 0.75, 0.25, 0.25, 0.3, 0.3]}
DEFAULT_SETUP = {"paper": "A4", "orientation": "portrait", "margins": MARGINS["normal"], "scale": 100,
                 "fit": None,        # None, "width" (all columns on one page wide) or "page" (whole sheet on one page)
                 "area": None,       # (r1, c1, r2, c2) to print; None: the used part of the sheet
                 "titles": None,     # (r1, r2): rows repeated at the top of every page
                 "gridlines": False, "center_h": False, "center_v": False,
                 "header": "", "footer": "",  # codes: &P page, &N pages, &D date, &T time, &F file, &A sheet
                 "breaks": [], "col_breaks": []}  # rows / columns that start a new page
DEFAULT_FONT, DEFAULT_SIZE = "Calibri", 11
BORDER_WIDTHS = {"thin": 0.5, "medium": 1.0, "thick": 1.75, "dashed": 0.5, "dotted": 0.5, "double": 0.5}


def setup_of(sheet):
    """A sheet's page setup with every key present."""
    return dict(DEFAULT_SETUP, **(sheet.page_setup or {}))


def paper_size(setup):
    w, h = PAPERS.get(setup["paper"], PAPERS["A4"])
    return (h, w) if setup["orientation"] == "landscape" else (w, h)


def header_text(text, page, pages, sheet, file_name):
    now = dt.datetime.now()
    for code, value in (("&P", str(page)), ("&N", str(pages)), ("&D", now.strftime("%d/%m/%Y")),
                        ("&T", now.strftime("%H:%M")), ("&F", file_name or "Book1"), ("&A", sheet)):
        text = text.replace(code, value)
    return text


def title_rows_text(titles):
    """(r1, r2) -> '$2:$3' as Excel writes print titles."""
    return f"${titles[0] + 1}:${titles[1] + 1}" if titles else ""


def parse_title_rows(text):
    """'$2:$3', '2:3' or '2' -> (1, 2); '' -> None; anything else -> False."""
    parts = text.replace("$", "").replace(" ", "").split(":")
    if parts == [""]:
        return None
    if len(parts) > 2 or not all(p.isdigit() and int(p) >= 1 for p in parts):
        return False
    a, b = int(parts[0]), int(parts[-1])
    return min(a, b) - 1, max(a, b) - 1


def wrap(text, width, measure):
    """Lines of text no wider than width (words longer than a line are cut)."""
    lines = []
    for para in text.split("\n"):
        line = ""
        for word in para.split(" "):
            candidate = f"{line} {word}" if line else word
            if measure(candidate) <= width or not line and len(word) <= 1:
                line = candidate
                continue
            if line:
                lines.append(line)
            while len(word) > 1 and measure(word) > width:
                cut = len(word) - 1
                while cut > 1 and measure(word[:cut]) > width:
                    cut -= 1
                lines.append(word[:cut])
                word = word[cut:]
            line = word
        lines.append(line)
    return lines


# ------------------------------------------------------------------ pages

def used_rect(wb, sheet):
    """The block that holds anything printable: values, fills, borders, merges and charts."""
    sh = wb.sheets[sheet]
    rects = [tuple(m) for m in sh.merges]
    if sh.cells:  # min/max over the keys in C: a big sheet has millions of cells
        rows, cols = sh.cells.extent()
        rects.append((min(sh.cells)[0], min(map(itemgetter(1), sh.cells)), rows - 1, cols - 1))
    printable = {}  # formats are shared between cells: judge each one once
    keys = [k for k, f in sh.formats.items()
            if printable.setdefault(id(f), bool(f.get("bg") or any(f.get(b) for b in fe.BORDER_KEYS)))]
    if keys:
        rects.append((min(keys)[0], min(map(itemgetter(1), keys)), max(keys)[0], max(map(itemgetter(1), keys))))
    for chart in sh.charts:
        r, c = chart["anchor"][:2]
        w, h = chart["size"]
        c2, x = c, sh.col_widths.get(c, fe.DEFAULT_COL_WIDTH) - chart["anchor"][2]
        while x < w and c2 < fe.MAX_COLS - 1:
            c2 += 1
            x += sh.col_widths.get(c2, fe.DEFAULT_COL_WIDTH)
        r2, y = r, sh.row_heights.get(r, fe.DEFAULT_ROW_HEIGHT) - chart["anchor"][3]
        while y < h and r2 < fe.MAX_ROWS - 1:
            r2 += 1
            y += sh.row_heights.get(r2, fe.DEFAULT_ROW_HEIGHT)
        rects.append((r, c, r2, c2))
    if not rects:
        return None
    return (min(r[0] for r in rects), min(r[1] for r in rects), max(r[2] for r in rects), max(r[3] for r in rects))


class Layout:
    """How a sheet falls onto pages: pages [(rows, cols)] - rows starting with the repeated title rows
    on the pages that do not show them anyway - the scale, and the paper geometry."""

    def __init__(self, wb, sheet, setup):
        self.wb, self.sheet, self.setup = wb, sheet, setup
        sh = wb.sheets[sheet]
        self.paper = paper_size(setup)
        top, bottom, left, right = (m * 72 for m in setup["margins"][:4])
        self.body = (left, top, self.paper[0] - right, self.paper[1] - bottom)
        self.rect = setup["area"] or used_rect(wb, sheet)
        self.pages, self.leads, self.scale, self.title_rows = [], [], 1.0, []
        if self.rect is None:
            return
        r1, c1, r2, c2 = self.rect
        titles = setup["titles"]
        self.title_rows = [r for r in range(titles[0], titles[1] + 1) if not sh.row_hidden(r)] if titles else []
        rows = [r for r in range(r1, r2 + 1) if not sh.row_hidden(r)]
        cols = [c for c in range(c1, c2 + 1) if c not in sh.hidden_cols]
        if not cols or not rows:
            return
        self.height = {r: sh.row_heights.get(r, fe.DEFAULT_ROW_HEIGHT) * PT for r in rows + self.title_rows}
        self.width = {c: sh.col_widths.get(c, fe.DEFAULT_COL_WIDTH) * PT for c in cols}
        body_w, body_h = self.body[2] - self.body[0], self.body[3] - self.body[1]
        title_h = sum(self.height[r] for r in self.title_rows) if self.title_rows and rows[0] > self.title_rows[-1] else 0
        total_w, total_h = sum(self.width.values()), sum(self.height[r] for r in rows) + title_h
        if setup["fit"] == "width":
            self.scale = min(1.0, body_w / total_w)
        elif setup["fit"] == "page":
            self.scale = min(1.0, body_w / total_w, body_h / max(total_h, 1))
        else:
            self.scale = max(0.1, min(4.0, (setup["scale"] or 100) / 100))
        s = self.scale
        col_breaks = set(setup["col_breaks"])
        col_pages, line, used = [], [], 0.0
        for c in cols:
            if line and (used + self.width[c] * s > body_w + 0.01 or c in col_breaks):
                col_pages.append(line)
                line, used = [], 0.0
            line.append(c)
            used += self.width[c] * s
        col_pages.append(line)
        breaks = set(setup["breaks"])
        last_title = self.title_rows[-1] if self.title_rows else -1
        row_pages, line, lead, used = [], [], 0, 0.0
        for r in rows:
            if len(line) > lead and (used + self.height[r] * s > body_h + 0.01 or r in breaks):
                row_pages.append((line, lead))
                line, lead, used = [], 0, 0.0
            if not line and self.title_rows and r > last_title:  # repeat the titles unless the page shows them
                line, lead = list(self.title_rows), len(self.title_rows)
                used = sum(self.height[t] for t in line) * s
            line.append(r)
            used += self.height[r] * s
        row_pages.append((line, lead))
        self.pages = [(rws, cls) for cls in col_pages for rws, _ in row_pages]  # down, then over (Excel's order)
        self.leads = [n for _ in col_pages for _, n in row_pages]  # repeated title rows heading each page

    def breaks(self):
        """(rows, columns) that start a new page, for drawing page-break lines on the grid."""
        if not self.pages:
            return set(), set()
        rows = {rws[n] for (rws, _), n in zip(self.pages, self.leads)} - {self.pages[0][0][self.leads[0]]}
        cols = {cls[0] for _, cls in self.pages} - {self.pages[0][1][0]}
        return rows, cols


# ------------------------------------------------------------------ surfaces

class PdfSurface:
    """Drawing on a pdf.Page; fonts are (family, bold, italic)."""

    def __init__(self, page, library):
        self.page, self.library = page, library
        self._fonts = {}

    def font(self, spec):
        font = self._fonts.get(spec)
        if font is None:
            font = self._fonts[spec] = self.library.find(*spec) or self.library.find("Arial")
        return font

    def measure(self, text, spec, size):
        return self.page.doc.measure(text, self.font(spec), size)

    def ascent(self, spec, size):
        return self.font(spec).ascent * size / 1000

    def descent(self, spec, size):
        return -self.font(spec).descent * size / 1000

    def text(self, x, y, text, spec, size, color):
        self.page.text(x, y, text, self.font(spec), size, color)

    def picture(self, x0, y0, x1, y1, picture):
        doc = self.page.doc
        img = doc.images.get(picture["id"])
        if img is None:
            info = images.pdf_image(base64.b64decode(picture["image"]), picture.get("format", "png"))
            if info is None:
                self.page.rect(x0, y0, x1, y1, fill="#f2f2f2", stroke="#a6a6a6", width=0.5)
                return
            img = doc.add_image(picture["id"], info)
        self.page.image(x0, y0, x1, y1, img)

    def __getattr__(self, name):  # rect, line, polyline, polygon, ellipse, arc, clip, unclip
        return getattr(self.page, name)


class ChartCanvas:
    """The bit of the Tk canvas API that charts.draw_chart uses, drawing onto a surface instead."""

    class Font:
        def __init__(self, surface, spec, size):
            self.surface, self.spec, self.size = surface, spec, size

        def measure(self, text):
            return self.surface.measure(text, self.spec, self.size)

        def metrics(self, what):
            return (self.surface.ascent(self.spec, self.size) + self.surface.descent(self.spec, self.size)) * 1.15

    def __init__(self, surface):
        self.s = surface
        self.boxes = {}

    def font(self, size, bold=False):
        return self.Font(self.s, (DEFAULT_FONT, bold, False), size)

    def create_rectangle(self, x0, y0, x1, y1, fill="", outline="#000000", width=1, **kw):
        self.s.rect(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1), fill=fill or None,
                    stroke=outline or None, width=width * PT)

    def create_line(self, *coords, fill="#000000", width=1, **kw):
        points = list(zip(coords[0::2], coords[1::2]))
        self.s.polyline(points, fill or "#000000", width * PT)

    def create_polygon(self, *coords, fill="", outline="", **kw):
        self.s.polygon(list(zip(coords[0::2], coords[1::2])), fill=fill or None, stroke=outline or None, width=0.5)

    def create_oval(self, x0, y0, x1, y1, fill="", outline="", **kw):
        self.s.ellipse(x0, y0, x1, y1, fill=fill or None, stroke=outline or None, width=0.5)

    def create_arc(self, x0, y0, x1, y1, start=0, extent=90, style="pieslice", fill="", outline="", width=1, **kw):
        self.s.arc(x0, y0, x1, y1, start, extent, fill=fill or None, stroke=outline or None, width=width * PT,
                   style=style)

    def create_text(self, x, y, text="", anchor="center", font=None, fill="#000000", width=None, **kw):
        lines = wrap(text, width, font.measure) if width else text.split("\n")
        ls = font.metrics("linespace")
        w, h = max((font.measure(line) for line in lines), default=0), ls * len(lines)
        x0 = x - w / 2 if anchor in ("center", "n", "s") else x - w if "e" in anchor else x
        y0 = y - h / 2 if anchor in ("center", "w", "e") else y - h if "s" in anchor else y
        for i, line in enumerate(lines):
            lx = x0 + (w - font.measure(line)) / 2 if anchor in ("center", "n", "s") else x0
            self.s.text(lx, y0 + i * ls + font.surface.ascent(font.spec, font.size), line, font.spec, font.size,
                        fill or "#000000")
        self.boxes[len(self.boxes) + 1] = (x0, y0, x0 + w, y0 + h)
        return len(self.boxes)

    def bbox(self, item):
        return self.boxes[item]


# ------------------------------------------------------------------ drawing a page

def render_page(surface, layout, index, file_name=""):
    """Draw page `index` of the layout (header, cells, charts, footer) on the surface."""
    wb, sheet, setup = layout.wb, layout.sheet, layout.setup
    sh = wb.sheets[sheet]
    all_rows, cols = layout.pages[index]
    rows = all_rows[layout.leads[index]:]  # the page's own rows, below any repeated titles
    s = layout.scale
    bx0, by0, bx1, by1 = layout.body
    width = sum(layout.width[c] for c in cols) * s
    height = sum(layout.height[r] for r in all_rows) * s
    x0 = bx0 + ((bx1 - bx0 - width) / 2 if setup["center_h"] else 0)
    y0 = by0 + ((by1 - by0 - height) / 2 if setup["center_v"] else 0)
    xs, x = {}, x0
    for c in cols:
        xs[c] = (x, x + layout.width[c] * s)
        x += layout.width[c] * s
    ys, y = {}, y0
    for r in all_rows:
        ys[r] = (y, y + layout.height[r] * s)
        y += layout.height[r] * s

    def box(r, c):
        m = sh.merge_at(r, c)
        if not m:
            return xs[c][0], ys[r][0], xs[c][1], ys[r][1]
        mr = [rr for rr in all_rows if m[0] <= rr <= m[2]]
        mc = [cc for cc in cols if m[1] <= cc <= m[3]]
        return xs[mc[0]][0], ys[mr[0]][0], xs[mc[-1]][1], ys[mr[-1]][1]

    cells = []
    for r in all_rows:
        for c in cols:
            m = sh.merge_at(r, c)
            if m and (r, c) != (next(rr for rr in all_rows if m[0] <= rr <= m[2]),
                                next(cc for cc in cols if m[1] <= cc <= m[3])):
                continue  # inside a merged block: drawn with the block's first cell on this page
            cells.append((r, c, m))
    styles = {(r, c): wb.cond_format(sheet, r, c) or {} for r, c, _ in cells} if sh.cond_formats else {}

    for r, c, m in cells:  # fills
        fmt = sh.formats.get((r, c)) or {}
        style = styles.get((r, c), {})
        bg = style.get("bg") or fmt.get("bg")
        if bg:
            surface.rect(*box(r, c), fill=bg)
        bar = style.get("databar")
        if bar:
            a, b, cc, d = box(r, c)
            surface.rect(a + 1, b + 1.5, a + 1 + (cc - a - 3) * bar[0], d - 1.5, fill=_mix(bar[1], "#ffffff", 0.45),
                         stroke=bar[1], width=0.4)
    if setup["gridlines"]:
        for r, c, m in cells:
            surface.rect(*box(r, c), stroke="#bfbfbf", width=0.25)

    def font_of(fmt, style):
        size = (fmt.get("size") or DEFAULT_SIZE) * s
        spec = (fmt.get("font") or DEFAULT_FONT, bool(style.get("bold", fmt.get("bold"))),
                bool(style.get("italic", fmt.get("italic"))))
        return spec, size

    for r, c, m in cells:  # text
        value, numfmt = wb.cell_display(sheet, r, c)
        if value is None or value == "":
            continue
        fmt = sh.formats.get((r, c)) or {}
        style = styles.get((r, c), {})
        spec, size = font_of(fmt, style)
        a, b, cc, d = box(r, c)
        pad = 2 * s
        indent = (fmt.get("indent") or 0) * 9 * PT * s
        if isinstance(value, (bool, fe.CellError)):
            text, align = fe.format_value(value, numfmt), "center"
        elif fe.is_number(value):
            text, align = fe.format_value(value, numfmt), "right"
        else:
            text, align = str(value), "left"
        align = fmt.get("align") or align
        avail = cc - a - 2 * pad - indent
        if fe.is_number(value) and surface.measure(text, spec, size) > avail:
            text = "#" * max(1, int(avail // max(0.1, surface.measure("#", spec, size))))
        clip_right = cc
        if not m and align == "left" and isinstance(value, str) and not fmt.get("wrap"):
            k = cols.index(c)  # run on over empty cells to the right, as on screen
            while (surface.measure(text, spec, size) > clip_right - a - 2 * pad - indent and k + 1 < len(cols)
                   and (r, cols[k + 1]) not in sh.cells and not sh.merge_at(r, cols[k + 1])):
                k += 1
                clip_right = xs[cols[k]][1]
        if fmt.get("wrap") or "\n" in text:
            lines = wrap(text, max(4.0, avail), lambda t: surface.measure(t, spec, size))
        else:
            lines = [text.replace("\n", " ")]
        asc, desc = surface.ascent(spec, size), surface.descent(spec, size)
        ls = (asc + desc) * 1.1
        block = ls * len(lines)
        valign = fmt.get("valign") or "bottom"
        top = {"top": b + 1, "center": (b + d - block) / 2}.get(valign, d - 1.5 - block)
        color = style.get("fg") or fmt.get("fg") or "#000000"
        surface.clip(a, b, clip_right, d)
        for i, line in enumerate(lines):
            w = surface.measure(line, spec, size)
            lx = {"center": (a + cc - w) / 2, "right": cc - pad - indent - w}.get(align, a + pad + indent)
            base = top + i * ls + asc
            surface.text(lx, base, line, spec, size, color)
            if fmt.get("underline"):
                surface.line(lx, base + desc * 0.4, lx + w, base + desc * 0.4, color, max(0.4, size / 16))
                if fmt.get("underline") == "double":
                    surface.line(lx, base + desc * 0.9, lx + w, base + desc * 0.9, color, max(0.4, size / 16))
            if fmt.get("strike"):
                surface.line(lx, base - asc * 0.3, lx + w, base - asc * 0.3, color, max(0.4, size / 16))
        surface.unclip()

    for r, c, m in cells:  # borders
        fmt = sh.formats.get((r, c)) or {}
        a, b, cc, d = box(r, c)
        for key, (p0, p1) in (("bt", ((a, b), (cc, b))), ("bb", ((a, d), (cc, d))), ("bl", ((a, b), (a, d))),
                              ("br", ((cc, b), (cc, d)))):
            spec = fmt.get(key)
            if m and spec is None:  # a merged block's edges come from the cells along them
                edge = {"bt": (m[0], c), "bb": (m[2], c), "bl": (r, m[1]), "br": (r, m[3])}[key]
                spec = (sh.formats.get(edge) or {}).get(key)
            if not spec:
                continue
            style, color = spec[0], spec[1] if len(spec) > 1 else "#000000"
            w = BORDER_WIDTHS.get(style, 0.5) * max(0.6, s)
            dash = {"dashed": (3 * s, 2 * s), "dotted": (0.8 * s, 1.2 * s)}.get(style)
            if style == "double":
                off = 1.0 * max(0.6, s)
                dx, dy = (0, off) if key in ("bt", "bb") else (off, 0)
                surface.line(p0[0] - dx, p0[1] - dy, p1[0] - dx, p1[1] - dy, color, w)
                surface.line(p0[0] + dx, p0[1] + dy, p1[0] + dx, p1[1] + dy, color, w)
            else:
                surface.line(p0[0], p0[1], p1[0], p1[1], color, w, dash)

    # charts lying over this page's columns and rows
    def edge_x(c):  # left edge of column c, also for columns off this page
        first = cols[0]
        if c >= first:
            return x0 + sum(sh.col_widths.get(k, fe.DEFAULT_COL_WIDTH) * PT * s for k in range(first, c)
                            if k not in sh.hidden_cols)
        return x0 - sum(sh.col_widths.get(k, fe.DEFAULT_COL_WIDTH) * PT * s for k in range(c, first)
                        if k not in sh.hidden_cols)

    def edge_y(r):
        first = rows[0]
        base = ys[first][0]
        if r >= first:
            return base + sum(sh.row_heights.get(k, fe.DEFAULT_ROW_HEIGHT) * PT * s for k in range(first, r)
                              if not sh.row_hidden(k))
        return base - sum(sh.row_heights.get(k, fe.DEFAULT_ROW_HEIGHT) * PT * s for k in range(r, first)
                          if not sh.row_hidden(k))

    if sh.charts:
        top_body = ys[rows[0]][0]
        surface.clip(x0, top_body, x0 + width, y0 + height)
        canvas = ChartCanvas(surface)
        for chart in sh.charts:
            r, c, dx, dy = chart["anchor"]
            cx, cy = edge_x(c) + dx * PT * s, edge_y(r) + dy * PT * s
            w, h = chart["size"][0] * PT * s, chart["size"][1] * PT * s
            if cx > x0 + width or cx + w < x0 or cy > y0 + height or cy + h < top_body:
                continue
            if chart["type"] == fe.PICTURE:
                surface.picture(cx, cy, cx + w, cy + h, chart)
                continue
            labels, series = wb.chart_data(sheet, chart)
            charts.draw_chart(canvas, (cx, cy, cx + w, cy + h), chart["type"], chart.get("title"), labels, series,
                              canvas.font(14 * s), canvas.font(9 * s), PT * s)
        surface.unclip()

    # header and footer, centred in the top and bottom margins
    total = len(layout.pages)
    pw, ph = layout.paper
    spec, size = (DEFAULT_FONT, False, False), 10.0
    for text, y in ((setup["header"], setup["margins"][4] * 72 + surface.ascent(spec, size)),
                    (setup["footer"], ph - setup["margins"][5] * 72 - surface.descent(spec, size))):
        if text:
            text = header_text(text, index + 1, total, sheet, file_name)
            surface.text((pw - surface.measure(text, spec, size)) / 2, y, text, spec, size, "#000000")


def _mix(c1, c2, t):
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


# ------------------------------------------------------------------ export

def export_pdf(wb, sheets, path, library, file_name="", selection=None):
    """Write the sheets (names) to a PDF file; `selection` (sheet, rect) prints just that block.
    Returns the number of pages."""
    doc = pdf.PdfDocument(library, os.path.splitext(file_name or os.path.basename(path))[0])
    for sheet in sheets:
        setup = setup_of(wb.sheets[sheet])
        if selection and selection[0] == sheet:
            setup = dict(setup, area=selection[1])
        layout = Layout(wb, sheet, setup)
        for i in range(len(layout.pages)):
            page = doc.new_page(*layout.paper)
            render_page(PdfSurface(page, library), layout, i, file_name)
    if not doc.pages:
        doc.new_page(*paper_size(DEFAULT_SETUP))
    doc.save(path)
    return len(doc.pages)
