"""A small PDF writer for PyExcel's export: vector drawing and text in embedded TrueType fonts.

Coordinates are in points with the origin at the top-left corner of the page (y grows downwards, as
on screen); the writer flips them. Text is set in Type 0 / CIDFontType2 fonts subset from installed
TrueType fonts (see ttf.py), with a ToUnicode map so it can be searched and copied; characters the
chosen font lacks are drawn with a fallback font.
"""

import datetime as dt
import math
import zlib


def _n(x):
    """A number as PDF writes it."""
    if isinstance(x, int):
        return str(x)
    text = f"{x:.3f}".rstrip("0").rstrip(".")
    return text if text not in ("", "-0") else "0"


def _rgb(color):
    return " ".join(_n(int(color[i:i + 2], 16) / 255) for i in (1, 3, 5))


def _text_string(text):
    return "<FEFF" + text.encode("utf-16-be").hex().upper() + ">"


class PdfFont:
    """A font used in the document; remembers the glyphs drawn (for subsetting and ToUnicode)."""

    def __init__(self, font, resource, obj_id):
        self.font, self.resource, self.obj_id = font, resource, obj_id
        self.used = {}  # glyph id -> characters

    def encode(self, text):
        out = []
        for ch in text:
            gid = self.font.glyph(ch)
            self.used.setdefault(gid, ch)
            out.append(f"{gid:04X}")
        return "<" + "".join(out) + ">"


class Page:
    def __init__(self, doc, width, height):
        self.doc, self.width, self.height = doc, width, height
        self.ops, self.fonts = [], set()

    def _y(self, y):
        return self.height - y

    def _paint(self, fill, stroke, width, dash=None):
        ops = []
        if fill:
            ops.append(f"{_rgb(fill)} rg")
        if stroke:
            ops.append(f"{_rgb(stroke)} RG {_n(width)} w")
            ops.append(f"[{' '.join(_n(d) for d in dash)}] 0 d" if dash else "[] 0 d")
        return ops, ("B" if fill and stroke else "f" if fill else "S")

    def rect(self, x0, y0, x1, y1, fill=None, stroke=None, width=1.0):
        if not (fill or stroke):
            return
        ops, paint = self._paint(fill, stroke, width)
        self.ops += ops + [f"{_n(x0)} {_n(self._y(y1))} {_n(x1 - x0)} {_n(y1 - y0)} re {paint}"]

    def polyline(self, points, color="#000000", width=1.0, dash=None, closed=False, fill=None):
        if len(points) < 2:
            return
        ops, paint = self._paint(fill, color, width, dash)
        path = [f"{_n(points[0][0])} {_n(self._y(points[0][1]))} m"]
        path += [f"{_n(x)} {_n(self._y(y))} l" for x, y in points[1:]]
        self.ops += ops + path + [("h " if closed else "") + paint]

    def line(self, x0, y0, x1, y1, color="#000000", width=1.0, dash=None):
        self.polyline([(x0, y0), (x1, y1)], color, width, dash)

    def polygon(self, points, fill=None, stroke=None, width=1.0):
        if not (fill or stroke) or len(points) < 3:
            return
        ops, paint = self._paint(fill, stroke, width)
        path = [f"{_n(points[0][0])} {_n(self._y(points[0][1]))} m"]
        path += [f"{_n(x)} {_n(self._y(y))} l" for x, y in points[1:]]
        self.ops += ops + path + ["h " + paint]

    def _arc_path(self, cx, cy, rx, ry, start, extent, move=True):
        """Bezier curves along an elliptical arc; angles in degrees, counter-clockwise as in Tk."""
        segments = max(1, math.ceil(abs(extent) / 90))
        step = math.radians(extent) / segments
        a = math.radians(start)
        cyp = self._y(cy)
        path = []
        if move:
            path.append(f"{_n(cx + rx * math.cos(a))} {_n(cyp + ry * math.sin(a))} m")
        k = 4 / 3 * math.tan(step / 4)
        for _ in range(segments):
            b = a + step
            p1 = (math.cos(a) - k * math.sin(a), math.sin(a) + k * math.cos(a))
            p2 = (math.cos(b) + k * math.sin(b), math.sin(b) - k * math.cos(b))
            path.append(" ".join(_n(v) for v in (cx + rx * p1[0], cyp + ry * p1[1], cx + rx * p2[0], cyp + ry * p2[1],
                                                  cx + rx * math.cos(b), cyp + ry * math.sin(b))) + " c")
            a = b
        return path

    def ellipse(self, x0, y0, x1, y1, fill=None, stroke=None, width=1.0):
        if not (fill or stroke):
            return
        ops, paint = self._paint(fill, stroke, width)
        self.ops += ops + self._arc_path((x0 + x1) / 2, (y0 + y1) / 2, (x1 - x0) / 2, (y1 - y0) / 2, 0, 360) + \
            ["h " + paint]

    def arc(self, x0, y0, x1, y1, start, extent, fill=None, stroke=None, width=1.0, style="pieslice"):
        cx, cy, rx, ry = (x0 + x1) / 2, (y0 + y1) / 2, (x1 - x0) / 2, (y1 - y0) / 2
        if style == "arc":
            ops, paint = self._paint(None, stroke or "#000000", width)
            self.ops += ops + self._arc_path(cx, cy, rx, ry, start, extent) + [paint]
            return
        ops, paint = self._paint(fill, stroke, width)
        a, cyp = math.radians(start), self._y(cy)
        path = [f"{_n(cx)} {_n(cyp)} m", f"{_n(cx + rx * math.cos(a))} {_n(cyp + ry * math.sin(a))} l"]
        self.ops += ops + path + self._arc_path(cx, cy, rx, ry, start, extent, move=False) + ["h " + paint]

    def text(self, x, y, text, font, size, color="#000000"):
        """Draw text with its baseline at y, starting at x, falling back to other fonts per character."""
        self.ops.append(f"{_rgb(color)} rg")
        for pfont, run in self.doc.runs(text, font):
            self.fonts.add(pfont)
            self.ops.append(f"BT /{pfont.resource} {_n(size)} Tf {_n(x)} {_n(self._y(y))} Td {pfont.encode(run)} Tj ET")
            x += self.doc.run_width(pfont.font, run, size)

    def clip(self, x0, y0, x1, y1):
        """Restrict drawing to a rectangle until unclip()."""
        self.ops.append(f"q {_n(x0)} {_n(self._y(y1))} {_n(x1 - x0)} {_n(y1 - y0)} re W n")

    def unclip(self):
        self.ops.append("Q")


class PdfDocument:
    def __init__(self, library, title=""):
        self.library, self.title = library, title
        self.pages, self.fonts = [], {}
        self._next_id = 3  # 1: catalog, 2: page tree

    def _reserve(self, count=1):
        first = self._next_id
        self._next_id += count
        return first

    def new_page(self, width, height):
        page = Page(self, width, height)
        self.pages.append(page)
        return page

    def _pdf_font(self, font):
        key = (font.path, font.member)
        pf = self.fonts.get(key)
        if pf is None:
            pf = self.fonts[key] = PdfFont(font, f"F{len(self.fonts) + 1}", self._reserve(5))
        return pf

    def runs(self, text, font):
        """[(PdfFont, text)]: stretches of text each drawn with one font."""
        out = []
        for ch in text:
            use = font
            if not font.has(ch) and not ch.isspace():
                use = self.library.fallback_for(ch) or font
            pf = self._pdf_font(use)
            if out and out[-1][0] is pf:
                out[-1][1].append(ch)
            else:
                out.append((pf, [ch]))
        return [(pf, "".join(chars)) for pf, chars in out]

    def run_width(self, font, text, size):
        return sum(font.advance(font.glyph(ch)) for ch in text) * size / 1000

    def measure(self, text, font, size):
        """Width in points of text set in font (with fallbacks) at size."""
        total = 0.0
        for ch in text:
            use = font
            if not font.has(ch) and not ch.isspace():
                use = self.library.fallback_for(ch) or font
            total += use.advance(use.glyph(ch))
        return total * size / 1000

    # writing
    def save(self, path):
        objects = {}

        def put(obj_id, body):
            objects[obj_id] = body if isinstance(body, bytes) else body.encode("latin-1")

        def stream(obj_id, data, extra=""):
            packed = zlib.compress(data, 9)
            put(obj_id, f"<< /Length {len(packed)} /Filter /FlateDecode{extra} >>\nstream\n".encode("latin-1")
                + packed + b"\nendstream")

        page_ids = []
        for page in self.pages:
            pid = self._reserve(2)
            cid = pid + 1
            page_ids.append(pid)
            fonts = " ".join(f"/{pf.resource} {pf.obj_id} 0 R" for pf in sorted(page.fonts, key=lambda f: f.resource))
            put(pid, f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {_n(page.width)} {_n(page.height)}] "
                     f"/Resources << /Font << {fonts} >> >> /Contents {cid} 0 R >>")
            stream(cid, "\n".join(page.ops).encode("latin-1"))
        for n, pf in enumerate(self.fonts.values()):
            self._write_font(pf, n, put, stream)
        put(1, "<< /Type /Catalog /Pages 2 0 R >>")
        put(2, f"<< /Type /Pages /Kids [{' '.join(f'{p} 0 R' for p in page_ids)}] /Count {len(page_ids)} >>")
        info = self._reserve()
        stamp = dt.datetime.now().strftime("D:%Y%m%d%H%M%S")
        put(info, f"<< /Title {_text_string(self.title)} /Producer (PyExcel) /CreationDate ({stamp}) >>")
        out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
        offsets = {}
        for obj_id in sorted(objects):
            offsets[obj_id] = len(out)
            out += f"{obj_id} 0 obj\n".encode("latin-1") + objects[obj_id] + b"\nendobj\n"
        xref = len(out)
        size = max(objects) + 1
        out += f"xref\n0 {size}\n0000000000 65535 f \n".encode("latin-1")
        for obj_id in range(1, size):
            out += (f"{offsets[obj_id]:010d} 00000 n \n" if obj_id in offsets else "0000000000 65535 f \n").encode()
        out += f"trailer\n<< /Size {size} /Root 1 0 R /Info {info} 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode("latin-1")
        with open(path, "wb") as f:
            f.write(out)

    def _write_font(self, pf, n, put, stream):
        font = pf.font
        type0, cid, desc, file_id, tounicode = range(pf.obj_id, pf.obj_id + 5)
        tag = "".join(chr(65 + (n // 26 ** i) % 26) for i in range(6))
        name = f"{tag}+{font.postscript}"
        stream(file_id, font.subset(set(pf.used)))
        flags = 32 | (64 if font.style & 2 else 0)
        put(desc, f"<< /Type /FontDescriptor /FontName /{name} /Flags {flags} /FontBBox [{' '.join(map(str, font.bbox))}] "
                  f"/ItalicAngle {_n(font.italic_angle)} /Ascent {font.ascent} /Descent {font.descent} "
                  f"/CapHeight {font.cap_height} /StemV {120 if font.style & 1 else 80} /FontFile2 {file_id} 0 R >>")
        widths = " ".join(f"{gid} [{_n(round(font.advance(gid), 1))}]" for gid in sorted(pf.used))
        put(cid, f"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{name} "
                 f"/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> "
                 f"/FontDescriptor {desc} 0 R /W [{widths}] /CIDToGIDMap /Identity >>")
        put(type0, f"<< /Type /Font /Subtype /Type0 /BaseFont /{name} /Encoding /Identity-H "
                   f"/DescendantFonts [{cid} 0 R] /ToUnicode {tounicode} 0 R >>")
        pairs = [f"<{gid:04X}> <{ch.encode('utf-16-be').hex().upper()}>" for gid, ch in sorted(pf.used.items())]
        blocks = []
        for i in range(0, len(pairs), 100):
            chunk = pairs[i:i + 100]
            blocks.append(f"{len(chunk)} beginbfchar\n" + "\n".join(chunk) + "\nendbfchar")
        cmap = ("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n"
                "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n"
                "/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\n"
                "endcodespacerange\n" + "\n".join(blocks) +
                "\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend")
        stream(tounicode, cmap.encode("latin-1"))
