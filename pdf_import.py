"""PDF pages as sheets (Tiện ích > Nhập PDF), the way PDF-to-Excel converters rebuild a page: the cells'
edges follow the page - column edges where table lines and text columns start, row edges where table
lines and text lines are - so each piece of text lands in the cell at its place. Table lines become cell
borders, filled boxes cell fills, other drawings shapes, pictures pictures. Text in a boxed table cell
goes in that cell (merged across the grid columns it spans), with its alignment.

Everything here works on data from pdfread; the app writes the result into the workbook.
"""

import base64
import bisect
import math
import re

import formula_engine as fe
import pdfread
import shapes

PX = 96 / 72  # pixels per point
MAX_POINTS = 80000  # points of drawings kept per page (a detailed CAD drawing would be slow to show)


class Options:
    def __init__(self, pages=None, one_sheet=False, borders=True, drawings=True, pictures=True, numbers=True,
                 families=()):
        self.pages, self.one_sheet = pages, one_sheet
        self.borders, self.drawings, self.pictures, self.numbers = borders, drawings, pictures, numbers
        # the fonts installed here: a PDF's font is used when one of them is it ("ArialMT" is Arial)
        self.families = {_squash(f): f for f in families}


class SheetData:
    """What becomes one sheet."""

    def __init__(self, name):
        self.name = name
        self.cells, self.formats = {}, {}
        self.col_widths, self.row_heights = {}, {}
        self.merges, self.objects = [], []
        self.landscape = False


class Result:
    def __init__(self):
        self.sheets = []
        self.pages = self.text_pages = self.cells = self.lines = self.shapes = self.pictures = 0
        self.scanned = []      # numbers of pages with pictures but no text
        self.skipped_pictures = 0
        self.trimmed = 0       # pages whose drawings were cut short


def page_kinds(doc, pages, limit=8):
    """(pages with text, pages with only pictures) among the first `limit` pages: a quick look before
    importing, to tell a scanned PDF."""
    text = scanned = 0
    for page in pages[:limit]:
        try:
            spans, paths, images = page.read(limit=60000)
        except (pdfread.PDFError, ValueError, TypeError):
            continue
        if any(s.text.strip() for s in spans):
            text += 1
        elif images:
            scanned += 1
    return text, scanned


def parse_pages(text, count):
    """Page numbers (from 0) in "1-3, 5, 8-" (all when empty); None when it can't be read."""
    text = text.strip()
    if not text:
        return list(range(count))
    out = []
    for part in re.split(r"[,;\s]+", text):
        if not part:
            continue
        m = re.fullmatch(r"(\d*)\s*-\s*(\d*)|(\d+)", part)
        if not m:
            return None
        if m.group(3):
            a = b = int(m.group(3))
        else:
            a = int(m.group(1)) if m.group(1) else 1
            b = int(m.group(2)) if m.group(2) else count
        if a < 1 or b < a:
            return None
        out += [p - 1 for p in range(a, min(b, count) + 1) if p - 1 not in out]
    return out or None


# ------------------------------------------------------------------ one page's pieces, in pixels from the top-left

class _Phrase:
    __slots__ = ("x0", "x1", "top", "bottom", "base", "size", "text", "bold", "italic", "color", "font", "angle",
                 "col_x", "box", "right")

    def __init__(self, x0, x1, top, bottom, base, size, text, bold, italic, color, font, angle):
        self.x0, self.x1, self.top, self.bottom, self.base = x0, x1, top, bottom, base
        self.size, self.text, self.bold, self.italic = size, text, bold, italic
        self.color, self.font, self.angle = color, font, angle
        self.col_x, self.box, self.right = x0, None, False


def _pieces(page, dy, doc, options):
    """Phrases, horizontal and vertical lines, fills, other drawings and pictures of a page, moved down by
    dy pixels (pages placed one below the other)."""
    spans, paths, images = page.read()
    H = page.height
    phrases = []
    for s in spans:
        text = s.text.replace(" ", " ").replace("\x00", "")
        if not text.strip():
            continue
        size = max(1.0, s.size) * PX
        a = s.angle % 360
        x0, x1 = s.x0 * PX, s.x1 * PX
        y0, y1 = (H - s.y0) * PX + dy, (H - s.y1) * PX + dy
        if a < 1 or a > 359:
            if x1 < x0:
                x0, x1 = x1, x0
            phrases.append(_Phrase(x0, max(x1, x0 + 1), y0 - 0.8 * size, y0 + 0.2 * size, y0, size, text, s.bold,
                                   s.italic, s.color, s.font, 0))
        else:
            turn = 90 if 45 <= a < 135 else -90 if 225 <= a < 315 else 0
            phrases.append(_Phrase(min(x0, x1) - 0.8 * size, max(x0, x1) + 0.2 * size, min(y0, y1), max(y0, y1),
                                   max(y0, y1), size, text, s.bold, s.italic, s.color, s.font, turn))
    hlines, vlines, fills, others = [], [], [], []
    for p in paths:
        width = max(p.width, 0.1) * PX
        for pts, closed in p.subpaths:
            q = [(x * PX, (H - y) * PX + dy) for x, y in pts]
            rect = _rect(q, closed)
            if rect is not None:
                x0, y0, x1, y1 = rect
                fill = p.fill if p.fill and p.fill != "#ffffff" else None
                if p.fill and (y1 - y0 <= 3 or x1 - x0 <= 3):  # a line drawn as a thin box
                    if y1 - y0 <= 3 and x1 - x0 > 3:
                        hlines.append(((y0 + y1) / 2, x0, x1, max(y1 - y0, 0.5), p.fill))
                    elif x1 - x0 <= 3 and y1 - y0 > 3:
                        vlines.append(((x0 + x1) / 2, y0, y1, max(x1 - x0, 0.5), p.fill))
                    continue
                if fill:
                    fills.append((x0, y0, x1, y1, fill))
                if p.stroke:
                    hlines += [(y0, x0, x1, width, p.stroke), (y1, x0, x1, width, p.stroke)]
                    vlines += [(x0, y0, y1, width, p.stroke), (x1, y0, y1, width, p.stroke)]
                continue
            straight = all(abs(a[0] - b[0]) < 0.8 or abs(a[1] - b[1]) < 0.8 for a, b in zip(q, q[1:] + (q[:1] if closed else [])))
            if straight and p.stroke and not p.fill and len(q) <= 2:  # (a bent line - a connector - stays a drawing)
                ring = q + (q[:1] if closed else [])
                for a, b in zip(ring, ring[1:]):
                    if abs(a[1] - b[1]) < 0.8 and abs(a[0] - b[0]) > 1:
                        hlines.append(((a[1] + b[1]) / 2, min(a[0], b[0]), max(a[0], b[0]), width, p.stroke))
                    elif abs(a[0] - b[0]) < 0.8 and abs(a[1] - b[1]) > 1:
                        vlines.append(((a[0] + b[0]) / 2, min(a[1], b[1]), max(a[1], b[1]), width, p.stroke))
                continue
            others.append((q, closed, p.stroke, p.fill if p.fill != "#ffffff" or p.stroke is None else p.fill,
                           max(p.width, 0.1)))
    pictures = []
    for im in images:
        x0, y0, x1, y1 = im.box
        pictures.append((x0 * PX, (H - y1) * PX + dy, x1 * PX, (H - y0) * PX + dy, im.stream))
    return phrases, hlines, vlines, fills, others, pictures


def _rect(q, closed):
    """(x0, y0, x1, y1) when the points make an upright rectangle."""
    pts = q[:-1] if len(q) == 5 and math.dist(q[0], q[-1]) < 0.5 else q
    if len(pts) != 4 or not closed and len(q) != 5:
        return None
    xs = sorted({round(p[0], 1) for p in pts})
    ys = sorted({round(p[1], 1) for p in pts})
    if len(xs) > 2 or len(ys) > 2:
        return None
    for a, b in zip(pts, pts[1:] + pts[:1]):
        if abs(a[0] - b[0]) > 0.5 and abs(a[1] - b[1]) > 0.5:
            return None
    return min(xs), min(ys), max(xs), max(ys)


# ------------------------------------------------------------------ phrases

def _merge_phrases(phrases, vlines):
    """Pieces of text on one line close together become one phrase (PDFs often place each word or letter
    on its own); a table line between them keeps them apart."""
    lines = {}
    for p in sorted(phrases, key=lambda p: (p.angle, round(p.base), p.x0)):
        key = (p.angle, round(p.base / max(2.0, p.size * 0.3)))
        lines.setdefault(key, []).append(p)
    walls = sorted(vlines)
    out = []
    for group in lines.values():
        group.sort(key=lambda p: p.x0)
        cur = None
        for p in group:
            if cur is not None and p.angle == 0 and cur.angle == 0:
                gap = p.x0 - cur.x1
                same = abs(p.size - cur.size) < 0.25 * cur.size and abs(p.base - cur.base) < 0.35 * cur.size
                if same and -0.3 * cur.size < gap < 0.35 * cur.size and not _wall_between(walls, cur.x1, p.x0, p.base):
                    space = gap > 0.12 * cur.size and not cur.text.endswith(" ") and not p.text.startswith(" ")
                    cur.text += (" " if space else "") + p.text
                    cur.x1 = max(cur.x1, p.x1)
                    cur.top, cur.bottom = min(cur.top, p.top), max(cur.bottom, p.bottom)
                    continue
            if cur is not None:
                out.append(cur)
            cur = p
        if cur is not None:
            out.append(cur)
    for p in out:
        p.text = p.text.strip()
        p.col_x = p.x0
    return [p for p in out if p.text]


def _wall_between(walls, a, b, y):
    i = bisect.bisect_left(walls, (a - 1,))
    while i < len(walls) and walls[i][0] <= b + 1:
        x, y0, y1 = walls[i][:3]
        if y0 - 1 <= y <= y1 + 1:
            return True
        i += 1
    return False


def _boxes(phrases, hlines, vlines, width):
    """The table cell (left, right, top, bottom lines) around each phrase that has lines on its left and
    right (and above and below)."""
    vs = sorted(vlines)
    hs = sorted(hlines)
    for p in phrases:
        if p.angle:
            continue
        mid = (p.top + p.bottom) / 2
        left = _nearest(vs, p.x0 + 2, mid, back=True)
        right = _nearest(vs, p.x1 - 2, mid, back=False)
        if left is None or right is None or right - left < 2 or right - left > 0.9 * width:
            continue
        cx = (p.x0 + p.x1) / 2
        top = _nearest(hs, p.top + 2, cx, back=True)
        bottom = _nearest(hs, p.bottom - 2, cx, back=False)
        p.box = (left, right, top, bottom)


def _nearest(lines, pos, across, back):
    """The position of the nearest line before (back) or after `pos` that crosses `across`."""
    i = bisect.bisect_right(lines, (pos, math.inf)) - 1 if back else bisect.bisect_left(lines, (pos, -math.inf))
    steps = 0
    while 0 <= i < len(lines) and steps < 400:
        at, a, b = lines[i][:3]
        if (at <= pos if back else at >= pos) and a - 1 <= across <= b + 1:
            return at
        i += -1 if back else 1
        steps += 1
    return None


# ------------------------------------------------------------------ the grid

def _cluster(values, tol):
    """Sorted representatives of values closer than tol: strong values (lines) win over weak ones."""
    out = []
    for v, strong in sorted(values):
        if out and v - out[-1][2] <= tol:
            group = out[-1]
            if strong and not group[1]:
                group[0], group[1] = v, True
            group[2] = v
        else:
            out.append([v, strong, v])
    return [g[0] for g in out]


def _text_columns(phrases):
    """Columns of a table without lines: text of one column overlaps across the rows whatever its alignment
    (left, centred, right), so the columns are the runs of overlapping text in rows of three or more pieces.
    Each phrase gets the left edge of its column (col_x); right-aligned ones are marked."""
    lines = {}
    for p in phrases:
        if p.box is None and not p.angle:
            lines.setdefault(round(p.base / max(2.0, p.size * 0.3)), []).append(p)
    rows = [line for line in lines.values() if len(line) >= 3] or [line for line in lines.values() if len(line) >= 2]
    members = sorted((p for line in rows for p in line), key=lambda p: p.x0)
    columns = []  # [left, right, [phrases]]
    for p in members:
        if columns and p.x0 <= columns[-1][1] + 1:
            columns[-1][1] = max(columns[-1][1], p.x1)
            columns[-1][2].append(p)
        else:
            columns.append([p.x0, p.x1, [p]])
    for left, right, group in columns:
        for p in group:
            p.col_x = left
            p.right = len(group) >= 3 and abs(p.x1 - right) < 2 and p.x0 - left > 3
    member_ids = {id(p) for p in members}
    for p in phrases:
        if p.box is None and not p.angle and id(p) not in member_ids:
            over = [c for c in columns if c[0] < p.x1 - 1 and p.x0 + 1 < c[1]]
            p.col_x = over[0][0] if len(over) == 1 else p.x0


_NUM = re.compile(r"[(+-]?[$¥€£₫]?\s?(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?\)?%?")


def _number(text):
    """(value, number format or None) of text that is clearly a number: 1234, -12.5, 1,234,567.89, (1,234)."""
    t = text.strip()
    if not t or len(t) > 30 or not _NUM.fullmatch(t) or t.endswith("%") or re.search(r"[$¥€£₫]", t):
        return None
    neg = t.startswith("-") or (t.startswith("(") and t.endswith(")"))
    body = t.strip("()+-").replace(" ", "")
    if body.startswith("0") and len(body.split(".")[0]) > 1:  # 007, 0123: codes, not numbers
        return None
    digits = body.replace(",", "")
    try:
        value = float(digits)
    except ValueError:
        return None
    decimals = len(digits.split(".")[1]) if "." in digits else 0
    fmt = None
    if "," in body:
        fmt = "#,##0" + ("." + "0" * decimals if decimals else "")
    if t.startswith("("):
        fmt = (fmt or "0") + ";(" + (fmt or "0") + ")"
    return (-value if neg else value), fmt


def _raw(text, convert):
    """The raw cell text: a number when it clearly is one, otherwise text kept as written."""
    if convert:
        found = _number(text)
        if found is not None:
            value, fmt = found
            return fe.number_text(int(value) if float(value).is_integer() and abs(value) < 1e15 else value), fmt
    value = fe.parse_literal(text)[0]
    if not isinstance(value, str) or text.startswith(("=", "'")):  # (a date, TRUE...) stays as written
        return "'" + text, None
    return text, None


def _squash(name):
    return re.sub(r"[^a-z0-9]", "", name.lower())


def _family(name, families):
    """The installed font family a PDF font is (its name without style words), or None."""
    name = re.sub(r"[-,](Bold|Italic|Regular|Oblique|BoldItalic|BoldOblique|Light|Medium|Semibold|Black|Roman|Book)"
                  r"(MT)?$", "", name or "")
    name = re.sub(r"(PS)?MT$", "", name)
    found = families.get(_squash(name))
    return None if found is None or _squash(found) == "calibri" else found


class _Grid:
    """Column and row edges in pixels and finding the cell of a point."""

    def __init__(self, xs, ys):
        self.xs, self.ys = xs, ys

    def col(self, x, slack=2.0):
        return max(0, min(len(self.xs) - 2, bisect.bisect_right(self.xs, x + slack) - 1))

    def row(self, y, slack=2.0):
        return max(0, min(len(self.ys) - 2, bisect.bisect_right(self.ys, y + slack) - 1))

    def anchor(self, x, y):
        c, r = self.col(x, 0.0), self.row(y, 0.0)
        return r, c, max(0, round(x - self.xs[c])), max(0, round(y - self.ys[r]))


def build(canvas_w, canvas_h, pieces, options, sd, result, doc):
    """Lay pieces (from _pieces, several pages may be put together) out on sheet data sd."""
    phrases, hlines, vlines, fills, others, pictures = pieces
    phrases = _merge_phrases(phrases, vlines)
    phrases, labels = _labels(phrases, others)
    _boxes(phrases, hlines, vlines, canvas_w)
    _text_columns(phrases)
    for p in phrases:
        # text beside a table, crossed by the height of one of its lines (the line is a row edge): the text
        # goes in the row below that line, where its letters mostly are, instead of a sliver of a row above it
        if p.box is None and not p.angle:
            inside = [y for y, x0, x1, _, _ in hlines if p.top + 2 < y < p.base - 1 and (x1 < p.x0 or x0 > p.x1)]
            if inside:
                p.top = max(inside)
    xs = [(0.0, True), (canvas_w, True)] + [(v[0], True) for v in vlines]
    ys = [(0.0, True), (canvas_h, True)] + [(h[0], True) for h in hlines]
    for x0, y0, x1, y1, _ in fills:  # filled cells mark the table's columns too
        if (x1 - x0) * (y1 - y0) < 0.3 * canvas_w * canvas_h:
            xs += [(x0, False), (x1, False)]
    for p in phrases:
        if p.box is None:
            xs.append((p.col_x, False))
            ys.append((p.top, False))
        else:
            left, right, top, bottom = p.box
            if top is None or bottom is None:
                ys.append((p.top, False))
    grid = _Grid(_cluster(xs, 4.0), _cluster(ys, 2.0))
    if len(grid.xs) < 2:
        grid.xs.append(grid.xs[0] + fe.DEFAULT_COL_WIDTH)
    if len(grid.ys) < 2:
        grid.ys.append(grid.ys[0] + fe.DEFAULT_ROW_HEIGHT)
    for c in range(len(grid.xs) - 1):
        if c < fe.MAX_COLS:
            sd.col_widths[c] = max(fe.MIN_COL_WIDTH, round(grid.xs[c + 1] - grid.xs[c]))
    for r in range(len(grid.ys) - 1):
        if r < fe.MAX_ROWS:
            sd.row_heights[r] = max(4, round(grid.ys[r + 1] - grid.ys[r]))
    _place_text(phrases, grid, options, sd, result)
    _fit_numbers(sd)
    if options.borders:
        others = others + _place_lines(hlines, vlines, grid, sd, result)
        _place_fills(fills, grid, canvas_w, canvas_h, sd)
    if options.drawings:
        _place_drawings(others, grid, sd, result)
        _place_labels(labels, grid, sd, options)
    else:
        for group in labels.values():
            for p in group:
                key = (grid.row(p.top), grid.col(p.x0))
                sd.cells.setdefault(key, p.text)
    if options.pictures:
        _place_pictures(pictures, grid, sd, result, doc)


def _fmt_of(p, convert_fmt, families=None):
    fmt = {}
    size = round(p.size / PX * 2) / 2
    if abs(size - 11) > 0.25:
        fmt["size"] = int(size) if size.is_integer() else size
    if p.bold:
        fmt["bold"] = True
    if p.italic:
        fmt["italic"] = True
    if p.color and p.color != "#000000":
        fmt["fg"] = p.color
    family = _family(p.font, families or {})
    if family:
        fmt["font"] = family
    if p.angle:
        fmt["rotate"] = 90 if p.angle > 0 else -90
    if convert_fmt:
        fmt["num"] = convert_fmt
    return fmt


def _place_text(phrases, grid, options, sd, result):
    cells = {}
    boxed = {}
    for p in phrases:
        if p.box is not None:
            boxed.setdefault(p.box, []).append(p)
            continue
        key = (grid.row(p.top), grid.col(p.col_x))
        cells.setdefault(key, []).append(p)
    taken = set()
    for (left, right, top, bottom), group in boxed.items():
        group.sort(key=lambda p: (round(p.base), p.x0))
        lines, last = [], None
        for p in group:
            if last is not None and abs(p.base - last.base) < 0.5 * p.size:
                lines[-1].append(p)
            else:
                lines.append([p])
            last = p
        text = "\n".join(" ".join(q.text for q in line) for line in lines)
        r0 = grid.row(top) if top is not None else grid.row(group[0].top)
        r1 = grid.row(bottom, -2.0) if bottom is not None else grid.row(group[-1].top)
        c0, c1 = grid.col(left), grid.col(right, -2.0)
        r1, c1 = max(r0, r1), max(c0, c1)
        raw, numfmt = _raw(text, options.numbers and len(lines) == 1)
        fmt = _fmt_of(group[0], numfmt, options.families)
        number = numfmt is not None or not raw.startswith("'") and fe.parse_literal(raw)[0] != raw
        x0, x1 = min(p.x0 for p in group), max(p.x1 for p in group)
        lgap, rgap = x0 - left, right - x1
        if lgap > 3 and abs(lgap - rgap) <= max(3.0, 0.12 * (right - left)):
            fmt["align"] = "center"
        elif rgap < lgap and rgap < 8:
            if not number:
                fmt["align"] = "right"
        elif number:
            fmt["align"] = "left"
        if top is not None and bottom is not None:
            y0, y1 = min(p.top for p in group), max(p.bottom for p in group)
            tgap, bgap = y0 - top, bottom - y1
            if abs(tgap - bgap) <= max(2.0, 0.15 * (bottom - top)):
                fmt["valign"] = "center"
            elif tgap < bgap:
                fmt["valign"] = "top"
        if len(lines) > 1:
            fmt["wrap"] = True
        if (r0, c0) in sd.cells:
            sd.cells[(r0, c0)] += "\n" + text
            continue
        sd.cells[(r0, c0)] = raw
        if fmt:
            sd.formats[(r0, c0)] = fmt
        if (r1 > r0 or c1 > c0) and not any((r, c) in taken for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)):
            sd.merges.append((r0, c0, r1, c1))
            taken.update((r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1))
        result.cells += 1
    for key, group in cells.items():
        if key in sd.cells:
            continue
        group.sort(key=lambda p: (round(p.base), p.x0))
        text = " ".join(p.text for p in group)
        raw, numfmt = _raw(text, options.numbers)
        fmt = _fmt_of(group[0], numfmt, options.families)
        if group[0].right and numfmt is None and fe.parse_literal(raw)[0] == raw:
            fmt["align"] = "right"
        if not group[0].angle:
            fmt["valign"] = "top"  # where the line sits on the page; the row reaches down to the next one
        sd.cells[key] = raw
        if fmt:
            sd.formats[key] = fmt
        result.cells += 1


def _fit_numbers(sd):
    """Widen columns whose numbers would show as ##### (fonts here may be wider than the PDF's)."""
    need = {}
    for (r, c), raw in sd.cells.items():
        if raw.startswith("'"):
            continue
        value = fe.parse_literal(raw)[0]
        if not isinstance(value, float):
            continue
        fmt = sd.formats.get((r, c)) or {}
        shown = fe.format_number(value, fmt["num"]) if fmt.get("num") else fe.format_general(value)
        size = (fmt.get("size") or 11) * PX
        width = len(shown) * size * 0.6 + 10
        merged = next((m for m in sd.merges if m[0] == r and m[1] == c), None)
        if merged:
            width -= sum(sd.col_widths.get(k, 0) for k in range(c + 1, merged[3] + 1))
        need[c] = max(need.get(c, 0), width)
    for c, width in need.items():
        if sd.col_widths.get(c, 0) < width:
            sd.col_widths[c] = round(width)


def _labels(phrases, others):
    """Text drawn inside a filled drawing (a flowchart box...) stays above it as a text box (cells lie under
    drawings, so a cell would be hidden): {drawing index: [phrases]}; the other phrases."""
    filled = []
    for i, (q, closed, stroke, fill, width) in enumerate(others):
        if closed and fill and len(q) >= 3:
            xs, ys = [x for x, _ in q], [y for _, y in q]
            filled.append((min(xs), min(ys), max(xs), max(ys), i, q))
    if not filled:
        return phrases, {}
    labels, rest = {}, []
    for p in phrases:
        cx, cy = (p.x0 + p.x1) / 2, (p.top + p.bottom) / 2
        inside = None
        for x0, y0, x1, y1, i, q in filled:
            if x0 <= cx <= x1 and y0 <= cy <= y1 and shapes.inside(cx, cy, q):
                inside = i  # the last one drawn: the topmost
        if inside is None:
            rest.append(p)
        else:
            labels.setdefault(inside, []).append(p)
    return rest, labels


def _place_labels(labels, grid, sd, options):
    for group in labels.values():
        group.sort(key=lambda p: (round(p.base), p.x0))
        lines, last = [], None
        for p in group:
            if last is not None and abs(p.base - last.base) < 0.5 * p.size:
                lines[-1].append(p)
            else:
                lines.append([p])
            last = p
        x0, x1 = min(p.x0 for p in group), max(p.x1 for p in group)
        y0, y1 = min(p.top for p in group), max(p.bottom for p in group)
        first = group[0]
        font = {"size": round(first.size / PX * 2) / 2, "color": first.color or "#000000"}
        if first.bold:
            font["bold"] = True
        if first.italic:
            font["italic"] = True
        family = _family(first.font, options.families)
        if family:
            font["font"] = family
        box = (x0 - 11, y0 - 6, x1 + 11, y1 + 6)
        obj = shapes.new_shape("textBox", grid.anchor(box[0], box[1]), (round(box[2] - box[0]), round(box[3] - box[1])),
                               dict(shapes.TEXT_BOX_STYLE, align="center", valign="middle"), id=fe.new_picture_id(),
                               name=f"Chữ từ PDF {len(sd.objects) + 1}",
                               text="\n".join(" ".join(q.text for q in line) for line in lines), font=font)
        sd.objects.append(obj)


def _style(width):
    return "thin" if width < 1.6 else "medium" if width < 2.6 else "thick"


def _place_lines(hlines, vlines, grid, sd, result):
    """Lines along the cells' edges become borders; lines that don't run along whole cells (a connector
    between flowchart boxes...) are returned to be kept as drawings."""
    xs, ys = grid.xs, grid.ys
    rest = []
    for y, x0, x1, w, color in hlines:
        r = bisect.bisect_left(ys, y - 2.5)
        cols = []
        if r < len(ys) and abs(ys[r] - y) <= 2.5:
            c = bisect.bisect_left(xs, x0 - 2.5)
            while c < len(xs) - 1 and xs[c + 1] <= x1 + 2.5:
                cols.append(c)
                c += 1
        if sum(xs[c + 1] - xs[c] for c in cols) < 0.8 * (x1 - x0) or (r > 0 and r - 1 >= len(ys) - 1):
            rest.append(([(x0, y), (x1, y)], False, color, None, w / PX))
            continue
        row, key = (r - 1, "bb") if r > 0 else (r, "bt")
        for c in cols:
            sd.formats.setdefault((row, c), {})[key] = [_style(w), color or "#000000"]
        result.lines += 1
    for x, y0, y1, w, color in vlines:
        c = bisect.bisect_left(xs, x - 2.5)
        rows = []
        if c < len(xs) and abs(xs[c] - x) <= 2.5:
            r = bisect.bisect_left(ys, y0 - 2.5)
            while r < len(ys) - 1 and ys[r + 1] <= y1 + 2.5:
                rows.append(r)
                r += 1
        if sum(ys[r + 1] - ys[r] for r in rows) < 0.8 * (y1 - y0) or (c > 0 and c - 1 >= len(xs) - 1):
            rest.append(([(x, y0), (x, y1)], False, color, None, w / PX))
            continue
        col, key = (c - 1, "br") if c > 0 else (c, "bl")
        for r in rows:
            sd.formats.setdefault((r, col), {})[key] = [_style(w), color or "#000000"]
        result.lines += 1
    return rest


def _place_fills(fills, grid, width, height, sd):
    xs, ys = grid.xs, grid.ys
    for x0, y0, x1, y1, color in fills:
        if (x1 - x0) * (y1 - y0) > 0.6 * width * height:
            continue  # a page background
        c0 = bisect.bisect_left(xs, x0 - 2.5)
        r0 = bisect.bisect_left(ys, y0 - 2.5)
        for r in range(r0, len(ys) - 1):
            if ys[r + 1] > y1 + 2.5:
                break
            for c in range(c0, len(xs) - 1):
                if xs[c + 1] > x1 + 2.5:
                    break
                sd.formats.setdefault((r, c), {})["bg"] = color


def _place_drawings(others, grid, sd, result):
    groups = {}
    count = 0
    for q, closed, stroke, fill, width in others:
        count += len(q)
        if count > MAX_POINTS:
            result.trimmed += 1
            break
        groups.setdefault((stroke, fill, round(width * 4) / 4), []).append((q, closed))
    for (stroke, fill, width), parts in groups.items():
        xs = [x for q, _ in parts for x, _ in q]
        ys = [y for q, _ in parts for _, y in q]
        x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
        w, h = max(x1 - x0, 1.0), max(y1 - y0, 1.0)
        paths = [([((x - x0) / w, (y - y0) / h) for x, y in q], closed, "norm" if fill and closed else "none")
                 for q, closed in parts]
        obj = shapes.new_shape("custom", grid.anchor(x0, y0), (max(1, round(w)), max(1, round(h))),
                               {"fill": fill, "line": stroke, "line_w": width}, id=fe.new_picture_id(),
                               name=f"Hình từ PDF {len(sd.objects) + 1}", paths=paths)
        sd.objects.append(obj)
        result.shapes += 1


def _place_pictures(pictures, grid, sd, result, doc):
    for x0, y0, x1, y1, stream in pictures:
        if x1 - x0 < 2 or y1 - y0 < 2:
            continue
        try:
            found = pdfread.image_data(doc, stream)
        except (pdfread.PDFError, ValueError, TypeError, IndexError, KeyError, MemoryError):
            found = None
        if found is None:
            result.skipped_pictures += 1
            continue
        fmt, data, native = found
        sd.objects.append({"type": fe.PICTURE, "image": base64.b64encode(data).decode("ascii"), "format": fmt,
                           "id": fe.new_picture_id(), "name": f"Ảnh từ PDF {result.pictures + 1}",
                           "anchor": grid.anchor(x0, y0), "size": (max(2, round(x1 - x0)), max(2, round(y1 - y0))),
                           "native": list(native)})
        result.pictures += 1


# ------------------------------------------------------------------ the whole import

def convert(doc, pages, options, name, progress=None):
    """Result of importing the pages (indices into doc.pages()) of a document; progress(done, total) is
    called after each page and may return False to stop."""
    result = Result()
    all_pages = doc.pages()
    chosen = [all_pages[i] for i in pages if 0 <= i < len(all_pages)]
    result.pages = len(chosen)
    gap = fe.DEFAULT_ROW_HEIGHT

    def note(page_no, pieces):
        phrases, _, _, _, _, pictures = pieces
        if phrases:
            result.text_pages += 1
        elif pictures:
            result.scanned.append(page_no + 1)

    if options.one_sheet:
        sd = SheetData(name)
        merged = ([], [], [], [], [], [])
        width = dy = 0.0
        for k, (i, page) in enumerate(zip(pages, chosen)):
            pieces = _pieces(page, dy, doc, options)
            note(i, pieces)
            for into, part in zip(merged, pieces):
                into.extend(part)
            width = max(width, page.width * PX)
            dy += page.height * PX + gap
            if progress and progress(k + 1, len(chosen)) is False:
                break
        sd.landscape = bool(chosen) and chosen[0].width > chosen[0].height
        build(width, max(dy - gap, 1.0), merged, options, sd, result, doc)
        result.sheets.append(sd)
        return result
    for k, (i, page) in enumerate(zip(pages, chosen)):
        sd = SheetData(f"{name} - Tr.{i + 1}" if len(chosen) > 1 else name)
        pieces = _pieces(page, 0.0, doc, options)
        note(i, pieces)
        sd.landscape = page.width > page.height
        build(page.width * PX, page.height * PX, pieces, options, sd, result, doc)
        result.sheets.append(sd)
        if progress and progress(k + 1, len(chosen)) is False:
            break
    return result
