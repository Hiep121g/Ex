"""Cells copied in another program, read from the clipboard formats spreadsheets exchange.

Excel puts a copied block on the clipboard as "XML Spreadsheet" (the SpreadsheetML of Excel 2003) besides
text, HTML and pictures: values with their types, formulas (in R1C1 form), every cell's style - font, size,
bold, italic, underline, colours, fill, borders, alignment, wrapping, rotation, number format - merged
cells, column widths, notes and links. That is read first. Otherwise an HTML table is read (Google Sheets,
LibreOffice, web pages, Word, Excel when the XML is missing): text, inline and class CSS, colspan/rowspan
and column widths. Both give a Block, which the app pastes like its own copied cells.
"""

import copy
import datetime as dt
import html
import html.parser
import re
import xml.etree.ElementTree as ET

import formula_engine as fe

PT = 96 / 72  # pixels per point


class Block:
    """A pasted block: h x w cells. raws[(i, j)] is what goes in the cell (a formula when there is one,
    else the value as raw text), values[(i, j)] the value alone, formulas[(i, j)] (formula text in R1C1
    form, or None) and formats[(i, j)] the cell format (dicts shared between cells of one style)."""

    def __init__(self, h=0, w=0):
        self.h, self.w = h, w
        self.values, self.formats, self.formulas = {}, {}, {}
        self.merges = []          # (i1, j1, i2, j2) in the block
        self.col_widths = {}      # j -> pixels
        self.default_width = None  # pixels of the columns col_widths leaves out
        self.row_heights = {}     # i -> pixels
        self.notes, self.links = {}, {}
        self.source = ""

    def empty(self, i, j):
        return (i, j) not in self.values and (i, j) not in self.formulas

    def transposed(self):
        """The block turned rows to columns (sizes are dropped: the columns are what were rows)."""
        out = copy.copy(self)
        out.h, out.w = self.w, self.h
        for name, value in vars(self).items():
            if isinstance(value, dict) and value and isinstance(next(iter(value)), tuple):
                setattr(out, name, {(j, i): v for (i, j), v in value.items()})
        out.merges = [(j1, i1, j2, i2) for i1, j1, i2, j2 in self.merges]
        out.col_widths, out.row_heights = {}, {}
        return out

    def fit_content(self):
        """Shrink the block to the cells that have something (a copied whole column has a million rows)."""
        keys = list(self.values) + list(self.formulas) + list(self.formats)
        self.h = max([k[0] + 1 for k in keys] + [m[2] + 1 for m in self.merges] + [1])
        self.w = max([k[1] + 1 for k in keys] + [m[3] + 1 for m in self.merges] + [1])

    def raw(self, i, j, row, col, names=()):
        """What to put in the cell when the block's (i, j) lands on sheet cell (row, col): its formula,
        with references for that place, or its value (see a1_formula; names: the workbook's names)."""
        f = self.formulas.get((i, j))
        if f:
            a1 = a1_formula(f, row, col, names)
            if a1:
                return a1
        return self.values.get((i, j))


# ------------------------------------------------------------------ R1C1 formulas

_R1C1_RE = re.compile(r"""
    (?P<string>"(?:[^"]|"")*")
  | (?P<sheet>(?:'(?:[^']|'')+'|(?<![\w.])[^\W\d][\w.]*)!)
  | (?<![\w.$])(?P<ref>R(?P<r>\[-?\d+\]|\d+)?C(?P<c>\[-?\d+\]|\d+)?|R(?P<ro>\[-?\d+\]|\d+)|C(?P<co>\[-?\d+\]|\d+))
    (?![\w(.\[])
  | (?P<other>\[)
""", re.VERBOSE)


def _part(text, here):
    """(index from 0, absolute) of R1C1's row or column part: '' (this one), '[n]' (n away) or 'n'."""
    if not text:
        return here, False
    if text.startswith("["):
        return here + int(text[1:-1]), False
    return int(text) - 1, True


def a1_formula(formula, row, col, names=()):
    """An R1C1 formula (as Excel's XML Spreadsheet has it) in A1 form for the cell (row, col) - or None when
    it can't be pasted as a formula here (the value is pasted instead): references to other sheets or
    workbooks (their names differ in PyExcel), whole rows or columns, references off the sheet, functions
    PyExcel lacks and names not among `names` (the workbook's, upper case)."""
    text = formula if formula.startswith("=") else "=" + formula
    out, pos, ok = [], 0, True
    for m in _R1C1_RE.finditer(text):
        out.append(text[pos:m.start()])
        pos = m.end()
        if m.group("string"):
            out.append(m.group("string"))
        elif m.group("sheet") or m.group("other") or m.group("ro") is not None or m.group("co") is not None:
            ok = False
            break
        else:
            r, rabs = _part(m.group("r"), row)
            c, cabs = _part(m.group("c"), col)
            if not (0 <= r < fe.MAX_ROWS and 0 <= c < fe.MAX_COLS):
                ok = False
                break
            out.append(f"{'$' if cabs else ''}{fe.col_name(c)}{'$' if rabs else ''}{r + 1}")
    if not ok:
        return None
    out.append(text[pos:])
    a1 = fe.formula_from_file("".join(out))
    if fe.parse_formula(a1)[0] == "err":
        return None
    tokens = [t for t in fe.tokenize(a1[1:]) if t.kind != "ws"]
    for k, t in enumerate(tokens):
        if t.kind == "ident":
            called = k + 1 < len(tokens) and tokens[k + 1].text == "("
            upper = t.text.upper()
            if (upper not in _known_functions() if called else upper not in ("TRUE", "FALSE") and upper not in names):
                return None
    return a1


_FUNCTIONS = None


def _known_functions():
    global _FUNCTIONS
    if _FUNCTIONS is None:
        _FUNCTIONS = set(fe.function_names())
    return _FUNCTIONS


# ------------------------------------------------------------------ values

def _serial(text):
    """Excel's serial number of an XML Spreadsheet DateTime ("2026-10-22T13:30:00.000"); times alone come as
    1899-12-31Thh:mm, and dates before March 1900 count Excel's 29 February 1900."""
    try:
        d = dt.datetime.strptime(text[:19], "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        return None
    base = dt.datetime(1899, 12, 31) if d < dt.datetime(1900, 3, 1) else dt.datetime(1899, 12, 30)
    delta = d - base
    return delta.days + delta.seconds / 86400


def text_raw(text):
    """Raw cell text keeping `text` as text (quoted when PyExcel would read it as a number, date...)."""
    if not text:
        return None
    if text[:1] in "='" or fe.parse_literal(text) != (text, None):
        return "'" + text
    return text


# ------------------------------------------------------------------ XML Spreadsheet

SS = "{urn:schemas-microsoft-com:office:spreadsheet}"
X = "{urn:schemas-microsoft-com:office:excel}"
NAMED_FORMATS = {"General": None, "General Number": None, "Fixed": "0.00", "Standard": "#,##0.00",
                 "Percent": "0.00%", "Scientific": "0.00E+00", "General Date": "dd/mm/yyyy hh:mm",
                 "Long Date": "dd/mm/yyyy", "Medium Date": "dd/mm/yyyy", "Short Date": "dd/mm/yyyy",
                 "Long Time": "hh:mm:ss", "Medium Time": "hh:mm", "Short Time": "hh:mm", "Currency": "#,##0.00",
                 "Euro Currency": "#,##0.00 €", "Yes/No": None, "True/False": None, "On/Off": None, "@": "@"}
BORDER_LINES = {"Continuous": "thin", "Double": "double", "Dash": "dashed", "Dot": "dotted", "DashDot": "dashed",
                "DashDotDot": "dotted", "SlantDashDot": "dashed"}


def _a(el, name):
    return el.get(SS + name) if el is not None else None


def _color(value):
    if not value:
        return None
    value = value.strip().lower()
    if re.fullmatch(r"#[0-9a-f]{6}", value):
        return value
    if re.fullmatch(r"#[0-9a-f]{3}", value):
        return "#" + "".join(ch * 2 for ch in value[1:])
    return CSS_COLORS.get(value)


def _style_xml(style, base):
    """A cell format from an XML Spreadsheet <Style>, on top of `base` (its parent's)."""
    fmt = dict(base)
    font = style.find(SS + "Font")
    if font is not None:
        if _a(font, "FontName"):
            fmt["font"] = _a(font, "FontName")
        if _a(font, "Size"):
            fmt["size"] = float(_a(font, "Size"))
        for key, attr in (("bold", "Bold"), ("italic", "Italic"), ("strike", "StrikeThrough")):
            if _a(font, attr) is not None:
                fmt[key] = _a(font, attr) == "1"
        if _a(font, "Underline") is not None:
            fmt["underline"] = ("double" if "Double" in _a(font, "Underline") else
                                "single" if _a(font, "Underline") != "None" else None)
        if _a(font, "Color"):
            fmt["fg"] = _color(_a(font, "Color"))
    fill = style.find(SS + "Interior")
    if fill is not None and (_a(fill, "Pattern") or "Solid") != "None" and _color(_a(fill, "Color")):
        fmt["bg"] = _color(_a(fill, "Color"))
    al = style.find(SS + "Alignment")
    if al is not None:
        h = _a(al, "Horizontal")
        if h:
            fmt["align"] = {"Left": "left", "Right": "right", "Center": "center", "CenterAcrossSelection": "center",
                            "Distributed": "center", "Justify": "left", "Fill": "left"}.get(h)
        v = _a(al, "Vertical")
        if v:
            fmt["valign"] = {"Top": "top", "Center": "center", "Justify": "center", "Distributed": "center"}.get(v)
        if _a(al, "WrapText") is not None:
            fmt["wrap"] = _a(al, "WrapText") == "1"
        if _a(al, "Indent"):
            fmt["indent"] = int(float(_a(al, "Indent")))
        if _a(al, "VerticalText") == "1":
            fmt["rotate"] = 255
        elif _a(al, "Rotate"):
            fmt["rotate"] = int(float(_a(al, "Rotate")))
    borders = style.find(SS + "Borders")
    if borders is not None:
        for b in borders.findall(SS + "Border"):
            key = {"Top": "bt", "Bottom": "bb", "Left": "bl", "Right": "br"}.get(_a(b, "Position"))
            line = BORDER_LINES.get(_a(b, "LineStyle"))
            if key and line:
                weight = int(float(_a(b, "Weight") or 1))
                if line == "thin":
                    line = "thick" if weight >= 3 else "medium" if weight == 2 else "thin"
                fmt[key] = [line, _color(_a(b, "Color")) or "#000000"]
    number = style.find(SS + "NumberFormat")
    if number is not None and _a(number, "Format") is not None:
        code = _a(number, "Format")
        fmt["num"] = NAMED_FORMATS[code] if code in NAMED_FORMATS else fe.excel_number_format(code)
    prot = style.find(SS + "Protection")
    if prot is not None:
        if _a(prot, "Protected") is not None:
            fmt["unlocked"] = _a(prot, "Protected") == "0"
        if prot.get(X + "HideFormula") is not None:
            fmt["hide_formula"] = prot.get(X + "HideFormula") == "1"
    return fmt


def clean_format(fmt):
    """A format without the defaults (Calibri 11, black, bottom-aligned...); None when nothing is left."""
    out = {}
    for key, value in fmt.items():
        if not value:
            continue
        if (key == "font" and value == "Calibri") or (key == "size" and value == 11) or (
                key == "fg" and value == "#000000"):
            continue
        if key == "size":
            value = int(value) if float(value).is_integer() else value
        out[key] = value
    return out or None


def from_spreadsheet_xml(data):
    """The Block in an XML Spreadsheet (bytes or text), or None when it has no table."""
    try:
        root = ET.fromstring(data.rstrip(b"\0") if isinstance(data, bytes) else data)
    except ET.ParseError:
        return None
    raw_styles = {s.get(SS + "ID"): s for s in root.iter(SS + "Style")}
    made = {}

    def style_of(sid):
        if sid in made:
            return made[sid]
        style = raw_styles.get(sid)
        made[sid] = {}  # (a parent loop ends here)
        if style is None:
            fmt = {}
        else:
            parent = style.get(SS + "Parent")
            base = style_of(parent) if parent else (style_of("Default") if sid != "Default" else {})
            fmt = _style_xml(style, base)
        made[sid] = fmt
        return fmt
    cleaned = {}

    def format_of(sid):
        if sid not in cleaned:
            cleaned[sid] = clean_format(style_of(sid))
        return cleaned[sid]

    sheet = root.find(SS + "Worksheet")
    table = sheet.find(SS + "Table") if sheet is not None else None
    if table is None:
        return None
    block = Block()
    block.source = "Excel"
    if _a(table, "DefaultColumnWidth"):
        block.default_width = round(float(_a(table, "DefaultColumnWidth")) * PT)
    j = -1
    for col in table.findall(SS + "Column"):
        j = int(_a(col, "Index")) - 1 if _a(col, "Index") else j + 1
        span = int(_a(col, "Span") or 0)
        if _a(col, "Width"):
            for k in range(j, j + span + 1):
                block.col_widths[k] = max(fe.MIN_COL_WIDTH, round(float(_a(col, "Width")) * PT))
        j += span
    i = -1
    for row in table.findall(SS + "Row"):
        i = int(_a(row, "Index")) - 1 if _a(row, "Index") else i + 1
        if _a(row, "Height"):
            block.row_heights[i] = round(float(_a(row, "Height")) * PT)
        j = -1
        for cell in row.findall(SS + "Cell"):
            j = int(_a(cell, "Index")) - 1 if _a(cell, "Index") else j + 1
            fmt = format_of(_a(cell, "StyleID") or "Default")
            if fmt:
                block.formats[(i, j)] = fmt
            across, down = int(_a(cell, "MergeAcross") or 0), int(_a(cell, "MergeDown") or 0)
            if across or down:
                block.merges.append((i, j, i + down, j + across))
            data = cell.find(SS + "Data")
            if data is not None:
                kind, text = _a(data, "Type"), "".join(data.itertext())
                if kind == "Number":
                    try:
                        value = fe.number_text(float(text))
                    except ValueError:
                        value = text_raw(text)
                elif kind == "DateTime":
                    serial = _serial(text)
                    value = fe.number_text(serial) if serial is not None else text_raw(text)
                elif kind == "Boolean":
                    value = "TRUE" if text.strip() in ("1", "true", "TRUE") else "FALSE"
                elif kind == "Error":
                    value = text
                else:
                    value = text_raw(text) if (fmt or {}).get("num") != "@" else text
                if value is not None:
                    block.values[(i, j)] = value
            if _a(cell, "Formula"):
                block.formulas[(i, j)] = _a(cell, "Formula")
            if _a(cell, "HRef"):
                block.links[(i, j)] = {"target": _a(cell, "HRef")}
            note = cell.find(SS + "Comment")
            if note is not None:
                text = "".join(note.itertext()).strip()
                if text:
                    block.notes[(i, j)] = text
            j += across
    rows = [k[0] for k in block.values] + [k[0] for k in block.formats] + [m[2] for m in block.merges]
    cols = [k[1] for k in block.values] + [k[1] for k in block.formats] + [m[3] for m in block.merges]
    count_r, count_c = _a(table, "ExpandedRowCount"), _a(table, "ExpandedColumnCount")
    block.h = max([i + 1] + [r + 1 for r in rows] + ([int(count_r)] if count_r else []))
    block.w = max([1] + [c + 1 for c in cols] + ([int(count_c)] if count_c else []))
    return block if block.h > 0 else None


# ------------------------------------------------------------------ HTML

CSS_COLORS = {"black": "#000000", "windowtext": "#000000", "white": "#ffffff", "window": "#ffffff",
              "red": "#ff0000", "green": "#008000", "blue": "#0000ff", "yellow": "#ffff00", "gray": "#808080",
              "grey": "#808080", "silver": "#c0c0c0", "maroon": "#800000", "navy": "#000080", "olive": "#808000",
              "purple": "#800080", "teal": "#008080", "lime": "#00ff00", "aqua": "#00ffff", "cyan": "#00ffff",
              "fuchsia": "#ff00ff", "magenta": "#ff00ff", "orange": "#ffa500", "darkred": "#8b0000",
              "darkblue": "#00008b", "darkgreen": "#006400", "lightgray": "#d3d3d3", "lightgrey": "#d3d3d3",
              "darkgray": "#a9a9a9", "darkgrey": "#a9a9a9", "pink": "#ffc0cb", "brown": "#a52a2a",
              "gold": "#ffd700", "lightblue": "#add8e6", "lightgreen": "#90ee90", "lightyellow": "#ffffe0"}


def _css_color(value):
    value = (value or "").strip().lower()
    m = re.fullmatch(r"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+)\s*)?\)", value)
    if m:
        if m.group(4) is not None and float(m.group(4)) == 0:
            return None  # transparent
        return "#%02x%02x%02x" % tuple(min(255, int(m.group(k))) for k in (1, 2, 3))
    return _color(value)


def _css_unescape(value):
    """A CSS value with its quotes and backslash escapes removed (Excel writes mso-number-format:"\\#\\,\\#\\#0")."""
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
        value = value[1:-1]
    return re.sub(r"\\(.)", r"\1", value)


def parse_css(text):
    """{property: value} of a declaration block ("font-weight:700; color:red")."""
    out = {}
    for part in re.split(r";(?=(?:[^\"']|\"[^\"]*\"|'[^']*')*$)", text or ""):
        name, sep, value = part.partition(":")
        if sep and name.strip():
            out[name.strip().lower()] = value.strip()
    return out


def parse_stylesheet(text):
    """{selector: {property: value}} of the simple rules in a <style> (td, .xl65, td.name)."""
    rules = {}
    text = re.sub(r"/\*.*?\*/|<!--|-->", " ", text, flags=re.S)
    for m in re.finditer(r"([^{}@]+)\{([^{}]*)\}", text):
        decl = parse_css(m.group(2))
        for selector in m.group(1).split(","):
            selector = selector.strip().lower()
            if selector:
                rules.setdefault(selector, {}).update(decl)
    return rules


def _length_px(value):
    """Pixels of an HTML length: a number alone or in px, or in pt; None when there is none."""
    m = re.match(r"([\d.]+)\s*(pt|px)?", value.strip())
    if not m or not m.group(1).strip("."):
        return None
    n = float(m.group(1))
    return round(n * PT if m.group(2) == "pt" else n) or None


def _length_pt(value):
    m = re.match(r"([\d.]+)\s*(pt|px|em)?", value.strip())
    if not m:
        return None
    n = float(m.group(1))
    return n if m.group(2) in (None, "pt") else n * 0.75 if m.group(2) == "px" else n * 11


def _border(value):
    """[style, colour] of a CSS border shorthand ("0.5pt solid windowtext", "1px solid #000"), or None."""
    style, width, color = "solid", 0.75, "#000000"
    for token in re.findall(r"rgba?\([^)]*\)|\S+", value.lower()):
        if token in ("none", "hidden"):
            return None
        if token in ("solid", "double", "dashed", "dotted", "groove", "ridge", "inset", "outset"):
            style = token
        elif token in ("thin", "medium", "thick", "hairline"):
            width = {"thin": 0.75, "medium": 1.5, "thick": 2.25, "hairline": 0.5}[token]
        elif re.match(r"[\d.]+(pt|px|em)?$", token):
            width = _length_pt(token) or 0
            if width == 0:
                return None
        elif _css_color(token):
            color = _css_color(token)
    if style == "double":
        return ["double", color]
    if style in ("dashed", "dotted"):
        return [style, color]
    return ["thin" if width <= 0.8 else "medium" if width <= 1.6 else "thick", color]


def _format_css(css, fmt):
    """Apply the CSS properties of a cell to a format dict (in place)."""
    for name, value in css.items():
        low = value.lower().strip()
        if name == "font-family":
            family = _css_unescape(value.split(",")[0])
            if family:
                fmt["font"] = family
        elif name == "font-size":
            size = _length_pt(low)
            if size:
                fmt["size"] = round(size * 2) / 2
        elif name == "font-weight":
            fmt["bold"] = low in ("bold", "bolder") or (low.isdigit() and int(low) >= 600)
        elif name == "font-style":
            fmt["italic"] = low in ("italic", "oblique")
        elif name == "text-decoration" or name == "text-decoration-line":
            fmt["underline"] = "single" if "underline" in low else None
            fmt["strike"] = "line-through" in low
        elif name == "text-underline-style":
            if "double" in low:
                fmt["underline"] = "double"
        elif name == "color":
            fmt["fg"] = _css_color(low) or fmt.get("fg")
        elif name in ("background", "background-color"):
            color = next((c for c in map(_css_color, re.findall(r"rgba?\([^)]*\)|#\w+|[a-z]+", low)) if c), None)
            if color or low in ("none", "transparent"):
                fmt["bg"] = color
        elif name == "text-align":
            fmt["align"] = {"left": "left", "right": "right", "center": "center", "centeracross": "center",
                            "justify": "left"}.get(low)
        elif name == "vertical-align":
            fmt["valign"] = {"top": "top", "middle": "center", "center": "center"}.get(low)
        elif name == "white-space":
            fmt["wrap"] = low in ("normal", "pre-wrap", "pre-line")
        elif name == "mso-rotate":
            try:
                fmt["rotate"] = int(float(low)) if low != "0" else 0
            except ValueError:
                pass
        elif name == "mso-char-indent-count":
            try:
                fmt["indent"] = int(float(low))
            except ValueError:
                pass
        elif name == "mso-number-format":
            code = _css_unescape(value)
            fmt["num"] = NAMED_FORMATS[code] if code in NAMED_FORMATS else fe.excel_number_format(code)
        elif name == "border":
            for key in fe.BORDER_KEYS:
                fmt[key] = _border(value)
        elif name in ("border-top", "border-bottom", "border-left", "border-right"):
            fmt[{"top": "bt", "bottom": "bb", "left": "bl", "right": "br"}[name[7:]]] = _border(value)


class _TableReader(html.parser.HTMLParser):
    """The first table of an HTML document: rows of cells with their attributes and text, the <col>
    widths and the style sheet."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.styles, self.cols, self.rows = [], [], []
        self.table_style = None
        self.depth = 0           # nesting of tables (cells of tables inside cells are text of the outer cell)
        self.done = False
        self.cell = None
        self.in_style = False
        self.marks = []          # formatting tags open inside the cell: (tag, css)
        self.google = False

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "style":
            self.in_style = True
            self.styles.append("")
            return
        if tag == "google-sheets-html-origin":
            self.google = True
        if self.done:
            return
        if tag == "table":
            self.depth += 1
            if self.depth == 1:
                self.table_style = parse_css(a.get("style"))
            return
        if self.depth != 1:
            return
        if tag == "col":
            px = _length_px(a.get("width") or parse_css(a.get("style")).get("width") or "")
            for _ in range(_int(a.get("span"))):
                self.cols.append(px)
        elif tag == "tr":
            self.rows.append([])
        elif tag in ("td", "th"):
            if not self.rows:
                self.rows.append([])
            self.cell = {"attrs": a, "tag": tag, "text": [], "runs": [], "link": None}
            self.rows[-1].append(self.cell)
            self.marks = []
        elif self.cell is not None:
            if tag == "br":
                self.cell["text"].append(_BREAK)
            elif tag in ("p", "div") and self.cell["text"] and "".join(self.cell["text"]).strip():
                self.cell["text"].append(_BREAK)
            elif tag == "a" and a.get("href"):
                self.cell["link"] = a.get("href")
            if tag in ("b", "strong", "i", "em", "u", "s", "strike", "del", "font", "span", "a"):
                css = parse_css(a.get("style"))
                if tag in ("b", "strong"):
                    css["font-weight"] = "700"
                elif tag in ("i", "em"):
                    css["font-style"] = "italic"
                elif tag == "u":
                    css["text-decoration"] = "underline"
                elif tag in ("s", "strike", "del"):
                    css["text-decoration"] = "line-through"
                elif tag == "font":
                    if a.get("color"):
                        css["color"] = a["color"]
                    if a.get("face"):
                        css["font-family"] = a["face"]
                self.marks.append((tag, css))

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)

    def handle_endtag(self, tag):
        if tag == "style":
            self.in_style = False
            return
        if self.done:
            return
        if tag == "table":
            self.depth -= 1
            if self.depth == 0 and self.rows:
                self.done = True
            return
        if self.depth != 1:
            return
        if tag in ("td", "th"):
            self.cell = None
        elif self.cell is not None and self.marks:
            for k in range(len(self.marks) - 1, -1, -1):
                if self.marks[k][0] == tag:
                    del self.marks[k:]
                    break

    def handle_data(self, data):
        if self.in_style:
            self.styles[-1] += data
        elif self.cell is not None and not self.done:
            self.cell["text"].append(data)
            if data.strip():
                css = {}
                for _, mark in self.marks:
                    css.update(mark)
                self.cell["runs"].append(css)


_BREAK = "\x00"  # a <br> in a cell's text (the line breaks of the HTML source are only spaces)


def _html_text(parts):
    text = re.sub(r"[ \t\r\n\f]+", " ", "".join(parts))
    lines = [line.strip() for line in text.split(_BREAK)]
    return "\n".join(lines).strip("\n").replace("\xa0", " ")


def html_fragment(data):
    """The HTML of Windows' "HTML Format" (a header of byte offsets, then the page) as text."""
    if isinstance(data, bytes):
        m = re.search(rb"StartHTML:(\d+)", data[:300])
        start = int(m.group(1)) if m and int(m.group(1)) > 0 else 0
        data = data[start:].rstrip(b"\0")
        if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
            return data.decode("utf-16")
        return data.decode("utf-8", "replace")
    return data


def from_html(data):
    """The Block of the first table in an HTML document (bytes of Windows' "HTML Format" or text), or None."""
    reader = _TableReader()
    try:
        reader.feed(html_fragment(data))
        reader.close()
    except (AssertionError, ValueError):
        return None
    if not reader.rows or not any(reader.rows):
        return None
    rules = {}
    for sheet in reader.styles:
        for selector, decl in parse_stylesheet(sheet).items():
            rules.setdefault(selector, {}).update(decl)
    base = {}
    if reader.table_style:
        _format_css({k: v for k, v in reader.table_style.items() if k.startswith("font")}, base)
    td_rule = dict(rules.get("td", {}))
    if reader.google:  # Google Sheets' style sheet draws its gridlines as cell borders
        for key in [k for k in td_rule if k.startswith("border")]:
            del td_rule[key]
    _format_css(td_rule, base)
    block = Block()
    block.source = "Google Sheets" if reader.google else "HTML"
    taken = set()
    made = {}
    for i, row in enumerate(reader.rows):
        j = 0
        for cell in row:
            while (i, j) in taken:
                j += 1
            a = cell["attrs"]
            rs, cs = max(1, _int(a.get("rowspan"))), max(1, _int(a.get("colspan")))
            if rs > 1 or cs > 1:
                block.merges.append((i, j, i + rs - 1, j + cs - 1))
            taken.update((i + di, j + dj) for di in range(rs) for dj in range(cs))
            css = {}
            for cls in (a.get("class") or "").split():
                css.update(rules.get("." + cls.lower(), {}))
                css.update(rules.get("td." + cls.lower(), {}))
            css.update(parse_css(a.get("style")))
            google = _json(a.get("data-sheets-numberformat"))
            if isinstance(google, dict) and isinstance(google.get("2"), str):
                css["mso-number-format"] = '"' + google["2"] + '"'
            runs = cell["runs"]
            if runs:  # formatting all the cell's text shares (<b>...</b> around it all)
                shared = dict(runs[0])
                for run in runs[1:]:
                    shared = {k: v for k, v in shared.items() if run.get(k) == v}
                css.update(shared)
            if cell["tag"] == "th" and "font-weight" not in css:
                css["font-weight"] = "700"
            key = tuple(sorted(css.items()))
            fmt = made.get(key)
            if fmt is None:
                full = dict(base)
                _format_css(css, full)
                if not full.get("align") and a.get("align"):
                    full["align"] = {"left": "left", "center": "center", "right": "right"}.get(a["align"].lower())
                fmt = made[key] = clean_format(full)
            if fmt:
                block.formats[(i, j)] = fmt
            text = _html_text(cell["text"])
            value = _typed_value(a, text)
            if value is None and text:  # the text as if typed ("1,234" is a number), unless it is Text
                value = text if (fmt or {}).get("num") == "@" or text[:1] not in "='" else "'" + text
            if value is not None:
                block.values[(i, j)] = value
            if (a.get("data-sheets-formula") or "").startswith("="):  # Google Sheets: R[-1]C[0] form
                block.formulas[(i, j)] = a["data-sheets-formula"]
            elif a.get("x:fmla"):
                pass  # (Excel's A1 formula for where the cell was: not usable elsewhere)
            if cell["link"]:
                block.links[(i, j)] = {"target": cell["link"]}
            j += cs
        block.w = max(block.w, j)
    block.h = len(reader.rows)
    block.w = max([block.w] + [m[3] + 1 for m in block.merges])
    for j, px in enumerate(reader.cols):
        if px:
            block.col_widths[j] = max(fe.MIN_COL_WIDTH, px)
    return block


def _json(text):
    if not text:
        return None
    try:
        import json
        return json.loads(text)
    except ValueError:
        return None


def _typed_value(attrs, text):
    """The value an HTML cell gives with its type: Excel's x:num, Google Sheets' data-sheets-value."""
    number = attrs.get("x:num")
    if number is not None:
        try:
            return fe.number_text(float(number.strip() or text.replace(",", "")))
        except ValueError:
            pass
    google = _json(attrs.get("data-sheets-value"))
    if isinstance(google, dict):
        if isinstance(google.get("3"), (int, float)) and not isinstance(google.get("3"), bool):
            return fe.number_text(float(google["3"]))
        if "4" in google:
            return "TRUE" if google["4"] else "FALSE"
        if isinstance(google.get("2"), str):
            return text_raw(google["2"])
    return None


def _int(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return 1
