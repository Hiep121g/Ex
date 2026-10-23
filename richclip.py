"""Copied cells in the clipboard formats Excel offers besides plain text.

Excel puts a block of cells on the clipboard as an HTML table, an RTF table and a picture (an
enhanced metafile plus a bitmap) along with tab-separated text; programs pick the richest format they
understand. AutoCAD, for one, pastes nothing useful from text alone. This module builds those
formats from a "table" description and, on Windows, writes them all to the clipboard.

A table is {"cols": [width px], "rows": [height px], "cells": [[cell]], "merges": [(r1, c1, r2, c2)]},
sizes in pixels at 96 DPI and merges in table coordinates. A cell is {"text": str, "align": "left" |
"center" | "right", "fmt": cell format dict as stored in a Sheet (font, size, bold, fg, bg, borders...)}.
"""

import html
import re
import sys

DEFAULT_FONT = "Calibri"
DEFAULT_SIZE = 11
BORDER_PT = {"thin": 0.5, "medium": 1.0, "thick": 1.5, "double": 2.25, "dashed": 0.5, "dotted": 0.5}
BORDER_PX = {"medium": 2, "thick": 3}
EDGES = (("bt", "top"), ("bb", "bottom"), ("bl", "left"), ("br", "right"))


def _merge_info(table):
    """({anchor: (rowspan, colspan)}, set of cells hidden under a merge)."""
    spans, covered = {}, set()
    for r1, c1, r2, c2 in table["merges"]:
        spans[(r1, c1)] = (r2 - r1 + 1, c2 - c1 + 1)
        covered.update((r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if (r, c) != (r1, c1))
    return spans, covered


# ------------------------------------------------------------------ HTML

def _css(cell):
    fmt = cell["fmt"]
    number = fmt.get("num")
    style = [f"font-family:'{fmt.get('font') or DEFAULT_FONT}'", f"font-size:{fmt.get('size') or DEFAULT_SIZE}pt",
             f"text-align:{cell['align']}", f"vertical-align:{'middle' if fmt.get('valign') == 'center' else fmt.get('valign') or 'bottom'}",
             "white-space:normal" if fmt.get("wrap") else "white-space:nowrap"]
    if fmt.get("bold"):
        style.append("font-weight:700")
    if fmt.get("italic"):
        style.append("font-style:italic")
    lines = [d for key, d in (("underline", "underline"), ("strike", "line-through")) if fmt.get(key)]
    if lines:
        style.append("text-decoration:" + " ".join(lines))
    if fmt.get("fg"):
        style.append(f"color:{fmt['fg']}")
    if fmt.get("bg"):
        style.append(f"background:{fmt['bg']}")
    for key, side in EDGES:
        spec = fmt.get(key)
        if spec:
            kind = {"double": "double", "dashed": "dashed", "dotted": "dotted"}.get(spec[0], "solid")
            style.append(f"border-{side}:{BORDER_PT.get(spec[0], 0.5)}pt {kind} {spec[1] if len(spec) > 1 else '#000000'}")
    if number:  # Excel's own property: the number format (backslash-escaped, as Excel writes it)
        style.append('mso-number-format:"' + re.sub(r"([^\w\s])", r"\\\1", number) + '"')
    return ";".join(style)


def html_table(table):
    spans, covered = _merge_info(table)
    out = ['<table border=0 cellpadding=0 cellspacing=0 style="border-collapse:collapse;table-layout:fixed">']
    out += [f'<col width={w} style="width:{w * 0.75:g}pt">' for w in table["cols"]]
    for r, (height, row) in enumerate(zip(table["rows"], table["cells"])):
        out.append(f'<tr height={height} style="height:{height * 0.75:g}pt">')
        for c, cell in enumerate(row):
            if (r, c) in covered:
                continue
            rs, cs = spans.get((r, c), (1, 1))
            attrs = (f" rowspan={rs}" if rs > 1 else "") + (f" colspan={cs}" if cs > 1 else "")
            text = html.escape(cell["text"]).replace("\n", "<br>")
            if cell.get("value") is not None:  # the number itself, which Excel reads instead of the text shown
                attrs += f' x:num="{cell["value"]!r}"'
            out.append(f'<td{attrs} style="{html.escape(_css(cell))}">{text}</td>')
        out.append("</tr>")
    out.append("</table>")
    return "\n".join(out)


def html_clipboard(table):
    """The table in Windows' "HTML Format": UTF-8 behind a header giving byte offsets."""
    header = "Version:0.9\r\nStartHTML:{:010d}\r\nEndHTML:{:010d}\r\nStartFragment:{:010d}\r\nEndFragment:{:010d}\r\n"
    before = ('<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"></head><body>'
              "<!--StartFragment-->")
    after = "<!--EndFragment--></body></html>"
    fragment = html_table(table).encode("utf-8")
    start_html = len(header.format(0, 0, 0, 0))
    start_frag = start_html + len(before.encode("utf-8"))
    end_frag = start_frag + len(fragment)
    end_html = end_frag + len(after.encode("utf-8"))
    return (header.format(start_html, end_html, start_frag, end_frag).encode("ascii") + before.encode("utf-8")
            + fragment + after.encode("utf-8"))


# ------------------------------------------------------------------ RTF

def _rtf_text(text):
    out = []
    for ch in text:
        if ch in "\\{}":
            out.append("\\" + ch)
        elif ch == "\n":
            out.append("\\line ")
        elif ord(ch) < 128:
            out.append(ch)
        else:  # RTF takes 16-bit signed code units, with a "?" for readers without Unicode
            for unit in _utf16_units(ch):
                out.append(f"\\u{unit - 65536 if unit > 32767 else unit}?")
    return "".join(out)


def _utf16_units(ch):
    data = ch.encode("utf-16-le")
    return [int.from_bytes(data[i:i + 2], "little") for i in range(0, len(data), 2)]


def rtf_table(table):
    """The table as RTF (twips = pixels at 96 DPI x 15), with merges, borders, colours and fonts."""
    spans, covered = _merge_info(table)
    fonts, colors = [DEFAULT_FONT], ["#000000"]

    def font(name):
        name = name or DEFAULT_FONT
        if name not in fonts:
            fonts.append(name)
        return fonts.index(name)

    def color(value):
        if value not in colors:
            colors.append(value)
        return colors.index(value) + 1  # \cf0 / \cb0 mean "automatic"

    owner = {}  # cell -> the merge anchor it belongs to
    for (r1, c1), (rs, cs) in spans.items():
        for r in range(r1, r1 + rs):
            for c in range(c1, c1 + cs):
                owner[(r, c)] = (r1, c1)
    body = []
    for r, (height, row) in enumerate(zip(table["rows"], table["cells"])):
        body.append(f"\\trowd\\trgaph45\\trleft0\\trrh{height * 15}")
        right = 0
        for c, cell in enumerate(row):
            fmt = cell["fmt"]
            anchor = owner.get((r, c))
            if anchor:
                rs, cs = spans[anchor]
                if cs > 1:
                    body.append("\\clmgf" if c == anchor[1] else "\\clmrg")
                if rs > 1:
                    body.append("\\clvmgf" if r == anchor[0] else "\\clvmrg")
            body.append({"top": "\\clvertalt", "center": "\\clvertalc"}.get(fmt.get("valign"), "\\clvertalb"))
            for key, side in EDGES:
                spec = fmt.get(key)
                if spec:
                    kind = {"double": "\\brdrdb", "dashed": "\\brdrdash", "dotted": "\\brdrdot"}.get(spec[0], "\\brdrs")
                    width = round(BORDER_PT.get(spec[0], 0.5) * 20)
                    body.append(f"\\clbrdr{side[0]}{kind}\\brdrw{width}\\brdrcf{color(spec[1] if len(spec) > 1 else '#000000')}")
            if fmt.get("bg"):
                body.append(f"\\clcbpat{color(fmt['bg'])}")
            right += table["cols"][c] * 15
            body.append(f"\\cellx{right}")
        for c, cell in enumerate(row):
            if (r, c) in covered:
                body.append("\\pard\\intbl\\cell")
                continue
            fmt = cell["fmt"]
            style = f"\\f{font(fmt.get('font'))}\\fs{round((fmt.get('size') or DEFAULT_SIZE) * 2)}"
            style += "\\b" if fmt.get("bold") else ""
            style += "\\i" if fmt.get("italic") else ""
            style += {"double": "\\uldb", None: ""}.get(fmt.get("underline"), "\\ul")
            style += "\\strike" if fmt.get("strike") else ""
            style += f"\\cf{color(fmt['fg'])}" if fmt.get("fg") else ""
            align = {"center": "\\qc", "right": "\\qr"}.get(cell["align"], "\\ql")
            body.append(f"\\pard\\intbl{align}{{{style} {_rtf_text(cell['text'])}}}\\cell")
        body.append("\\row")
    # (names like 游ゴシック go through \\u escapes too: RTF is 7-bit)
    font_table = "".join(f"{{\\f{i}\\fnil\\fcharset0 {_rtf_text(name)};}}" for i, name in enumerate(fonts))
    color_table = ";" + "".join(f"\\red{int(c[1:3], 16)}\\green{int(c[3:5], 16)}\\blue{int(c[5:7], 16)};"
                                for c in colors)
    return (f"{{\\rtf1\\ansi\\ansicpg1252\\deff0{{\\fonttbl{font_table}}}{{\\colortbl{color_table}}}\n"
            + "\n".join(body) + "\n}")


# ------------------------------------------------------------------ picture (Windows GDI)

def spill_right(table, r, c):
    """Index of the last column that the text of cell (r, c) may run over, as in Excel: a left-aligned,
    unwrapped text runs on over the empty cells to its right."""
    cell, row = table["cells"][r][c], table["cells"][r]
    if cell["align"] != "left" or cell["fmt"].get("wrap") or not cell["text"]:
        return c
    _, covered = _merge_info(table)
    merged = covered | {(m[0], m[1]) for m in table["merges"]}
    end = c
    while end + 1 < len(row) and not row[end + 1]["text"] and (r, end + 1) not in merged:
        end += 1
    return end


def _colorref(color):
    return int(color[1:3], 16) | int(color[3:5], 16) << 8 | int(color[5:7], 16) << 16


class _Gdi:
    """The few user32 / gdi32 / kernel32 calls used here, with 64-bit-safe prototypes."""

    def __init__(self, ctypes):
        from ctypes import wintypes as w

        self.ctypes = ctypes
        self.RECT = w.RECT
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        handle, prect = ctypes.c_void_p, ctypes.POINTER(w.RECT)
        protos = {
            (user32, "OpenClipboard"): ([w.HWND], w.BOOL), (user32, "EmptyClipboard"): ([], w.BOOL),
            (user32, "CloseClipboard"): ([], w.BOOL), (user32, "SetClipboardData"): ([w.UINT, handle], handle),
            (user32, "RegisterClipboardFormatW"): ([w.LPCWSTR], w.UINT),
            (user32, "GetClipboardData"): ([w.UINT], handle), (user32, "IsClipboardFormatAvailable"): ([w.UINT], w.BOOL),
            (kernel32, "GlobalSize"): ([handle], ctypes.c_size_t),
            (user32, "GetDC"): ([w.HWND], handle), (user32, "ReleaseDC"): ([w.HWND, handle], ctypes.c_int),
            (user32, "FillRect"): ([handle, prect, handle], ctypes.c_int),
            (user32, "DrawTextW"): ([handle, w.LPCWSTR, ctypes.c_int, prect, w.UINT], ctypes.c_int),
            (gdi32, "GetDeviceCaps"): ([handle, ctypes.c_int], ctypes.c_int),
            (gdi32, "CreateEnhMetaFileW"): ([handle, w.LPCWSTR, prect, w.LPCWSTR], handle),
            (gdi32, "CloseEnhMetaFile"): ([handle], handle), (gdi32, "DeleteEnhMetaFile"): ([handle], w.BOOL),
            (gdi32, "PlayEnhMetaFile"): ([handle, handle, prect], w.BOOL),
            (gdi32, "CreateFontW"): ([ctypes.c_int] * 5 + [w.DWORD] * 8 + [w.LPCWSTR], handle),
            (gdi32, "CreatePen"): ([ctypes.c_int, ctypes.c_int, w.COLORREF], handle),
            (gdi32, "CreateSolidBrush"): ([w.COLORREF], handle),
            (gdi32, "SelectObject"): ([handle, handle], handle), (gdi32, "DeleteObject"): ([handle], w.BOOL),
            (gdi32, "SetTextColor"): ([handle, w.COLORREF], w.COLORREF),
            (gdi32, "SetBkMode"): ([handle, ctypes.c_int], ctypes.c_int),
            (gdi32, "MoveToEx"): ([handle, ctypes.c_int, ctypes.c_int, ctypes.c_void_p], w.BOOL),
            (gdi32, "LineTo"): ([handle, ctypes.c_int, ctypes.c_int], w.BOOL),
            (gdi32, "CreateCompatibleDC"): ([handle], handle), (gdi32, "DeleteDC"): ([handle], w.BOOL),
            (gdi32, "CreateCompatibleBitmap"): ([handle, ctypes.c_int, ctypes.c_int], handle),
            (kernel32, "GlobalAlloc"): ([w.UINT, ctypes.c_size_t], handle),
            (kernel32, "GlobalLock"): ([handle], handle), (kernel32, "GlobalUnlock"): ([handle], w.BOOL),
            (kernel32, "GlobalFree"): ([handle], handle),
        }
        for (dll, name), (args, result) in protos.items():
            func = getattr(dll, name)
            func.argtypes, func.restype = args, result
            setattr(self, name, func)

    def rect(self, x0, y0, x1, y1):
        return self.RECT(round(x0), round(y0), round(x1), round(y1))


def _draw(gdi, dc, table, sx, sy):
    """Paint the table on a device context whose units are pixels at sx, sy times 96 DPI."""
    xs, ys = [0.0], [0.0]
    for w in table["cols"]:
        xs.append(xs[-1] + w * sx)
    for h in table["rows"]:
        ys.append(ys[-1] + h * sy)
    spans, covered = _merge_info(table)
    owner = {}
    for (r1, c1), (rs, cs) in spans.items():
        for r in range(r1, r1 + rs):
            for c in range(c1, c1 + cs):
                owner[(r, c)] = (r1, c1)

    def box(r, c):
        rs, cs = spans.get((r, c), (1, 1))
        return xs[c], ys[r], xs[c + cs], ys[r + rs]

    cells = [(r, c, cell) for r, row in enumerate(table["cells"]) for c, cell in enumerate(row)
             if (r, c) not in covered]
    for r, c, cell in cells:
        if cell["fmt"].get("bg"):
            brush = gdi.CreateSolidBrush(_colorref(cell["fmt"]["bg"]))
            gdi.FillRect(dc, gdi.ctypes.byref(gdi.rect(*box(r, c))), brush)
            gdi.DeleteObject(brush)
    gdi.SetBkMode(dc, 1)  # TRANSPARENT
    pad = 3 * sx
    for r, c, cell in cells:
        if not cell["text"]:
            continue
        fmt = cell["fmt"]
        x0, y0, x1, y1 = box(r, c)
        if (r, c) not in spans:
            x1 = xs[spill_right(table, r, c) + 1]
        font = gdi.CreateFontW(-round((fmt.get("size") or DEFAULT_SIZE) * 96 * sy / 72), 0, 0, 0,
                               700 if fmt.get("bold") else 400, bool(fmt.get("italic")), bool(fmt.get("underline")),
                               bool(fmt.get("strike")), 1, 0, 0, 5, 0, fmt.get("font") or DEFAULT_FONT)
        old = gdi.SelectObject(dc, font)
        gdi.SetTextColor(dc, _colorref(fmt.get("fg") or "#000000"))
        flags = 0x800 | {"center": 0x1, "right": 0x2}.get(cell["align"], 0x0)  # DT_NOPREFIX | alignment
        if fmt.get("wrap") or "\n" in cell["text"]:
            flags |= 0x10  # DT_WORDBREAK (multi-line text starts at the top)
        else:
            flags |= 0x20 | {"top": 0x0, "center": 0x4}.get(fmt.get("valign"), 0x8)  # DT_SINGLELINE | v-align
        text = cell["text"].replace("\n", "\r\n")
        gdi.DrawTextW(dc, text, -1, gdi.ctypes.byref(gdi.rect(x0 + pad, y0 + 1, x1 - pad, y1 - 1)), flags)
        gdi.DeleteObject(gdi.SelectObject(dc, old))
    for r, c, cell in cells:
        x0, y0, x1, y1 = box(r, c)
        for key, _ in EDGES:
            spec = cell["fmt"].get(key)
            if not spec:
                continue
            style, color = spec[0], _colorref(spec[1] if len(spec) > 1 else "#000000")
            width = max(1, round(BORDER_PX.get(style, 1) * sx))
            pen = gdi.CreatePen({"dashed": 1, "dotted": 2}.get(style, 0), width, color)
            old = gdi.SelectObject(dc, pen)
            line = {"bt": (x0, y0, x1, y0), "bb": (x0, y1, x1, y1), "bl": (x0, y0, x0, y1), "br": (x1, y0, x1, y1)}[key]
            offsets = (0, 2 * sx) if style == "double" else (0,)
            for off in offsets:
                dx, dy = {"bt": (0, off), "bb": (0, -off), "bl": (off, 0), "br": (-off, 0)}[key]
                gdi.MoveToEx(dc, round(line[0] + dx), round(line[1] + dy), None)
                gdi.LineTo(dc, round(line[2] + dx), round(line[3] + dy))
            gdi.DeleteObject(gdi.SelectObject(dc, old))
    return round(xs[-1]) + 1, round(ys[-1]) + 1


def _set_data(gdi, fmt, data):
    """Hand a block of bytes to the clipboard under format `fmt` (the clipboard then owns it)."""
    handle = gdi.GlobalAlloc(0x0002, len(data))  # GMEM_MOVEABLE
    if not handle:
        raise OSError("GlobalAlloc")
    ptr = gdi.GlobalLock(handle)
    gdi.ctypes.memmove(ptr, data, len(data))
    gdi.GlobalUnlock(handle)
    if not gdi.SetClipboardData(fmt, handle):
        gdi.GlobalFree(handle)
        raise OSError("SetClipboardData")


def _pictures(gdi, table):
    """(enhanced metafile, bitmap) handles showing the table, drawn at the screen's DPI."""
    screen = gdi.GetDC(None)
    try:
        sx, sy = gdi.GetDeviceCaps(screen, 88) / 96, gdi.GetDeviceCaps(screen, 90) / 96  # LOGPIXELSX / Y
        width = round(sum(table["cols"]) * sx) + 1
        height = round(sum(table["rows"]) * sy) + 1
        # the metafile frame is in 0.01 mm, measured on the reference (screen) device
        mm_x = gdi.GetDeviceCaps(screen, 4) / max(1, gdi.GetDeviceCaps(screen, 8))  # HORZSIZE / HORZRES
        mm_y = gdi.GetDeviceCaps(screen, 6) / max(1, gdi.GetDeviceCaps(screen, 10))  # VERTSIZE / VERTRES
        frame = gdi.rect(0, 0, width * mm_x * 100, height * mm_y * 100)
        dc = gdi.CreateEnhMetaFileW(screen, None, gdi.ctypes.byref(frame), None)
        if not dc:
            raise OSError("CreateEnhMetaFileW")
        _draw(gdi, dc, table, sx, sy)
        emf = gdi.CloseEnhMetaFile(dc)
        if not emf:
            raise OSError("CloseEnhMetaFile")
        mem = gdi.CreateCompatibleDC(screen)
        bitmap = gdi.CreateCompatibleBitmap(screen, width, height)
        old = gdi.SelectObject(mem, bitmap)
        white = gdi.CreateSolidBrush(0xFFFFFF)
        area = gdi.rect(0, 0, width, height)
        gdi.FillRect(mem, gdi.ctypes.byref(area), white)
        gdi.DeleteObject(white)
        gdi.PlayEnhMetaFile(mem, emf, gdi.ctypes.byref(area))
        gdi.SelectObject(mem, old)
        gdi.DeleteDC(mem)
        return emf, bitmap
    finally:
        gdi.ReleaseDC(None, screen)


def copy_to_clipboard(hwnd, text, table):
    """Put `text` and the table as HTML, RTF and a picture on the Windows clipboard, on behalf of window
    `hwnd`. Returns False (writing nothing) off Windows; raises OSError when the clipboard fails."""
    if sys.platform != "win32":
        return False
    import ctypes

    gdi = _Gdi(ctypes)
    emf, bitmap = _pictures(gdi, table)
    for _ in range(10):  # another program may hold the clipboard for a moment
        if gdi.OpenClipboard(hwnd):
            break
        import time
        time.sleep(0.02)
    else:
        gdi.DeleteEnhMetaFile(emf)
        gdi.DeleteObject(bitmap)
        raise OSError("OpenClipboard")
    try:
        gdi.EmptyClipboard()
        _set_data(gdi, 13, (text.replace("\r\n", "\n").replace("\n", "\r\n") + "\0").encode("utf-16-le"))  # CF_UNICODETEXT
        _set_data(gdi, gdi.RegisterClipboardFormatW("HTML Format"), html_clipboard(table) + b"\0")
        _set_data(gdi, gdi.RegisterClipboardFormatW("Rich Text Format"), rtf_table(table).encode("ascii") + b"\0")
        if not gdi.SetClipboardData(14, emf):  # CF_ENHMETAFILE (Windows adds the old metafile format itself)
            gdi.DeleteEnhMetaFile(emf)
        if not gdi.SetClipboardData(2, bitmap):  # CF_BITMAP (and Windows adds CF_DIB)
            gdi.DeleteObject(bitmap)
    finally:
        gdi.CloseClipboard()
    return True


def read_cells(hwnd):
    """The cell formats another program put on the Windows clipboard: {"xml": Excel's XML Spreadsheet,
    "html": HTML Format} (bytes; those present), or {} (also off Windows)."""
    if sys.platform != "win32":
        return {}
    import ctypes

    gdi = _Gdi(ctypes)
    kinds = {key: gdi.RegisterClipboardFormatW(name) for key, name in (("xml", "XML Spreadsheet"),
                                                                        ("html", "HTML Format"))}
    kinds = {key: kind for key, kind in kinds.items() if kind and gdi.IsClipboardFormatAvailable(kind)}
    if not kinds:
        return {}
    for _ in range(10):  # another program may hold the clipboard for a moment
        if gdi.OpenClipboard(hwnd):
            break
        import time
        time.sleep(0.02)
    else:
        return {}
    found = {}
    try:
        for key, kind in kinds.items():
            handle = gdi.GetClipboardData(kind)
            ptr = gdi.GlobalLock(handle) if handle else None
            if not ptr:
                continue
            try:
                found[key] = ctypes.string_at(ptr, gdi.GlobalSize(handle)).rstrip(b"\0")
            finally:
                gdi.GlobalUnlock(handle)
    finally:
        gdi.CloseClipboard()
    return found


def read_image(hwnd):
    """PNG data of a picture on the Windows clipboard (a screenshot, a picture copied from a web page or
    Paint), or None when there is none (or not on Windows)."""
    if sys.platform != "win32":
        return None
    import ctypes
    import images

    gdi = _Gdi(ctypes)
    png = gdi.RegisterClipboardFormatW("PNG")
    kinds = [k for k in (png, 17, 8) if gdi.IsClipboardFormatAvailable(k)]  # PNG, CF_DIBV5, CF_DIB
    if not kinds or not gdi.OpenClipboard(hwnd):
        return None
    try:
        for kind in kinds:
            handle = gdi.GetClipboardData(kind)
            if not handle:
                continue
            size = gdi.GlobalSize(handle)
            ptr = gdi.GlobalLock(handle)
            if not ptr:
                continue
            try:
                data = ctypes.string_at(ptr, size)
            finally:
                gdi.GlobalUnlock(handle)
            if kind == png and data[:8] == b"\x89PNG\r\n\x1a\n":
                return data
            if kind != png:
                found = images.png_from_dib(data)
                if found:
                    return found
    finally:
        gdi.CloseClipboard()
    return None
