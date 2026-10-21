"""Excel-style ribbon widgets for PyExcel.

Everything is drawn with tkinter primitives (no image files): the green tab
strip, groups with captions and dialog launchers, buttons with vector icons
(large, small, split and dropdown variants), colour palettes, galleries and
tooltips.
"""

import math
import os
import tkinter as tk
import tkinter.font as tkfont

RIBBON_BG = "#f3f3f3"
TAB_BAR_BG = "#217346"
TAB_HOVER_BG = "#2d8656"
HOVER_BG = "#dcdcdc"
HOVER_SOFT = "#e8e8e8"
PRESSED_BG = "#c4c4c4"
CHECKED_BG = "#c6c6c6"
CHECKED_LINE = "#a8a8a8"
GROUP_LINE = "#d4d4d4"
GROUP_FG = "#5f5f5f"
TEXT_FG = "#262626"

DARK = "#444444"
GRAY = "#7a7a7a"
BLUE = "#2b579a"
GREEN = "#217346"
RED = "#c42b1c"

THEME_COLORS = ["#ffffff", "#000000", "#e7e6e6", "#44546a", "#4472c4", "#ed7d31", "#a5a5a5", "#ffc000",
                "#5b9bd5", "#70ad47"]
STANDARD_COLORS = ["#c00000", "#ff0000", "#ffc000", "#ffff00", "#92d050", "#00b050", "#00b0f0", "#0070c0",
                   "#002060", "#7030a0"]


def mix(c1, c2, t):
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def theme_palette():
    """Excel's 10 x 6 theme grid: base colours, then lighter and darker variants."""
    rows = [THEME_COLORS]
    for step in range(5):
        row = []
        for base in THEME_COLORS:
            if base == "#ffffff":
                row.append(mix(base, "#000000", (0.05, 0.15, 0.25, 0.35, 0.5)[step]))
            elif base == "#000000":
                row.append(mix(base, "#ffffff", (0.5, 0.35, 0.25, 0.15, 0.05)[step]))
            elif step < 3:
                row.append(mix(base, "#ffffff", (0.8, 0.6, 0.4)[step]))
            else:
                row.append(mix(base, "#000000", (0.25, 0.5)[step - 3]))
        rows.append(row)
    return rows


class RibbonStyle:
    """Sizes and fonts shared by the ribbon widgets."""

    def __init__(self, root, scale):
        families = set(tkfont.families(root))
        self.family = "Segoe UI" if "Segoe UI" in families else "Arial"
        self.scale = scale
        self.bg = RIBBON_BG
        self.font = tkfont.Font(root=root, family=self.family, size=-round(12 * scale))
        self.small_font = tkfont.Font(root=root, family=self.family, size=-round(11 * scale))
        self.tab_font = tkfont.Font(root=root, family=self.family, size=-round(12 * scale))
        self.row_h = round(23 * scale)
        self.content_h = 3 * self.row_h
        self.label_h = round(17 * scale)

    def px(self, v):
        return round(v * self.scale)


# ------------------------------------------------------------------ icons

def _glyph(cv, P, k, x, y, text, size, color=DARK, weight="normal", slant="roman", family="Segoe UI",
           underline=False, angle=0):
    font = (family, -max(6, round(size * k)), weight) if slant == "roman" else (family, -max(6, round(size * k)), weight, "italic")
    if underline:
        font = font + ("underline",)
    cv.create_text(*P(x, y), text=text, font=font, fill=color, angle=angle)


def _lines(cv, P, k, rows, color=DARK):
    for x1, y, x2 in rows:
        cv.create_line(*P(x1, y, x2, y), fill=color, width=max(1, round(k)))


def _grid(cv, P, k, x0, y0, x1, y1, n=3, color=GRAY, fill=None):
    cv.create_rectangle(*P(x0, y0, x1, y1), outline=color, fill=fill or "")
    for i in range(1, n):
        cv.create_line(*P(x0 + (x1 - x0) * i / n, y0, x0 + (x1 - x0) * i / n, y1), fill=color)
        cv.create_line(*P(x0, y0 + (y1 - y0) * i / n, x1, y0 + (y1 - y0) * i / n), fill=color)


def _arrow(cv, P, k, x1, y1, x2, y2, color=BLUE, width=1.5):
    cv.create_line(*P(x1, y1, x2, y2), fill=color, width=max(1, round(width * k)), arrow="last",
                   arrowshape=(round(4 * k), round(5 * k), round(2 * k)))


def _icon_paste(cv, P, k, c):
    cv.create_rectangle(*P(2, 2.5, 11.5, 15), fill="#d9b278", outline="#8a6a3a")
    cv.create_rectangle(*P(4.5, 1, 9, 4), fill="#8c8c8c", outline="#5c5c5c")
    cv.create_rectangle(*P(7, 6, 15, 15.5), fill="#ffffff", outline="#6b6b6b")
    _lines(cv, P, k, [(8.5, 8.5, 13.5), (8.5, 10.5, 13.5), (8.5, 12.5, 12)], "#9a9a9a")


def _icon_cut(cv, P, k, c):
    w = max(1, round(1.2 * k))
    cv.create_line(*P(5, 10, 11, 1.5), fill=DARK, width=w)
    cv.create_line(*P(11, 10, 5, 1.5), fill=DARK, width=w)
    cv.create_oval(*P(2, 9.5, 6.5, 14), outline=BLUE, width=w)
    cv.create_oval(*P(9.5, 9.5, 14, 14), outline=BLUE, width=w)


def _icon_copy(cv, P, k, c):
    cv.create_rectangle(*P(2, 1.5, 10, 11.5), fill="#ffffff", outline=GRAY)
    cv.create_rectangle(*P(6, 4.5, 14, 14.5), fill="#ffffff", outline=BLUE)
    _lines(cv, P, k, [(7.5, 7.5, 12.5), (7.5, 9.5, 12.5), (7.5, 11.5, 11)], "#8aa4cf")


def _icon_painter(cv, P, k, c):
    cv.create_rectangle(*P(2, 1.5, 13, 6), fill="#ffd65c", outline="#b38f00")
    cv.create_rectangle(*P(3.5, 6, 11.5, 8), fill="#a6a6a6", outline="#7a7a7a")
    cv.create_line(*P(7.5, 8, 7.5, 15), fill="#8a5a2b", width=max(2, round(2.2 * k)))


def _icon_grow(cv, P, k, c, grow=True):
    _glyph(cv, P, k, 6, 9.5, "A", 13 if grow else 10, DARK)
    if grow:
        cv.create_polygon(*P(11, 6, 13.5, 2.5, 16, 6), fill=DARK)
    else:
        cv.create_polygon(*P(11, 3, 13.5, 6.5, 16, 3), fill=DARK)


def _icon_borders(cv, P, k, c):
    cv.create_rectangle(*P(2, 2, 14, 14), outline="#a0a0a0", dash=".")
    cv.create_line(*P(8, 2, 8, 14), fill="#a0a0a0", dash=".")
    cv.create_line(*P(2, 8, 14, 8), fill="#a0a0a0", dash=".")
    cv.create_line(*P(1.5, 14, 14.5, 14), fill="#000000", width=max(2, round(1.6 * k)))


def _icon_fill(cv, P, k, c):
    cv.create_polygon(*P(2.5, 6.5, 7.5, 1.5, 12.5, 6.5, 7.5, 11.5), fill="#ffffff", outline=DARK)
    cv.create_line(*P(4, 6.5, 11, 6.5), fill=GRAY)
    cv.create_polygon(*P(13, 7.5, 14.5, 10.5, 11.5, 10.5), fill=DARK)
    cv.create_rectangle(*P(1, 12.5, 15, 15.5), fill=c or "#ffff00", outline="")


def _icon_font_color(cv, P, k, c):
    _glyph(cv, P, k, 8, 6.5, "A", 12, DARK)
    cv.create_rectangle(*P(1, 12.5, 15, 15.5), fill=c or "#ff0000", outline="")


def _icon_valign(cv, P, k, c, where):
    cv.create_rectangle(*P(1.5, 1.5, 14.5, 14.5), outline="#c0c0c0")
    y = {"top": 4, "center": 6.5, "bottom": 9}[where]
    _lines(cv, P, k, [(4, y, 12), (4, y + 3, 12)])


def _icon_halign(cv, P, k, c, where):
    rows = []
    for i, (y, long) in enumerate(((3, True), (6, False), (9, True), (12, False))):
        width = 12 if long else 8
        x1 = {"left": 2, "center": 8 - width / 2, "right": 14 - width}[where]
        rows.append((x1, y, x1 + width))
    _lines(cv, P, k, rows)


def _icon_indent(cv, P, k, c, inc):
    _lines(cv, P, k, [(2, 2.5, 14), (7, 6, 14), (7, 9.5, 14), (2, 13, 14)])
    if inc:
        cv.create_polygon(*P(1.5, 5, 5, 7.75, 1.5, 10.5), fill=BLUE)
    else:
        cv.create_polygon(*P(5, 5, 1.5, 7.75, 5, 10.5), fill=BLUE)


def _icon_orientation(cv, P, k, c):
    _glyph(cv, P, k, 6.5, 7.5, "ab", 9, DARK, angle=45)
    _arrow(cv, P, k, 9, 14.5, 15, 8.5, BLUE, 1)


def _icon_wrap(cv, P, k, c):
    _lines(cv, P, k, [(2, 3, 14), (2, 13, 6)])
    cv.create_line(*P(2, 8, 12.5, 8), fill=DARK, width=max(1, round(k)))
    cv.create_arc(*P(9.5, 8, 15, 13), start=-90, extent=180, style="arc", outline=BLUE, width=max(1, round(k)))
    cv.create_polygon(*P(12.5, 11, 9.5, 13, 12.5, 15), fill=BLUE)


def _icon_merge(cv, P, k, c):
    cv.create_rectangle(*P(1.5, 3, 14.5, 13), outline=GRAY)
    cv.create_line(*P(3.5, 8, 12.5, 8), fill=BLUE, width=max(1, round(1.4 * k)), arrow="both",
                   arrowshape=(round(3.5 * k), round(4.5 * k), round(2 * k)))


def _icon_accounting(cv, P, k, c):
    cv.create_rectangle(*P(1, 3.5, 12.5, 11), fill="#c6e0b4", outline="#548235")
    cv.create_oval(*P(4.5, 5, 9, 9.5), outline="#548235")
    cv.create_oval(*P(8.5, 10, 15, 13), fill="#ffd966", outline="#bf9000")
    cv.create_oval(*P(8.5, 12, 15, 15), fill="#ffd966", outline="#bf9000")


def _icon_decimals(cv, P, k, c, inc):
    if inc:
        _glyph(cv, P, k, 10, 4.5, ".0", 8, DARK)
        _glyph(cv, P, k, 9, 11.5, ".00", 8, DARK)
        _arrow(cv, P, k, 4, 4.5, 4, 12, BLUE, 1)
    else:
        _glyph(cv, P, k, 9, 4.5, ".00", 8, DARK)
        _glyph(cv, P, k, 10, 11.5, ".0", 8, DARK)
        _arrow(cv, P, k, 3, 4.5, 3, 12, BLUE, 1)


def _icon_cond_format(cv, P, k, c):
    fills = ["#f8696b", "#ffffff", "#63be7b", "#ffffff", "#ffeb84", "#ffffff", "#63be7b", "#ffffff", "#f8696b"]
    for i in range(3):
        for j in range(3):
            cv.create_rectangle(*P(1.5 + 4.5 * j, 2 + 4 * i, 6 + 4.5 * j, 6 + 4 * i), fill=fills[i * 3 + j],
                                outline="#9a9a9a")
    cv.create_rectangle(*P(1.5, 14, 11, 15.5), fill=BLUE, outline="")


def _icon_format_table(cv, P, k, c):
    for i in range(4):
        fill = "#4472c4" if i == 0 else ("#d9e1f2" if i % 2 else "#ffffff")
        cv.create_rectangle(*P(1.5, 2 + 3.2 * i, 14.5, 5.2 + 3.2 * i), fill=fill, outline="#8ea9db")
    for x in (5.8, 10.2):
        cv.create_line(*P(x, 2, x, 14.8), fill="#8ea9db")


def _icon_cells(cv, P, k, c, kind):
    _grid(cv, P, k, 1.5, 1.5, 11.5, 11.5, 2, GRAY, "#ffffff")
    if kind == "insert":
        cv.create_line(*P(13, 9, 13, 15.5), fill=GREEN, width=max(2, round(1.8 * k)))
        cv.create_line(*P(9.75, 12.25, 16, 12.25), fill=GREEN, width=max(2, round(1.8 * k)))
    elif kind == "delete":
        cv.create_line(*P(10, 9.5, 15.5, 15), fill=RED, width=max(2, round(1.8 * k)))
        cv.create_line(*P(15.5, 9.5, 10, 15), fill=RED, width=max(2, round(1.8 * k)))
    else:
        cv.create_rectangle(*P(9, 9, 15.5, 15.5), fill="#ffd65c", outline="#b38f00")
        _lines(cv, P, k, [(10.5, 11.5, 14), (10.5, 13.5, 14)], "#7f6000")


def _icon_autosum(cv, P, k, c):
    _glyph(cv, P, k, 8, 8, "Σ", 15, DARK)


def _icon_fill_down(cv, P, k, c):
    cv.create_rectangle(*P(3, 1.5, 13, 14.5), fill="#ddebf7", outline="#7f9fcf")
    _arrow(cv, P, k, 8, 3.5, 8, 13, BLUE, 2)


def _icon_clear(cv, P, k, c):
    cv.create_polygon(*P(1.5, 11, 6.5, 6, 10, 9.5, 5, 14.5), fill="#f4b6c2", outline="#a05a6a")
    cv.create_polygon(*P(6.5, 6, 11, 1.5, 14.5, 5, 10, 9.5), fill="#ffffff", outline="#a05a6a")
    cv.create_line(*P(5, 14.5, 14, 14.5), fill="#a05a6a")


def _icon_sort_filter(cv, P, k, c):
    _glyph(cv, P, k, 4, 4, "A", 7, BLUE, "bold")
    _glyph(cv, P, k, 4, 11.5, "Z", 7, DARK, "bold")
    _arrow(cv, P, k, 8.5, 2, 8.5, 14, DARK, 1)
    cv.create_polygon(*P(10, 6, 16, 6, 13.8, 9.5, 13.8, 14, 12.2, 15, 12.2, 9.5), fill="#7f9fcf", outline=BLUE)


def _icon_sort(cv, P, k, c, desc):
    top, bottom = ("Z", "A") if desc else ("A", "Z")
    _glyph(cv, P, k, 5, 4, top, 7, BLUE, "bold")
    _glyph(cv, P, k, 5, 11.5, bottom, 7, DARK, "bold")
    _arrow(cv, P, k, 11.5, 2, 11.5, 14, DARK, 1.2)


def _icon_filter(cv, P, k, c, clear=False):
    cv.create_polygon(*P(2, 2.5, 14, 2.5, 9.5, 8, 9.5, 14, 6.5, 15.5, 6.5, 8), fill="#7f9fcf", outline=BLUE)
    if clear:
        cv.create_line(*P(10.5, 9.5, 15.5, 14.5), fill=RED, width=max(2, round(1.6 * k)))
        cv.create_line(*P(15.5, 9.5, 10.5, 14.5), fill=RED, width=max(2, round(1.6 * k)))


def _icon_find(cv, P, k, c):
    cv.create_oval(*P(1.5, 1.5, 10, 10), outline=BLUE, width=max(2, round(1.6 * k)), fill="#ffffff")
    cv.create_line(*P(9, 9, 14.5, 14.5), fill=DARK, width=max(2, round(2.6 * k)))


def _icon_function(cv, P, k, c):
    _glyph(cv, P, k, 8, 8, "fx", 13, c or DARK, slant="italic", family="Cambria")


def _icon_book(cv, P, k, c):
    cv.create_rectangle(*P(2.5, 1.5, 13.5, 14.5), fill=c or "#ddebf7", outline=GRAY)
    cv.create_line(*P(4.5, 1.5, 4.5, 14.5), fill=GRAY)
    _glyph(cv, P, k, 9, 8, "fx", 8, DARK, slant="italic", family="Cambria")


def _icon_save(cv, P, k, c):
    color = c or BLUE
    cv.create_rectangle(*P(2, 2, 14, 14), outline=color, width=max(1, round(1.2 * k)))
    cv.create_rectangle(*P(5, 2, 11, 6.5), outline=color)
    cv.create_rectangle(*P(4.5, 9, 11.5, 14), outline=color)


def _icon_undo(cv, P, k, c, redo=False):
    color = c or BLUE
    if redo:
        cv.create_arc(*P(2, 4, 13, 15), start=0, extent=160, style="arc", outline=color, width=max(1, round(1.6 * k)))
        cv.create_polygon(*P(10, 9.5, 13.5, 6.5, 15.5, 11), fill=color)
    else:
        cv.create_arc(*P(3, 4, 14, 15), start=20, extent=160, style="arc", outline=color, width=max(1, round(1.6 * k)))
        cv.create_polygon(*P(6, 9.5, 2.5, 6.5, 0.5, 11), fill=color)


def _icon_sheet(cv, P, k, c):
    _grid(cv, P, k, 1.5, 1.5, 12, 12, 3, GRAY, "#ffffff")
    cv.create_line(*P(13, 9, 13, 15.5), fill=GREEN, width=max(2, round(1.8 * k)))
    cv.create_line(*P(9.75, 12.25, 16, 12.25), fill=GREEN, width=max(2, round(1.8 * k)))


def _icon_rowcol(cv, P, k, c, kind):
    _grid(cv, P, k, 1.5, 1.5, 14.5, 14.5, 3, GRAY, "#ffffff")
    if kind in ("row", "row_del"):
        cv.create_rectangle(*P(1.5, 5.8, 14.5, 10.2), fill="#9fc5e8" if kind == "row" else "#f4b6c2", outline=GRAY)
    else:
        cv.create_rectangle(*P(5.8, 1.5, 10.2, 14.5), fill="#9fc5e8" if kind == "col" else "#f4b6c2", outline=GRAY)


def _icon_calendar(cv, P, k, c):
    cv.create_rectangle(*P(1.5, 3, 14.5, 14.5), fill="#ffffff", outline=GRAY)
    cv.create_rectangle(*P(1.5, 3, 14.5, 6), fill=RED, outline="")
    _glyph(cv, P, k, 8, 10.5, "21", 7, DARK, "bold")


def _icon_clock(cv, P, k, c):
    cv.create_oval(*P(1.5, 1.5, 14.5, 14.5), fill="#ffffff", outline=GRAY)
    cv.create_line(*P(8, 8, 8, 3.5), fill=DARK, width=max(1, round(1.2 * k)))
    cv.create_line(*P(8, 8, 11.5, 9.5), fill=DARK, width=max(1, round(1.2 * k)))


def _icon_show_formulas(cv, P, k, c):
    cv.create_rectangle(*P(1.5, 1.5, 14.5, 14.5), fill="#ffffff", outline=GRAY)
    _glyph(cv, P, k, 8, 8, "f=", 9, BLUE, slant="italic", family="Cambria")


def _icon_calc(cv, P, k, c):
    cv.create_rectangle(*P(3, 1, 13, 15), fill="#ffffff", outline=GRAY)
    cv.create_rectangle(*P(4.5, 2.5, 11.5, 5.5), fill="#c6e0b4", outline="")
    for i in range(3):
        for j in range(3):
            cv.create_rectangle(*P(4.5 + 2.6 * j, 7 + 2.5 * i, 6 + 2.6 * j, 8.4 + 2.5 * i), fill=GRAY, outline="")


def _icon_zoom(cv, P, k, c, label="+"):
    cv.create_oval(*P(1.5, 1.5, 11, 11), outline=BLUE, width=max(1, round(1.4 * k)), fill="#ffffff")
    cv.create_line(*P(10, 10, 14.5, 14.5), fill=DARK, width=max(2, round(2.4 * k)))
    _glyph(cv, P, k, 6.25, 6.25, label, 7 if len(label) > 1 else 9, DARK, "bold")


def _icon_remove_dup(cv, P, k, c):
    _grid(cv, P, k, 1.5, 1.5, 10, 14.5, 2, GRAY, "#ffffff")
    cv.create_rectangle(*P(1.5, 5.8, 10, 10.2), fill="#f4b6c2", outline=GRAY)
    cv.create_line(*P(11, 6, 15.5, 10.5), fill=RED, width=max(2, round(1.6 * k)))
    cv.create_line(*P(15.5, 6, 11, 10.5), fill=RED, width=max(2, round(1.6 * k)))


def _icon_help(cv, P, k, c):
    cv.create_oval(*P(1.5, 1.5, 14.5, 14.5), fill=BLUE, outline="")
    _glyph(cv, P, k, 8, 8, "?", 10, "#ffffff", "bold")


def _icon_info(cv, P, k, c):
    cv.create_oval(*P(1.5, 1.5, 14.5, 14.5), fill=GREEN, outline="")
    _glyph(cv, P, k, 8, 8, "i", 10, "#ffffff", "bold", family="Georgia")


def _icon_freeze(cv, P, k, c):
    _grid(cv, P, k, 1.5, 1.5, 14.5, 14.5, 4, GRAY, "#ffffff")
    cv.create_rectangle(*P(1.5, 1.5, 14.5, 4.75), fill="#bdd7ee", outline=GRAY)
    cv.create_rectangle(*P(1.5, 1.5, 4.75, 14.5), fill="#bdd7ee", outline=GRAY)
    cv.create_line(*P(1.5, 4.75, 14.5, 4.75), fill=BLUE, width=max(1, round(1.4 * k)))
    cv.create_line(*P(4.75, 1.5, 4.75, 14.5), fill=BLUE, width=max(1, round(1.4 * k)))


def _icon_validation(cv, P, k, c):
    cv.create_rectangle(*P(1.5, 3, 12.5, 8), fill="#ffffff", outline=GRAY)
    cv.create_rectangle(*P(9, 3, 12.5, 8), fill="#e8e8e8", outline=GRAY)
    cv.create_polygon(*P(9.75, 4.75, 11.75, 4.75, 10.75, 6.5), fill=DARK, outline="")
    cv.create_line(*P(5, 12, 7.5, 14.5, 14.5, 7.5), fill=GREEN, width=max(2, round(1.8 * k)))
    cv.create_line(*P(1.5, 10, 4.5, 13), fill=RED, width=max(1, round(1.4 * k)))
    cv.create_line(*P(4.5, 10, 1.5, 13), fill=RED, width=max(1, round(1.4 * k)))


def _icon_note(cv, P, k, c, delete=False):
    cv.create_polygon(*P(2, 2, 14, 2, 14, 10.5, 10.5, 14, 2, 14), fill="#fff2a8", outline="#b89b00")
    cv.create_polygon(*P(10.5, 14, 10.5, 10.5, 14, 10.5), fill="#e6cf4f", outline="#b89b00")
    _lines(cv, P, k, [(4, 5, 12), (4, 7.5, 12), (4, 10, 9)], GRAY)
    if delete:
        cv.create_line(*P(9, 9, 15.5, 15.5), fill=RED, width=max(2, round(1.6 * k)))
        cv.create_line(*P(15.5, 9, 9, 15.5), fill=RED, width=max(2, round(1.6 * k)))


def _icon_note_step(cv, P, k, c, forward):
    cv.create_polygon(*P(5, 2, 14, 2, 14, 9, 11, 12, 5, 12), fill="#fff2a8", outline="#b89b00")
    if forward:
        _arrow(cv, P, k, 1, 14, 8, 14, BLUE)
    else:
        _arrow(cv, P, k, 8, 14, 1, 14, BLUE)


CHART_COLORS = ("#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47")


def _icon_chart(cv, P, k, c, kind="column"):
    if kind == "pie":
        cv.create_arc(*P(1.5, 1.5, 14.5, 14.5), start=90, extent=-130, fill=CHART_COLORS[0], outline="#ffffff")
        cv.create_arc(*P(1.5, 1.5, 14.5, 14.5), start=-40, extent=-110, fill=CHART_COLORS[1], outline="#ffffff")
        cv.create_arc(*P(1.5, 1.5, 14.5, 14.5), start=-150, extent=-120, fill=CHART_COLORS[3], outline="#ffffff")
        return
    cv.create_line(*P(1.5, 1.5, 1.5, 14.5, 14.5, 14.5), fill=GRAY)
    if kind == "line":
        cv.create_line(*P(2.5, 11, 6, 7, 9, 9, 14, 3), fill=CHART_COLORS[0], width=max(2, round(1.6 * k)))
        cv.create_line(*P(2.5, 13, 6, 11, 9, 12, 14, 8), fill=CHART_COLORS[1], width=max(2, round(1.6 * k)))
    elif kind == "bar":
        for i, (w, color) in enumerate(((9, CHART_COLORS[0]), (12, CHART_COLORS[1]), (6, CHART_COLORS[0]))):
            cv.create_rectangle(*P(1.5, 2.5 + i * 4, 1.5 + w, 5 + i * 4), fill=color, outline="")
    else:
        for i, (h, color) in enumerate(((6, CHART_COLORS[0]), (10, CHART_COLORS[1]), (8, CHART_COLORS[0]),
                                        (12, CHART_COLORS[1]))):
            cv.create_rectangle(*P(3 + i * 3, 14.5 - h, 5 + i * 3, 14.5), fill=color, outline="")


# menu icons ----------------------------------------------------------------------------------------

def _dots(cv, P, k, x0, y0, x1, y1, color="#a6a6a6"):
    """A dotted horizontal or vertical line (one-unit dots, Excel's 'no border here' grid)."""
    n = round(max(abs(x1 - x0), abs(y1 - y0)) / 2)
    for i in range(n + 1):
        x, y = x0 + (x1 - x0) * i / max(1, n), y0 + (y1 - y0) * i / max(1, n)
        cv.create_rectangle(*P(x - 0.45, y - 0.45, x + 0.45, y + 0.45), fill=color, outline="")


_BORDER_EDGES = {"bottom": "B", "top": "T", "left": "L", "right": "R", "none": "", "all": "TBLRHV",
                 "outside": "TBLR", "thick_outside": "tblr", "bottom_double": "d", "thick_bottom": "b",
                 "top_bottom": "TB", "top_thick_bottom": "Tb", "top_double_bottom": "Td"}


def _icon_border(cv, P, k, c, kind):
    """Excel's border-menu pictures: a dotted 2 x 2 grid with the borders `kind` draws in solid black
    (upper case thin, lower case thick, d a double bottom line)."""
    lines = {"T": (2, 2, 14, 2), "B": (2, 14, 14, 14), "L": (2, 2, 2, 14), "R": (14, 2, 14, 14),
             "H": (2, 8, 14, 8), "V": (8, 2, 8, 14)}
    for seg in lines.values():
        _dots(cv, P, k, *seg)
    for code in _BORDER_EDGES[kind]:
        if code == "d":
            for y in (12.6, 14.4):
                cv.create_line(*P(1.6, y, 14.4, y), fill="#000000", width=max(1, round(0.9 * k)))
            continue
        x0, y0, x1, y1 = lines[code.upper()]
        thick = code.islower()
        grow = 0.6 if thick else 0.4
        cv.create_line(*P(x0 - grow * (x0 != x1), y0 - grow * (y0 != y1), x1 + grow * (x0 != x1), y1 + grow * (y0 != y1)),
                       fill="#000000", width=max(2 if thick else 1, round((2.3 if thick else 1.15) * k)))


def _icon_line_style(cv, P, k, c, style):
    color = "#262626"
    if style == "double":
        for y in (6.5, 9.5):
            cv.create_line(*P(1, y, 15, y), fill=color, width=max(1, round(k)))
    elif style == "dotted":
        _dots(cv, P, k, 1.5, 8, 14.5, 8, color)
    elif style == "dashed":
        for x in (1, 6, 11):
            cv.create_line(*P(x, 8, x + 3.5, 8), fill=color, width=max(1, round(1.2 * k)))
    else:
        width = {"thin": 1.1, "medium": 2.2, "thick": 3.4}[style]
        cv.create_line(*P(1, 8, 15, 8), fill=color, width=max(1, round(width * k)))


def _pencil(cv, P, k, ox=0.0, oy=0.0, s=1.0):
    def Q(*xy):
        return P(*[ox + v * s if i % 2 == 0 else oy + v * s for i, v in enumerate(xy)])
    cv.create_polygon(*Q(4, 11, 11, 4, 13.5, 6.5, 6.5, 13.5), fill="#ffd65c", outline="#8a6a3a")
    cv.create_polygon(*Q(11, 4, 12.5, 2.5, 15, 5, 13.5, 6.5), fill="#e07b7b", outline="#8a6a3a")
    cv.create_polygon(*Q(4, 11, 6.5, 13.5, 2.5, 15), fill="#5c5c5c", outline="")


def _icon_line_color(cv, P, k, c):
    _pencil(cv, P, k, 0, -1.5, 0.95)
    cv.create_rectangle(*P(1, 13.5, 15, 15.6), fill=c or "#000000", outline="")


def _clipboard(cv, P, k):
    cv.create_rectangle(*P(1.5, 2.5, 10, 14.5), fill="#d9b278", outline="#8a6a3a")
    cv.create_rectangle(*P(3.8, 1, 7.7, 4), fill="#8c8c8c", outline="#5c5c5c")


def _icon_paste_as(cv, P, k, c, kind):
    """The clipboard with a sheet showing what is pasted: formulas, values, formats, transposed, special."""
    _clipboard(cv, P, k)
    cv.create_rectangle(*P(6.5, 6.5, 15.5, 15.5), fill="#ffffff", outline="#6b6b6b")
    if kind == "formulas":
        _glyph(cv, P, k, 11, 10.8, "fx", 7.5, BLUE, slant="italic", family="Cambria")
    elif kind == "values":
        _glyph(cv, P, k, 11, 11, "123", 5.5, DARK, "bold")
    elif kind == "formats":
        cv.create_rectangle(*P(8, 8, 14, 10.5), fill="#ffd65c", outline="#b38f00")
        cv.create_line(*P(11, 10.5, 11, 14.5), fill="#8a5a2b", width=max(2, round(1.6 * k)))
    elif kind == "transpose":
        cv.create_line(*P(8.5, 8, 8.5, 13.5, 13.5, 13.5), fill=BLUE, width=max(1, round(1.2 * k)), arrow="last",
                       arrowshape=(round(3 * k), round(3.5 * k), round(1.6 * k)))
        _lines(cv, P, k, [(10.5, 8.5, 14), (10.5, 10.5, 14)], "#9a9a9a")
    else:
        _lines(cv, P, k, [(8, 9, 14), (8, 11, 14), (8, 13, 12)], "#9a9a9a")
        cv.create_oval(*P(12, 11.5, 15.5, 15), fill=GREEN, outline="")


def _icon_underline_double(cv, P, k, c):
    _glyph(cv, P, k, 8, 6.5, "U", 11, DARK)
    _lines(cv, P, k, [(3.5, 12.5, 12.5), (3.5, 14.5, 12.5)])


def _icon_merge_kind(cv, P, k, c, kind):
    if kind == "across":
        for y0 in (2, 8.5):
            cv.create_rectangle(*P(1.5, y0, 14.5, y0 + 5.5), outline=GRAY, fill="#ffffff")
            cv.create_line(*P(4, y0 + 2.75, 12, y0 + 2.75), fill=BLUE, width=max(1, round(1.1 * k)), arrow="both",
                           arrowshape=(round(2.5 * k), round(3 * k), round(1.4 * k)))
    elif kind == "cells":
        cv.create_rectangle(*P(1.5, 3, 14.5, 13), outline=GRAY, fill="#ddebf7")
        _dots(cv, P, k, 8, 3, 8, 13, "#9fc5e8")
    else:  # unmerge
        _grid(cv, P, k, 1.5, 3, 14.5, 13, 2, GRAY, "#ffffff")
        _arrow(cv, P, k, 6.5, 8, 2.5, 8, BLUE, 1.1)
        _arrow(cv, P, k, 9.5, 8, 13.5, 8, BLUE, 1.1)


def _icon_text_angle(cv, P, k, c, angle):
    if angle == 255:  # stacked letters
        _glyph(cv, P, k, 6, 4.5, "a", 8, DARK)
        _glyph(cv, P, k, 6, 10.5, "b", 8, DARK)
        _arrow(cv, P, k, 12, 2, 12, 14.5, BLUE, 1)
        return
    rad = math.radians(angle)
    dx, dy = math.cos(rad), -math.sin(rad)
    px, py = -dy, dx  # perpendicular, towards the lower side of the text
    _glyph(cv, P, k, 8 - 1.8 * px, 8 - 1.8 * py, "ab", 8.5, DARK, angle=angle)
    _arrow(cv, P, k, 8 - 5.5 * dx + 4 * px, 8 - 5.5 * dy + 4 * py, 8 + 5.5 * dx + 4 * px, 8 + 5.5 * dy + 4 * py, BLUE, 1)


def _icon_cf_kind(cv, P, k, c, kind):
    """Conditional-formatting pictures: highlight rules, top/bottom rules, data bars, colour scales, manage."""
    if kind == "highlight":
        _grid(cv, P, k, 1.5, 1.5, 14.5, 14.5, 3, GRAY, "#ffffff")
        cv.create_rectangle(*P(5.83, 5.83, 10.17, 10.17), fill="#ffc7ce", outline="#9c0006")
    elif kind == "topbottom":
        for i, (w, color) in enumerate(((12, "#f8696b"), (9, "#bfbfbf"), (6, "#bfbfbf"))):
            cv.create_rectangle(*P(2, 2.5 + 4.2 * i, 2 + w, 5.5 + 4.2 * i), fill=color, outline="")
        _glyph(cv, P, k, 12, 12.5, "10", 6, DARK, "bold")
    elif kind == "databar":
        for i, w in enumerate((11, 6, 8.5)):
            cv.create_rectangle(*P(1.5, 1.5 + 4.5 * i, 14.5, 5.5 + 4.5 * i), fill="#ffffff", outline="#bfbfbf")
            cv.create_rectangle(*P(2.2, 2.2 + 4.5 * i, 2.2 + w, 4.8 + 4.5 * i), fill=c or "#638ec6", outline="")
    elif kind == "scale":
        colors = (c or "#63be7b,#ffeb84,#f8696b").split(",")
        if len(colors) == 2:
            colors = [colors[0], mix(colors[0], colors[1], 0.5), colors[1]]
        for i, color in enumerate(colors):
            cv.create_rectangle(*P(2.5, 1.5 + 4.4 * i, 13.5, 5.9 + 4.4 * i), fill=color, outline="#a6a6a6")
    else:  # manage
        for i, color in enumerate(("#f8696b", "#63be7b", "#638ec6")):
            cv.create_rectangle(*P(1.5, 2 + 4.5 * i, 5, 5.5 + 4.5 * i), fill=color, outline="")
            cv.create_line(*P(6.5, 3.75 + 4.5 * i, 14.5, 3.75 + 4.5 * i), fill=DARK, width=max(1, round(k)))


def _icon_cf_rule(cv, P, k, c, text):
    cv.create_rectangle(*P(1.5, 3, 14.5, 13), fill="#ffc7ce", outline="#9c0006")
    _glyph(cv, P, k, 8, 8, text, 9 if len(text) < 2 else 7, "#9c0006", "bold")


def _icon_rank(cv, P, k, c, text, up):
    _glyph(cv, P, k, 6, 8, text, 7 if len(text) > 1 else 9, DARK, "bold")
    if up:
        _arrow(cv, P, k, 13, 14, 13, 2, GREEN, 1.6)
    else:
        _arrow(cv, P, k, 13, 2, 13, 14, RED, 1.6)


def _icon_stat(cv, P, k, c, kind):
    """Autosum variants: average, count, maximum, minimum."""
    if kind == "count":
        _glyph(cv, P, k, 8, 8, "123", 7, BLUE, "bold")
        return
    heights = (6, 11, 8)
    pick = {"max": 1, "min": 0}.get(kind)
    for i, h in enumerate(heights):
        color = BLUE if i == pick or kind == "avg" else "#bfbfbf"
        cv.create_rectangle(*P(2 + 4.3 * i, 14.5 - h, 5.5 + 4.3 * i, 14.5), fill=color, outline="")
    if kind == "avg":
        cv.create_line(*P(1, 6.3, 15, 6.3), fill=RED, width=max(1, round(1.2 * k)))


def _icon_sort_custom(cv, P, k, c):
    _lines(cv, P, k, [(1.5, 3, 9), (1.5, 6.5, 9), (1.5, 10, 9), (1.5, 13.5, 9)])
    _arrow(cv, P, k, 12.5, 2, 12.5, 14.5, BLUE, 1.4)


def _icon_filter_variant(cv, P, k, c, kind):
    cv.create_polygon(*P(1.5, 2, 12, 2, 8, 7, 8, 12.5, 5.5, 14, 5.5, 7), fill="#7f9fcf", outline=BLUE)
    if kind == "reapply":
        cv.create_arc(*P(8.5, 8.5, 15.5, 15.5), start=90, extent=270, style="arc", outline=GREEN,
                      width=max(1, round(1.4 * k)))
        cv.create_polygon(*P(11, 7.3, 13.4, 8.8, 11, 10.3), fill=GREEN)
    else:  # filter by the selected cell's value
        cv.create_rectangle(*P(9.5, 9.5, 15.5, 15.5), fill="#ffd65c", outline="#b38f00")


def _icon_fill_dir(cv, P, k, c, direction):
    vertical = direction in ("down", "up")
    box = (3, 1.5, 13, 14.5) if vertical else (1.5, 3, 14.5, 13)
    cv.create_rectangle(*P(*box), fill="#ddebf7", outline="#7f9fcf")
    ends = {"down": (8, 3.5, 8, 13), "up": (8, 12.5, 8, 3), "right": (3.5, 8, 13, 8), "left": (12.5, 8, 3, 8)}
    _arrow(cv, P, k, *ends[direction], BLUE, 2)


def _icon_fill_series(cv, P, k, c):
    for i, digit in enumerate("123"):
        cv.create_rectangle(*P(3, 1.5 + 4.4 * i, 13, 5.5 + 4.4 * i), fill="#ddebf7", outline="#7f9fcf")
        _glyph(cv, P, k, 8, 3.6 + 4.4 * i, digit, 5.5, BLUE, "bold")


def _eraser(cv, P, k, ox, oy, s):
    def Q(*xy):
        return P(*[ox + v * s if i % 2 == 0 else oy + v * s for i, v in enumerate(xy)])
    cv.create_polygon(*Q(1.5, 11, 6.5, 6, 10, 9.5, 5, 14.5), fill="#f4b6c2", outline="#a05a6a")
    cv.create_polygon(*Q(6.5, 6, 11, 1.5, 14.5, 5, 10, 9.5), fill="#ffffff", outline="#a05a6a")


def _icon_clear_what(cv, P, k, c, what):
    if what == "formats":
        _glyph(cv, P, k, 5, 5, "A", 9, RED, "bold")
        cv.create_rectangle(*P(1, 9, 8, 10.5), fill="#ffd65c", outline="")
    else:  # contents
        _lines(cv, P, k, [(1.5, 3, 9), (1.5, 6, 9), (1.5, 9, 6)], GRAY)
    _eraser(cv, P, k, 5, 5, 0.7)


def _icon_replace(cv, P, k, c):
    _glyph(cv, P, k, 5, 4.5, "ab", 7, GRAY)
    _glyph(cv, P, k, 10.5, 12, "ac", 7, BLUE, "bold")
    cv.create_line(*P(2.5, 8, 2.5, 12, 6, 12), fill=DARK, width=max(1, round(k)), arrow="last",
                   arrowshape=(round(2.5 * k), round(3 * k), round(1.4 * k)))


def _icon_goto(cv, P, k, c):
    _grid(cv, P, k, 5, 5, 15, 15, 2, GRAY, "#ffffff")
    cv.create_rectangle(*P(10, 10, 15, 15), fill="#c6e0b4", outline=GREEN)
    _arrow(cv, P, k, 1.5, 1.5, 9.5, 9.5, BLUE, 1.6)


def _icon_size(cv, P, k, c, axis, auto):
    """Row height / column width (auto: fit to the contents)."""
    if axis == "row":
        _lines(cv, P, k, [(1.5, 3.5, 14.5), (1.5, 12.5, 14.5)], GRAY)
        line = (5 if auto else 8, 4.5, 5 if auto else 8, 11.5)
    else:
        cv.create_line(*P(3.5, 1.5, 3.5, 14.5), fill=GRAY, width=max(1, round(k)))
        cv.create_line(*P(12.5, 1.5, 12.5, 14.5), fill=GRAY, width=max(1, round(k)))
        line = (4.5, 5 if auto else 8, 11.5, 5 if auto else 8)
    cv.create_line(*P(*line), fill=BLUE, width=max(1, round(1.3 * k)), arrow="both",
                   arrowshape=(round(2.5 * k), round(3 * k), round(1.5 * k)))
    if auto:
        _glyph(cv, P, k, 10.5 if axis == "row" else 8, 8 if axis == "row" else 11.5, "A", 7, DARK, "bold")


def _icon_hide(cv, P, k, c, axis, show):
    _grid(cv, P, k, 1.5, 1.5, 14.5, 14.5, 3, GRAY, "#ffffff")
    band = (1.5, 5.83, 14.5, 10.17) if axis == "row" else (5.83, 1.5, 10.17, 14.5)
    cv.create_rectangle(*P(*band), fill="#9fc5e8" if show else "#a6a6a6", outline=GRAY)
    if not show:
        x0, y0, x1, y1 = band
        cv.create_line(*P(x0 + 1, y0 + 1, x1 - 1, y1 - 1), fill="#ffffff", width=max(1, round(k)))


def _sheet(cv, P, k, x0, y0, x1, y1):
    _grid(cv, P, k, x0, y0, x1, y1, 3, GRAY, "#ffffff")
    cv.create_rectangle(*P(x0, y0, x1, y0 + (y1 - y0) / 3), fill="#c6e0b4", outline=GRAY)


def _icon_sheet_op(cv, P, k, c, op):
    if op == "copy":
        _sheet(cv, P, k, 1, 1, 10.5, 10.5)
        _sheet(cv, P, k, 5.5, 5.5, 15, 15)
    elif op == "rename":
        _sheet(cv, P, k, 1, 1, 12, 12)
        _pencil(cv, P, k, 5, 5, 0.65)
    elif op == "delete":
        _sheet(cv, P, k, 1, 1, 12, 12)
        cv.create_line(*P(9.5, 9.5, 15.5, 15.5), fill=RED, width=max(2, round(1.8 * k)))
        cv.create_line(*P(15.5, 9.5, 9.5, 15.5), fill=RED, width=max(2, round(1.8 * k)))
    elif op in ("hide", "unhide"):
        _sheet(cv, P, k, 1, 1, 12, 12)
        cv.create_oval(*P(7, 9, 15.5, 14.5), fill="#ffffff", outline=DARK)
        cv.create_oval(*P(9.7, 10.2, 12.8, 13.3), fill=DARK, outline="")
        if op == "hide":
            cv.create_line(*P(7, 15.5, 15.5, 8), fill=RED, width=max(1, round(1.4 * k)))
    elif op == "tab_color":
        cv.create_polygon(*P(1, 4, 15, 4, 13.5, 11, 2.5, 11), fill="#ffffff", outline=GRAY)
        cv.create_rectangle(*P(2.5, 11, 13.5, 13.5), fill=c or "#ed7d31", outline="")
    else:  # move left / right
        cv.create_polygon(*P(4, 9, 12, 9, 11, 14, 5, 14), fill="#ffffff", outline=GRAY)
        if op == "left":
            _arrow(cv, P, k, 14, 4.5, 2, 4.5, BLUE, 1.6)
        else:
            _arrow(cv, P, k, 2, 4.5, 14, 4.5, BLUE, 1.6)


def _icon_file(cv, P, k, c, kind):
    if kind == "open":
        cv.create_polygon(*P(1, 3, 6, 3, 7.5, 5, 14, 5, 14, 14, 1, 14), fill="#ffd65c", outline="#b38f00")
        cv.create_polygon(*P(1, 14, 3.5, 8, 15.5, 8, 14, 14), fill="#ffe699", outline="#b38f00")
    elif kind == "new":
        cv.create_polygon(*P(3, 1, 10, 1, 13.5, 4.5, 13.5, 15, 3, 15), fill="#ffffff", outline=GRAY)
        cv.create_polygon(*P(10, 1, 10, 4.5, 13.5, 4.5), fill="#d9d9d9", outline=GRAY)
    elif kind == "save_as":
        _icon_save(cv, P, k, None)
        _pencil(cv, P, k, 5.5, 5.5, 0.65)
    else:  # exit
        cv.create_rectangle(*P(1.5, 1.5, 9, 14.5), fill="#ffffff", outline=GRAY)
        _arrow(cv, P, k, 5, 8, 15, 8, RED, 1.8)


def _icon_freeze_part(cv, P, k, c, part):
    _grid(cv, P, k, 1.5, 1.5, 14.5, 14.5, 4, GRAY, "#ffffff")
    band = (1.5, 1.5, 14.5, 4.75) if part == "row" else (1.5, 1.5, 4.75, 14.5)
    cv.create_rectangle(*P(*band), fill="#bdd7ee", outline=GRAY)
    edge = (1.5, 4.75, 14.5, 4.75) if part == "row" else (4.75, 1.5, 4.75, 14.5)
    cv.create_line(*P(*edge), fill=BLUE, width=max(1, round(1.4 * k)))


def _icon_validation_clear(cv, P, k, c):
    cv.create_rectangle(*P(1.5, 3, 12.5, 8), fill="#ffffff", outline=GRAY)
    cv.create_rectangle(*P(9, 3, 12.5, 8), fill="#e8e8e8", outline=GRAY)
    cv.create_polygon(*P(9.75, 4.75, 11.75, 4.75, 10.75, 6.5), fill=DARK, outline="")
    _eraser(cv, P, k, 5, 5, 0.7)


def _icon_delete_x(cv, P, k, c):
    cv.create_line(*P(3, 3, 13, 13), fill=RED, width=max(2, round(2.2 * k)))
    cv.create_line(*P(13, 3, 3, 13), fill=RED, width=max(2, round(2.2 * k)))


def _icon_chart_edit(cv, P, k, c):
    _icon_chart(cv, P, k, c)
    _pencil(cv, P, k, 6, 6, 0.6)


def _icon_by_color(cv, P, k, c, sort):
    """Sort by colour (coloured cells and an arrow) / filter by colour (funnel and colours)."""
    if sort:
        for i, color in enumerate(("#f8696b", "#ffd65c", "#5b9bd5")):
            cv.create_rectangle(*P(7.5, 1.5 + 4.6 * i, 14.5, 5.3 + 4.6 * i), fill=color, outline="#7f7f7f")
        _arrow(cv, P, k, 3.5, 2, 3.5, 14.5, DARK, 1.2)
        return
    cv.create_polygon(*P(1.5, 1.5, 12.5, 1.5, 8.5, 6.5, 8.5, 11, 5.5, 12.5, 5.5, 6.5), fill="#7f9fcf", outline=BLUE)
    for i, color in enumerate(("#f8696b", "#ffd65c", "#70ad47")):
        cv.create_oval(*P(9 + 2.2 * i - 1.6, 12.6, 9 + 2.2 * i + 1.6, 15.8), fill=color, outline="#7f7f7f")


def _icon_pivot(cv, P, k, c):
    _grid(cv, P, k, 1.5, 1.5, 14.5, 14.5, 3, GRAY, "#ffffff")
    cv.create_rectangle(*P(1.5, 1.5, 14.5, 5.83), fill="#bdd7ee", outline=GRAY)
    cv.create_rectangle(*P(1.5, 1.5, 5.83, 14.5), fill="#c6e0b4", outline=GRAY)
    cv.create_line(*P(8, 9.5, 8, 12.5, 11.5, 12.5), fill=BLUE, width=max(1, round(1.2 * k)), arrow="last",
                   arrowshape=(round(2.5 * k), round(3 * k), round(1.4 * k)))


def _icon_refresh(cv, P, k, c):
    w = max(2, round(1.6 * k))
    cv.create_arc(*P(2, 2, 14, 14), start=20, extent=150, style="arc", outline=GREEN, width=w)
    cv.create_arc(*P(2, 2, 14, 14), start=200, extent=150, style="arc", outline=GREEN, width=w)
    cv.create_polygon(*P(1, 7.5, 3.6, 11, 6, 7.5), fill=GREEN)
    cv.create_polygon(*P(15, 8.5, 12.4, 5, 10, 8.5), fill=GREEN)


def _icon_pivot_fields(cv, P, k, c):
    cv.create_rectangle(*P(1.5, 1.5, 14.5, 14.5), fill="#ffffff", outline=GRAY)
    for i, on in enumerate((True, False, True)):
        y = 3.5 + 4 * i
        cv.create_rectangle(*P(3, y, 6, y + 3), fill=BLUE if on else "#ffffff", outline=BLUE)
        cv.create_line(*P(7.5, y + 1.5, 13, y + 1.5), fill=DARK, width=max(1, round(k)))


def _page(cv, P, x0, y0, x1, y1, fill="#ffffff"):
    """A sheet of paper with a folded corner."""
    f = min(x1 - x0, y1 - y0) * 0.28
    cv.create_polygon(*P(x0, y0, x1 - f, y0, x1, y0 + f, x1, y1, x0, y1), fill=fill, outline=GRAY)
    cv.create_polygon(*P(x1 - f, y0, x1 - f, y0 + f, x1, y0 + f), fill="#d9d9d9", outline=GRAY)


def _icon_margins(cv, P, k, c):
    _page(cv, P, 2.5, 1, 13.5, 15)
    cv.create_rectangle(*P(4.5, 3.5, 11.5, 12.5), outline=BLUE, dash=(2, 1))
    _lines(cv, P, k, [(5.5, 6, 10.5), (5.5, 8, 10.5), (5.5, 10, 9)], GRAY)


def _icon_page_orient(cv, P, k, c, kind=None):
    """Orientation: a portrait and a landscape sheet, or just one of them."""
    if kind == "portrait":
        _page(cv, P, 3.5, 1, 12.5, 15)
        _lines(cv, P, k, [(5.5, 5, 10.5), (5.5, 7, 10.5), (5.5, 9, 9)], GRAY)
    elif kind == "landscape":
        _page(cv, P, 1, 3.5, 15, 12.5)
        _lines(cv, P, k, [(3, 7, 12), (3, 9, 10)], GRAY)
    else:
        _page(cv, P, 1.5, 1, 9.5, 12)
        _page(cv, P, 4.5, 7, 15, 14.5, "#deebf7")
        _lines(cv, P, k, [(6.5, 10, 12.5), (6.5, 12, 11)], GRAY)


def _icon_paper_size(cv, P, k, c):
    _page(cv, P, 4, 1, 14, 15)
    cv.create_line(*P(2, 1.5, 2, 14.5), fill=BLUE, width=max(1, round(k)), arrow="both",
                   arrowshape=(round(2.5 * k), round(3 * k), round(1.4 * k)))
    _glyph(cv, P, k, 9, 9.5, "A4", 6, BLUE, "bold")


def _icon_print_area(cv, P, k, c):
    _grid(cv, P, k, 1.5, 1.5, 14.5, 14.5, 4, GRAY, "#ffffff")
    cv.create_rectangle(*P(1.5, 1.5, 11.25, 11.25), fill="#deebf7", outline=BLUE, width=max(1, round(1.4 * k)),
                        dash=(3, 2))


def _icon_page_break(cv, P, k, c):
    _page(cv, P, 3, 1, 13, 7)
    _page(cv, P, 3, 9.5, 13, 15)
    cv.create_line(*P(0.5, 8.25, 15.5, 8.25), fill=BLUE, width=max(1, round(1.3 * k)), dash=(3, 2))


def _icon_print_titles(cv, P, k, c):
    for x, y in ((1, 1), (5, 4)):
        _page(cv, P, x, y, x + 10, y + 11)
        cv.create_rectangle(*P(x + 1.5, y + 1.5, x + 8.5, y + 3.5), fill="#bdd7ee", outline=BLUE)
    _lines(cv, P, k, [(6.5, 10, 13.5), (6.5, 12, 12)], GRAY)


def _icon_fit_page(cv, P, k, c):
    _page(cv, P, 3, 2.5, 13, 13.5)
    for x1, y1, x2, y2 in ((1, 1, 4.5, 4.5), (15, 1, 11.5, 4.5), (1, 15, 4.5, 11.5), (15, 15, 11.5, 11.5)):
        _arrow(cv, P, k, x1, y1, x2, y2, BLUE, 1.2)


def _icon_print(cv, P, k, c):
    cv.create_rectangle(*P(4, 1.5, 12, 6), fill="#ffffff", outline=GRAY)
    cv.create_rectangle(*P(1.5, 5.5, 14.5, 11.5), fill="#bfbfbf", outline=DARK)
    cv.create_oval(*P(11.5, 7, 13, 8.5), fill=GREEN, outline="")
    cv.create_rectangle(*P(4, 9.5, 12, 15), fill="#ffffff", outline=GRAY)
    _lines(cv, P, k, [(5.5, 11.5, 10.5), (5.5, 13.3, 9.5)], GRAY)


def _icon_print_preview(cv, P, k, c):
    _page(cv, P, 1.5, 1, 11.5, 14)
    _lines(cv, P, k, [(3.5, 4, 8.5), (3.5, 6, 8.5), (3.5, 8, 7)], GRAY)
    cv.create_oval(*P(6.5, 7.5, 12.5, 13.5), outline=BLUE, width=max(1, round(1.4 * k)), fill="#ffffff")
    cv.create_line(*P(12, 13, 15, 16), fill=DARK, width=max(2, round(2 * k)))


def _icon_pdf(cv, P, k, c):
    _page(cv, P, 2.5, 1, 13.5, 15)
    cv.create_rectangle(*P(0.5, 7, 12, 13), fill=RED, outline="")
    _glyph(cv, P, k, 6.25, 10, "PDF", 5.5, "#ffffff", "bold")


def _icon_page_setup(cv, P, k, c):
    _page(cv, P, 1.5, 1, 11.5, 15)
    _lines(cv, P, k, [(3.5, 4.5, 8.5), (3.5, 6.5, 8.5)], GRAY)
    for i, x in enumerate((7.5, 10.5, 13.5)):
        cv.create_line(*P(x, 8, x, 15.5), fill=DARK, width=max(1, round(k)))
        y = (10, 13.5, 11.5)[i]
        cv.create_rectangle(*P(x - 1.2, y - 1, x + 1.2, y + 1), fill=BLUE, outline="")


def _outline_box(cv, P, k, x, y, sign):
    cv.create_rectangle(*P(x - 3, y - 3, x + 3, y + 3), fill="#ffffff", outline=GRAY)
    cv.create_line(*P(x - 1.7, y, x + 1.9, y), fill=DARK, width=max(1, round(k)))
    if sign == "+":
        cv.create_line(*P(x, y - 1.7, x, y + 1.9), fill=DARK, width=max(1, round(k)))


def _icon_group(cv, P, k, c, ungroup=False):
    for i in range(4):
        y = 2.5 + 3.5 * i
        cv.create_rectangle(*P(7, y, 15, y + 2.5), fill="#ffffff" if i < 3 else "#e2efda", outline=GRAY)
    if ungroup:
        cv.create_line(*P(3, 2.5, 3, 10), fill=GRAY, width=max(1, round(k)), dash=(2, 2))
        _arrow(cv, P, k, 6.5, 6, 0.5, 6, RED, 1.3)
    else:
        cv.create_line(*P(5, 3.5, 3, 3.5, 3, 11.5), fill=BLUE, width=max(1, round(1.3 * k)))
        _outline_box(cv, P, k, 3, 13.5, "-")


def _icon_subtotal(cv, P, k, c):
    for i in range(4):
        y = 1.5 + 3.5 * i
        total = i == 2 or i == 3
        cv.create_rectangle(*P(6, y, 15, y + 3), fill="#dde8f6" if total else "#ffffff", outline=GRAY)
    _glyph(cv, P, k, 3, 5.5, "Σ", 9, BLUE, "bold")
    _outline_box(cv, P, k, 3, 12.5, "-")


def _icon_detail(cv, P, k, c, show):
    _grid(cv, P, k, 5.5, 1.5, 14.5, 14.5, 3, GRAY, "#ffffff")
    cv.create_rectangle(*P(5.5, 10.17, 14.5, 14.5), fill="#e2efda", outline=GRAY)
    _outline_box(cv, P, k, 2.5, 12.3, "+" if show else "-")


def _tag(cv, P, k, x, y, s=1.0, fill="#ffd65c"):
    """A name tag: a label with a punched hole."""
    cv.create_polygon(*P(x, y + 2.5 * s, x + 2.5 * s, y, x + 9 * s, y, x + 9 * s, y + 5 * s, x + 2.5 * s, y + 5 * s),
                      fill=fill, outline="#b38f00")
    cv.create_oval(*P(x + 2 * s, y + 2 * s, x + 3.2 * s, y + 3.2 * s), fill="#ffffff", outline="#b38f00")


def _icon_name(cv, P, k, c, kind):
    if kind == "manager":
        _grid(cv, P, k, 1.5, 4.5, 14.5, 14.5, 3, GRAY, "#ffffff")
        _tag(cv, P, k, 1, 0.5, 1.1)
    elif kind == "define":
        _tag(cv, P, k, 0.5, 4, 1.6)
        _pencil(cv, P, k, 7, 6, 0.55)
    elif kind == "use":
        _glyph(cv, P, k, 4.5, 10.5, "fx", 8, BLUE, "bold", "italic", "Times New Roman")
        _tag(cv, P, k, 6, 1.5, 1.0)
    else:  # create from selection
        _grid(cv, P, k, 1.5, 1.5, 14.5, 14.5, 3, GRAY, "#ffffff")
        cv.create_rectangle(*P(1.5, 1.5, 14.5, 5.83), fill="#bdd7ee", outline=GRAY)
        _tag(cv, P, k, 6, 8.5, 0.95)


def _icon_picture(cv, P, k, c):
    cv.create_rectangle(*P(1.5, 2.5, 14.5, 13.5), fill="#ffffff", outline=GRAY)
    cv.create_rectangle(*P(3, 4, 13, 12), fill="#ddebf7", outline="")
    cv.create_polygon(*P(3, 12, 6.5, 7, 9, 10, 10.5, 8.5, 13, 12), fill="#70ad47", outline="")
    cv.create_oval(*P(9.5, 4.8, 11.9, 7.2), fill="#ffc000", outline="")


def _icon_link(cv, P, k, c):
    w = max(2, round(1.7 * k))
    for x, y in ((2.5, 7.5), (7.5, 3.5)):
        cv.create_polygon(*P(x + 1.5, y + 5, x, y + 3.5, x + 4, y - 0.5, x + 5.5, y + 1), fill="", outline=BLUE,
                          width=w)
    cv.create_line(*P(6, 10, 10, 6), fill=BLUE, width=w)


def _icon_launcher(cv, P, k, c):
    cv.create_line(*P(2, 2, 2, 12), fill=GRAY)
    cv.create_line(*P(2, 2, 12, 2), fill=GRAY)
    _arrow(cv, P, k, 5, 5, 13, 13, GRAY, 1)


ICONS = {
    "paste": _icon_paste, "cut": _icon_cut, "copy": _icon_copy, "painter": _icon_painter,
    "grow_font": lambda *a: _icon_grow(*a, True), "shrink_font": lambda *a: _icon_grow(*a, False),
    "bold": lambda cv, P, k, c: _glyph(cv, P, k, 8, 8, "B", 13, DARK, "bold"),
    "italic": lambda cv, P, k, c: _glyph(cv, P, k, 8, 8, "I", 13, DARK, slant="italic", family="Times New Roman"),
    "underline": lambda cv, P, k, c: _glyph(cv, P, k, 8, 7.5, "U", 12, DARK, underline=True),
    "strike": lambda cv, P, k, c: (_glyph(cv, P, k, 8, 8, "ab", 11, DARK),
                                   cv.create_line(*P(1.5, 9, 14.5, 9), fill=RED, width=max(1, round(k)))),
    "borders": _icon_borders, "fill": _icon_fill, "font_color": _icon_font_color,
    "valign_top": lambda *a: _icon_valign(*a, "top"), "valign_center": lambda *a: _icon_valign(*a, "center"),
    "valign_bottom": lambda *a: _icon_valign(*a, "bottom"),
    "align_left": lambda *a: _icon_halign(*a, "left"), "align_center": lambda *a: _icon_halign(*a, "center"),
    "align_right": lambda *a: _icon_halign(*a, "right"),
    "indent_dec": lambda *a: _icon_indent(*a, False), "indent_inc": lambda *a: _icon_indent(*a, True),
    "orientation": _icon_orientation, "wrap": _icon_wrap, "merge": _icon_merge,
    "accounting": _icon_accounting,
    "percent": lambda cv, P, k, c: _glyph(cv, P, k, 8, 8, "%", 13, DARK, "bold"),
    "comma": lambda cv, P, k, c: _glyph(cv, P, k, 8, 5, ",", 20, DARK, "bold"),
    "dec_inc": lambda *a: _icon_decimals(*a, True), "dec_dec": lambda *a: _icon_decimals(*a, False),
    "cond_format": _icon_cond_format, "format_table": _icon_format_table,
    "insert_cells": lambda *a: _icon_cells(*a, "insert"), "delete_cells": lambda *a: _icon_cells(*a, "delete"),
    "format_cells": lambda *a: _icon_cells(*a, "format"),
    "autosum": _icon_autosum, "fill_down": _icon_fill_down, "clear": _icon_clear,
    "sort_filter": _icon_sort_filter, "sort_az": lambda *a: _icon_sort(*a, False),
    "sort_za": lambda *a: _icon_sort(*a, True), "filter": _icon_filter,
    "filter_clear": lambda *a: _icon_filter(*a, True), "find": _icon_find,
    "function": _icon_function, "book": _icon_book, "save": _icon_save, "undo": _icon_undo,
    "redo": lambda *a: _icon_undo(*a, True), "sheet": _icon_sheet,
    "insert_row": lambda *a: _icon_rowcol(*a, "row"), "insert_col": lambda *a: _icon_rowcol(*a, "col"),
    "delete_row": lambda *a: _icon_rowcol(*a, "row_del"), "delete_col": lambda *a: _icon_rowcol(*a, "col_del"),
    "calendar": _icon_calendar, "clock": _icon_clock, "show_formulas": _icon_show_formulas, "calc": _icon_calc,
    "zoom": _icon_zoom, "zoom100": lambda *a: _icon_zoom(*a, label="1:1"),
    "zoom_sel": lambda *a: _icon_zoom(*a, label="□"),
    "remove_dup": _icon_remove_dup, "help": _icon_help, "info": _icon_info, "launcher": _icon_launcher,
    "freeze": _icon_freeze, "validation": _icon_validation, "note": _icon_note,
    "note_delete": lambda *a: _icon_note(*a, delete=True),
    "note_prev": lambda *a: _icon_note_step(*a, False), "note_next": lambda *a: _icon_note_step(*a, True),
    "chart_column": _icon_chart, "chart_bar": lambda *a: _icon_chart(*a, kind="bar"),
    "chart_line": lambda *a: _icon_chart(*a, kind="line"), "chart_pie": lambda *a: _icon_chart(*a, kind="pie"),
}
ICONS.update({f"border_{kind}": (lambda kind: lambda *a: _icon_border(*a, kind))(kind) for kind in _BORDER_EDGES})
ICONS.update({f"line_{style}": (lambda style: lambda *a: _icon_line_style(*a, style))(style)
              for style in ("thin", "medium", "thick", "double", "dashed", "dotted")})
ICONS.update({f"paste_{kind}": (lambda kind: lambda *a: _icon_paste_as(*a, kind))(kind)
              for kind in ("formulas", "values", "formats", "transpose", "special")})
ICONS.update({f"merge_{kind}": (lambda kind: lambda *a: _icon_merge_kind(*a, kind))(kind)
              for kind in ("across", "cells", "unmerge")})
ICONS.update({f"angle_{name}": (lambda angle: lambda *a: _icon_text_angle(*a, angle))(angle)
              for name, angle in (("ccw", 45), ("cw", -45), ("vertical", 255), ("up", 90), ("down", -90))})
ICONS.update({f"cf_{kind}": (lambda kind: lambda *a: _icon_cf_kind(*a, kind))(kind)
              for kind in ("highlight", "topbottom", "databar", "scale", "manage")})
ICONS.update({f"cf_rule_{name}": (lambda text: lambda *a: _icon_cf_rule(*a, text))(text)
              for name, text in (("gt", ">"), ("lt", "<"), ("between", "↔"), ("eq", "="), ("text", "ab"),
                                 ("dup", "2×"))})
ICONS.update({f"rank_{name}": (lambda text, up: lambda *a: _icon_rank(*a, text, up))(text, up)
              for name, text, up in (("top", "10", True), ("top_pct", "%", True), ("bottom", "10", False),
                                     ("bottom_pct", "%", False), ("above", "x̄", True), ("below", "x̄", False))})
ICONS.update({f"stat_{kind}": (lambda kind: lambda *a: _icon_stat(*a, kind))(kind)
              for kind in ("avg", "count", "max", "min")})
ICONS.update({f"fill_{d}": (lambda d: lambda *a: _icon_fill_dir(*a, d))(d) for d in ("right", "up", "left")})
ICONS.update({f"clear_{what}": (lambda what: lambda *a: _icon_clear_what(*a, what))(what)
              for what in ("formats", "contents")})
ICONS.update({f"{axis}_{kind}": (lambda axis, auto: lambda *a: _icon_size(*a, axis, auto))(axis, kind == "autofit")
              for axis in ("row", "col") for kind in ("size", "autofit")})
ICONS.update({f"{verb}_{axis}": (lambda axis, show: lambda *a: _icon_hide(*a, axis, show))(axis, verb == "unhide")
              for axis in ("row", "col") for verb in ("hide", "unhide")})
ICONS.update({f"sheet_{op}": (lambda op: lambda *a: _icon_sheet_op(*a, op))(op)
              for op in ("copy", "rename", "delete", "tab_color", "left", "right", "hide", "unhide")})
ICONS.update({f"file_{kind}": (lambda kind: lambda *a: _icon_file(*a, kind))(kind)
              for kind in ("new", "open", "save_as", "exit")})
ICONS.update({
    "line_color": _icon_line_color, "underline_double": _icon_underline_double, "sort_custom": _icon_sort_custom,
    "filter_reapply": lambda *a: _icon_filter_variant(*a, "reapply"),
    "filter_value": lambda *a: _icon_filter_variant(*a, "value"), "fill_series": _icon_fill_series,
    "replace": _icon_replace, "goto": _icon_goto, "freeze_row": lambda *a: _icon_freeze_part(*a, "row"),
    "freeze_col": lambda *a: _icon_freeze_part(*a, "col"), "validation_clear": _icon_validation_clear,
    "delete_x": _icon_delete_x, "chart_edit": _icon_chart_edit,
    "sort_color": lambda *a: _icon_by_color(*a, True), "filter_color": lambda *a: _icon_by_color(*a, False),
    "pivot": _icon_pivot, "refresh": _icon_refresh, "pivot_fields": _icon_pivot_fields,
    "margins": _icon_margins, "page_orient": _icon_page_orient,
    "portrait": lambda *a: _icon_page_orient(*a, "portrait"), "landscape": lambda *a: _icon_page_orient(*a, "landscape"),
    "paper_size": _icon_paper_size, "print_area": _icon_print_area, "page_break": _icon_page_break,
    "print_titles": _icon_print_titles, "fit_page": _icon_fit_page, "print": _icon_print,
    "print_preview": _icon_print_preview, "pdf": _icon_pdf, "page_setup": _icon_page_setup,
    "group": _icon_group, "ungroup": lambda *a: _icon_group(*a, ungroup=True), "subtotal": _icon_subtotal,
    "show_detail": lambda *a: _icon_detail(*a, True), "hide_detail": lambda *a: _icon_detail(*a, False),
})
ICONS.update({"picture": _icon_picture, "link": _icon_link})
ICONS.update({f"name_{kind}": (lambda kind: lambda *a: _icon_name(*a, kind))(kind)
              for kind in ("manager", "define", "use", "create")})


def draw_icon(cv, name, x, y, size, color=None):
    """Draw icon `name` with its top-left corner at (x, y), `size` pixels square."""
    func = ICONS.get(name)
    if func is None:
        return
    k = size / 16.0

    def P(*coords):
        return [x + v * k if i % 2 == 0 else y + v * k for i, v in enumerate(coords)]
    func(cv, P, k, color)


class MenuIcons:
    """Icons for tk.Menu entries, cut from the sprite sheets made by tools/make_menu_icons.py.

    get(name) gives a PhotoImage; a blank one for names without a picture (so labels line up), or None
    when the sheets cannot be loaded at all (menus then simply have no pictures)."""

    SIZES = (16, 20, 24, 32)
    COLUMNS = 16

    def __init__(self, root, scale, folder):
        self.root = root
        self.size = min(self.SIZES, key=lambda s: abs(s - 16 * scale))
        self._cache = {}
        try:
            with open(os.path.join(folder, "menu_icons.txt"), encoding="utf-8") as f:
                self.index = {name: i for i, name in enumerate(f.read().split("\n")) if name}
            self.sheet = tk.PhotoImage(master=root, file=os.path.join(folder, f"menu_{self.size}.png"))
        except (OSError, tk.TclError):
            self.index, self.sheet = {}, None
        self.gap = max(3, self.size // 4)  # transparent margin between picture and label
        self.blank = tk.PhotoImage(master=root, width=self.size + self.gap, height=self.size)

    def get(self, name):
        if self.sheet is None:
            return None
        if name not in self.index:
            return self.blank
        image = self._cache.get(name)
        if image is None:
            s, i = self.size, self.index[name]
            x, y = (i % self.COLUMNS) * s, (i // self.COLUMNS) * s
            image = self._cache[name] = tk.PhotoImage(master=self.root, width=s + self.gap, height=s)
            image.tk.call(image, "copy", self.sheet, "-from", x, y, x + s, y + s, "-to", 0, 0)
        return image


# ------------------------------------------------------------------ widgets

class Tooltip:
    def __init__(self, widget, text, font=None, delay=550):
        self.widget, self.text, self.font, self.delay = widget, text, font, delay
        self.job = self.tip = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, e=None):
        self._hide()
        self.job = self.widget.after(self.delay, self._show)

    def _show(self):
        self.job = None
        if not self.widget.winfo_viewable():
            return
        tip = self.tip = tk.Toplevel(self.widget)
        tip.overrideredirect(True)
        tip.attributes("-topmost", True)
        tk.Label(tip, text=self.text, bg="#ffffff", fg="#333333", bd=1, relief="solid", padx=7, pady=4,
                 justify="left", font=self.font).pack()
        tip.update_idletasks()
        x = min(self.widget.winfo_rootx(), tip.winfo_screenwidth() - tip.winfo_reqwidth() - 4)
        tip.geometry(f"+{x}+{self.widget.winfo_rooty() + self.widget.winfo_height() + 4}")

    def _hide(self, e=None):
        if self.job:
            self.widget.after_cancel(self.job)
            self.job = None
        if self.tip:
            self.tip.destroy()
            self.tip = None


class RibbonButton(tk.Canvas):
    """A ribbon button drawn on a canvas.

    large: icon above a label (use '\\n' for two lines); otherwise icon left of the label.
    menu: a tk.Menu, or a callable(button) that opens its own dropdown below the button.
    split: with a menu, the arrow part opens it while the rest runs `command`.
    """

    def __init__(self, master, ui, *, text="", icon=None, command=None, menu=None, large=False, split=False,
                 tooltip=None, color=None, bg=None, fg=TEXT_FG, hover_bg=HOVER_BG, height=None):
        self.ui, self.text, self.icon, self.command = ui, text, icon, command
        self.menu, self.large, self.split, self.color = menu, large, split and menu is not None, color
        self.bg, self.fg, self.hover_bg = bg or ui.bg, fg, hover_bg
        self.checked = self.pressed = False
        self.hover = None
        self.enabled = True
        self.arrow_w = round(11 * ui.scale) if menu else 0
        self.icon_px = round((32 if large else 16) * ui.scale)
        height = height or (ui.content_h if large else ui.row_h)
        super().__init__(master, width=self._width(), height=height, bg=self.bg, highlightthickness=0, bd=0,
                         takefocus=0)
        self.bind("<Enter>", self._on_motion)
        self.bind("<Motion>", self._on_motion)
        self.bind("<Leave>", lambda e: self._set_hover(None))
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        if tooltip:
            Tooltip(self, tooltip, ui.small_font)
        self.draw()

    def _width(self):
        s, font = self.ui.scale, self.ui.font
        if self.large:
            lines = self.text.split("\n") if self.text else [""]
            widths = [font.measure(line) for line in lines]
            if len(lines) > 1:  # the arrow follows the second line; one-line labels get it on a line of its own
                widths[-1] += self.arrow_w
            return max([self.icon_px] + widths) + round(8 * s)
        return round(4 * s) + (self.icon_px if self.icon else 0) + \
            (round(5 * s) + font.measure(self.text) if self.text else 0) + self.arrow_w + round(4 * s)

    def set_text(self, text):
        """Change the label (the ribbon hides labels when the window is narrow)."""
        if text != self.text:
            self.text = text
            self.config(width=self._width())
            self.draw()

    def set_checked(self, value):
        if bool(value) != self.checked:
            self.checked = bool(value)
            self.draw()

    def set_color(self, color):
        self.color = color
        self.draw()

    def set_enabled(self, value):
        self.enabled = bool(value)
        self.draw()

    def _boxes(self):
        w, h = int(self["width"]), int(self["height"])
        if self.large:
            split_y = round(4 * self.ui.scale) + self.icon_px + round(3 * self.ui.scale)
            return (0, 0, w - 1, split_y), (0, split_y + 1, w - 1, h - 1)
        return (0, 0, w - self.arrow_w - 1, h - 1), (w - self.arrow_w, 0, w - 1, h - 1)

    def draw(self):
        self.delete("all")
        w, h, s = int(self["width"]), int(self["height"]), self.ui.scale
        if self.checked:
            self.create_rectangle(0, 0, w - 1, h - 1, fill=CHECKED_BG, outline=CHECKED_LINE)
        if self.hover and self.enabled:
            hot = PRESSED_BG if self.pressed else self.hover_bg
            if self.split:
                main, arrow = self._boxes()
                active, other = (arrow, main) if self.hover == "arrow" else (main, arrow)
                self.create_rectangle(*other, fill=HOVER_SOFT if self.bg == RIBBON_BG else self.hover_bg, outline="")
                self.create_rectangle(*active, fill=hot, outline="")
            else:
                self.create_rectangle(0, 0, w - 1, h - 1, fill=hot, outline="")
        fg = self.fg if self.enabled else "#a0a0a0"
        if self.large:
            iy = round(4 * s)
            draw_icon(self, self.icon, (w - self.icon_px) // 2, iy, self.icon_px, self.color)
            lines = self.text.split("\n") if self.text else []
            line_h = self.ui.font.metrics("linespace")
            y = iy + self.icon_px + round(4 * s) + line_h // 2
            inline = len(lines) > 1
            for i, line in enumerate(lines):
                last = i == len(lines) - 1
                tw = self.ui.font.measure(line) + (self.arrow_w if last and inline and self.menu else 0)
                x = (w - tw) // 2
                self.create_text(x, y, text=line, anchor="w", font=self.ui.font, fill=fg)
                if last and inline and self.menu:
                    self._triangle(x + self.ui.font.measure(line) + self.arrow_w // 2 + 1, y)
                y += line_h
            if self.menu and not inline:
                self._triangle(w // 2, y)
        else:
            x = round(4 * s)
            if self.icon:
                draw_icon(self, self.icon, x, (h - self.icon_px) // 2, self.icon_px, self.color)
                x += self.icon_px + round(5 * s)
            if self.text:
                self.create_text(x, h // 2, text=self.text, anchor="w", font=self.ui.font, fill=fg)
            if self.menu:
                self._triangle(w - self.arrow_w // 2 - 1, h // 2)

    def _triangle(self, cx, cy):
        d = max(2, round(2.6 * self.ui.scale))
        color = "#ffffff" if self.bg != RIBBON_BG else DARK
        self.create_polygon(cx - d, cy - d // 2, cx + d, cy - d // 2, cx, cy + d - d // 2, fill=color, outline="")

    def _part(self, x, y):
        if not self.menu:
            return "main"
        if not self.split:
            return "arrow"
        arrow = self._boxes()[1]
        return "arrow" if arrow[0] <= x <= arrow[2] and arrow[1] <= y <= arrow[3] else "main"

    def _set_hover(self, part):
        if part != self.hover:
            self.hover = part
            if part is None:
                self.pressed = False
            self.draw()

    def _on_motion(self, e):
        self._set_hover(self._part(e.x, e.y))

    def _on_press(self, e):
        if not self.enabled:
            return
        if self._part(e.x, e.y) == "arrow":
            self.open_menu()
            return
        self.pressed = True
        self.draw()

    def _on_release(self, e):
        if not self.pressed:
            return
        self.pressed = False
        self.draw()
        if 0 <= e.x < int(self["width"]) and 0 <= e.y < int(self["height"]) and self.command:
            self.command()

    def open_menu(self):
        self._set_hover(None)
        if callable(self.menu):
            self.menu(self)
        else:
            self.menu.tk_popup(self.winfo_rootx(), self.winfo_rooty() + self.winfo_height())
            self.menu.grab_release()


class Popup(tk.Toplevel):
    """Borderless dropdown window that closes on Escape or a click outside it."""

    def __init__(self, master, x, y):
        super().__init__(master)
        self.withdraw()
        self.overrideredirect(True)
        self.configure(bg="#c6c6c6")
        self.inner = tk.Frame(self, bg="#ffffff")
        self.inner.pack(padx=1, pady=1, fill="both", expand=True)
        self._pos = (x, y)
        self.bind("<Escape>", lambda e: self.close())

    def show(self):
        self.update_idletasks()
        x, y = self._pos
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        x = max(0, min(x, self.winfo_screenwidth() - w - 4))
        if y + h > self.winfo_screenheight() - 40:
            y = max(0, y - h - 30)
        self.geometry(f"+{x}+{y}")
        self.deiconify()
        self.lift()
        self.attributes("-topmost", True)
        self.focus_force()
        self._prev_grab = self.grab_current()  # e.g. a modal dialog the popup was opened from
        self.grab_set()
        self.bind("<ButtonPress>", self._on_click, add="+")
        return self

    def _on_click(self, e):
        x0, y0 = self.winfo_rootx(), self.winfo_rooty()
        if not (x0 <= e.x_root < x0 + self.winfo_width() and y0 <= e.y_root < y0 + self.winfo_height()):
            self.close()

    def close(self):
        prev = getattr(self, "_prev_grab", None)
        try:
            self.grab_release()
        finally:
            self.destroy()
            if prev is not None and prev.winfo_exists():
                prev.grab_set()
                prev.focus_set()


def _hover_link(widget, normal, hot):
    widget.bind("<Enter>", lambda e: widget.config(bg=hot), add="+")
    widget.bind("<Leave>", lambda e: widget.config(bg=normal), add="+")


def color_popup(button, ui, on_pick, none_label, on_more):
    """Excel's colour dropdown: 'no colour' entry, theme grid, standard colours, 'more colours'."""
    pop = Popup(button.winfo_toplevel(), button.winfo_rootx(), button.winfo_rooty() + button.winfo_height())
    f = pop.inner
    pad = ui.px(6)

    def pick(color):
        pop.close()
        on_pick(color)

    def entry(text, command):
        lbl = tk.Label(f, text=text, bg="#ffffff", fg=TEXT_FG, anchor="w", font=ui.font, padx=pad, pady=ui.px(3))
        lbl.pack(fill="x", padx=2, pady=1)
        _hover_link(lbl, "#ffffff", HOVER_BG)
        lbl.bind("<ButtonRelease-1>", lambda e: command())

    entry("  " + none_label, lambda: pick(None))
    size, gap = ui.px(14), ui.px(3)

    def swatch_grid(title, rows):
        tk.Label(f, text=title, bg="#f0f0f0", fg=TEXT_FG, anchor="w", font=ui.small_font, padx=pad).pack(fill="x", pady=(4, 2))
        grid = tk.Frame(f, bg="#ffffff")
        grid.pack(padx=pad, pady=(0, 4))
        for i, row in enumerate(rows):
            for j, color in enumerate(row):
                sw = tk.Frame(grid, width=size, height=size, bg=color, highlightthickness=1,
                              highlightbackground="#d0d0d0" if color != "#ffffff" else "#bdbdbd")
                sw.grid(row=i, column=j, padx=gap // 2, pady=(gap if i == 1 else 0, 0))
                sw.bind("<Enter>", lambda e, w=sw: w.config(highlightbackground="#f29436"))
                sw.bind("<Leave>", lambda e, w=sw, c=color: w.config(highlightbackground="#d0d0d0" if c != "#ffffff" else "#bdbdbd"))
                sw.bind("<ButtonRelease-1>", lambda e, c=color: pick(c))
                Tooltip(sw, color.upper(), ui.small_font, 700)

    swatch_grid("Màu chủ đề", theme_palette())
    swatch_grid("Màu chuẩn", [STANDARD_COLORS])
    tk.Frame(f, bg=GROUP_LINE, height=1).pack(fill="x", padx=2)
    entry("  Màu khác...", lambda: (pop.close(), on_more()))
    return pop.show()


def gallery_popup(button, ui, sections, on_pick, tile_w=None):
    """Dropdown gallery: sections of (title, [(label, preview_spec, value)]) shown as styled tiles.

    preview_spec is a dict with "bg", "fg", "bold", "italic", "border" or a callable(canvas, w, h)."""
    pop = Popup(button.winfo_toplevel(), button.winfo_rootx(), button.winfo_rooty() + button.winfo_height())
    f = pop.inner
    tile_w = tile_w or ui.px(96)
    tile_h = ui.px(26)

    def pick(value):
        pop.close()
        on_pick(value)

    for title, items in sections:
        tk.Label(f, text=title, bg="#f0f0f0", fg=TEXT_FG, anchor="w", font=ui.small_font,
                 padx=ui.px(6)).pack(fill="x", pady=(2, 2))
        grid = tk.Frame(f, bg="#ffffff")
        grid.pack(padx=ui.px(6), pady=(0, ui.px(4)), anchor="w")
        per_row = max(1, round(ui.px(560) / (tile_w + ui.px(6))))
        for n, (label, spec, value) in enumerate(items):
            tile = StyleTile(grid, ui, label, spec, tile_w, tile_h, lambda v=value: pick(v))
            tile.grid(row=n // per_row, column=n % per_row, padx=ui.px(2), pady=ui.px(2))
    return pop.show()


class StyleTile(tk.Canvas):
    """A small preview of a cell style (or a custom drawing) that reacts to hover and click."""

    def __init__(self, master, ui, label, spec, w, h, command):
        super().__init__(master, width=w, height=h, bg="#ffffff", highlightthickness=1,
                         highlightbackground="#d6d6d6", bd=0, cursor="hand2")
        self.command = command
        if callable(spec):
            spec(self, w, h)
        else:
            font = tkfont.Font(family=ui.family, size=-round(12 * ui.scale),
                               weight="bold" if spec.get("bold") else "normal",
                               slant="italic" if spec.get("italic") else "roman")
            self._font = font
            self.create_rectangle(0, 0, w, h, fill=spec.get("bg") or "#ffffff", outline="")
            if spec.get("border"):
                self.create_rectangle(1, 1, w - 2, h - 2, outline=spec["border"], width=2 if spec.get("double") else 1)
            if spec.get("underline"):
                self.create_line(2, h - 3, w - 2, h - 3, fill=spec["underline"], width=2)
            text = label if font.measure(label) < w - 8 else label[:max(1, int(len(label) * (w - 12) / font.measure(label)))] + "…"
            self.create_text(5, h // 2, text=text, anchor="w", font=font, fill=spec.get("fg") or "#000000")
        self.bind("<Enter>", lambda e: self.config(highlightbackground="#f29436"))
        self.bind("<Leave>", lambda e: self.config(highlightbackground="#d6d6d6"))
        self.bind("<ButtonRelease-1>", lambda e: self.command())
        Tooltip(self, label, ui.small_font, 700)


class RibbonGroup(tk.Frame):
    def __init__(self, page, ui, title, launcher=None, launcher_tip=None):
        super().__init__(page, bg=ui.bg)
        self.ui = ui
        self.pack(side="left", fill="y")
        self.body = tk.Frame(self, bg=ui.bg)
        self.body.pack(side="top", fill="both", expand=True, padx=ui.px(3), pady=(ui.px(2), 0))
        caption_w = ui.small_font.measure(title) + ui.px(10) + (ui.px(32) if launcher else 0)
        foot = tk.Frame(self, bg=ui.bg, height=ui.label_h, width=caption_w)  # the group is never narrower
        foot.pack(side="bottom", fill="x")
        foot.pack_propagate(False)
        tk.Label(foot, text=title, bg=ui.bg, fg=GROUP_FG, font=ui.small_font).pack(side="left", expand=True)
        if launcher:
            size = ui.px(13)
            btn = tk.Canvas(foot, width=size, height=size, bg=ui.bg, highlightthickness=0, bd=0)
            draw_icon(btn, "launcher", 0, 0, size)
            btn.pack(side="right", padx=(0, ui.px(3)))
            btn.bind("<ButtonRelease-1>", lambda e: launcher())
            btn.bind("<Enter>", lambda e: btn.config(bg=HOVER_BG))
            btn.bind("<Leave>", lambda e: btn.config(bg=ui.bg))
            if launcher_tip:
                Tooltip(btn, launcher_tip, ui.small_font)
        tk.Frame(page, bg=GROUP_LINE, width=1).pack(side="left", fill="y", pady=ui.px(5))

    def large(self, **kw):
        btn = RibbonButton(self.body, self.ui, large=True, **kw)
        btn.pack(side="left", fill="y", padx=1)
        return btn

    def column(self):
        """A stack of up to three small controls."""
        frame = tk.Frame(self.body, bg=self.ui.bg)
        frame.pack(side="left", fill="y", padx=1)
        return frame

    def rows(self, n=3):
        """Up to three horizontal rows of small controls."""
        outer = tk.Frame(self.body, bg=self.ui.bg)
        outer.pack(side="left", fill="y", padx=1)
        rows = []
        for _ in range(n):
            row = tk.Frame(outer, bg=self.ui.bg, height=self.ui.row_h)
            row.pack(side="top", fill="x", anchor="w")
            rows.append(row)
        return rows

    def small(self, parent, side="left", **kw):
        btn = RibbonButton(parent, self.ui, **kw)
        btn.pack(side=side, anchor="w")
        return btn

    def gap(self, parent, width=6):
        tk.Frame(parent, bg=self.ui.bg, width=self.ui.px(width)).pack(side="left")


class Ribbon(tk.Frame):
    """Green tab strip with a quick-access toolbar, plus the active tab's page of groups."""

    def __init__(self, master, ui):
        super().__init__(master, bg=ui.bg)
        self.ui = ui
        self.bar = tk.Frame(self, bg=TAB_BAR_BG)
        self.bar.pack(side="top", fill="x")
        self.qat = tk.Frame(self.bar, bg=TAB_BAR_BG)
        self.qat.pack(side="left", padx=(ui.px(4), ui.px(10)))
        self.tab_strip = tk.Frame(self.bar, bg=TAB_BAR_BG)
        self.tab_strip.pack(side="left", fill="y")
        self.extra = tk.Frame(self.bar, bg=TAB_BAR_BG)
        self.extra.pack(side="left", fill="both", expand=True, padx=(ui.px(16), ui.px(8)))
        self.body = tk.Frame(self, bg=ui.bg, height=ui.content_h + ui.label_h + ui.px(4))
        self.body.pack(side="top", fill="x")
        self.body.pack_propagate(False)
        self.line = tk.Frame(self, bg=GROUP_LINE, height=1)
        self.line.pack(side="top", fill="x")
        self.pages, self.labels = {}, {}
        self.current = None
        self.collapsed = False

    def add_tab(self, name, command=None):
        """A tab with a page of groups; with `command` the tab only runs it (like 'File')."""
        label = tk.Label(self.tab_strip, text=name, bg=TAB_BAR_BG, fg="#ffffff", font=self.ui.tab_font,
                         padx=self.ui.px(11), pady=self.ui.px(4), cursor="hand2")
        label.pack(side="left", fill="y")
        self.labels[name] = label
        label.bind("<Enter>", lambda e: label.config(bg=TAB_HOVER_BG) if name != self.current else None)
        label.bind("<Leave>", lambda e: label.config(bg=TAB_BAR_BG) if name != self.current else None)
        if command:
            label.bind("<ButtonRelease-1>", lambda e: command(label))
            return None
        page = tk.Frame(self.body, bg=self.ui.bg)
        self.pages[name] = page
        label.bind("<ButtonRelease-1>", lambda e: self.select(name))
        label.bind("<Double-Button-1>", lambda e: self.toggle_collapse())
        if self.current is None:
            self.select(name)
        return page

    def group(self, page, title, launcher=None, launcher_tip=None):
        return RibbonGroup(page, self.ui, title, launcher, launcher_tip)

    def select(self, name):
        if self.collapsed:
            self.toggle_collapse()
        for n, page in self.pages.items():
            if n != name:
                page.pack_forget()
        self.pages[name].pack(side="left", fill="y", padx=(self.ui.px(4), 0))
        self.current = name
        for n, label in self.labels.items():
            active = n == name
            label.config(bg=self.ui.bg if active else TAB_BAR_BG, fg=TAB_BAR_BG if active else "#ffffff")

    def toggle_collapse(self):
        self.collapsed = not self.collapsed
        if self.collapsed:
            self.body.pack_forget()
        else:
            self.body.pack(side="top", fill="x", before=self.line)
