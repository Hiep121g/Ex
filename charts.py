"""Chart drawing for PyExcel: column, bar, line and pie charts drawn with tkinter canvas items.

The workbook decides what a chart plots (Workbook.chart_data); this module only lays it out:
title, legend, value axis with round ticks, category labels and the marks themselves.
"""

import math

import formula_engine as fe

COLORS = ("#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47", "#264478", "#9e480e", "#636363",
          "#997300")
FRAME = "#d9d9d9"
AXIS = "#bfbfbf"
GRIDLINE = "#e7e7e7"
LABEL = "#595959"
TITLE = "#404040"
NAMES = {"column": "Cột", "bar": "Thanh", "line": "Đường", "pie": "Tròn"}


def nice_ticks(lo, hi, count=5):
    """Evenly spaced round values (steps of 1, 2, 2.5 or 5 times a power of ten) covering lo..hi and 0."""
    lo, hi = min(lo, 0.0), max(hi, 0.0)
    if lo == hi:
        hi = 1.0
    raw = (hi - lo) / count
    mag = 10 ** math.floor(math.log10(raw))
    step = next(m * mag for m in (1, 2, 2.5, 5, 10) if m * mag >= raw * (1 - 1e-9))
    start = math.floor(lo / step + 1e-9) * step
    n = max(1, round(math.ceil(hi / step - 1e-9) - start / step))
    return [round(start + i * step, 10) for i in range(n + 1)]


def tick_label(v, step):
    """Axis text for a tick: thousands separators, and as many decimals as the step needs."""
    decimals = next((d for d in range(7) if abs(round(step, d) - step) < 1e-9), 6)
    return fe.format_number(v, "#,##0" + ("." + "0" * decimals if decimals else ""))


def _fit(text, font, width):
    """Text cut with an ellipsis to fit `width` pixels."""
    if font.measure(text) <= width:
        return text
    while text and font.measure(text + "…") > width:
        text = text[:-1]
    return text + "…" if text else ""


def draw_chart(cv, box, kind, title, labels, series, title_font, font, k):
    """Draw a chart filling canvas box (x0, y0, x1, y1); k is the pixel scale (DPI times zoom)."""
    x0, y0, x1, y1 = box
    pad = 10 * k
    cv.create_rectangle(x0, y0, x1, y1, fill="#ffffff", outline=FRAME)
    if not title and len(series) == 1:
        title = series[0][0]  # like Excel, a single series names the chart
    top = y0 + pad
    if title:
        item = cv.create_text((x0 + x1) / 2, top, text=title, anchor="n", font=title_font, fill=TITLE,
                              width=max(10, x1 - x0 - 2 * pad), justify="center")
        top = cv.bbox(item)[3] + pad / 2
    if kind == "pie":
        values = series[0][1] if series else []
        slices = [(labels[i], v) for i, v in enumerate(values) if v is not None and v > 0]
        names = [name for name, _ in slices]
    else:
        slices = None
        names = [name for name, _ in series] if len(series) > 1 else []
    bottom = y1 - pad - _draw_legend(cv, names, x0 + pad, x1 - pad, y1 - pad, font, k)
    if bottom - top < 20 * k or x1 - x0 < 60 * k:
        return
    if kind == "pie":
        _draw_pie(cv, (x0 + pad, top, x1 - pad, bottom), slices, font, k)
    else:
        _draw_axes(cv, (x0 + pad, top, x1 - pad, bottom), kind, labels, series, font, k)


def _no_data(cv, area, font):
    cv.create_text((area[0] + area[2]) / 2, (area[1] + area[3]) / 2, text="Không có dữ liệu số để vẽ",
                   font=font, fill=LABEL)


def _draw_legend(cv, names, left, right, bottom, font, k):
    """Legend centred along the bottom, wrapping onto up to three lines; returns the height used."""
    if not names:
        return 0
    ls, swatch, gap = font.metrics("linespace"), 8 * k, 14 * k
    items = [(i, _fit(name, font, (right - left) * 0.6)) for i, name in enumerate(names)]
    lines, line, width = [], [], 0
    for i, text in items:
        w = swatch + 4 * k + font.measure(text)
        if line and width + gap + w > right - left:
            lines.append((line, width))
            line, width = [], 0
        width += (gap if line else 0) + w
        line.append((i, text, w))
    lines.append((line, width))
    lines = lines[:3]
    height = len(lines) * ls + 4 * k
    y = bottom - height + 4 * k + ls / 2
    for line, width in lines:
        x = (left + right - width) / 2
        for i, text, w in line:
            color = COLORS[i % len(COLORS)]
            cv.create_rectangle(x, y - swatch / 2, x + swatch, y + swatch / 2, fill=color, outline="")
            cv.create_text(x + swatch + 4 * k, y, text=text, anchor="w", font=font, fill=LABEL)
            x += w + gap
        y += ls
    return height


def _draw_pie(cv, area, slices, font, k):
    total = sum(v for _, v in slices)
    if not total:
        _no_data(cv, area, font)
        return
    left, top, right, bottom = area
    r = max(10.0, min(right - left, bottom - top) / 2 - 4 * k)
    cx, cy = (left + right) / 2, (top + bottom) / 2
    start = 90.0
    for i, (_, v) in enumerate(slices):
        extent = -360.0 * v / total
        color = COLORS[i % len(COLORS)]
        if extent <= -359.99:
            cv.create_oval(cx - r, cy - r, cx + r, cy + r, fill=color, outline="#ffffff")
        else:
            cv.create_arc(cx - r, cy - r, cx + r, cy + r, start=start, extent=extent, fill=color,
                          outline="#ffffff", width=max(1, round(k)))
        if extent < -18:
            mid = math.radians(start + extent / 2)
            cv.create_text(cx + 0.62 * r * math.cos(mid), cy - 0.62 * r * math.sin(mid),
                           text=f"{round(100 * v / total)}%", font=font, fill="#ffffff")
        start += extent


def _draw_axes(cv, area, kind, labels, series, font, k):
    values = [v for _, vals in series for v in vals if v is not None]
    if not values or not labels:
        _no_data(cv, area, font)
        return
    ticks = nice_ticks(min(values), max(values))
    lo, hi, step = ticks[0], ticks[-1], ticks[1] - ticks[0]
    texts = [tick_label(t, step) for t in ticks]
    left, top, right, bottom = area
    ls = font.metrics("linespace")
    n = len(labels)
    horizontal = kind == "bar"  # bars: categories up the left side, values along the bottom
    if horizontal:
        cat_w = min(max(font.measure(t) for t in labels), (right - left) * 0.35)
        left += cat_w + 6 * k
        right -= font.measure(texts[-1]) / 2
        bottom -= ls + 4 * k
    else:
        left += max(font.measure(t) for t in texts) + 6 * k
        bottom -= ls + 4 * k
        top += ls / 2
    if right - left < 10 or bottom - top < 10:
        return

    def at(v):  # pixel along the value axis
        frac = (v - lo) / (hi - lo)
        return left + frac * (right - left) if horizontal else bottom - frac * (bottom - top)

    for t, text in zip(ticks, texts):
        p = at(t)
        if horizontal:
            cv.create_line(p, top, p, bottom, fill=GRIDLINE)
            cv.create_text(p, bottom + 3 * k, text=text, anchor="n", font=font, fill=LABEL)
        else:
            cv.create_line(left, p, right, p, fill=GRIDLINE)
            cv.create_text(left - 5 * k, p, text=text, anchor="e", font=font, fill=LABEL)
    zero = at(0.0)
    if horizontal:
        cv.create_line(zero, top, zero, bottom, fill=AXIS)
        slot = (bottom - top) / n
    else:
        cv.create_line(left, zero, right, zero, fill=AXIS)
        slot = (right - left) / n

    # category labels: every m-th one when they would not fit side by side
    if horizontal:
        every = max(1, math.ceil(ls / slot))
    else:
        every = max(1, math.ceil((max(font.measure(t) for t in labels) + 6 * k) / slot))
    for i in range(0, n, every):
        if horizontal:
            cv.create_text(left - 6 * k, bottom - (i + 0.5) * slot, text=_fit(labels[i], font, cat_w),
                           anchor="e", font=font, fill=LABEL)
        else:
            cv.create_text(left + (i + 0.5) * slot, bottom + 3 * k, text=_fit(labels[i], font, slot * every - 4 * k),
                           anchor="n", font=font, fill=LABEL)

    if kind == "line":
        dot = max(2.0, 3 * k)
        for j, (_, vals) in enumerate(series):
            color = COLORS[j % len(COLORS)]
            run = []
            for i, v in enumerate(vals + [None]):
                if v is None:
                    if len(run) > 1:
                        cv.create_line(*[c for p in run for c in p], fill=color, width=max(2, round(2.25 * k)))
                    run = []
                    continue
                run.append((left + (i + 0.5) * slot, at(v)))
            for i, v in enumerate(vals):
                if v is not None:
                    x, y = left + (i + 0.5) * slot, at(v)
                    cv.create_oval(x - dot, y - dot, x + dot, y + dot, fill=color, outline=color)
        return
    group = slot * 0.65
    thick = group / len(series)
    for j, (_, vals) in enumerate(series):
        color = COLORS[j % len(COLORS)]
        for i, v in enumerate(vals):
            if v is None:
                continue
            a = (i * slot + (slot - group) / 2 + j * thick)
            if horizontal:
                y_hi = bottom - a
                cv.create_rectangle(min(zero, at(v)), y_hi - thick, max(zero, at(v)), y_hi, fill=color, outline="")
            else:
                x = left + a
                cv.create_rectangle(x, min(zero, at(v)), x + thick, max(zero, at(v)), fill=color, outline="")
