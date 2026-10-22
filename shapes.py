"""Shapes drawn over the cells (Insert > Shapes) and the diagram tools (Diagram tab): outlines, connection
points, text areas, connector routes, arrowheads and hit testing.

A shape is a dict kept with the sheet's charts and pictures (see formula_engine.SHAPE):

    {"type": "shape", "id": str, "kind": a KINDS key, "name": str,
     "anchor": (row, col, dx, dy), "size": (w, h),        # the box, pixels at 96 DPI, as charts
     "rot": degrees clockwise, "flip_h": bool, "flip_v": bool,
     "fill": "#rrggbb" or None, "line": "#rrggbb" or None, "line_w": points, "dash": None or a DASHES key,
     "text": str, "font": {"size", "bold", "italic", "underline", "color", "font"}, "align", "valign",
     # lines and connectors (kind in LINE_KINDS): the line runs from the box's top-left to its bottom-right
     # corner, flips swap them; arrowheads at the start ("head") and end ("tail"); an end glued to another
     # shape's connection point follows it: "start"/"end": {"id": shape id, "site": index}
     "head": None or an ARROWS key, "tail": ..., "start": ..., "end": ...,
     # callouts: where the tail points, as fractions of the box
     "tip": (fx, fy)}

Shape dicts are never changed in place (undo keeps copies): changes make a new dict.
Everything here works in one coordinate system given by the caller (canvas pixels, PDF points...); `px` is
the number of those units per pixel at 96 DPI.
"""

import math
from collections import namedtuple

Geometry = namedtuple("Geometry", "paths sites text")  # paths: [(points, closed, fill mode)]
# fill modes: "norm" the shape's fill, "lighten" / "darken" a lighter / darker shade of it (the top of a
# cylinder), "none" only a line (the extra lines of a flowchart shape)

SHAPE = "shape"
LINE_KINDS = ("line", "elbow", "curve")
DASHES = {"dash": (8, 4), "dot": (2, 3), "dashdot": (8, 3, 2, 3), "longdash": (14, 5)}
DASH_NAMES = {None: "Liền", "dash": "Gạch", "dot": "Chấm", "dashdot": "Gạch chấm", "longdash": "Gạch dài"}
ARROWS = {"triangle": "Tam giác", "arrow": "Mũi tên hở", "stealth": "Mũi nhọn", "diamond": "Hình thoi",
          "oval": "Tròn"}

_STEPS = 64  # points on a whole ellipse
CALLOUT_TIP = (0.25, 1.4)  # where a new callout points (fractions of its box)


# ------------------------------------------------------------------ building outlines

def _arc(cx, cy, rx, ry, a0, a1, steps=None):
    """Points of an elliptical arc from angle a0 to a1 (degrees; 0 = right, 90 = down)."""
    n = steps or max(3, round(abs(a1 - a0) / 360 * _STEPS))
    out = []
    for i in range(n + 1):
        a = math.radians(a0 + (a1 - a0) * i / n)
        out.append((cx + rx * math.cos(a), cy + ry * math.sin(a)))
    return out


def _cubic(p0, c1, c2, p1, steps=16):
    """Points of a cubic Bezier curve after p0."""
    out = []
    for i in range(1, steps + 1):
        t = i / steps
        m = 1 - t
        a, b, c, d = m * m * m, 3 * m * m * t, 3 * m * t * t, t * t * t
        out.append((a * p0[0] + b * c1[0] + c * c2[0] + d * p1[0], a * p0[1] + b * c1[1] + c * c2[1] + d * p1[1]))
    return out


def _quad(p0, c, p1, steps=16):
    """Points of a quadratic Bezier curve after p0."""
    out = []
    for i in range(1, steps + 1):
        t = i / steps
        m = 1 - t
        out.append((m * m * p0[0] + 2 * m * t * c[0] + t * t * p1[0], m * m * p0[1] + 2 * m * t * c[1] + t * t * p1[1]))
    return out


def _tlbr(w, h, top=None, left=None, bottom=None, right=None):
    """The four usual connection points (Excel's order: top, left, bottom, right), middles of the sides
    unless given."""
    return [(top or (w / 2, 0)) + (270,), (left or (0, h / 2)) + (180,), (bottom or (w / 2, h)) + (90,),
            (right or (w, h / 2)) + (0,)]


def _poly(points, w, h, text=None, sites=None, extra=()):
    return Geometry([(points, True, "norm")] + list(extra), sites if sites is not None else _tlbr(w, h),
                    text or (0, 0, w, h))


def _round_rect_points(w, h, r):
    r = max(0.0, min(r, w / 2, h / 2))
    return ([(r, 0), (w - r, 0)] + _arc(w - r, r, r, r, 270, 360)[1:] + [(w, h - r)] + _arc(w - r, h - r, r, r, 0, 90)[1:]
            + [(r, h)] + _arc(r, h - r, r, r, 90, 180)[1:] + [(0, r)] + _arc(r, r, r, r, 180, 270)[1:-1])


def _ellipse_sites(w, h):
    """Excel's eight points of an ellipse: top, top-left, left, bottom-left, bottom, bottom-right, right, top-right."""
    c = math.cos(math.radians(45))
    il, ir, it, ib = w / 2 - w / 2 * c, w / 2 + w / 2 * c, h / 2 - h / 2 * c, h / 2 + h / 2 * c
    return [(w / 2, 0, 270), (il, it, 225), (0, h / 2, 180), (il, ib, 135), (w / 2, h, 90), (ir, ib, 45),
            (w, h / 2, 0), (ir, it, 315)]


def _ellipse(w, h):
    c = math.cos(math.radians(45))
    return Geometry([(_arc(w / 2, h / 2, w / 2, h / 2, 0, 360)[:-1], True, "norm")], _ellipse_sites(w, h),
                    (w / 2 - w / 2 * c, h / 2 - h / 2 * c, w / 2 + w / 2 * c, h / 2 + h / 2 * c))


def _cylinder(w, h, ry):
    ry = min(ry, h / 2)
    body = [(0, ry), (0, h - ry)] + _arc(w / 2, h - ry, w / 2, ry, 180, 0)[1:] + [(w, ry)] + _arc(w / 2, ry, w / 2, ry, 0, -180)[1:-1]
    top = _arc(w / 2, ry, w / 2, ry, 0, 360)[:-1]
    return Geometry([(body, True, "norm"), (top, True, "lighten")], _tlbr(w, h), (0, 2 * ry, w, h - ry))


def _star(w, h, points, inner):
    out = []
    for i in range(points * 2):
        a = math.radians(-90 + 180 / points * i)
        f = 1 if i % 2 == 0 else inner
        out.append((w / 2 + w / 2 * f * math.cos(a), h / 2 + h / 2 * f * math.sin(a)))
    tips = out[0::2]
    sites = [(x, y, (math.degrees(math.atan2(y - h / 2, x - w / 2)) + 360) % 360) for x, y in tips]
    return Geometry([(out, True, "norm")], sites,
                    (w / 2 - w / 2 * inner * 0.8, h / 2 - h / 2 * inner * 0.8, w / 2 + w / 2 * inner * 0.8,
                     h / 2 + h / 2 * inner * 0.8))


def _document_outline(x0, y0, w, h):
    y1 = y0 + h * 17322 / 21600
    return [(x0, y0), (x0 + w, y0), (x0 + w, y1)] + _cubic(
        (x0 + w, y1), (x0 + w / 2, y1), (x0 + w / 2, y0 + h * 23922 / 21600), (x0, y0 + h * 20172 / 21600))


def _callout(w, h, tip, outline):
    """A box outline (a function of w, h giving its points, the first one top-left going clockwise) with a
    wedge pointing at tip (fractions of the box)."""
    tx, ty = tip[0] * w, tip[1] * h
    body = outline(w, h)
    if 0 <= tx <= w and 0 <= ty <= h:
        return body
    # the side the wedge leaves from: the one facing the tip
    dx, dy = (tx - w / 2) / max(w, 1e-9), (ty - h / 2) / max(h, 1e-9)
    if abs(dy) >= abs(dx):
        side = "b" if dy > 0 else "t"
        base = min(max(w / 2 + (tx - w / 2) / 2, w / 6), w * 5 / 6)
        half = w / 12
        a, b = ((base - half, h), (base + half, h)) if side == "b" else ((base - half, 0), (base + half, 0))
    else:
        side = "r" if dx > 0 else "l"
        base = min(max(h / 2 + (ty - h / 2) / 2, h / 6), h * 5 / 6)
        half = h / 12
        a, b = ((w, base - half), (w, base + half)) if side == "r" else ((0, base - half), (0, base + half))
    # insert the wedge where the outline crosses that side: going clockwise the first base point met is
    # the one further along the clockwise direction of the side
    order = {"t": (a, b), "r": (a, b), "b": (b, a), "l": (b, a)}[side]

    def along(p):  # position of a point along the clockwise outline
        x, y = p
        if side == "t":
            return abs(y) < 1e-6, x
        if side == "r":
            return abs(x - w) < 1e-6, y
        if side == "b":
            return abs(y - h) < 1e-6, -x
        return abs(x) < 1e-6, -y

    first, second = order
    key_first = along(first)[1]
    out, done = [], False
    n = len(body)
    for i in range(n):
        p, q = body[i], body[(i + 1) % n]
        out.append(p)
        on_p, kp = along(p)
        on_q, kq = along(q)
        if not done and on_p and on_q and kp <= key_first <= kq:
            out += [first, (tx, ty), second]
            done = True
    if not done:  # (a rounded corner took the spot) put it after the nearest outline point
        i = min(range(n), key=lambda k: math.dist(body[k], first))
        out = body[:i + 1] + [first, (tx, ty), second] + body[i + 1:]
    return out


def _ellipse_callout(w, h, tip):
    tx, ty = tip[0] * w, tip[1] * h
    a = math.degrees(math.atan2((ty - h / 2) / max(h, 1e-9), (tx - w / 2) / max(w, 1e-9)))
    if ((tx - w / 2) / (w / 2 or 1)) ** 2 + ((ty - h / 2) / (h / 2 or 1)) ** 2 <= 1:
        return _arc(w / 2, h / 2, w / 2, h / 2, 0, 360)[:-1]
    return _arc(w / 2, h / 2, w / 2, h / 2, a + 12, a + 348) + [(tx, ty)]


# logic gates: inputs on the left, output on the right; the body sits between short leads
def _gate(kind, w, h):
    bubble = kind in ("nand", "nor", "xnor", "not")
    r = min(0.05 * w, 0.09 * h) if bubble else 0.0
    x0, xb = 0.18 * w, 0.82 * w - 2 * r          # body from x0 to xb, then the bubble
    top, bot = 0.06 * h, 0.94 * h
    paths, ins = [], (0.3 * h, 0.7 * h)
    if kind in ("and", "nand"):
        xm = x0 + (xb - x0) * 0.42
        body = [(x0, top), (xm, top)] + _arc(xm, h / 2, xb - xm, (bot - top) / 2, 270, 450)[1:] + [(xm, bot), (x0, bot)]
        paths.append((body, True, "norm"))
        ends = [x0, x0]
    elif kind in ("or", "nor", "xor", "xnor"):
        shift = 0.07 * w if kind in ("xor", "xnor") else 0.0
        bx = x0 + shift
        bulge = 0.16 * (xb - x0)
        back = _quad((bx, top), (bx + 2 * bulge, h / 2), (bx, bot), 12)
        front_top = _cubic((bx, top), (bx + (xb - bx) * 0.55, top), (bx + (xb - bx) * 0.85, h * 0.25), (xb, h / 2), 14)
        front_bot = _cubic((xb, h / 2), (bx + (xb - bx) * 0.85, h * 0.75), (bx + (xb - bx) * 0.55, bot), (bx, bot), 14)
        body = [(bx, top)] + front_top + front_bot[:-1] + list(reversed(back))
        paths.append((body, True, "norm"))

        def back_x(y, x_start):  # x of a back curve (quadratic with y = t) at height y
            t = (y - top) / (bot - top)
            return x_start + 2 * t * (1 - t) * 2 * bulge
        if shift:
            extra = [(x0, top)] + _quad((x0, top), (x0 + 2 * bulge, h / 2), (x0, bot), 12)
            paths.append((extra, False, "none"))
            ends = [back_x(y, x0) for y in ins]
        else:
            ends = [back_x(y, bx) for y in ins]
    else:  # not, buffer: a triangle
        tri = [(x0, 0.1 * h), (xb, h / 2), (x0, 0.9 * h)]
        paths.append((tri, True, "norm"))
        ins, ends = (h / 2,), [x0]
    if bubble:
        paths.append((_arc(xb + r, h / 2, r, r, 0, 360)[:-1], True, "norm"))
    out_x = xb + 2 * r
    for y, end in zip(ins, ends):
        paths.append(([(0, y), (end, y)], False, "none"))
    paths.append(([(out_x, h / 2), (w, h / 2)], False, "none"))
    sites = [(0, y, 180) for y in ins] + [(w, h / 2, 0)]
    return Geometry(paths, sites, (x0, top, xb, bot))


def _flow_document(w, h):
    pts = _document_outline(0, 0, w, h)
    bottom = _cubic((w, h * 17322 / 21600), (w / 2, h * 17322 / 21600), (w / 2, h * 23922 / 21600),
                    (0, h * 20172 / 21600), 2)[0]
    return _poly(pts, w, h, (0, 0, w, h * 17322 / 21600), _tlbr(w, h, bottom=(w / 2, bottom[1])))


def _flow_multidocument(w, h):
    dw, dh = w * 0.86, h * 0.83
    back = _document_outline(w * 0.14, 0, dw, dh)
    mid = _document_outline(w * 0.07, h * 0.085, dw, dh)
    front = _document_outline(0, h * 0.17, dw, dh)
    paths = [(back, True, "norm"), (mid, True, "norm"), (front, True, "norm")]
    return Geometry(paths, _tlbr(w, h, top=(w * 0.57, 0), left=(0, h * 0.58), bottom=(dw / 2, h * 0.97),
                                 right=(w, h * 0.4)), (0, h * 0.17, dw, h * 0.17 + dh * 17322 / 21600))


def _builders():
    def rect(w, h):
        return _poly([(0, 0), (w, 0), (w, h), (0, h)], w, h)

    def round_rect(w, h, f=0.16667):
        r = min(w, h) * f
        return _poly(_round_rect_points(w, h, r), w, h, (r * 0.29, r * 0.29, w - r * 0.29, h - r * 0.29))

    def diamond(w, h):
        return _poly([(w / 2, 0), (w, h / 2), (w / 2, h), (0, h / 2)], w, h, (w / 4, h / 4, w * 3 / 4, h * 3 / 4))

    def parallelogram(w, h, o=None):
        o = min(w, h) * 0.25 if o is None else o
        return _poly([(o, 0), (w, 0), (w - o, h), (0, h)], w, h, (o, 0, w - o, h),
                     _tlbr(w, h, left=(o / 2, h / 2), right=(w - o / 2, h / 2)))

    def trapezoid(w, h):
        o = min(w, h) * 0.25
        return _poly([(o, 0), (w - o, 0), (w, h), (0, h)], w, h, (o * 2 / 3, h / 3, w - o * 2 / 3, h),
                     _tlbr(w, h, left=(o / 2, h / 2), right=(w - o / 2, h / 2)))

    def hexagon(w, h, o=None):
        o = min(w, h) * 0.25 if o is None else o
        return _poly([(0, h / 2), (o, 0), (w - o, 0), (w, h / 2), (w - o, h), (o, h)], w, h, (o, 0, w - o, h))

    def octagon(w, h):
        o = min(w, h) * 0.29289
        return _poly([(o, 0), (w - o, 0), (w, o), (w, h - o), (w - o, h), (o, h), (0, h - o), (0, o)], w, h,
                     (o / 2, o / 2, w - o / 2, h - o / 2))

    def plus(w, h):
        o = min(w, h) * 0.25
        return _poly([(o, 0), (w - o, 0), (w - o, o), (w, o), (w, h - o), (w - o, h - o), (w - o, h), (o, h), (o, h - o),
                      (0, h - o), (0, o), (o, o)], w, h, (0, o, w, h - o))

    def triangle(w, h):
        return _poly([(w / 2, 0), (w, h), (0, h)], w, h, (w / 4, h / 2, w * 3 / 4, h),
                     [(w / 2, 0, 270), (w / 4, h / 2, 180), (0, h, 90), (w / 2, h, 90), (w, h, 90), (w * 3 / 4, h / 2, 0)])

    def rt_triangle(w, h):
        return _poly([(0, 0), (w, h), (0, h)], w, h, (w / 12, h * 7 / 12, w * 7 / 12, h * 11 / 12),
                     [(0, 0, 270), (0, h / 2, 180), (0, h, 90), (w / 2, h, 90), (w, h, 90), (w / 2, h / 2, 0)])

    def pentagon(w, h):
        pts = [(w / 2, 0), (w, h * 0.382), (w * 0.809, h), (w * 0.191, h), (0, h * 0.382)]
        return _poly(pts, w, h, (w * 0.191, h * 0.382, w * 0.809, h),
                     [(w / 2, 0, 270), (0, h * 0.382, 180), (w * 0.191, h, 90), (w * 0.809, h, 90), (w, h * 0.382, 0)])

    def arrow_right(w, h):
        xh = w - min(w, h) * 0.5
        return _poly([(0, h / 4), (xh, h / 4), (xh, 0), (w, h / 2), (xh, h), (xh, h * 3 / 4), (0, h * 3 / 4)], w, h,
                     (0, h / 4, xh + (w - xh) / 2, h * 3 / 4), _tlbr(w, h, top=(xh, 0), bottom=(xh, h)))

    def arrow_left(w, h):
        xh = min(w, h) * 0.5
        return _poly([(w, h / 4), (xh, h / 4), (xh, 0), (0, h / 2), (xh, h), (xh, h * 3 / 4), (w, h * 3 / 4)], w, h,
                     (xh / 2, h / 4, w, h * 3 / 4), _tlbr(w, h, top=(xh, 0), bottom=(xh, h)))

    def arrow_up(w, h):
        yh = min(w, h) * 0.5
        return _poly([(w / 2, 0), (w, yh), (w * 3 / 4, yh), (w * 3 / 4, h), (w / 4, h), (w / 4, yh), (0, yh)], w, h,
                     (w / 4, yh / 2, w * 3 / 4, h), _tlbr(w, h, left=(0, yh), right=(w, yh)))

    def arrow_down(w, h):
        yh = h - min(w, h) * 0.5
        return _poly([(w / 4, 0), (w * 3 / 4, 0), (w * 3 / 4, yh), (w, yh), (w / 2, h), (0, yh), (w / 4, yh)], w, h,
                     (w / 4, 0, w * 3 / 4, yh + (h - yh) / 2), _tlbr(w, h, left=(0, yh), right=(w, yh)))

    def arrow_left_right(w, h):
        o = min(w, h) * 0.5
        return _poly([(0, h / 2), (o, 0), (o, h / 4), (w - o, h / 4), (w - o, 0), (w, h / 2), (w - o, h), (w - o, h * 3 / 4),
                      (o, h * 3 / 4), (o, h)], w, h, (o / 2, h / 4, w - o / 2, h * 3 / 4))

    def chevron(w, h):
        o = min(w, h) * 0.5
        return _poly([(0, 0), (w - o, 0), (w, h / 2), (w - o, h), (0, h), (o, h / 2)], w, h, (o, 0, w - o, h),
                     _tlbr(w, h, left=(o, h / 2)))

    def home_plate(w, h):
        o = min(w, h) * 0.5
        return _poly([(0, 0), (w - o, 0), (w, h / 2), (w - o, h), (0, h)], w, h, (0, 0, w - o / 2, h))

    def terminator(w, h):
        rx, ry = w * 3475 / 21600, h / 2
        pts = ([(rx, 0), (w - rx, 0)] + _arc(w - rx, ry, rx, ry, 270, 450)[1:] + [(rx, h)] + _arc(rx, ry, rx, ry, 90, 270)[1:-1])
        return _poly(pts, w, h, (w * 1018 / 21600, h * 3163 / 21600, w * 20582 / 21600, h * 18437 / 21600))

    def predefined(w, h):
        return _poly([(0, 0), (w, 0), (w, h), (0, h)], w, h, (w / 8, 0, w * 7 / 8, h),
                     extra=[([(w / 8, 0), (w / 8, h)], False, "none"), ([(w * 7 / 8, 0), (w * 7 / 8, h)], False, "none")])

    def internal_storage(w, h):
        return _poly([(0, 0), (w, 0), (w, h), (0, h)], w, h, (w / 8, h / 8, w, h),
                     extra=[([(w / 8, 0), (w / 8, h)], False, "none"), ([(0, h / 8), (w, h / 8)], False, "none")])

    def preparation(w, h):
        return hexagon(w, h, w / 5)

    def manual_input(w, h):
        return _poly([(0, h / 5), (w, 0), (w, h), (0, h)], w, h, (0, h / 5, w, h), _tlbr(w, h, top=(w / 2, h / 10)))

    def manual_operation(w, h):
        return _poly([(0, 0), (w, 0), (w * 4 / 5, h), (w / 5, h)], w, h, (w / 5, 0, w * 4 / 5, h),
                     _tlbr(w, h, left=(w / 10, h / 2), right=(w * 9 / 10, h / 2)))

    def offpage(w, h):
        return _poly([(0, 0), (w, 0), (w, h * 4 / 5), (w / 2, h), (0, h * 4 / 5)], w, h, (0, 0, w, h * 4 / 5),
                     _tlbr(w, h, left=(0, h * 2 / 5), right=(w, h * 2 / 5)))

    def delay(w, h):
        pts = [(0, 0), (w / 2, 0)] + _arc(w / 2, h / 2, w / 2, h / 2, 270, 450)[1:] + [(0, h)]
        return _poly(pts, w, h, (0, h * 0.1464, w * 0.8536, h * 0.8536))

    def display(w, h):
        pts = [(0, h / 2), (w / 6, 0), (w * 5 / 6, 0)] + _arc(w * 5 / 6, h / 2, w / 6, h / 2, 270, 450)[1:] + [(w / 6, h)]
        return _poly(pts, w, h, (w / 6, 0, w * 5 / 6, h))

    def merge(w, h):
        return _poly([(0, 0), (w, 0), (w / 2, h)], w, h, (w / 4, 0, w * 3 / 4, h / 2),
                     _tlbr(w, h, left=(w / 4, h / 2), right=(w * 3 / 4, h / 2)))

    def extract(w, h):
        return _poly([(w / 2, 0), (w, h), (0, h)], w, h, (w / 4, h / 2, w * 3 / 4, h),
                     _tlbr(w, h, left=(w / 4, h / 2), right=(w * 3 / 4, h / 2)))

    def sort(w, h):
        g = diamond(w, h)
        return Geometry(g.paths + [([(0, h / 2), (w, h / 2)], False, "none")], g.sites, g.text)

    def punched_card(w, h):
        return _poly([(w / 5, 0), (w, 0), (w, h), (0, h), (0, h / 5)], w, h, (0, h / 5, w, h))

    def circle_mark(cross):
        def build(w, h):
            g = _ellipse(w, h)
            c = math.cos(math.radians(45))
            if cross:
                lines = [([(w / 2 - w / 2 * c, h / 2 - h / 2 * c), (w / 2 + w / 2 * c, h / 2 + h / 2 * c)], False, "none"),
                         ([(w / 2 + w / 2 * c, h / 2 - h / 2 * c), (w / 2 - w / 2 * c, h / 2 + h / 2 * c)], False, "none")]
            else:
                lines = [([(w / 2, 0), (w / 2, h)], False, "none"), ([(0, h / 2), (w, h / 2)], False, "none")]
            return Geometry(g.paths + lines, g.sites, g.text)
        return build

    def callout(outline):
        def build(w, h, tip=CALLOUT_TIP):
            return _poly(_callout(w, h, tip, outline), w, h, (0, 0, w, h))
        return build

    def ellipse_callout(w, h, tip=CALLOUT_TIP):
        g = _ellipse(w, h)
        return Geometry([(_ellipse_callout(w, h, tip), True, "norm")], g.sites, g.text)

    def gate(kind):
        return lambda w, h: _gate(kind, w, h)

    rect_outline = lambda w, h: [(0, 0), (w, 0), (w, h), (0, h)]  # noqa: E731
    round_outline = lambda w, h: _round_rect_points(w, h, min(w, h) * 0.16667)  # noqa: E731
    return {
        "rect": rect, "roundRect": round_rect, "ellipse": _ellipse, "triangle": triangle, "rtTriangle": rt_triangle,
        "parallelogram": parallelogram, "trapezoid": trapezoid, "diamond": diamond, "pentagon": pentagon,
        "hexagon": hexagon, "octagon": octagon, "plus": plus, "can": lambda w, h: _cylinder(w, h, min(w, h) * 0.125),
        "star5": lambda w, h: _star(w, h, 5, 0.382), "star4": lambda w, h: _star(w, h, 4, 0.25),
        "star6": lambda w, h: _star(w, h, 6, 0.577),
        "rightArrow": arrow_right, "leftArrow": arrow_left, "upArrow": arrow_up, "downArrow": arrow_down,
        "leftRightArrow": arrow_left_right, "chevron": chevron, "homePlate": home_plate,
        "flowChartProcess": rect, "flowChartAlternateProcess": lambda w, h: round_rect(w, h, 1 / 6),
        "flowChartDecision": diamond, "flowChartInputOutput": lambda w, h: parallelogram(w, h, w / 5),
        "flowChartPredefinedProcess": predefined, "flowChartInternalStorage": internal_storage,
        "flowChartDocument": _flow_document, "flowChartMultidocument": _flow_multidocument,
        "flowChartTerminator": terminator, "flowChartPreparation": preparation,
        "flowChartManualInput": manual_input, "flowChartManualOperation": manual_operation,
        "flowChartConnector": _ellipse, "flowChartOffpageConnector": offpage, "flowChartPunchedCard": punched_card,
        "flowChartSummingJunction": circle_mark(True), "flowChartOr": circle_mark(False),
        "flowChartSort": sort, "flowChartExtract": extract, "flowChartMerge": merge, "flowChartDelay": delay,
        "flowChartMagneticDisk": lambda w, h: _cylinder(w, h, h / 6), "flowChartDisplay": display,
        "wedgeRectCallout": callout(rect_outline), "wedgeRoundRectCallout": callout(round_outline),
        "wedgeEllipseCallout": ellipse_callout,
        "textBox": rect,
        "gateAnd": gate("and"), "gateOr": gate("or"), "gateNot": gate("not"), "gateNand": gate("nand"),
        "gateNor": gate("nor"), "gateXor": gate("xor"), "gateXnor": gate("xnor"), "gateBuffer": gate("buffer"),
    }


_BUILD = _builders()

# kind -> (Vietnamese name, gallery group, Excel preset name or None, default size)
KINDS = {}
GROUPS = ("Đường", "Hình chữ nhật", "Hình cơ bản", "Mũi tên khối", "Lưu đồ", "Sao", "Chú thích", "Cổng logic")


def _kind(group, prst, size, *items):
    for key, name in items:
        KINDS[key] = (name, group, prst if prst is None else (key if prst is True else prst), size)


_kind("Đường", None, (150, 0), ("line", "Đường thẳng"), ("elbow", "Đường gấp khúc"), ("curve", "Đường cong"))
_kind("Hình chữ nhật", True, (120, 72), ("rect", "Hình chữ nhật"), ("roundRect", "Hình chữ nhật góc tròn"))
_kind("Hình cơ bản", True, (96, 96), ("textBox", "Hộp văn bản"), ("ellipse", "Hình bầu dục"),
      ("triangle", "Tam giác cân"), ("rtTriangle", "Tam giác vuông"), ("parallelogram", "Hình bình hành"),
      ("trapezoid", "Hình thang"), ("diamond", "Hình thoi"), ("pentagon", "Ngũ giác đều"), ("hexagon", "Lục giác"),
      ("octagon", "Bát giác"), ("plus", "Dấu cộng"), ("can", "Hình trụ"))
_kind("Mũi tên khối", True, (120, 60), ("rightArrow", "Mũi tên phải"), ("leftArrow", "Mũi tên trái"),
      ("leftRightArrow", "Mũi tên trái-phải"), ("chevron", "Chữ V"), ("homePlate", "Ngũ giác mũi tên"))
_kind("Mũi tên khối", True, (60, 120), ("upArrow", "Mũi tên lên"), ("downArrow", "Mũi tên xuống"))
_kind("Lưu đồ", True, (120, 60), ("flowChartProcess", "Xử lý"), ("flowChartAlternateProcess", "Xử lý thay thế"),
      ("flowChartInputOutput", "Nhập/Xuất dữ liệu"), ("flowChartPredefinedProcess", "Chương trình con"),
      ("flowChartInternalStorage", "Bộ nhớ trong"), ("flowChartDocument", "Tài liệu"),
      ("flowChartMultidocument", "Nhiều tài liệu"), ("flowChartPreparation", "Chuẩn bị"),
      ("flowChartManualInput", "Nhập thủ công"), ("flowChartManualOperation", "Thao tác thủ công"),
      ("flowChartPunchedCard", "Thẻ"), ("flowChartDelay", "Trì hoãn"), ("flowChartDisplay", "Hiển thị"))
_kind("Lưu đồ", True, (120, 50), ("flowChartTerminator", "Bắt đầu/Kết thúc"))
_kind("Lưu đồ", True, (120, 80), ("flowChartDecision", "Quyết định"), ("flowChartSort", "Sắp xếp"))
_kind("Lưu đồ", True, (40, 40), ("flowChartConnector", "Điểm nối"), ("flowChartSummingJunction", "Nút tổng"),
      ("flowChartOr", "Nút Hoặc"))
_kind("Lưu đồ", True, (48, 48), ("flowChartOffpageConnector", "Nối sang trang khác"), ("flowChartExtract", "Trích"),
      ("flowChartMerge", "Gộp"))
_kind("Lưu đồ", True, (80, 80), ("flowChartMagneticDisk", "Cơ sở dữ liệu"))
_kind("Sao", True, (96, 96), ("star4", "Sao 4 cánh"), ("star5", "Sao 5 cánh"), ("star6", "Sao 6 cánh"))
_kind("Chú thích", True, (130, 70), ("wedgeRectCallout", "Chú thích hình chữ nhật"),
      ("wedgeRoundRectCallout", "Chú thích góc tròn"), ("wedgeEllipseCallout", "Chú thích bầu dục"))
_kind("Cổng logic", None, (80, 50), ("gateAnd", "Cổng AND"), ("gateOr", "Cổng OR"), ("gateNand", "Cổng NAND"),
      ("gateNor", "Cổng NOR"), ("gateXor", "Cổng XOR"), ("gateXnor", "Cổng XNOR"))
_kind("Cổng logic", None, (70, 40), ("gateNot", "Cổng NOT"), ("gateBuffer", "Cổng đệm (Buffer)"))
KINDS["textBox"] = ("Hộp văn bản", "Hình cơ bản", "rect", (120, 30))
_PRESET_KIND = {v[2]: k for k, v in KINDS.items() if v[2] and k != "textBox"}

# Excel's look for a new shape (Office theme), and the plainer one of the Diagram tab
EXCEL_STYLE = {"fill": "#4472c4", "line": "#2f528f", "line_w": 1.0, "font": {"size": 11, "color": "#ffffff"}}
EXCEL_LINE = {"line": "#4472c4", "line_w": 1.0}
DIAGRAM_STYLE = {"fill": "#ffffff", "line": "#262626", "line_w": 1.25, "font": {"size": 11, "color": "#000000"}}
DIAGRAM_LINE = {"line": "#262626", "line_w": 1.25, "tail": "triangle"}
TEXT_BOX_STYLE = {"fill": None, "line": None, "line_w": 0.75, "font": {"size": 11, "color": "#000000"},
                  "align": "left", "valign": "top"}

# quick styles (Excel's Shape Styles gallery): (fill, line, text colour) per theme colour
THEME = ["#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47", "#262626"]


def _mix(c1, c2, t):
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def quick_styles():
    """Rows of {"fill", "line", "font": {"color"}} like Excel's Shape Styles: outline only, light fill,
    solid fill, dark fill."""
    rows = []
    rows.append([{"fill": "#ffffff", "line": c, "font": {"color": "#000000"}} for c in THEME])
    rows.append([{"fill": _mix(c, "#ffffff", 0.8), "line": c, "font": {"color": "#000000"}} for c in THEME])
    rows.append([{"fill": c, "line": _mix(c, "#000000", 0.35), "font": {"color": "#ffffff"}} for c in THEME])
    rows.append([{"fill": _mix(c, "#000000", 0.25), "line": None, "font": {"color": "#ffffff"}} for c in THEME])
    return rows


def name_of(kind):
    return KINDS.get(kind, ("Hình",))[0]


def is_line(shape):
    return shape.get("kind") in LINE_KINDS


def default_size(kind):
    return KINDS.get(kind, (None, None, None, (96, 96)))[3]


def new_shape(kind, anchor, size, style=None, **extra):
    """A shape dict of `kind` at anchor/size with a style (EXCEL_STYLE by default)."""
    if style is None:
        style = EXCEL_LINE if kind in LINE_KINDS else TEXT_BOX_STYLE if kind == "textBox" else EXCEL_STYLE
    shape = {"type": SHAPE, "id": None, "kind": kind, "anchor": tuple(anchor), "size": tuple(size), "rot": 0.0,
             "flip_h": False, "flip_v": False, "fill": None, "line": None, "line_w": 1.0, "dash": None, "text": "",
             "font": {"size": 11, "color": "#000000"}, "align": "center", "valign": "middle"}
    for key, value in style.items():
        shape[key] = dict(value) if isinstance(value, dict) else value
    if kind in LINE_KINDS:
        shape.update(fill=None, text="")
    shape.update(extra)
    return shape


# ------------------------------------------------------------------ placing a shape

def local_geometry(shape, w, h):
    """The Geometry of a shape in its own w x h box (no rotation or flips)."""
    kind = shape.get("kind")
    build = _BUILD.get(kind)
    if kind == "custom":
        paths = [([(x * w, y * h) for x, y in pts], closed, fill)
                 for pts, closed, fill in shape.get("paths", [])] or _BUILD["rect"](w, h).paths
        return Geometry(paths, _tlbr(w, h), (0, 0, w, h))
    if build is None:
        build = _BUILD["rect"]
    w, h = max(w, 1e-6), max(h, 1e-6)
    if kind in ("wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout") and shape.get("tip"):
        return build(w, h, tuple(shape["tip"]))
    return build(w, h)


def transform(shape, box):
    """A function mapping a point of the shape's own box to `box` (x0, y0, x1, y1), with flips then
    rotation about the centre; and the function for directions (degrees)."""
    x0, y0, x1, y1 = box
    w, h = x1 - x0, y1 - y0
    cx, cy = x0 + w / 2, y0 + h / 2
    a = math.radians(shape.get("rot") or 0.0)
    cos, sin = math.cos(a), math.sin(a)
    fh, fv = bool(shape.get("flip_h")), bool(shape.get("flip_v"))

    def point(x, y):
        if fh:
            x = w - x
        if fv:
            y = h - y
        dx, dy = x - w / 2, y - h / 2
        return cx + dx * cos - dy * sin, cy + dx * sin + dy * cos

    def angle(deg):
        if fh:
            deg = 180 - deg
        if fv:
            deg = -deg
        return (deg + (shape.get("rot") or 0.0)) % 360
    return point, angle


def outline(shape, box):
    """[(points, closed, fill mode)] of a (non-line) shape placed on box."""
    x0, y0, x1, y1 = box
    geo = local_geometry(shape, x1 - x0, y1 - y0)
    point, _ = transform(shape, box)
    return [([point(x, y) for x, y in pts], closed, fill) for pts, closed, fill in geo.paths]


def sites(shape, box):
    """[(x, y, outward direction in degrees)] where connectors can be glued, placed on box."""
    if is_line(shape):
        return []
    x0, y0, x1, y1 = box
    geo = local_geometry(shape, x1 - x0, y1 - y0)
    point, angle = transform(shape, box)
    return [point(x, y) + (angle(a),) for x, y, a in geo.sites]


def text_area(shape, box):
    """(centre x, centre y, width, height) of the shape's text area (before rotation)."""
    x0, y0, x1, y1 = box
    w, h = x1 - x0, y1 - y0
    tx0, ty0, tx1, ty1 = local_geometry(shape, w, h).text
    if shape.get("flip_h"):
        tx0, tx1 = w - tx1, w - tx0
    if shape.get("flip_v"):
        ty0, ty1 = h - ty1, h - ty0
    point, _ = transform(dict(shape, flip_h=False, flip_v=False), box)
    cx, cy = point((tx0 + tx1) / 2, (ty0 + ty1) / 2)
    return cx, cy, tx1 - tx0, ty1 - ty0


def extent(shape, box):
    """Bounds of everything a shape draws placed on box: a callout's tail and rotation included (lines:
    their box)."""
    if is_line(shape):
        return box
    return bounds([p for pts, _, _ in outline(shape, box) for p in pts] + [rotated_box(shape, box)[:2],
                                                                            rotated_box(shape, box)[2:]])


def bounds(points):
    xs, ys = [p[0] for p in points], [p[1] for p in points]
    return min(xs), min(ys), max(xs), max(ys)


def rotated_box(shape, box):
    """The axis-aligned bounds of the shape's box after rotation."""
    x0, y0, x1, y1 = box
    point, _ = transform(shape, box)
    return bounds([point(0, 0), point(x1 - x0, 0), point(x1 - x0, y1 - y0), point(0, y1 - y0)])


# ------------------------------------------------------------------ lines and connectors

def line_ends(shape, box):
    """(start, end) points of a line placed on box."""
    x0, y0, x1, y1 = box
    sx, ex = (x1, x0) if shape.get("flip_h") else (x0, x1)
    sy, ey = (y1, y0) if shape.get("flip_v") else (y0, y1)
    return (sx, sy), (ex, ey)


def box_for_ends(p0, p1):
    """(box, flip_h, flip_v) of a line from p0 to p1."""
    return ((min(p0[0], p1[0]), min(p0[1], p1[1]), max(p0[0], p1[0]), max(p0[1], p1[1])),
            p1[0] < p0[0], p1[1] < p0[1])


_DIRS = {0: (1, 0), 90: (0, 1), 180: (-1, 0), 270: (0, -1)}


def _axis_dir(angle):
    """The axis direction (unit vector) nearest to an angle."""
    return _DIRS[round((angle % 360) / 90) * 90 % 360]


def _crosses(a, b, box, margin=0.5):
    """Whether segment a-b (horizontal or vertical) runs through the inside of box."""
    if box is None:
        return False
    x0, y0, x1, y1 = box[0] + margin, box[1] + margin, box[2] - margin, box[3] - margin
    if x0 >= x1 or y0 >= y1:
        return False
    if abs(a[1] - b[1]) < 1e-9:  # horizontal
        lo, hi = sorted((a[0], b[0]))
        return y0 < a[1] < y1 and lo < x1 and hi > x0
    lo, hi = sorted((a[1], b[1]))
    return x0 < a[0] < x1 and lo < y1 and hi > y0


def simplify(points):
    """Drop repeated points and middle points of straight runs."""
    out = []
    for p in points:
        if out and abs(p[0] - out[-1][0]) < 1e-6 and abs(p[1] - out[-1][1]) < 1e-6:
            continue
        if len(out) >= 2:
            a, b = out[-2], out[-1]
            if (abs(a[0] - b[0]) < 1e-6 and abs(b[0] - p[0]) < 1e-6) or (abs(a[1] - b[1]) < 1e-6 and abs(b[1] - p[1]) < 1e-6):
                out[-1] = p
                continue
        out.append(p)
    return out


def elbow_route(p0, a0, p1, a1, box0=None, box1=None, stub=16.0):
    """An orthogonal route from p0 to p1 for an elbow connector. a0 / a1: the outward directions (degrees)
    of the connection points it is glued to (None: free end); box0 / box1: the bounds of those shapes,
    which the route goes around."""
    dx, dy = p1[0] - p0[0], p1[1] - p0[1]
    if a0 is None:
        a0 = (0 if dx >= 0 else 180) if abs(dx) >= abs(dy) else (90 if dy >= 0 else 270)
    if a1 is None:
        a1 = (180 if dx >= 0 else 0) if abs(dx) >= abs(dy) else (270 if dy >= 0 else 90)
    d0, d1 = _axis_dir(a0), _axis_dir(a1)
    free = min(stub, max(abs(dx), abs(dy)) / 3)  # a free end only needs to leave the right way
    s0 = stub if box0 is not None else free
    s1 = stub if box1 is not None else free
    A = (p0[0] + d0[0] * s0, p0[1] + d0[1] * s0)
    B = (p1[0] + d1[0] * s1, p1[1] + d1[1] * s1)
    boxes = [b for b in (box0, box1) if b is not None]
    candidates = [[(B[0], A[1])], [(A[0], B[1])]]
    mx, my = (A[0] + B[0]) / 2, (A[1] + B[1]) / 2
    candidates += [[(mx, A[1]), (mx, B[1])], [(A[0], my), (B[0], my)]]
    if boxes:
        right = max(b[2] for b in boxes) + stub
        left = min(b[0] for b in boxes) - stub
        bottom = max(b[3] for b in boxes) + stub
        top = min(b[1] for b in boxes) - stub
        for x in (right, left):
            candidates.append([(x, A[1]), (x, B[1])])
        for y in (bottom, top):
            candidates.append([(A[0], y), (B[0], y)])
        # leaving sideways, going round below/above the shapes
        for x in (right, left):
            for y in (bottom, top):
                candidates.append([(x, A[1]), (x, y), (B[0], y)])
                candidates.append([(A[0], y), (x, y), (x, B[1])])
    best, best_cost = None, None
    for corners in candidates:
        path = simplify([p0, A] + corners + [B, p1])
        ok = all(abs(a[0] - b[0]) < 1e-6 or abs(a[1] - b[1]) < 1e-6 for a, b in zip(path, path[1:]))
        if not ok:
            continue
        cost = sum(abs(a[0] - b[0]) + abs(a[1] - b[1]) for a, b in zip(path, path[1:]))
        cost += (len(path) - 2) * stub * 1.5
        segs = list(zip(path, path[1:]))
        for i, (a, b) in enumerate(segs):
            for k, bx in enumerate((box0, box1)):
                # the first / last segment touches its own shape at the glue point
                if bx is not None and _crosses(a, b, bx) and not (k == 0 and i == 0) and not (k == 1 and i == len(segs) - 1):
                    cost += 10000
        # leaving against the glue direction (back into the shape) is bad
        if len(path) >= 2:
            first = (path[1][0] - path[0][0], path[1][1] - path[0][1])
            if box0 is not None and first[0] * d0[0] + first[1] * d0[1] < -1e-6:
                cost += 5000
            last = (path[-2][0] - path[-1][0], path[-2][1] - path[-1][1])
            if box1 is not None and last[0] * d1[0] + last[1] * d1[1] < -1e-6:
                cost += 5000
        if best_cost is None or cost < best_cost - 1e-6:
            best, best_cost = path, cost
    return best or [p0, p1]


def curve_route(p0, a0, p1, a1):
    """A smooth curve from p0 to p1 leaving and arriving along the glue directions."""
    dist = math.dist(p0, p1)
    k = max(20.0, dist / 2.5)
    if a0 is None:
        a0 = math.degrees(math.atan2(p1[1] - p0[1], p1[0] - p0[0])) if dist else 0
        k0 = dist / 3
    else:
        k0 = k
    if a1 is None:
        a1 = math.degrees(math.atan2(p0[1] - p1[1], p0[0] - p1[0])) if dist else 180
        k1 = dist / 3
    else:
        k1 = k
    c1 = (p0[0] + math.cos(math.radians(a0)) * k0, p0[1] + math.sin(math.radians(a0)) * k0)
    c2 = (p1[0] + math.cos(math.radians(a1)) * k1, p1[1] + math.sin(math.radians(a1)) * k1)
    return [p0] + _cubic(p0, c1, c2, p1, 24)


def route(shape, p0, p1, a0=None, a1=None, box0=None, box1=None, px=1.0):
    """Points of a line or connector from p0 to p1."""
    kind = shape.get("kind")
    if kind == "elbow":
        return elbow_route(p0, a0, p1, a1, box0, box1, 16 * px)
    if kind == "curve":
        return curve_route(p0, a0, p1, a1)
    return [p0, p1]


def arrowhead(kind, tip, back, width):
    """(polygon points, filled, how far the line should stop short of the tip) of an arrowhead at `tip`
    for a line coming from `back`; width: the line's width."""
    length = max(7.0, width * 3.2)
    half = length * 0.5
    ang = math.atan2(tip[1] - back[1], tip[0] - back[0])
    ux, uy = math.cos(ang), math.sin(ang)
    nx, ny = -uy, ux

    def at(along, side):
        return tip[0] - ux * along + nx * side, tip[1] - uy * along + ny * side
    if kind == "arrow":
        return [at(length, half), tip, at(length, -half)], False, 0.0
    if kind == "stealth":
        return [tip, at(length, half), at(length * 0.65, 0), at(length, -half)], True, length * 0.6
    if kind == "diamond":
        return [tip, at(length / 2, half * 0.8), at(length, 0), at(length / 2, -half * 0.8)], True, length / 2
    if kind == "oval":
        r = length * 0.38
        c = at(r, 0)
        return [(c[0] + r * math.cos(t * math.pi / 8), c[1] + r * math.sin(t * math.pi / 8)) for t in range(16)], True, r
    return [tip, at(length, half), at(length, -half)], True, length * 0.85


def _cut(points, at_end, dist):
    """The polyline with `dist` taken off its start (at_end False) or end."""
    pts = list(points)
    if dist <= 0 or len(pts) < 2:
        return pts
    if at_end:
        pts.reverse()
    while len(pts) >= 2 and dist > 0:
        a, b = pts[0], pts[1]
        seg = math.dist(a, b)
        if seg > dist:
            t = dist / seg
            pts[0] = (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)
            break
        dist -= seg
        pts.pop(0)
    if at_end:
        pts.reverse()
    return pts


def with_arrows(points, head, tail, width):
    """(the polyline shortened under its arrowheads, [(polygon, filled)])."""
    heads = []
    pts = list(points)
    if len(pts) < 2:
        return pts, heads
    if tail in ARROWS:
        back = next((p for p in reversed(pts[:-1]) if math.dist(p, pts[-1]) > 1e-6), pts[0])
        poly, filled, short = arrowhead(tail, pts[-1], back, width)
        heads.append((poly, filled))
        pts = _cut(pts, True, short)
    if head in ARROWS:
        back = next((p for p in pts[1:] if math.dist(p, pts[0]) > 1e-6), pts[-1])
        poly, filled, short = arrowhead(head, pts[0], back, width)
        heads.append((poly, filled))
        pts = _cut(pts, False, short)
    return pts, heads


# ------------------------------------------------------------------ drawing

def line_width(shape, px):
    """Width of the shape's line in target units (its "line_w" is in points)."""
    return max(0.0, (shape.get("line_w") or 0.75) * 96 / 72 * px)


def fill_color(shape, mode):
    fill = shape.get("fill")
    if not fill or mode == "none":
        return None
    if mode == "lighten":
        return _mix(fill, "#ffffff", 0.35)
    if mode == "darken":
        return _mix(fill, "#000000", 0.2)
    return fill


def draw(target, shape, box, px, font=None, path=None):
    """Draw a shape with a Tk-canvas-like target (create_polygon, create_line, create_text). box is the
    shape's box in target units; px target units per 96-DPI pixel; font(spec) gives the target's font for
    the shape's text; path: the points of a line or connector when already routed (glued ends)."""
    color = shape.get("line")
    width = line_width(shape, px)
    dash = DASHES.get(shape.get("dash"))
    dash = tuple(max(1, round(d * max(1.0, width) / 1.5)) for d in dash) if dash else None
    if is_line(shape):
        pts = path or line_ends(shape, box)
        pts, heads = with_arrows(pts, shape.get("head"), shape.get("tail"), max(1.0, width))
        if color and len(pts) >= 2:
            target.create_line(*[v for p in pts for v in p], fill=color, width=max(1.0, width), dash=dash,
                               capstyle="round" if dash is None else "butt", joinstyle="round")
        for poly, filled in heads:
            if color:
                if filled:
                    target.create_polygon(*[v for p in poly for v in p], fill=color, outline=color, width=1)
                else:
                    target.create_line(*[v for p in poly for v in p], fill=color, width=max(1.0, width),
                                       capstyle="round", joinstyle="round")
        return
    for pts, closed, mode in outline(shape, box):
        flat = [v for p in pts for v in p]
        fill = fill_color(shape, mode)
        if closed and fill:
            target.create_polygon(*flat, fill=fill, outline=color or "", width=width if color else 0, dash=dash,
                                  joinstyle="miter")
        elif color and width > 0:
            if closed:
                flat += flat[:2]
            target.create_line(*flat, fill=color, width=width, dash=dash, joinstyle="miter")
    text = shape.get("text")
    if text and font is not None:
        draw_text(target, shape, box, px, font)


def draw_text(target, shape, box, px, font):
    cx, cy, w, h = text_area(shape, box)
    inset_x, inset_y = 9.6 * px, 4.8 * px
    align, valign = shape.get("align") or "center", shape.get("valign") or "middle"
    lx = {"left": -w / 2 + inset_x, "right": w / 2 - inset_x}.get(align, 0.0)
    ly = {"top": -h / 2 + inset_y, "bottom": h / 2 - inset_y}.get(valign, 0.0)
    a = math.radians(shape.get("rot") or 0.0)
    x, y = cx + lx * math.cos(a) - ly * math.sin(a), cy + lx * math.sin(a) + ly * math.cos(a)
    anchor = {"top": "n", "bottom": "s"}.get(valign, "") + {"left": "w", "right": "e"}.get(align, "")
    spec = dict(shape.get("font") or {})
    target.create_text(x, y, text=shape["text"], fill=spec.get("color") or "#000000", font=font(spec),
                       width=max(1.0, w - 2 * inset_x), justify=align, anchor=anchor or "center",
                       angle=-(shape.get("rot") or 0.0))


# ------------------------------------------------------------------ hit testing

def _near_segment(x, y, a, b, tol):
    ax, ay = a
    bx, by = b
    dx, dy = bx - ax, by - ay
    seg = dx * dx + dy * dy
    t = 0.0 if seg == 0 else max(0.0, min(1.0, ((x - ax) * dx + (y - ay) * dy) / seg))
    return math.hypot(x - (ax + t * dx), y - (ay + t * dy)) <= tol


def inside(x, y, pts):
    """Point in polygon (even-odd)."""
    hit = False
    n = len(pts)
    for i in range(n):
        x0, y0 = pts[i]
        x1, y1 = pts[(i + 1) % n]
        if (y0 > y) != (y1 > y) and x < (x1 - x0) * (y - y0) / (y1 - y0) + x0:
            hit = not hit
    return hit


def hit(shape, box, x, y, tol, path=None):
    """Whether a click at (x, y) falls on the shape: inside a filled (or text-holding) shape, or within
    tol of its lines."""
    if is_line(shape):
        pts = path or line_ends(shape, box)
        return any(_near_segment(x, y, a, b, tol) for a, b in zip(pts, pts[1:]))
    solid = bool(shape.get("fill")) or bool(shape.get("text")) or shape.get("kind") == "textBox"
    for pts, closed, _ in outline(shape, box):
        if closed and solid and inside(x, y, pts):
            return True
        ring = pts + [pts[0]] if closed else pts
        if any(_near_segment(x, y, a, b, tol) for a, b in zip(ring, ring[1:])):
            return True
    if solid:  # the text area of a shape without a fill
        bx0, by0, bx1, by1 = rotated_box(shape, box)
        return shape.get("kind") == "textBox" and bx0 <= x <= bx1 and by0 <= y <= by1
    return False


def handles(shape, box, px):
    """{name: (x, y)} of a selected shape's handles: the eight sizing handles ("l", "rt"...) and "rot",
    or "start"/"end" for a line; "tip" for a callout."""
    if is_line(shape):
        p0, p1 = line_ends(shape, box)
        return {"start": p0, "end": p1}
    x0, y0, x1, y1 = box
    w, h = x1 - x0, y1 - y0
    point, _ = transform(dict(shape, flip_h=False, flip_v=False), box)
    out = {"lt": point(0, 0), "t": point(w / 2, 0), "rt": point(w, 0), "r": point(w, h / 2), "rb": point(w, h),
           "b": point(w / 2, h), "lb": point(0, h), "l": point(0, h / 2), "rot": point(w / 2, -22 * px)}
    if shape.get("kind") in ("wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout"):
        tip = shape.get("tip") or CALLOUT_TIP
        tpoint, _ = transform(shape, box)
        out["tip"] = tpoint(tip[0] * w, tip[1] * h)
    return out


def resize(shape, box, handle, x, y, keep_ratio=False, least=4.0):
    """The new box (x0, y0, x1, y1) after dragging sizing handle `handle` of a (possibly rotated) shape to
    (x, y): the opposite side or corner stays where it is on the screen."""
    x0, y0, x1, y1 = box
    w, h = x1 - x0, y1 - y0
    cx, cy = x0 + w / 2, y0 + h / 2
    a = math.radians(shape.get("rot") or 0.0)
    cos, sin = math.cos(a), math.sin(a)
    # the pointer in the shape's own (unrotated) frame, origin at the box's top-left
    rx, ry = x - cx, y - cy
    lx, ly = rx * cos + ry * sin + w / 2, -rx * sin + ry * cos + h / 2
    nx0, ny0, nx1, ny1 = 0.0, 0.0, w, h
    if "l" in handle:
        nx0 = min(lx, w - least)
    if "r" in handle:
        nx1 = max(lx, least)
    if "t" in handle:
        ny0 = min(ly, h - least)
    if "b" in handle:
        ny1 = max(ly, least)
    if keep_ratio and len(handle) == 2 and w > 0 and h > 0:
        nw, nh = nx1 - nx0, ny1 - ny0
        f = max(nw / w, nh / h)
        nw, nh = w * f, h * f
        nx0, nx1 = (nx1 - nw, nx1) if "l" in handle else (nx0, nx0 + nw)
        ny0, ny1 = (ny1 - nh, ny1) if "t" in handle else (ny0, ny0 + nh)
    # the new centre, back on the screen
    mx, my = (nx0 + nx1) / 2 - w / 2, (ny0 + ny1) / 2 - h / 2
    ncx, ncy = cx + mx * cos - my * sin, cy + mx * sin + my * cos
    nw, nh = nx1 - nx0, ny1 - ny0
    return ncx - nw / 2, ncy - nh / 2, ncx + nw / 2, ncy + nh / 2


def tip_from_point(shape, box, x, y):
    """A callout's "tip" (fractions of its box) for the point (x, y)."""
    x0, y0, x1, y1 = box
    w, h = max(x1 - x0, 1e-6), max(y1 - y0, 1e-6)
    cx, cy = x0 + w / 2, y0 + h / 2
    a = math.radians(shape.get("rot") or 0.0)
    rx, ry = x - cx, y - cy
    lx, ly = rx * math.cos(a) + ry * math.sin(a) + w / 2, -rx * math.sin(a) + ry * math.cos(a) + h / 2
    if shape.get("flip_h"):
        lx = w - lx
    if shape.get("flip_v"):
        ly = h - ly
    return round(lx / w, 4), round(ly / h, 4)
