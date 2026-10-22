"""Shapes in .xlsx files: PyExcel's shape dicts (see shapes.py) as Excel's DrawingML and back.

Excel keeps a sheet's shapes in a drawing part (xl/drawings/drawingN.xml), each in an anchor that ties it
to a cell: <xdr:sp> for shapes (a preset outline such as "flowChartDecision", or a custom one given as
paths) and <xdr:cxnSp> for connectors, whose ends name the shapes and connection points they are glued
to. openpyxl drops all this, so formula_engine writes and reads it with the functions here.
"""

import bisect
import colorsys
import math
import re
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

import shapes

EMU = 9525  # EMUs per pixel at 96 DPI
PT = 12700  # EMUs per point
NS = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
_DECLARE = " ".join(f'xmlns:{k}="{v}"' for k, v in NS.items())
DRAWING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml"
DRAWING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
TAG = "PyExcel:"  # alt text marking shapes Excel has no preset for (logic gates), to read them back
THEME = {"dk1": "#000000", "lt1": "#ffffff", "tx1": "#000000", "bg1": "#ffffff", "dk2": "#44546a", "lt2": "#e7e6e6",
         "tx2": "#44546a", "bg2": "#e7e6e6", "accent1": "#4472c4", "accent2": "#ed7d31", "accent3": "#a5a5a5",
         "accent4": "#ffc000", "accent5": "#5b9bd5", "accent6": "#70ad47", "hlink": "#0563c1",
         "folHlink": "#954f72", "phClr": "#4472c4"}
_DASH_OUT = {"dash": "dash", "dot": "sysDot", "dashdot": "dashDot", "longdash": "lgDash"}
_DASH_IN = {"dash": "dash", "sysDash": "dash", "dot": "dot", "sysDot": "dot", "dashDot": "dashdot",
            "sysDashDot": "dashdot", "lgDash": "longdash", "lgDashDot": "dashdot", "lgDashDotDot": "dashdot",
            "sysDashDotDot": "dashdot"}
_THEME_LINE = {1: 0.5, 2: 1.0, 3: 1.5}  # the Office theme's line widths (lnRef idx)


class SheetGeometry:
    """Pixel positions (96 DPI) of a sheet's columns and rows: hidden ones have no width."""

    def __init__(self, sh, col_width, row_height, max_rows, max_cols):
        self.dw, self.dh, self.max_rows, self.max_cols = col_width, row_height, max_rows, max_cols
        widths = {c: w for c, w in sh.col_widths.items()}
        for c in sh.hidden_cols:
            widths[c] = 0
        heights = {r: h for r, h in sh.row_heights.items()}
        for r in set(sh.hidden_rows) | set(sh.filter["hidden"] if sh.filter else ()):
            heights[r] = 0
        self._cols = self._prefix(widths, col_width)
        self._rows = self._prefix(heights, row_height)

    @staticmethod
    def _prefix(sizes, default):
        keys = sorted(sizes)
        sums, total = [], 0
        for k in keys:
            sums.append(total)
            total += sizes[k] - default
        return keys, sums, total, sizes

    @staticmethod
    def _edge(table, i, default):
        keys, sums, total, sizes = table
        n = bisect.bisect_left(keys, i)
        extra = sums[n] if n < len(keys) else total
        return i * default + extra

    def col_x(self, c):
        return self._edge(self._cols, c, self.dw)

    def row_y(self, r):
        return self._edge(self._rows, r, self.dh)

    def _at(self, pos, edge, limit):
        lo, hi = 0, limit - 1
        while lo < hi:  # the last line whose edge is at or before pos
            mid = (lo + hi + 1) // 2
            if edge(mid) <= pos:
                lo = mid
            else:
                hi = mid - 1
        return lo

    def anchor(self, x, y):
        """(row, col, dx, dy) of the point (x, y)."""
        x, y = max(0.0, x), max(0.0, y)
        c = self._at(x, self.col_x, self.max_cols)
        r = self._at(y, self.row_y, self.max_rows)
        return r, c, round(x - self.col_x(c)), round(y - self.row_y(r))

    def box(self, obj):
        r, c, dx, dy = obj["anchor"]
        x, y = self.col_x(c) + dx, self.row_y(r) + dy
        return x, y, x + obj["size"][0], y + obj["size"][1]


# ------------------------------------------------------------------ writing

def _color(value):
    return f'<a:solidFill><a:srgbClr val="{value.lstrip("#").upper()}"/></a:solidFill>' if value else "<a:noFill/>"


def _line(obj):
    if not obj.get("line"):
        return "<a:ln><a:noFill/></a:ln>"
    out = [f'<a:ln w="{round((obj.get("line_w") or 0.75) * PT)}">{_color(obj["line"])}']
    if obj.get("dash") in _DASH_OUT:
        out.append(f'<a:prstDash val="{_DASH_OUT[obj["dash"]]}"/>')
    if shapes.is_line(obj):
        for end, key in (("headEnd", "head"), ("tailEnd", "tail")):
            if obj.get(key) in shapes.ARROWS:
                out.append(f'<a:{end} type="{obj[key]}"/>')
    out.append("</a:ln>")
    return "".join(out)


def _xfrm(obj, x, y, w, h, rot=True):
    attrs = ""
    if rot and obj.get("rot"):
        attrs += f' rot="{round((obj["rot"] % 360) * 60000)}"'
    if obj.get("flip_h"):
        attrs += ' flipH="1"'
    if obj.get("flip_v"):
        attrs += ' flipV="1"'
    return (f'<a:xfrm{attrs}><a:off x="{round(x * EMU)}" y="{round(y * EMU)}"/>'
            f'<a:ext cx="{round(w * EMU)}" cy="{round(h * EMU)}"/></a:xfrm>')


def _path_xml(paths, w, h, scale=EMU):
    out = []
    for pts, closed, fill in paths:
        if len(pts) < 2:
            continue
        attrs = f' w="{max(1, round(w * scale))}" h="{max(1, round(h * scale))}"'
        if fill == "none" or not closed:
            attrs += ' fill="none"'
        elif fill in ("lighten", "darken"):
            attrs += f' fill="{fill}"'
        seq = [f'<a:moveTo><a:pt x="{round(pts[0][0] * scale)}" y="{round(pts[0][1] * scale)}"/></a:moveTo>']
        seq += [f'<a:lnTo><a:pt x="{round(x * scale)}" y="{round(y * scale)}"/></a:lnTo>' for x, y in pts[1:]]
        if closed:
            seq.append("<a:close/>")
        out.append(f"<a:path{attrs}>{''.join(seq)}</a:path>")
    return "".join(out)


def _custom_geometry(obj, w, h):
    geo = shapes.local_geometry(obj, w, h)
    sites = "".join(f'<a:cxn ang="{round((a % 360) * 60000)}"><a:pos x="{round(x * EMU)}" y="{round(y * EMU)}"/></a:cxn>'
                    for x, y, a in geo.sites)
    l, t, r, b = (round(v * EMU) for v in geo.text)
    return (f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst>{sites}</a:cxnLst>'
            f'<a:rect l="{l}" t="{t}" r="{r}" b="{b}"/><a:pathLst>{_path_xml(geo.paths, w, h)}</a:pathLst></a:custGeom>')


def _geometry(obj, w, h):
    kind = obj.get("kind")
    prst = shapes.KINDS.get(kind, (None, None, None))[2]
    if kind == "rect" and obj.get("xl_prst"):
        prst = obj["xl_prst"]
    if not prst:
        return _custom_geometry(obj, w, h)
    av = ""
    if kind in ("wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout"):
        fx, fy = obj.get("tip") or shapes.CALLOUT_TIP
        av = (f'<a:gd name="adj1" fmla="val {round((fx - 0.5) * 100000)}"/>'
              f'<a:gd name="adj2" fmla="val {round((fy - 0.5) * 100000)}"/>')
    return f'<a:prstGeom prst="{prst}"><a:avLst>{av}</a:avLst></a:prstGeom>'


def _text(obj):
    font = obj.get("font") or {}
    attrs = f'lang="vi-VN" sz="{round((font.get("size") or 11) * 100)}"'
    if font.get("bold"):
        attrs += ' b="1"'
    if font.get("italic"):
        attrs += ' i="1"'
    if font.get("underline"):
        attrs += ' u="sng"'
    props = _color(font.get("color") or "#000000")
    if font.get("font"):
        props += f'<a:latin typeface={quoteattr(font["font"])}/><a:ea typeface={quoteattr(font["font"])}/>'
    algn = {"left": "l", "right": "r"}.get(obj.get("align"), "ctr")
    anchor = {"top": "t", "bottom": "b"}.get(obj.get("valign"), "ctr")
    paras = []
    for line in (obj.get("text") or "").split("\n"):
        run = f'<a:r><a:rPr {attrs}>{props}</a:rPr><a:t>{escape(line)}</a:t></a:r>' if line else ""
        paras.append(f'<a:p><a:pPr algn="{algn}"/>{run}<a:endParaRPr {attrs}>{props}</a:endParaRPr></a:p>')
    return (f'<xdr:txBody><a:bodyPr vertOverflow="overflow" horzOverflow="overflow" wrap="square" rtlCol="0" '
            f'anchor="{anchor}"/><a:lstStyle/>{"".join(paras)}</xdr:txBody>')


def _from(geom, x, y):
    r, c, dx, dy = geom.anchor(x, y)
    return (f'<xdr:from><xdr:col>{c}</xdr:col><xdr:colOff>{round(dx * EMU)}</xdr:colOff>'
            f'<xdr:row>{r}</xdr:row><xdr:rowOff>{round(dy * EMU)}</xdr:rowOff></xdr:from>')


def anchors_xml(objects, geom, first_id=1000):
    """The <xdr:oneCellAnchor> elements of the shapes among a sheet's floating objects (each declaring its
    namespaces, so they can join a drawing part written by someone else)."""
    boxes = [geom.box(o) if o.get("type") == shapes.SHAPE else (0, 0, 0, 0) for o in objects]
    paths = connector_paths(objects, boxes)
    ids, n = {}, first_id
    for o in objects:
        if o.get("type") == shapes.SHAPE:
            ids[o.get("id")] = n
            n += 1
    out = []
    for i, obj in enumerate(objects):
        if obj.get("type") != shapes.SHAPE:
            continue
        nid = ids[obj.get("id")]
        name = quoteattr(obj.get("name") or f"{shapes.name_of(obj.get('kind'))} {nid - first_id + 1}")
        if shapes.is_line(obj):
            out.append(_connector_xml(obj, paths[i], geom, nid, name, ids))
            continue
        x0, y0, x1, y1 = boxes[i]
        w, h = max(x1 - x0, 0.01), max(y1 - y0, 0.01)
        descr = f' descr="{TAG}{obj["kind"]}"' if not shapes.KINDS.get(obj.get("kind"), (0, 0, 1))[2] else ""
        box_attr = ' txBox="1"' if obj.get("kind") == "textBox" else ""
        sp = (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{nid}" name={name}{descr}/>'
              f'<xdr:cNvSpPr{box_attr}/></xdr:nvSpPr><xdr:spPr>{_xfrm(obj, x0, y0, w, h)}{_geometry(obj, w, h)}'
              f'{_color(obj.get("fill"))}{_line(obj)}</xdr:spPr>{_text(obj)}</xdr:sp>')
        out.append(f'<xdr:oneCellAnchor {_DECLARE}>{_from(geom, x0, y0)}'
                   f'<xdr:ext cx="{round(w * EMU)}" cy="{round(h * EMU)}"/>{sp}<xdr:clientData/></xdr:oneCellAnchor>')
    return "".join(out)


def _connector_xml(obj, path, geom, nid, name, ids):
    glue = ""
    for tag, end in (("stCxn", "start"), ("endCxn", "end")):
        g = obj.get(end)
        if g and g.get("id") in ids:
            glue += f'<a:{tag} id="{ids[g["id"]]}" idx="{int(g.get("site", 0))}"/>'
    if obj.get("kind") == "line" or len(path) == 2:
        (x0, y0, x1, y1), fh, fv = shapes.box_for_ends(path[0], path[-1])
        prst = {"elbow": "bentConnector3", "curve": "curvedConnector3"}.get(obj.get("kind"), "straightConnector1")
        geometry = f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>'
        xfrm = _xfrm(dict(obj, flip_h=fh, flip_v=fv, rot=0), x0, y0, x1 - x0, y1 - y0)
    else:  # the route as drawn, as a custom path in its bounds
        x0, y0, x1, y1 = shapes.bounds(path)
        w, h = max(x1 - x0, 0.01), max(y1 - y0, 0.01)
        local = [(x - x0, y - y0) for x, y in path]
        geometry = (f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
                    f'<a:pathLst>{_path_xml([(local, False, "none")], w, h)}</a:pathLst></a:custGeom>')
        xfrm = _xfrm(dict(obj, flip_h=False, flip_v=False, rot=0), x0, y0, w, h)
    w, h = max(x1 - x0, 0), max(y1 - y0, 0)
    descr = f' descr="{TAG}{obj["kind"]}"' if obj.get("kind") != "line" else ""
    sp = (f'<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="{nid}" name={name}{descr}/><xdr:cNvCxnSpPr>{glue}'
          f'</xdr:cNvCxnSpPr></xdr:nvCxnSpPr><xdr:spPr>{xfrm}{geometry}<a:noFill/>{_line(obj)}</xdr:spPr></xdr:cxnSp>')
    return (f'<xdr:oneCellAnchor {_DECLARE}>{_from(geom, x0, y0)}<xdr:ext cx="{round(w * EMU)}" cy="{round(h * EMU)}"/>'
            f'{sp}<xdr:clientData/></xdr:oneCellAnchor>')


def connector_paths(objects, boxes, px=1.0):
    """{index: points} of the lines among a sheet's objects placed on `boxes` (one per object, in any
    coordinates): glued ends follow the shapes they are glued to."""
    ids = {o.get("id"): i for i, o in enumerate(objects) if o.get("type") == shapes.SHAPE and o.get("id")}

    def glued(glue):
        i = ids.get((glue or {}).get("id"))
        if i is None:
            return None
        target = objects[i]
        if shapes.is_line(target):
            return None
        points = shapes.sites(target, boxes[i])
        k = glue.get("site", 0)
        if not isinstance(k, int) or not 0 <= k < len(points):
            return None
        return points[k], shapes.rotated_box(target, boxes[i])

    out = {}
    for i, obj in enumerate(objects):
        if obj.get("type") != shapes.SHAPE or not shapes.is_line(obj):
            continue
        p0, p1 = shapes.line_ends(obj, boxes[i])
        a0 = a1 = b0 = b1 = None
        g = glued(obj.get("start"))
        if g:
            (x, y, a0), b0 = g
            p0 = (x, y)
        g = glued(obj.get("end"))
        if g:
            (x, y, a1), b1 = g
            p1 = (x, y)
        out[i] = shapes.route(obj, p0, p1, a0, a1, b0, b1, px)
    return out


# ------------------------------------------------------------------ reading

def _q(tag):
    prefix, name = tag.split(":")
    return f"{{{NS[prefix]}}}{name}"


def _hex_hsl(color):
    r, g, b = (int(color[i:i + 2], 16) / 255 for i in (1, 3, 5))
    return colorsys.rgb_to_hls(r, g, b)


def _hsl_hex(h, l, s):
    r, g, b = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
    return "#%02x%02x%02x" % tuple(round(v * 255) for v in (r, g, b))


def _read_color(el, placeholder=None):
    """#rrggbb of a colour element's child (srgbClr, schemeClr, sysClr, prstClr) with its modifiers."""
    if el is None:
        return None
    for child in el:
        tag = child.tag.split("}")[1]
        if tag == "srgbClr":
            color = "#" + child.get("val", "000000").lower()
        elif tag == "schemeClr":
            val = child.get("val")
            color = placeholder if val == "phClr" and placeholder else THEME.get(val, "#000000")
        elif tag == "sysClr":
            color = "#" + (child.get("lastClr") or ("FFFFFF" if child.get("val") == "window" else "000000")).lower()
        elif tag == "prstClr":
            color = {"black": "#000000", "white": "#ffffff", "red": "#ff0000", "green": "#008000",
                     "blue": "#0000ff", "yellow": "#ffff00"}.get(child.get("val"), "#000000")
        else:
            continue
        for mod in child:
            name, v = mod.tag.split("}")[1], int(mod.get("val", "100000")) / 100000
            h, l, s = _hex_hsl(color)
            if name == "lumMod":
                color = _hsl_hex(h, l * v, s)
            elif name == "lumOff":
                color = _hsl_hex(h, l + v, s)
            elif name == "shade":
                color = shapes._mix(color, "#000000", 1 - v)
            elif name == "tint":
                color = shapes._mix(color, "#ffffff", 1 - v)
        return color
    return None


def _fill_of(sp_pr, style_ref, default):
    """Fill colour from spPr (solidFill / noFill / gradFill) or the style reference."""
    if sp_pr is not None:
        if sp_pr.find(_q("a:noFill")) is not None:
            return None
        solid = sp_pr.find(_q("a:solidFill"))
        if solid is not None:
            return _read_color(solid)
        grad = sp_pr.find(_q("a:gradFill"))
        if grad is not None:
            stop = grad.find(f"{_q('a:gsLst')}/{_q('a:gs')}")
            return _read_color(stop)
    if style_ref is not None and style_ref.get("idx", "0") not in ("0", ""):
        return _read_color(style_ref)
    return default


def _number(el, name, default=0):
    try:
        return int(el.get(name, default))
    except (TypeError, ValueError):
        return default


def _read_paths(cust):
    """[(points 0..1, closed, fill mode)] of a custGeom's paths (curves flattened)."""
    out = []
    for path in cust.iter(_q("a:path")):
        w, h = max(1, _number(path, "w", 1)), max(1, _number(path, "h", 1))
        fill = path.get("fill") or "norm"
        if fill not in ("none", "lighten", "darken", "lightenLess", "darkenLess"):
            fill = "norm"
        fill = {"lightenLess": "lighten", "darkenLess": "darken"}.get(fill, fill)
        pts, closed, current = [], False, (0, 0)

        def pt(e):
            return _number(e, "x") / w, _number(e, "y") / h
        for cmd in path:
            tag = cmd.tag.split("}")[1]
            ps = [pt(p) for p in cmd.findall(_q("a:pt"))]
            if tag == "moveTo" and ps:
                if len(pts) >= 2:
                    out.append((pts, False, fill))
                pts, current = [ps[0]], ps[0]
            elif tag == "lnTo" and ps:
                pts.append(ps[0])
                current = ps[0]
            elif tag == "cubicBezTo" and len(ps) == 3:
                pts += shapes._cubic(current, ps[0], ps[1], ps[2], 12)
                current = ps[2]
            elif tag == "quadBezTo" and len(ps) == 2:
                pts += shapes._quad(current, ps[0], ps[1], 12)
                current = ps[1]
            elif tag == "arcTo":
                wr, hr = _number(cmd, "wR") / w, _number(cmd, "hR") / h
                st, sw = _number(cmd, "stAng") / 60000, _number(cmd, "swAng") / 60000
                cx = current[0] - wr * math.cos(math.radians(st))
                cy = current[1] - hr * math.sin(math.radians(st))
                arc = shapes._arc(cx, cy, wr, hr, st, st + sw, max(4, round(abs(sw) / 8)))
                pts += arc[1:]
                current = arc[-1]
            elif tag == "close":
                closed = True
        if len(pts) >= 2:
            out.append((pts, closed, fill))
    return out


def _read_text(sp):
    body = sp.find(_q("xdr:txBody"))
    if body is None:
        return "", {}, {}
    lines, font, para = [], {}, {}
    for p in body.findall(_q("a:p")):
        parts = []
        for child in p:
            tag = child.tag.split("}")[1]
            if tag in ("r", "fld"):
                t = child.find(_q("a:t"))
                parts.append((t.text or "") if t is not None else "")
                rpr = child.find(_q("a:rPr"))
                if rpr is not None and not font:
                    font = _font_of(rpr)
            elif tag == "br":
                parts.append("\n")
            elif tag == "pPr" and not para:
                para = {"align": {"l": "left", "r": "right"}.get(child.get("algn"), "center")}
        lines.append("".join(parts))
        if not font:
            end = p.find(_q("a:endParaRPr"))
            if end is not None:
                font = _font_of(end)
    bpr = body.find(_q("a:bodyPr"))
    if bpr is not None:
        para["valign"] = {"t": "top", "b": "bottom"}.get(bpr.get("anchor"), "middle")
    return "\n".join(lines).strip("\n"), font, para


def _font_of(rpr):
    font = {}
    if rpr.get("sz"):
        size = _number(rpr, "sz", 1100) / 100
        font["size"] = int(size) if size.is_integer() else size
    for key, attr in (("bold", "b"), ("italic", "i")):
        if rpr.get(attr) in ("1", "true"):
            font[key] = True
    if rpr.get("u") not in (None, "none"):
        font["underline"] = True
    solid = rpr.find(_q("a:solidFill"))
    if solid is not None:
        font["color"] = _read_color(solid)
    latin = rpr.find(_q("a:latin"))
    if latin is not None and latin.get("typeface") and not latin.get("typeface").startswith("+"):
        font["font"] = latin.get("typeface")
    return font


def _shape_from_xml(sp, box, geom, new_id):
    """A shape dict from an <xdr:sp> or <xdr:cxnSp> placed on box (pixels), or None."""
    connector = sp.tag == _q("xdr:cxnSp")
    nv = sp.find(_q("xdr:nvCxnSpPr" if connector else "xdr:nvSpPr"))
    cnv = nv.find(_q("xdr:cNvPr")) if nv is not None else None
    sp_pr = sp.find(_q("xdr:spPr"))
    style = sp.find(_q("xdr:style"))
    refs = {k: style.find(_q(f"a:{k}")) if style is not None else None for k in ("lnRef", "fillRef", "fontRef")}
    prst_el = sp_pr.find(_q("a:prstGeom")) if sp_pr is not None else None
    cust = sp_pr.find(_q("a:custGeom")) if sp_pr is not None else None
    prst = prst_el.get("prst") if prst_el is not None else None
    descr = (cnv.get("descr") or "") if cnv is not None else ""
    xfrm = sp_pr.find(_q("a:xfrm")) if sp_pr is not None else None
    if connector or prst in ("line", "straightConnector1") or (prst or "").startswith(("bentConnector", "curvedConnector")):
        kind = "elbow" if (prst or "").startswith("bentConnector") else \
            "curve" if (prst or "").startswith("curvedConnector") else "line"
        if cust is not None:
            kind = "elbow"
        if descr.startswith(TAG) and descr[len(TAG):] in shapes.LINE_KINDS:
            kind = descr[len(TAG):]
    elif descr.startswith(TAG) and descr[len(TAG):] in shapes.KINDS:
        kind = descr[len(TAG):]
    elif prst:
        kind = shapes._PRESET_KIND.get(prst, "rect")
        cnv_sp = nv.find(_q("xdr:cNvSpPr")) if nv is not None else None
        if kind == "rect" and cnv_sp is not None and cnv_sp.get("txBox") == "1":
            kind = "textBox"
    elif cust is not None:
        kind = "custom"
    else:
        return None
    line = kind in shapes.LINE_KINDS
    ln = sp_pr.find(_q("a:ln")) if sp_pr is not None else None
    line_color = _fill_of(ln, refs["lnRef"], None) if ln is not None else (
        _read_color(refs["lnRef"]) if refs["lnRef"] is not None and refs["lnRef"].get("idx", "0") != "0" else None)
    line_w = (_number(ln, "w") / PT if ln is not None and ln.get("w") else
              _THEME_LINE.get(_number(refs["lnRef"], "idx"), 0.75) if refs["lnRef"] is not None else 0.75)
    obj = shapes.new_shape(kind, (0, 0, 0, 0), (0, 0), {}, id=new_id,
                           name=cnv.get("name") if cnv is not None else None,
                           fill=None if line else _fill_of(sp_pr, refs["fillRef"], None),
                           line=line_color, line_w=round(line_w, 2))
    if ln is not None:
        dash = ln.find(_q("a:prstDash"))
        if dash is not None:
            obj["dash"] = _DASH_IN.get(dash.get("val"))
        for tag, key in (("headEnd", "head"), ("tailEnd", "tail")):
            end = ln.find(_q(f"a:{tag}"))
            if end is not None and end.get("type") in shapes.ARROWS:
                obj[key] = end.get("type")
    if prst and kind == "rect" and prst != "rect":
        obj["xl_prst"] = prst
    if kind == "custom" and cust is not None:
        obj["paths"] = _read_paths(cust)
    if kind in ("wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout") and prst_el is not None:
        adj = {gd.get("name"): gd.get("fmla", "") for gd in prst_el.iter(_q("a:gd"))}
        try:
            obj["tip"] = (0.5 + int(adj.get("adj1", "val -20833").split()[1]) / 100000,
                          0.5 + int(adj.get("adj2", "val 62500").split()[1]) / 100000)
        except (IndexError, ValueError):
            pass
    if xfrm is not None:
        obj["rot"] = round(_number(xfrm, "rot") / 60000, 2) % 360 if not line else 0.0
        obj["flip_h"] = xfrm.get("flipH") in ("1", "true")
        obj["flip_v"] = xfrm.get("flipV") in ("1", "true")
    if not line:
        text, font, para = _read_text(sp)
        base = {"size": 11, "color": _read_color(refs["fontRef"]) if refs["fontRef"] is not None else "#000000"}
        obj["text"] = text
        obj["font"] = dict(base, **{k: v for k, v in font.items() if v is not None})
        obj.update(para)
    x0, y0, x1, y1 = box
    obj["anchor"] = geom.anchor(x0, y0)
    obj["size"] = (max(0, round(x1 - x0)), max(0, round(y1 - y0)))
    if connector and nv is not None:
        cxn = nv.find(_q("xdr:cNvCxnSpPr"))
        for tag, end in (("stCxn", "start"), ("endCxn", "end")):
            el = cxn.find(_q(f"a:{tag}")) if cxn is not None else None
            if el is not None:
                obj[end] = {"xl_id": el.get("id"), "site": _number(el, "idx")}
    obj["xl_nid"] = cnv.get("id") if cnv is not None else None
    return obj


def _xfrm_box(xfrm):
    off, ext = xfrm.find(_q("a:off")), xfrm.find(_q("a:ext"))
    if off is None or ext is None:
        return None
    return (_number(off, "x") / EMU, _number(off, "y") / EMU, _number(ext, "cx") / EMU, _number(ext, "cy") / EMU)


def _cell_point(el, geom):
    if el is None:
        return None
    vals = {k: int((el.find(_q(f"xdr:{k}")).text or "0")) if el.find(_q(f"xdr:{k}")) is not None else 0
            for k in ("col", "colOff", "row", "rowOff")}
    return geom.col_x(vals["col"]) + vals["colOff"] / EMU, geom.row_y(vals["row"]) + vals["rowOff"] / EMU


def shapes_from_drawing(data, geom, new_id):
    """The shapes and connectors of a drawing part (bytes) as shape dicts, connectors glued by id;
    pictures and charts are left to openpyxl. new_id() gives ids."""
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return []
    found = []

    def walk(el, origin, mapping):
        """el: an sp / cxnSp / grpSp; origin: where the anchor puts the top-left of the outermost thing;
        mapping(x, y, w, h) -> box in sheet pixels for a child's own xfrm box (groups)."""
        tag = el.tag
        if tag == _q("xdr:grpSp"):
            gpr = el.find(_q("xdr:grpSpPr"))
            gx = gpr.find(_q("a:xfrm")) if gpr is not None else None
            if gx is None:
                return
            off, ext = gx.find(_q("a:off")), gx.find(_q("a:ext"))
            choff, chext = gx.find(_q("a:chOff")), gx.find(_q("a:chExt"))
            if None in (off, ext, choff, chext):
                return
            o = (_number(off, "x") / EMU, _number(off, "y") / EMU)
            e = (_number(ext, "cx") / EMU, _number(ext, "cy") / EMU)
            co = (_number(choff, "x") / EMU, _number(choff, "y") / EMU)
            ce = (max(1e-6, _number(chext, "cx") / EMU), max(1e-6, _number(chext, "cy") / EMU))
            gbox = mapping(o[0], o[1], e[0], e[1])
            kx, ky = (gbox[2] - gbox[0]) / ce[0], (gbox[3] - gbox[1]) / ce[1]

            def inner(x, y, w, h):
                return (gbox[0] + (x - co[0]) * kx, gbox[1] + (y - co[1]) * ky,
                        gbox[0] + (x - co[0] + w) * kx, gbox[1] + (y - co[1] + h) * ky)
            for child in el:
                if child.tag in (_q("xdr:sp"), _q("xdr:cxnSp"), _q("xdr:grpSp")):
                    walk(child, origin, inner)
            return
        sp_pr = el.find(_q("xdr:spPr"))
        xfrm = sp_pr.find(_q("a:xfrm")) if sp_pr is not None else None
        xb = _xfrm_box(xfrm) if xfrm is not None else None
        if xb is None:
            return
        box = mapping(*xb)
        obj = _shape_from_xml(el, box, geom, new_id())
        if obj is not None:
            found.append(obj)

    for anchor in root:
        name = anchor.tag.split("}")[1]
        if name not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
            continue
        if name == "absoluteAnchor":
            pos = anchor.find(_q("xdr:pos"))
            origin = (_number(pos, "x") / EMU, _number(pos, "y") / EMU) if pos is not None else (0, 0)
        else:
            origin = _cell_point(anchor.find(_q("xdr:from")), geom) or (0, 0)
        for child in anchor:
            if child.tag not in (_q("xdr:sp"), _q("xdr:cxnSp"), _q("xdr:grpSp")):
                continue
            # the outermost thing's own xfrm offset sits at the anchor's cell position
            if child.tag == _q("xdr:grpSp"):
                gpr = child.find(_q("xdr:grpSpPr"))
                top = gpr.find(_q("a:xfrm")) if gpr is not None else None
            else:
                sp_pr = child.find(_q("xdr:spPr"))
                top = sp_pr.find(_q("a:xfrm")) if sp_pr is not None else None
            tb = _xfrm_box(top) if top is not None else None
            if tb is None:
                continue
            # placed at its anchor cell (column widths may differ from Excel's), sized as Excel drew it
            shift = (origin[0] - tb[0], origin[1] - tb[1])

            def mapping(x, y, w, h, shift=shift):
                return x + shift[0], y + shift[1], x + shift[0] + w, y + shift[1] + h
            walk(child, origin, mapping)
    by_xl = {o.pop("xl_nid"): o["id"] for o in found if o.get("xl_nid") is not None}
    for o in found:
        o.pop("xl_nid", None)
        for end in ("start", "end"):
            g = o.get(end)
            if g and "xl_id" in g:
                o[end] = {"id": by_xl[g["xl_id"]], "site": g["site"]} if g["xl_id"] in by_xl else None
    return found


def drawing_parts(z, sheet_part):
    """Paths of the drawing parts a worksheet part refers to (in an open .xlsx zip)."""
    folder, name = sheet_part.rsplit("/", 1)
    rels = f"{folder}/_rels/{name}.rels"
    try:
        xml = z.read(rels).decode("utf-8")
    except KeyError:
        return []
    out = []
    for m in re.finditer(r"<Relationship\b([^>]*)/?>", xml):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        if attrs.get("Type", "").endswith("/drawing"):
            target = attrs.get("Target", "")
            out.append(_resolve(folder, target))
    return out


def _resolve(folder, target):
    if target.startswith("/"):
        return target[1:]
    parts = folder.split("/")
    for piece in target.split("/"):
        if piece == "..":
            parts.pop()
        elif piece and piece != ".":
            parts.append(piece)
    return "/".join(parts)


def empty_drawing():
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr {_DECLARE}></xdr:wsDr>')
