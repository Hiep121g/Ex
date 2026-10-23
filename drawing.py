"""Shapes on the sheet - Insert > Shapes and the Diagram tab: drawing them over the cells, selecting,
moving, sizing and rotating them, typing their text, gluing connectors to them, and the commands that
format and arrange them. A mixin of excel_app.SpreadsheetApp, which supplies the grid geometry, undo and
the ribbon helpers.

Shapes live in the sheet's list of floating objects with the charts and pictures (Sheet.charts), so they
share one stacking order and the same anchoring to cells. Several objects can be selected: chart_sel is
the main one (whose handles show), more_sel the others.
"""

import copy
import math
import tkinter as tk

import formula_engine as fe
import shapes
from ribbon import Popup, Tooltip, color_popup

SHIFT, CONTROL, ALT = 0x0001, 0x0004, 0x20000
MODIFIERS = {"Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock", "Win_L", "Win_R"}
HANDLE_LINE = "#5f5f5f"
SELECT_LINE = "#8c8c8c"
GLUED = "#2fa84f"        # a line end glued to a shape
SITE = "#2f74d0"         # connection points shown while drawing a connector
GUIDE = "#e8457c"        # smart guides
ADJUST = "#ffd54a"       # the yellow handle of a callout's tail
FLOWCHART_MAIN = ("flowChartTerminator", "flowChartProcess", "flowChartDecision", "flowChartInputOutput",
                  "flowChartPredefinedProcess", "flowChartDocument", "flowChartMagneticDisk", "flowChartConnector")
GATES = ("gateAnd", "gateOr", "gateNot", "gateNand", "gateNor", "gateXor", "gateXnor", "gateBuffer")
ICON_STYLE = {"fill": "#ffffff", "line": "#404040", "line_w": 1.0}


def shape_icon(kind, style=None):
    """An icon drawing for ribbon.draw_icon: shape `kind` fitted into the 16 x 16 icon grid."""
    style = dict(ICON_STYLE, **(style or {}))

    def icon(cv, P, k, color):
        x0, y0, x1, y1 = P(1.5, 1.5, 14.5, 14.5)
        size = x1 - x0
        if kind in shapes.LINE_KINDS:
            shape = shapes.new_shape(kind, (0, 0, 0, 0), (0, 0), dict(style, fill=None), flip_v=True)
            p0, p1 = (x0, y1), (x1, y0)
            path = shapes.route(shape, p0, p1, 0 if kind == "elbow" else None, 180 if kind == "elbow" else None)
            shapes.draw(cv, shape, (0, 0, 0, 0), max(0.6, k * 0.75), None, path)
            return
        w, h = shapes.default_size(kind)
        f = min(size / max(w, 1), size / max(h, 1))
        w, h = w * f, h * f
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        shape = shapes.new_shape(kind, (0, 0, 0, 0), (w, h), style)
        if kind == "textBox":
            cv.create_rectangle(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2, outline="#7f7f7f", dash=(1, 1))
            cv.create_text(cx, cy, text="A", fill="#404040", font=("Segoe UI", -max(7, round(9 * k))))
            return
        shapes.draw(cv, shape, (cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2), max(0.6, k * 0.7))
    return icon


class DrawingTools:
    """Mixin: shapes and diagrams for SpreadsheetApp."""

    # ------------------------------------------------------------------ state and selection

    def _init_drawing(self):
        self.more_sel = set()
        self.insert_tool = None    # {"kind", "style", "sticky"}: what the next drag on the sheet draws
        self.object_mode = False   # Select Objects: dragging on the sheet selects shapes instead of cells
        self.snap_grid = False     # shapes snap to the cells' edges
        self.text_edit = None      # {"index", "id", "widget"} while a shape's text is typed
        self.obj_clip = None       # copied shapes: {"objects", "marker"}
        self.diagram_style = dict(shapes.DIAGRAM_STYLE)
        self.diagram_line = dict(shapes.DIAGRAM_LINE)
        self._boxes, self._shape_paths, self._id_index = [], {}, {}
        self._guides, self._hover_sites, self._tool_buttons = [], None, {}

    @property
    def chart_sel(self):
        return self.__dict__.get("_chart_sel")

    @chart_sel.setter
    def chart_sel(self, index):  # selecting one object (or none) drops the others
        self._chart_sel = index
        self.more_sel = set()

    def selected_objects(self):
        """Indices of the selected floating objects, the main one first."""
        n = len(self.sh.charts)
        main = self.chart_sel
        out = [main] if main is not None and main < n else []
        return out + sorted(i for i in self.more_sel if i < n and i != main)

    def selected_shapes(self, lines=None):
        """Selected shapes; lines=True only lines and connectors, False none of them."""
        out = []
        for i in self.selected_objects():
            obj = self.sh.charts[i]
            if obj["type"] == fe.SHAPE and (lines is None or shapes.is_line(obj) == lines):
                out.append(i)
        return out

    def select_objects(self, indices, main=None):
        indices = list(dict.fromkeys(i for i in indices if i is not None))
        if not indices:
            self.chart_sel = None
            return
        main = main if main in indices else indices[0]
        self.chart_sel = main
        self.more_sel = set(indices) - {main}

    def _objects_selected(self):
        return bool(self.selected_shapes()) or len(self.selected_objects()) > 1

    # ------------------------------------------------------------------ canvas <-> cells

    def _col_at_x(self, x):
        """(column, canvas x of its left edge) of canvas x, also outside the view (inverse of _sheet_x)."""
        c0, c1 = self.frozen()[2:]
        ox = self.pane_origin()[0]
        if c1 > c0 and x < ox:
            c, pos = c0, self.col_x(c0)
            while c < c1 - 1 and pos + self.col_w(c) <= x:
                pos += self.col_w(c)
                c += 1
            return c, pos
        c, pos = self.left, ox
        if x >= pos:
            while c < fe.MAX_COLS - 1 and pos + self.col_w(c) <= x:
                pos += self.col_w(c)
                c += 1
        else:
            while c > 0 and pos > x:
                c -= 1
                pos -= self.col_w(c)
        return c, pos

    def _row_at_y(self, y):
        """(row, canvas y of its top) of canvas y, also outside the view (inverse of _sheet_y)."""
        r0, r1 = self.frozen()[:2]
        oy = self.pane_origin()[1]
        if r1 > r0 and y < oy:
            r, pos = r0, self.row_y(r0)
            while r < r1 - 1 and pos + self.row_h(r) <= y:
                pos += self.row_h(r)
                r += 1
            return r, pos
        r, pos = self.top, oy
        far = self.canvas.winfo_height() + 4000
        step = max(1, self.default_row_px)
        if y >= pos:
            while r < fe.MAX_ROWS - 1 and pos + self.row_h(r) <= y:
                pos += self.row_h(r)
                r += 1
                if abs(pos) > far:  # _sheet_y estimates rows this far away
                    k = min(fe.MAX_ROWS - 1 - r, int((y - pos) // step))
                    return r + k, pos + k * step
        else:
            while r > 0 and pos > y:
                r -= 1
                pos -= self.row_h(r)
                if abs(pos) > far:
                    k = min(r, int(math.ceil((pos - y) / step)))
                    return r - k, pos - k * step
        return r, pos

    def _anchor_for(self, x, y):
        c, cx = self._col_at_x(x)
        r, ry = self._row_at_y(y)
        return r, c, max(0, round((x - cx) / self.px)), max(0, round((y - ry) / self.px))

    def _set_box(self, index, box, base=None, **changes):
        """Place object `index` (from `base`, a version of it, when given) on canvas box (x0, y0, x1, y1)."""
        x0, y0, x1, y1 = box
        obj = base if base is not None else self.sh.charts[index]
        size = (max(0, round((x1 - x0) / self.px)), max(0, round((y1 - y0) / self.px)))
        self.sh.charts[index] = dict(obj, anchor=self._anchor_for(x0, y0), size=size, **changes)

    # ------------------------------------------------------------------ geometry of the sheet's shapes

    def _prepare_shapes(self):
        """Canvas boxes of the sheet's objects, shape indices by id and the routes of lines (glued ends
        follow their shapes) - for drawing and for finding what is under the mouse."""
        objs = self.sh.charts
        self._boxes = [self._chart_box(o) for o in objs]
        self._id_index = {o.get("id"): i for i, o in enumerate(objs) if o["type"] == fe.SHAPE and o.get("id")}
        self._shape_paths = {i: self._route(o, self._boxes[i]) for i, o in enumerate(objs)
                             if o["type"] == fe.SHAPE and shapes.is_line(o)}

    def _glue_end(self, glue):
        """(x, y, direction, bounds of the shape) of the connection point a line end is glued to."""
        if not glue:
            return None
        i = self._id_index.get(glue.get("id"))
        if i is None or i >= len(self._boxes):
            return None
        target = self.sh.charts[i]
        if target["type"] != fe.SHAPE or shapes.is_line(target):
            return None
        box = self._boxes[i]
        points = shapes.sites(target, box)
        k = glue.get("site", 0)
        if not isinstance(k, int) or not 0 <= k < len(points):
            return None
        x, y, a = points[k]
        return x, y, a, shapes.rotated_box(target, box)

    def _route(self, obj, box):
        p0, p1 = shapes.line_ends(obj, box)
        a0 = a1 = b0 = b1 = None
        g = self._glue_end(obj.get("start"))
        if g:
            p0, a0, b0 = g[:2], g[2], g[3]
        g = self._glue_end(obj.get("end"))
        if g:
            p1, a1, b1 = g[:2], g[2], g[3]
        return shapes.route(obj, p0, p1, a0, a1, b0, b1, self.px)

    def _visual_box(self, i):
        """Axis-aligned canvas bounds of object i as drawn."""
        obj, box = self.sh.charts[i], self._boxes[i]
        if obj["type"] != fe.SHAPE:
            return box
        if shapes.is_line(obj):
            return shapes.bounds(self._shape_paths.get(i) or shapes.line_ends(obj, box))
        return shapes.rotated_box(obj, box)

    def _shape_font(self, spec):
        return self.font_for({"font": spec.get("font"), "size": spec.get("size") or 11, "bold": spec.get("bold"),
                              "italic": spec.get("italic"), "underline": spec.get("underline")})

    def _sync_connectors(self):
        """Store where glued connectors run now (their ends follow the shapes they are glued to); glue to
        shapes that are gone is dropped."""
        self._prepare_shapes()
        for i, obj in enumerate(self.sh.charts):
            if obj["type"] != fe.SHAPE or not shapes.is_line(obj) or not (obj.get("start") or obj.get("end")):
                continue
            changes = {end: None for end in ("start", "end") if obj.get(end) and self._glue_end(obj[end]) is None}
            path = self._shape_paths[i]
            box, fh, fv = shapes.box_for_ends(path[0], path[-1])
            self._set_box(i, box, flip_h=fh, flip_v=fv, **changes)
        self._prepare_shapes()

    # ------------------------------------------------------------------ drawing on the grid

    def _draw_shape(self, cv, i, obj, box, W, H):
        """Draw shape i; False when it lies outside the canvas."""
        path = self._shape_paths.get(i)
        x0, y0, x1, y1 = shapes.bounds(path) if path else shapes.rotated_box(obj, box)
        m = 40 * self.px
        if x1 + m < self.header_w or x0 - m > W or y1 + m < self.header_h or y0 - m > H:
            return False
        if self.text_edit and self.text_edit["index"] == i:
            obj = dict(obj, text="")
        shapes.draw(cv, obj, box, self.px, self._shape_font, path)
        return True

    def _draw_drawing_overlay(self, cv, W, H):
        """Selection handles, connection points, smart guides, the selection band and the shape being drawn."""
        s = max(3, round(3.5 * self.px))
        d = self.drag or {}
        if d.get("mode") == "shape_draw" and d.get("moved"):
            obj, box, path = self._draw_preview(d)
            shapes.draw(cv, obj, box, self.px, self._shape_font, path)
        sel = self.selected_objects()
        for i in sel:
            obj, box = self.sh.charts[i], self._boxes[i]
            main = i == self.chart_sel
            if obj["type"] != fe.SHAPE:
                x0, y0, x1, y1 = box
                cv.create_rectangle(x0, y0, x1, y1, outline="#7a7a7a", width=max(1, round(self.px)))
                if main and len(sel) == 1:
                    for hx, hy in self._chart_handles(box).values():
                        cv.create_rectangle(hx - s, hy - s, hx + s, hy + s, fill="#ffffff", outline=HANDLE_LINE)
                continue
            if shapes.is_line(obj):
                path = self._shape_paths.get(i) or list(shapes.line_ends(obj, box))
                for end, (hx, hy) in (("start", path[0]), ("end", path[-1])):
                    glued = self._glue_end(obj.get(end)) is not None
                    cv.create_oval(hx - s, hy - s, hx + s, hy + s, fill=GLUED if glued else "#ffffff", outline=HANDLE_LINE)
                continue
            hs = shapes.handles(obj, box, self.px)
            ring = [hs[k] for k in ("lt", "rt", "rb", "lb")]
            cv.create_polygon(*[v for p in ring for v in p], fill="", outline=SELECT_LINE, dash=(3, 2))
            if not main:
                continue
            for name, (hx, hy) in hs.items():
                if name == "rot":
                    cv.create_line(*hs["t"], hx, hy, fill=SELECT_LINE)
                    cv.create_oval(hx - s - 1, hy - s - 1, hx + s + 1, hy + s + 1, fill="#ffffff", outline=HANDLE_LINE)
                    cv.create_arc(hx - s + 1, hy - s + 1, hx + s - 1, hy + s - 1, start=30, extent=270, style="arc",
                                  outline=HANDLE_LINE)
                elif name == "tip":
                    cv.create_polygon(hx, hy - s, hx + s, hy, hx, hy + s, hx - s, hy, fill=ADJUST, outline=HANDLE_LINE)
                else:
                    cv.create_oval(hx - s, hy - s, hx + s, hy + s, fill="#ffffff", outline=HANDLE_LINE)
        if self._hover_sites is not None and self._hover_sites < len(self.sh.charts):
            target = self.sh.charts[self._hover_sites]
            r = max(2, round(3 * self.px))
            for x, y, _ in shapes.sites(target, self._boxes[self._hover_sites]):
                cv.create_oval(x - r, y - r, x + r, y + r, fill="#ffffff", outline=SITE, width=max(1, round(self.px)))
        for kind, pos, a, b in self._guides:
            if kind == "v":
                cv.create_line(pos, a, pos, b, fill=GUIDE, dash=(4, 2))
            else:
                cv.create_line(a, pos, b, pos, fill=GUIDE, dash=(4, 2))
        if d.get("mode") == "shape_band":
            (x0, y0), (x1, y1) = d["start"], d["now"]
            cv.create_rectangle(x0, y0, x1, y1, outline="#4a7ebb", fill="#cfe0f5", stipple="gray25")
            cv.create_rectangle(x0, y0, x1, y1, outline="#4a7ebb")
        if self.text_edit:
            self._place_text_editor()

    def _shape_hit(self, x, y):
        """(index, box, handle or None) of what a click at canvas (x, y) falls on among the shapes and the
        selected objects' handles - or None."""
        tol = max(4, 6 * self.px)
        main = self.chart_sel
        if main is not None and main < len(self.sh.charts) and main < len(self._boxes):
            obj = self.sh.charts[main]
            if obj["type"] == fe.SHAPE:
                box = self._boxes[main]
                if shapes.is_line(obj):
                    path = self._shape_paths.get(main) or list(shapes.line_ends(obj, box))
                    found = {"start": path[0], "end": path[-1]}
                else:
                    found = shapes.handles(obj, box, self.px)
                for name, (hx, hy) in found.items():
                    if abs(x - hx) <= tol and abs(y - hy) <= tol:
                        return main, box, name
        return None

    def _object_hit(self, i, x, y):
        """Whether a click falls on object i (shapes by their outline, others by their box)."""
        obj, box = self.sh.charts[i], self._boxes[i]
        if obj["type"] != fe.SHAPE:
            return box[0] <= x <= box[2] and box[1] <= y <= box[3]
        return shapes.hit(obj, box, x, y, max(4, 5 * self.px), self._shape_paths.get(i))

    def _site_near(self, x, y, exclude=()):
        """(shape index, site number, (x, y)) of the connection point nearest to (x, y) within reach, and
        the index of the shape under (x, y) (whose points to show) - for gluing line ends."""
        reach = max(8, 12 * self.px)
        best, under = None, None
        for i in range(len(self.sh.charts) - 1, -1, -1):
            obj = self.sh.charts[i]
            if i in exclude or obj["type"] != fe.SHAPE or shapes.is_line(obj) or i >= len(self._boxes):
                continue
            box = self._boxes[i]
            bx0, by0, bx1, by1 = shapes.rotated_box(obj, box)
            if not (bx0 - reach <= x <= bx1 + reach and by0 - reach <= y <= by1 + reach):
                continue
            if under is None:
                under = i
            for k, (sx, sy, _) in enumerate(shapes.sites(obj, box)):
                dist = math.hypot(sx - x, sy - y)
                if dist <= reach and (best is None or dist < best[0]):
                    best = (dist, i, k, (sx, sy))
        return (best[1:] if best else None), under

    # ------------------------------------------------------------------ mouse

    def _drawing_press(self, e):
        """A press on the sheet that concerns shapes: True when handled here."""
        if self.text_edit:
            self._finish_text_edit()
        if self.insert_tool:
            self.commit_edit()
            self.canvas.focus_set()
            self._start_draw(e.x, e.y, e.state)
            return True
        if self._point_mode_ok():
            return False
        prot = self.sh.protection
        if prot and "objects" not in prot.get("allow", ()):
            return False
        self._prepare_shapes()
        shift, ctrl = bool(e.state & SHIFT), bool(e.state & CONTROL)
        sel = self.selected_objects()
        hit = self._shape_hit(e.x, e.y)
        if hit is not None:
            self.commit_edit()
            self.canvas.focus_set()
            self._start_handle_drag(*hit, e)
            return True
        index = next((i for i in range(len(self.sh.charts) - 1, -1, -1) if self._object_hit(i, e.x, e.y)), None)
        if index is None:
            if self.object_mode:
                self.commit_edit()
                self.canvas.focus_set()
                keep = sel if (shift or ctrl) else []
                self.select_objects(keep)
                self.drag = {"mode": "shape_band", "start": (e.x, e.y), "now": (e.x, e.y), "keep": keep}
                self.redraw()
                return True
            if self.more_sel or self.selected_shapes():
                self.chart_sel = None
                self.redraw()
            return False
        obj = self.sh.charts[index]
        if obj["type"] != fe.SHAPE and not (shift or ctrl) and (len(sel) <= 1 or index not in sel):
            return False  # a chart or picture alone: moved and sized as before
        self.commit_edit()
        self.canvas.focus_set()
        if shift and index in sel:
            self.select_objects([i for i in sel if i != index])
            self.redraw()
            return True
        if shift or ctrl:
            if index not in sel:
                self.select_objects(sel + [index], main=index)
                self._start_move(e, dup=False)
            else:  # Ctrl on a selected object: dragging copies it, a click deselects it
                self._start_move(e, dup=True, toggle=index)
            return True
        if index not in sel:
            self.chart_sel = index
        self._start_move(e, dup=False)
        return True

    def _start_move(self, e, dup=False, toggle=None):
        self._sync_connectors()
        idx = self.selected_objects()
        self.drag = {"mode": "shape_move", "indices": idx, "boxes": {i: self._boxes[i] for i in idx},
                     "objs": {i: self.sh.charts[i] for i in idx}, "start": (e.x, e.y), "moved": False,
                     "dup": dup or bool(e.state & CONTROL), "toggle": toggle, "before": self.wb.snapshot(),
                     "sel": self._sel_state()}
        self.redraw()

    def _start_handle_drag(self, index, box, handle, e):
        self._sync_connectors()
        obj = self.sh.charts[index]
        box = self._boxes[index]
        mode = {"rot": "shape_rotate", "start": "shape_end", "end": "shape_end", "tip": "shape_tip"}.get(handle,
                                                                                                       "shape_resize")
        path = self._shape_paths.get(index)
        self.drag = {"mode": mode, "index": index, "handle": handle, "box": box, "obj": obj, "start": (e.x, e.y),
                     "path": path, "before": self.wb.snapshot(), "sel": self._sel_state(), "moved": False}

    def _start_draw(self, x, y, state):
        tool = self.insert_tool
        kind = tool["kind"]
        glue = None
        self._prepare_shapes()
        if kind in shapes.LINE_KINDS:
            near, under = self._site_near(x, y)
            if near:
                i, k, (x, y) = near
                glue = {"id": self.sh.charts[i]["id"], "site": k}
            self._hover_sites = under
        else:
            x, y = self._snap_point(x, y, ())[0]
        self.drag = {"mode": "shape_draw", "kind": kind, "style": tool["style"], "p0": (x, y), "p1": (x, y),
                     "glue0": glue, "glue1": None, "moved": False}
        self.redraw()

    def _drawing_drag(self, x, y, state):
        d = self.drag
        mode = d["mode"]
        shift, ctrl = bool(state & SHIFT), bool(state & CONTROL)
        self._guides = []
        if mode == "shape_band":
            d["now"] = (x, y)
        elif mode == "shape_move":
            self._drag_move(d, x, y, shift)
        elif mode == "shape_resize":
            obj = d["obj"]
            if not obj.get("rot"):
                (x, y), self._guides = self._snap_point(x, y, {d["index"]})
            box = shapes.resize(obj, d["box"], d["handle"], x, y, keep_ratio=shift, least=4 * self.px)
            if ctrl:  # from the centre
                bx0, by0, bx1, by1 = d["box"]
                cx, cy = (bx0 + bx1) / 2, (by0 + by1) / 2
                hw, hh = (box[2] - box[0]) / 2, (box[3] - box[1]) / 2
                box = (cx - hw, cy - hh, cx + hw, cy + hh)
            self._set_box(d["index"], box, base=obj)
        elif mode == "shape_rotate":
            bx0, by0, bx1, by1 = d["box"]
            angle = (math.degrees(math.atan2(y - (by0 + by1) / 2, x - (bx0 + bx1) / 2)) + 90) % 360
            if shift:
                angle = round(angle / 15) * 15 % 360
            self.sh.charts[d["index"]] = dict(d["obj"], rot=round(angle, 1))
        elif mode == "shape_tip":
            self.sh.charts[d["index"]] = dict(d["obj"], tip=shapes.tip_from_point(d["obj"], d["box"], x, y))
        elif mode == "shape_end":
            self._drag_line_end(d, x, y, shift)
        elif mode == "shape_draw":
            self._drag_draw(d, x, y, shift)
        d["moved"] = True
        self.redraw()

    def _drag_move(self, d, x, y, shift):
        sx, sy = d["start"]
        dx, dy = x - sx, y - sy
        if not d["moved"] and abs(dx) < 3 and abs(dy) < 3:
            return
        if shift:
            if abs(dx) > abs(dy):
                dy = 0
            else:
                dx = 0
        dx, dy, self._guides = self._snap_offset(d, dx, dy)
        moving = {self.sh.charts[i].get("id") for i in d["indices"]}
        for i in d["indices"]:
            obj = d["objs"][i]
            x0, y0, x1, y1 = d["boxes"][i]
            changes = {}
            if obj["type"] == fe.SHAPE and shapes.is_line(obj):  # glue to shapes left behind lets go
                changes = {end: None for end in ("start", "end") if obj.get(end) and obj[end].get("id") not in moving}
            self._set_box(i, (x0 + dx, y0 + dy, x1 + dx, y1 + dy), base=obj, **changes)

    def _drag_line_end(self, d, x, y, shift):
        index, obj = d["index"], d["obj"]
        path = d["path"] or list(shapes.line_ends(obj, d["box"]))
        other = path[-1] if d["handle"] == "start" else path[0]
        near, under = self._site_near(x, y, exclude={index})
        glue = None
        if near:
            i, k, (x, y) = near
            glue = {"id": self.sh.charts[i]["id"], "site": k}
        elif shift:  # at a multiple of 45 degrees
            ang = round(math.atan2(y - other[1], x - other[0]) / (math.pi / 4)) * (math.pi / 4)
            dist = math.hypot(x - other[0], y - other[1])
            x, y = other[0] + dist * math.cos(ang), other[1] + dist * math.sin(ang)
        self._hover_sites = under
        p0, p1 = ((x, y), other) if d["handle"] == "start" else (other, (x, y))
        box, fh, fv = shapes.box_for_ends(p0, p1)
        self._set_box(index, box, base=obj, flip_h=fh, flip_v=fv, **{d["handle"]: glue})

    def _drag_draw(self, d, x, y, shift):
        if d["kind"] in shapes.LINE_KINDS:
            near, under = self._site_near(x, y)
            d["glue1"] = None
            if near:
                i, k, (x, y) = near
                d["glue1"] = {"id": self.sh.charts[i]["id"], "site": k}
            elif shift:
                p0 = d["p0"]
                ang = round(math.atan2(y - p0[1], x - p0[0]) / (math.pi / 4)) * (math.pi / 4)
                dist = math.hypot(x - p0[0], y - p0[1])
                x, y = p0[0] + dist * math.cos(ang), p0[1] + dist * math.sin(ang)
            self._hover_sites = under
        else:
            (x, y), self._guides = self._snap_point(x, y, ())
            if shift:  # square / circle
                p0 = d["p0"]
                side = max(abs(x - p0[0]), abs(y - p0[1]))
                x = p0[0] + math.copysign(side, x - p0[0] or 1)
                y = p0[1] + math.copysign(side, y - p0[1] or 1)
        d["p1"] = (x, y)

    def _draw_preview(self, d):
        """(shape, box, route) of the shape a drag is drawing."""
        kind, p0, p1 = d["kind"], d["p0"], d["p1"]
        obj = shapes.new_shape(kind, (0, 0, 0, 0), (0, 0), d["style"])
        if kind in shapes.LINE_KINDS:
            box, fh, fv = shapes.box_for_ends(p0, p1)
            obj.update(flip_h=fh, flip_v=fv, start=d["glue0"], end=d["glue1"])
            return obj, box, self._route(obj, box)
        box = (min(p0[0], p1[0]), min(p0[1], p1[1]), max(p0[0], p1[0]), max(p0[1], p1[1]))
        return obj, box, None

    def _drawing_release(self, d, e):
        self._guides = []
        self._hover_sites = None
        mode = d["mode"]
        if mode == "shape_band":
            (x0, y0), (x1, y1) = d["start"], d["now"]
            x0, x1 = sorted((x0, x1))
            y0, y1 = sorted((y0, y1))
            self._prepare_shapes()
            found = [i for i in range(len(self.sh.charts)) if self._inside_band(i, x0, y0, x1, y1)]
            self.select_objects(d["keep"] + [i for i in found if i not in d["keep"]])
            self.redraw()
            return
        if mode == "shape_draw":
            self._finish_draw(d)
            return
        if mode == "shape_move" and d["dup"] and d["moved"]:
            self._finish_copy_move(d)
            return
        if not d["moved"]:
            if mode == "shape_move" and d.get("toggle") is not None:
                self.select_objects([i for i in self.selected_objects() if i != d["toggle"]])
            self.redraw()
            return
        self._sync_connectors()
        names = {"shape_move": "Di chuyển đối tượng", "shape_resize": "Đổi cỡ hình", "shape_rotate": "Xoay hình",
                 "shape_tip": "Đổi hình chú thích", "shape_end": "Nối đường"}
        self._push_snapshot(d["before"], d["sel"], names.get(mode, "Hình"))
        self.redraw()

    def _inside_band(self, i, x0, y0, x1, y1):
        if i >= len(self._boxes):
            return False
        bx0, by0, bx1, by1 = self._visual_box(i)
        return x0 <= bx0 and y0 <= by0 and bx1 <= x1 and by1 <= y1

    def _finish_draw(self, d):
        kind = d["kind"]
        p0, p1 = d["p0"], d["p1"]
        line = kind in shapes.LINE_KINDS
        if not d["moved"] or math.hypot(p1[0] - p0[0], p1[1] - p0[1]) < 4:
            w, h = shapes.default_size(kind)
            p1 = (p0[0] + w * self.px, p0[1] + (0 if line else h * self.px))
            d = dict(d, glue1=None)
        obj, box, _ = self._draw_preview(dict(d, p1=p1))
        obj["id"] = fe.new_picture_id()
        obj["name"] = f"{shapes.name_of(kind)} {sum(1 for o in self.sh.charts if o['type'] == fe.SHAPE) + 1}"
        prev, self._undo_name = self._undo_name, self._undo_name or ("Vẽ đường nối" if line else "Chèn hình")
        try:
            def change():
                self.sh.charts.append(obj)
                self._set_box(len(self.sh.charts) - 1, box)
            done = self._structural(change)
        finally:
            self._undo_name = prev
        tool = self.insert_tool
        if tool and not tool.get("sticky"):
            self.set_insert_tool(None)
        if done:
            self.chart_sel = len(self.sh.charts) - 1
            if tool and tool.get("after"):
                tool["after"]()
            if kind == "textBox":
                self.edit_shape_text(self.chart_sel)
        self.redraw()

    def _finish_copy_move(self, d):
        """Ctrl+drag: the dragged objects stay and copies land where they were dropped."""
        copies = self._copies([self.sh.charts[i] for i in d["indices"]])
        self.wb.restore(d["before"])  # the originals back where they were

        def change():
            self.sh.charts.extend(copies)
        prev, self._undo_name = self._undo_name, "Sao chép đối tượng"
        try:
            self._structural(change)
        finally:
            self._undo_name = prev
        n = len(self.sh.charts)
        self.select_objects(range(n - len(copies), n))
        self.redraw()

    # ------------------------------------------------------------------ snapping and smart guides

    def _other_bounds(self, exclude):
        out = []
        for j, obj in enumerate(self.sh.charts):
            if j in exclude or j >= len(self._boxes) or (obj["type"] == fe.SHAPE and shapes.is_line(obj)):
                continue
            out.append(self._visual_box(j))
        return out

    def _snap_offset(self, d, dx, dy):
        """Offsets that line the moved objects up with other shapes (edges or centres within a few pixels,
        shown with guides) or, with Snap to Grid, with the cells' edges."""
        boxes = [d["boxes"][i] if self.sh.charts[i]["type"] != fe.SHAPE or not shapes.is_line(d["objs"][i])
                 else None for i in d["indices"]]
        boxes = [b for b in boxes if b]
        if not boxes:
            return dx, dy, []
        rot = [shapes.rotated_box(d["objs"][i], d["boxes"][i]) if d["objs"][i]["type"] == fe.SHAPE else d["boxes"][i]
               for i in d["indices"] if not (d["objs"][i]["type"] == fe.SHAPE and shapes.is_line(d["objs"][i]))]
        gx0, gy0 = min(b[0] for b in rot) + dx, min(b[1] for b in rot) + dy
        gx1, gy1 = max(b[2] for b in rot) + dx, max(b[3] for b in rot) + dy
        adj_x, adj_y, guides = self._align_to_others((gx0, gy0, gx1, gy1), set(d["indices"]))
        if adj_x is None and self.snap_grid:
            adj_x = self._grid_delta(gx0, 0)
        if adj_y is None and self.snap_grid:
            adj_y = self._grid_delta(gy0, 1)
        return dx + (adj_x or 0), dy + (adj_y or 0), guides

    def _align_to_others(self, box, exclude):
        """(x adjustment, y adjustment, guides) lining box up with the other objects' edges and centres."""
        tol = max(4, 6 * self.scale)
        x0, y0, x1, y1 = box
        xs, ys = (x0, (x0 + x1) / 2, x1), (y0, (y0 + y1) / 2, y1)
        best_x = best_y = None
        for b in self._other_bounds(exclude):
            for tx in (b[0], (b[0] + b[2]) / 2, b[2]):
                for sx in xs:
                    delta = tx - sx
                    if abs(delta) <= tol and (best_x is None or abs(delta) < abs(best_x[0])):
                        best_x = (delta, tx, b)
            for ty in (b[1], (b[1] + b[3]) / 2, b[3]):
                for sy in ys:
                    delta = ty - sy
                    if abs(delta) <= tol and (best_y is None or abs(delta) < abs(best_y[0])):
                        best_y = (delta, ty, b)
        guides = []
        if best_x:
            b = best_x[2]
            guides.append(("v", best_x[1], min(b[1], y0) - 8, max(b[3], y1) + 8))
        if best_y:
            b = best_y[2]
            guides.append(("h", best_y[1], min(b[0], x0) - 8, max(b[2], x1) + 8))
        return (best_x[0] if best_x else None), (best_y[0] if best_y else None), guides

    def _grid_delta(self, pos, axis):
        """How far canvas position `pos` is from the nearest cell edge (columns: axis 0, rows: 1)."""
        if axis == 0:
            c, edge = self._col_at_x(pos)
            nxt = edge + self.col_w(c)
        else:
            r, edge = self._row_at_y(pos)
            nxt = edge + self.row_h(r)
        return (edge if pos - edge <= nxt - pos else nxt) - pos

    def _snap_point(self, x, y, exclude):
        """A dragged point lined up with other shapes or the grid: ((x, y), guides)."""
        adj_x, adj_y, guides = self._align_to_others((x, y, x, y), exclude)
        if adj_x is None and self.snap_grid:
            adj_x = self._grid_delta(x, 0)
        if adj_y is None and self.snap_grid:
            adj_y = self._grid_delta(y, 1)
        return (x + (adj_x or 0), y + (adj_y or 0)), guides

    def _drawing_motion(self, e):
        """Cursor (or None to let the grid choose) for the mouse at (e.x, e.y); shows connection points
        while a connector tool is chosen."""
        if self.insert_tool:
            sites = None
            if self.insert_tool["kind"] in shapes.LINE_KINDS:
                self._prepare_shapes()
                sites = self._site_near(e.x, e.y)[1]
            if sites != self._hover_sites:
                self._hover_sites = sites
                self.redraw()
            return "crosshair"
        if not self.sh.charts:
            return None
        hit = self._shape_hit(e.x, e.y)
        if hit is not None:
            return {"rot": "exchange", "start": "crosshair", "end": "crosshair", "tip": "hand2", "l": "sb_h_double_arrow",
                    "r": "sb_h_double_arrow", "t": "sb_v_double_arrow", "b": "sb_v_double_arrow",
                    "lt": "top_left_corner", "rb": "bottom_right_corner", "rt": "top_right_corner",
                    "lb": "bottom_left_corner"}.get(hit[2], "fleur")
        if len(self._boxes) != len(self.sh.charts):
            return None
        for i in range(len(self.sh.charts) - 1, -1, -1):
            if self.sh.charts[i]["type"] == fe.SHAPE and self._object_hit(i, e.x, e.y):
                return "fleur"
        return None

    # ------------------------------------------------------------------ keys

    def _drawing_key(self, e):
        """Keys while shapes are selected (or a drawing tool is chosen): True when handled here."""
        ks = e.keysym
        if self.insert_tool and ks == "Escape":
            self.set_insert_tool(None)
            return True
        sel = self.selected_objects()
        if not sel or ks in MODIFIERS or not self._objects_selected():
            return False
        shift, ctrl, alt = bool(e.state & SHIFT), bool(e.state & CONTROL), bool(e.state & ALT)
        arrows = {"Up": (0, -1), "Down": (0, 1), "Left": (-1, 0), "Right": (1, 0)}
        if ks in ("Delete", "BackSpace"):
            self.delete_objects()
        elif ks == "Escape":
            self.chart_sel = None
            self.redraw()
        elif ks in arrows:
            dx, dy = arrows[ks]
            if alt and dx:
                self.rotate_shapes(15 * dx)
            elif shift:
                self.resize_shapes(5 * dx, -5 * dy)
            else:
                step = 1 if ctrl else 5
                self.nudge_objects(dx * step, dy * step)
        elif ks in ("Return", "KP_Enter", "F2"):
            self.edit_shape_text()
        elif ks in ("Tab", "ISO_Left_Tab"):
            n = len(self.sh.charts)
            self.chart_sel = (sel[0] + (-1 if shift or ks == "ISO_Left_Tab" else 1)) % n
            self.redraw()
        elif e.char and e.char.isprintable() and not ctrl:
            self.edit_shape_text(initial=e.char)
        else:
            self.chart_sel = None  # other keys go back to the cells
            self.redraw()
            return False
        return True

    # ------------------------------------------------------------------ text in shapes

    def edit_shape_text(self, index=None, initial=None):
        """Type the text of a shape (Enter / F2 / typing / double-click on it): a box over the shape's text
        area until Esc or a click elsewhere."""
        if index is None:
            found = [i for i in self.selected_shapes(lines=False)]
            index = found[0] if found else None
        if index is None:
            self.root.bell()
            return
        obj = self.sh.charts[index]
        if obj["type"] != fe.SHAPE or shapes.is_line(obj):
            self.root.bell()
            return
        self._finish_text_edit()
        self.chart_sel = index
        spec = obj.get("font") or {}
        fg = spec.get("color") or "#000000"
        widget = tk.Text(self.canvas, font=self._shape_font(spec), fg=fg, bg=obj.get("fill") or "#ffffff",
                         insertbackground=fg, relief="flat", bd=0, highlightthickness=1,
                         highlightbackground="#9a9a9a", highlightcolor="#4a7ebb", wrap="word", undo=True,
                         padx=max(1, round(3 * self.px)), pady=max(1, round(2 * self.px)))
        align = obj.get("align") or "center"
        widget.tag_configure("text", justify=align)
        widget.insert("1.0", obj.get("text", "") if initial is None else initial)
        widget.tag_add("text", "1.0", "end")
        widget.mark_set("insert", "end-1c")
        widget.bind("<KeyRelease>", lambda ev: widget.tag_add("text", "1.0", "end"))
        widget.bind("<Escape>", lambda ev: (self._finish_text_edit(), "break")[1])
        widget.bind("<Control-Return>", lambda ev: (self._finish_text_edit(), "break")[1])
        widget.bind("<FocusOut>", lambda ev: self.root.after_idle(self._finish_text_edit_if_away))
        self.text_edit = {"index": index, "id": obj.get("id"), "widget": widget, "sheet": self.sheet}
        self.redraw()
        widget.focus_set()

    def _place_text_editor(self):
        te = self.text_edit
        i = self._id_index.get(te["id"])
        if i is None or te["sheet"] != self.sheet:
            return
        obj, box = self.sh.charts[i], self._boxes[i]
        cx, cy, w, h = shapes.text_area(obj, box)
        font = self._shape_font(obj.get("font") or {})
        w = max(w, 60 * self.px)
        h = max(h, font.metrics("linespace") * 1.6)
        te["widget"].place(x=round(cx - w / 2), y=round(cy - h / 2), width=round(w), height=round(h))

    def _finish_text_edit_if_away(self):
        te = self.text_edit
        if te and str(self.root.tk.call("focus")) != str(te["widget"]):
            self._finish_text_edit()

    def _finish_text_edit(self):
        te = self.text_edit
        if not te:
            return
        self.text_edit = None
        try:
            text = te["widget"].get("1.0", "end-1c")
            te["widget"].destroy()
        except tk.TclError:
            return
        sh = self.wb.sheets.get(te["sheet"])
        index = next((i for i, o in enumerate(sh.charts) if o.get("id") == te["id"]), None) if sh else None
        if index is None:
            return
        if text != sh.charts[index].get("text", ""):
            obj = sh.charts[index]
            prev, self._undo_name = self._undo_name, "Văn bản trong hình"
            try:
                self._structural(lambda: sh.charts.__setitem__(index, dict(obj, text=text)))
            finally:
                self._undo_name = prev
        if te["sheet"] == self.sheet:
            self.chart_sel = index
        self.canvas.focus_set()
        self.redraw()

    # ------------------------------------------------------------------ commands on the selection

    def _change_objects(self, name, func, indices=None):
        """Replace each selected object (or `indices`) by func(index, object) - None leaves it - as one
        undo step named `name`, keeping the selection."""
        indices = self.selected_objects() if indices is None else indices
        if not indices:
            self.root.bell()
            return False
        keep, main = self.selected_objects(), self.chart_sel
        prev, self._undo_name = self._undo_name, self._undo_name or name

        def change():
            for i in indices:
                new = func(i, self.sh.charts[i])
                if new is not None:
                    self.sh.charts[i] = new
            self._sync_connectors()
        try:
            done = self._structural(change)
        finally:
            self._undo_name = prev
        self.select_objects(keep, main)
        self.redraw()
        return done

    def _change_shapes(self, name, lines=None, **changes):
        """Set `changes` (shape keys) on the selected shapes; a dict value is merged (the font)."""
        targets = self.selected_shapes(lines)
        if not targets:
            self.root.bell()
            return

        def apply(i, obj):
            new = dict(obj)
            for key, value in changes.items():
                new[key] = dict(obj.get(key) or {}, **value) if isinstance(value, dict) else value
            return new
        self._change_objects(name, apply, targets)

    def nudge_objects(self, dx, dy):
        """Move the selected objects by (dx, dy) pixels at 96 DPI."""
        self._prepare_shapes()
        boxes = {i: self._boxes[i] for i in self.selected_objects()}
        k = self.px

        def move(i, obj):
            x0, y0, x1, y1 = boxes[i]
            self._set_box(i, (x0 + dx * k, y0 + dy * k, x1 + dx * k, y1 + dy * k), base=obj)
            return self.sh.charts[i]
        self._change_objects("Di chuyển đối tượng", move)

    def resize_shapes(self, dw, dh):
        self._prepare_shapes()
        k = self.px

        def grow(i, obj):
            if obj["type"] == fe.SHAPE and shapes.is_line(obj):
                return None
            x0, y0, x1, y1 = self._boxes[i]
            self._set_box(i, (x0, y0, max(x0 + 4 * k, x1 + dw * k), max(y0 + 4 * k, y1 + dh * k)), base=obj)
            return self.sh.charts[i]
        self._change_objects("Đổi cỡ hình", grow)

    def rotate_shapes(self, degrees):
        self._change_objects("Xoay hình", lambda i, o: dict(o, rot=((o.get("rot") or 0) + degrees) % 360)
                             if o["type"] == fe.SHAPE and not shapes.is_line(o) else None)

    def flip_shapes(self, axis):
        key = "flip_h" if axis == "h" else "flip_v"
        self._change_objects("Lật hình", lambda i, o: dict(o, **{key: not o.get(key), "rot": -(o.get("rot") or 0) % 360})
                             if o["type"] == fe.SHAPE else None)

    def delete_objects(self):
        sel = set(self.selected_objects())
        if not sel:
            return
        self._sync_connectors()
        gone = {self.sh.charts[i].get("id") for i in sel}

        def change():
            kept = []
            for i, obj in enumerate(self.sh.charts):
                if i in sel:
                    continue
                if obj["type"] == fe.SHAPE and shapes.is_line(obj):
                    obj = dict(obj, **{end: None for end in ("start", "end") if obj.get(end) and obj[end].get("id") in gone})
                kept.append(obj)
            self.sh.charts = kept
        self.chart_sel = None
        prev, self._undo_name = self._undo_name, self._undo_name or "Xóa đối tượng"
        try:
            self._structural(change)
        finally:
            self._undo_name = prev
        self.redraw()

    def _copies(self, objs):
        """Copies of objects with new ids; glue between them follows, glue to others lets go."""
        ids = {}
        out = []
        for obj in objs:
            new = copy.deepcopy(obj)
            if new["type"] == fe.SHAPE:
                ids[obj.get("id")] = new["id"] = fe.new_picture_id()
            elif new["type"] == fe.PICTURE:
                new["id"] = fe.new_picture_id()
            out.append(new)
        for new in out:
            if new["type"] == fe.SHAPE and shapes.is_line(new):
                for end in ("start", "end"):
                    glue = new.get(end)
                    new[end] = dict(glue, id=ids[glue["id"]]) if glue and glue.get("id") in ids else None
        return out

    def copy_objects(self, cut=False):
        sel = self.selected_objects()
        if not sel:
            return
        self._sync_connectors()
        objs = [copy.deepcopy(self.sh.charts[i]) for i in sel]
        names = "; ".join(o.get("name") or shapes.name_of(o.get("kind")) for o in objs if o["type"] == fe.SHAPE)
        marker = names or "[đối tượng]"
        try:  # other programs get the shapes' names; PyExcel knows its own copy by that text
            self.root.clipboard_clear()
            self.root.clipboard_append(marker)
        except tk.TclError:
            pass
        self.obj_clip = {"objects": objs, "marker": marker, "pastes": 0}
        if self.clip:
            self.clip = None
        if cut:
            self.delete_objects()

    def objects_on_clipboard(self):
        """Whether Ctrl+V pastes copied shapes: they were copied last (the clipboard still holds their mark)."""
        if not self.obj_clip:
            return False
        try:
            return self.root.clipboard_get() == self.obj_clip["marker"]
        except tk.TclError:
            return False

    def paste_objects(self, offset=None):
        clip = self.obj_clip
        if not clip:
            return
        clip["pastes"] += 1
        step = 12 * clip["pastes"] if offset is None else offset
        copies = self._copies(clip["objects"])
        for obj in copies:
            r, c, dx, dy = obj["anchor"]
            obj["anchor"] = (r, c, dx + step, dy + step)

        def change():
            self.sh.charts.extend(copies)
        prev, self._undo_name = self._undo_name, self._undo_name or "Dán đối tượng"
        try:
            self._structural(change)
        finally:
            self._undo_name = prev
        n = len(self.sh.charts)
        self.select_objects(range(n - len(copies), n))
        self.redraw()

    def duplicate_objects(self):
        """Ctrl+D: copies of the selected objects just below and right of them."""
        sel = self.selected_objects()
        if not sel:
            return
        self._sync_connectors()
        copies = self._copies([self.sh.charts[i] for i in sel])
        for obj in copies:
            r, c, dx, dy = obj["anchor"]
            obj["anchor"] = (r, c, dx + 12, dy + 12)
        prev, self._undo_name = self._undo_name, self._undo_name or "Nhân bản đối tượng"
        try:
            self._structural(lambda: self.sh.charts.extend(copies))
        finally:
            self._undo_name = prev
        n = len(self.sh.charts)
        self.select_objects(range(n - len(copies), n))
        self.redraw()

    def select_all_objects(self):
        if self.sh.charts:
            self.select_objects(range(len(self.sh.charts)), self.chart_sel)
            self.redraw()

    def order_objects(self, where):
        """Bring to front / send to back ("front", "back") or one step ("forward", "backward")."""
        sel = set(self.selected_objects())
        if not sel:
            self.root.bell()
            return
        objs = list(self.sh.charts)
        order = list(range(len(objs)))
        if where == "front":
            order = [i for i in order if i not in sel] + [i for i in order if i in sel]
        elif where == "back":
            order = [i for i in order if i in sel] + [i for i in order if i not in sel]
        elif where == "forward":
            for k in range(len(order) - 2, -1, -1):
                if order[k] in sel and order[k + 1] not in sel:
                    order[k], order[k + 1] = order[k + 1], order[k]
        else:
            for k in range(1, len(order)):
                if order[k] in sel and order[k - 1] not in sel:
                    order[k], order[k - 1] = order[k - 1], order[k]
        main = self.chart_sel
        names = {"front": "Đưa lên trên cùng", "back": "Đưa xuống dưới cùng", "forward": "Đưa lên một lớp",
                 "backward": "Đưa xuống một lớp"}
        prev, self._undo_name = self._undo_name, self._undo_name or names[where]
        try:
            self._structural(lambda: setattr(self.sh, "charts", [objs[i] for i in order]))
        finally:
            self._undo_name = prev
        self.select_objects([order.index(i) for i in sel], order.index(main) if main in sel else None)
        self.redraw()

    def align_objects(self, how):
        """Align the selected objects' left / centre / right edges or top / middle / bottom (to the
        selection's bounds), or spread them evenly ("hspace", "vspace")."""
        self._prepare_shapes()
        sel = [i for i in self.selected_objects()
               if not (self.sh.charts[i]["type"] == fe.SHAPE and shapes.is_line(self.sh.charts[i])
                       and (self.sh.charts[i].get("start") or self.sh.charts[i].get("end")))]
        if len(sel) < 2 or (how in ("hspace", "vspace") and len(sel) < 3):
            self.root.bell()
            return
        vis = {i: self._visual_box(i) for i in sel}
        x0, y0 = min(b[0] for b in vis.values()), min(b[1] for b in vis.values())
        x1, y1 = max(b[2] for b in vis.values()), max(b[3] for b in vis.values())
        moves = {}
        if how in ("hspace", "vspace"):
            axis = 0 if how == "hspace" else 1
            order = sorted(sel, key=lambda i: (vis[i][axis] + vis[i][axis + 2]) / 2)
            total = sum(vis[i][axis + 2] - vis[i][axis] for i in order)
            gap = ((x1 - x0) if axis == 0 else (y1 - y0)) - total
            gap /= len(order) - 1
            pos = x0 if axis == 0 else y0
            for i in order:
                b = vis[i]
                moves[i] = (pos - b[0], 0) if axis == 0 else (0, pos - b[1])
                pos += b[axis + 2] - b[axis] + gap
        else:
            for i, b in vis.items():
                moves[i] = {"left": (x0 - b[0], 0), "center": ((x0 + x1) / 2 - (b[0] + b[2]) / 2, 0),
                            "right": (x1 - b[2], 0), "top": (0, y0 - b[1]),
                            "middle": (0, (y0 + y1) / 2 - (b[1] + b[3]) / 2), "bottom": (0, y1 - b[3])}[how]
        boxes = {i: self._boxes[i] for i in sel}

        def move(i, obj):
            dx, dy = moves[i]
            bx0, by0, bx1, by1 = boxes[i]
            self._set_box(i, (bx0 + dx, by0 + dy, bx1 + dx, by1 + dy), base=obj)
            return self.sh.charts[i]
        self._change_objects("Căn chỉnh đối tượng", move, sel)

    def set_shape_size(self, width, height):
        """Size (pixels at 96 DPI) of the selected shapes, keeping their centres."""
        self._prepare_shapes()

        def size(i, obj):
            if obj["type"] == fe.SHAPE and shapes.is_line(obj):
                return None
            x0, y0, x1, y1 = self._boxes[i]
            cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
            w, h = (width or (x1 - x0) / self.px) * self.px / 2, (height or (y1 - y0) / self.px) * self.px / 2
            self._set_box(i, (cx - w, cy - h, cx + w, cy + h), base=obj)
            return self.sh.charts[i]
        self._change_objects("Đổi cỡ hình", size)

    def change_shape_kind(self, kind):
        """Edit Shape > Change Shape: the selected shapes take another outline, keeping everything else."""
        if kind in shapes.LINE_KINDS:
            self._change_shapes("Đổi kiểu đường nối", lines=True, kind=kind)
        else:
            self._change_shapes("Đổi hình", lines=False, kind=kind)

    def apply_quick_style(self, style):
        targets = self.selected_shapes()
        if not targets:
            self.root.bell()
            return

        def apply(i, obj):
            if shapes.is_line(obj):
                return dict(obj, line=style.get("line") or style.get("fill"))
            return dict(obj, fill=style.get("fill"), line=style.get("line"),
                        font=dict(obj.get("font") or {}, color=style["font"]["color"]))
        self._change_objects("Kiểu hình", apply, targets)

    # what the Home tab's font and fill buttons do to selected shapes

    def _shape_fmt(self):
        """The main selected shape's text formatting in cell-format terms (for the Home tab)."""
        sel = self.selected_shapes(lines=False)
        if not sel:
            return None
        obj = self.sh.charts[sel[0]]
        spec = obj.get("font") or {}
        return {"bold": spec.get("bold"), "italic": spec.get("italic"), "underline": spec.get("underline"),
                "size": spec.get("size") or 11, "font": spec.get("font"), "fg": spec.get("color"),
                "bg": obj.get("fill"), "align": obj.get("align") or "center",
                "valign": {"middle": "center"}.get(obj.get("valign"), obj.get("valign"))}

    def format_shapes_from_cells(self, update):
        """Apply a cell-format update (bold, size, font colour, fill, alignment...) to the selected shapes."""
        font, other = {}, {}
        for key, value in update.items():
            if key in ("bold", "italic"):
                font[key] = bool(value)
            elif key == "underline":
                font[key] = bool(value)
            elif key == "size":
                font[key] = value or 11
            elif key == "font":
                font[key] = value or None
            elif key == "fg":
                font["color"] = value or "#000000"
            elif key == "bg":
                other["fill"] = value or None
            elif key == "align":
                other["align"] = value if value in ("left", "right") else "center"
            elif key == "valign":
                other["valign"] = {"center": "middle"}.get(value, value) or "middle"
        if not font and not other:
            self.root.bell()
            return
        if font:
            other["font"] = font
        self._change_shapes("Định dạng hình", lines=None if "fill" not in other else False, **other)

    # ------------------------------------------------------------------ tools (ribbon)

    def set_insert_tool(self, kind, style=None, sticky=False, after=None, diagram=False, key=None):
        """Choose what the next drag on the sheet draws (None: back to selecting)."""
        self._finish_text_edit()
        if kind is None:
            self.insert_tool = None
        else:
            if style is None:
                style = (shapes.EXCEL_LINE if kind in shapes.LINE_KINDS else
                         shapes.TEXT_BOX_STYLE if kind == "textBox" else shapes.EXCEL_STYLE)
            self.insert_tool = {"kind": kind, "style": style, "sticky": sticky, "after": after, "diagram": diagram,
                                "key": key or kind}
            self.chart_sel = None
            self.set_mode("Kéo trên trang tính để vẽ — Esc để thôi")
        if kind is None:
            self.set_mode("Sẵn sàng")
        self._hover_sites = None
        tool = self.insert_tool
        for key, button in self._tool_buttons.items():
            if isinstance(key, tuple):
                try:
                    button.set_checked(bool(tool) and key == (tool.get("key", tool["kind"]), tool["diagram"]))
                except tk.TclError:
                    pass
        self.canvas.config(cursor="crosshair" if kind else "")
        self._cursor = "crosshair" if kind else ""
        self.redraw()

    def diagram_tool(self, kind):
        """A tool of the Diagram tab: shapes and connectors in the diagram style; connectors stay chosen
        so several can be drawn one after another (Esc ends)."""
        key = kind
        line = kind in shapes.LINE_KINDS or kind.startswith("line_")
        if line:
            base = dict(self.diagram_line)
            if kind == "line_plain":
                kind, base = "line", dict(base, tail=None)
            elif kind == "line_double":
                kind, base = "line", dict(base, head="triangle", tail="triangle")
            style = base
        elif kind == "textBox":
            style = dict(shapes.TEXT_BOX_STYLE, align="center", valign="middle")
        else:
            style = dict(self.diagram_style)
        self.set_insert_tool(kind, style, sticky=line, diagram=True, key=key)

    def toggle_object_mode(self, button=None):
        self.object_mode = not self.object_mode
        if button is not None:
            button.set_checked(self.object_mode)
        self.set_mode("Chọn đối tượng: kéo để chọn nhiều hình" if self.object_mode else "Sẵn sàng")

    def toggle_snap(self, button=None):
        self.snap_grid = not self.snap_grid
        if button is not None:
            button.set_checked(self.snap_grid)

    def add_connected_shape(self, direction):
        """Quick add: a copy of the selected shape next to it (direction "down", "right", "up", "left"),
        joined to it by an elbow arrow."""
        sel = self.selected_shapes(lines=False)
        if len(sel) != 1:
            self.root.bell()
            return
        i = sel[0]
        self._prepare_shapes()
        obj, box = self.sh.charts[i], self._boxes[i]
        w, h = box[2] - box[0], box[3] - box[1]
        gap = 40 * self.px
        dx, dy = {"down": (0, h + gap), "up": (0, -h - gap), "right": (w + gap, 0), "left": (-w - gap, 0)}[direction]
        new = dict(copy.deepcopy(obj), id=fe.new_picture_id(), text="",
                   name=f"{shapes.name_of(obj.get('kind'))} {sum(1 for o in self.sh.charts if o['type'] == fe.SHAPE) + 1}")
        site = {"down": (2, 0), "up": (0, 2), "right": (3, 1), "left": (1, 3)}[direction]  # top, left, bottom, right
        line = shapes.new_shape("elbow", (0, 0, 0, 0), (0, 0), self.diagram_line, id=fe.new_picture_id(),
                                start={"id": obj.get("id"), "site": site[0]}, end={"id": new["id"], "site": site[1]})
        line["name"] = f"{shapes.name_of('elbow')} {sum(1 for o in self.sh.charts if o['type'] == fe.SHAPE) + 2}"
        if len(shapes.sites(obj, box)) < 4 or obj.get("kind", "").startswith("gate"):
            line["start"] = line["end"] = None

        def change():
            self.sh.charts.append(new)
            self._set_box(len(self.sh.charts) - 1, (box[0] + dx, box[1] + dy, box[2] + dx, box[3] + dy))
            self.sh.charts.append(line)
            self._sync_connectors()
        prev, self._undo_name = self._undo_name, "Thêm hình nối tiếp"
        try:
            done = self._structural(change)
        finally:
            self._undo_name = prev
        if done:
            self.chart_sel = len(self.sh.charts) - 2
            self.redraw()
            self.edit_shape_text(self.chart_sel)

    def new_diagram_sheet(self):
        """A sheet for a diagram: added after the others, with Select Objects and Snap to Grid on."""
        n = 1
        while f"Sơ đồ {n}" in self.wb.sheets:
            n += 1
        name = f"Sơ đồ {n}"
        prev, self._undo_name = self._undo_name, "Trang sơ đồ mới"
        try:
            done = self._structural(lambda: self.wb.add_sheet(name, index=len(self.wb.sheets)))
        finally:
            self._undo_name = prev
        if done:
            self.switch_sheet(name)
            if not self.object_mode:
                self.toggle_object_mode(self._tool_buttons.get("object_mode"))
            if not self.snap_grid:
                self.toggle_snap(self._tool_buttons.get("snap"))

    def _fit_diagram_tab(self):
        """Labels of the Diagram tab's small buttons go when the window is too narrow for them."""
        page = self.ribbon.pages.get("Sơ đồ")
        if page is None:
            return
        avail = self.ribbon.body.winfo_width() - self.ui.px(8)
        if avail < 50:
            return
        if not self._diagram_compact:
            page.update_idletasks()
            self._diagram_full_w = page.winfo_reqwidth()
        compact = self._diagram_full_w > avail
        if compact != self._diagram_compact:
            self._diagram_compact = compact
            for button, text in self._diagram_labels:
                button.set_text("" if compact else text)

    def _shape_tiles(self, group, kinds, cols):
        """Shape tools of the Diagram tab as pictures big enough to tell apart (an in-ribbon gallery)."""
        ui = self.ui
        frame = tk.Frame(group.body, bg=ui.bg)
        frame.pack(side="left", fill="y", padx=ui.px(2), pady=ui.px(1))
        w, h = ui.px(36), ui.px(30)
        for n, kind in enumerate(kinds):
            tile = _Tile(frame, ui, kind, w, h, lambda k=kind: self.diagram_tool(k))
            tile.grid(row=n // cols, column=n % cols, padx=1, pady=1)
            self._tool_buttons[(kind, True)] = tile
            self.commands.append((shapes.name_of(kind), lambda k=kind: self.diagram_tool(k)))
            self.command_icons.setdefault(shapes.name_of(kind), f"shape_{kind}")

    # ------------------------------------------------------------------ galleries and menus

    def shape_gallery(self, button, on_pick, groups=shapes.GROUPS, style=None):
        """Excel's Shapes dropdown: the shapes by group as small pictures."""
        ui = self.ui
        pop = Popup(button.winfo_toplevel(), button.winfo_rootx(), button.winfo_rooty() + button.winfo_height())
        f = pop.inner
        size = ui.px(26)
        per_row = 12

        def pick(kind):
            pop.close()
            on_pick(kind)
        for group in groups:
            kinds = [k for k, v in shapes.KINDS.items() if v[1] == group]
            if not kinds:
                continue
            tk.Label(f, text=group, bg="#f0f0f0", fg="#262626", anchor="w", font=ui.small_font,
                     padx=ui.px(6)).pack(fill="x", pady=(ui.px(2), ui.px(2)))
            grid = tk.Frame(f, bg="#ffffff")
            grid.pack(padx=ui.px(6), pady=(0, ui.px(4)), anchor="w")
            for n, kind in enumerate(kinds):
                cv = tk.Canvas(grid, width=size, height=size, bg="#ffffff", highlightthickness=1,
                               highlightbackground="#ffffff", bd=0, cursor="hand2")
                cv.grid(row=n // per_row, column=n % per_row, padx=1, pady=1)
                k = size / 16
                shape_icon(kind, style)(cv, lambda *c: [v * k for v in c], k, None)
                cv.bind("<Enter>", lambda e, w=cv: w.config(highlightbackground="#f29436"))
                cv.bind("<Leave>", lambda e, w=cv: w.config(highlightbackground="#ffffff"))
                cv.bind("<ButtonRelease-1>", lambda e, kd=kind: pick(kd))
                Tooltip(cv, shapes.name_of(kind), ui.small_font, 500)
        return pop.show()

    def quick_style_gallery(self, button):
        ui = self.ui
        pop = Popup(button.winfo_toplevel(), button.winfo_rootx(), button.winfo_rooty() + button.winfo_height())
        f = pop.inner
        w, h = ui.px(44), ui.px(30)
        tk.Label(f, text="Kiểu hình", bg="#f0f0f0", fg="#262626", anchor="w", font=ui.small_font,
                 padx=ui.px(6)).pack(fill="x", pady=(ui.px(2), ui.px(2)))
        grid = tk.Frame(f, bg="#ffffff")
        grid.pack(padx=ui.px(6), pady=(0, ui.px(6)))

        def pick(style):
            pop.close()
            self.diagram_style.update(fill=style["fill"], line=style["line"],
                                      font=dict(self.diagram_style.get("font") or {}, color=style["font"]["color"]))
            if self.selected_shapes():
                self.apply_quick_style(style)
        for r, row in enumerate(shapes.quick_styles()):
            for c, style in enumerate(row):
                cv = tk.Canvas(grid, width=w, height=h, bg="#ffffff", highlightthickness=1,
                               highlightbackground="#d6d6d6", bd=0, cursor="hand2")
                cv.grid(row=r, column=c, padx=2, pady=2)
                cv.create_rectangle(4, 4, w - 4, h - 4, fill=style["fill"] or "", outline=style["line"] or "",
                                    width=max(1, round(ui.scale)))
                cv.create_text(w / 2, h / 2, text="Abc", fill=style["font"]["color"], font=ui.small_font)
                cv.bind("<Enter>", lambda e, x=cv: x.config(highlightbackground="#f29436"))
                cv.bind("<Leave>", lambda e, x=cv: x.config(highlightbackground="#d6d6d6"))
                cv.bind("<ButtonRelease-1>", lambda e, s=style: pick(s))
        return pop.show()

    def _shape_color(self, button, key):
        """Fill / outline colour dropdown for the selected shapes (and the diagram style for new ones)."""
        def picked(color):
            if key == "fill":
                self.diagram_style["fill"] = color
                self._change_shapes("Màu tô hình", lines=False, fill=color)
            elif key == "line":
                self.diagram_style["line"] = color
                self.diagram_line["line"] = color or "#262626"
                self._change_shapes("Màu đường viền", line=color)
            else:
                self.diagram_style["font"] = dict(self.diagram_style.get("font") or {}, color=color or "#000000")
                self._change_shapes("Màu chữ", lines=False, font={"color": color or "#000000"})

        def more():
            from tkinter import colorchooser
            chosen = colorchooser.askcolor(parent=self.root, title="Màu khác")
            if chosen and chosen[1]:
                picked(chosen[1])
        none = {"fill": "Không tô", "line": "Không viền", "text": "Tự động"}[key]
        color_popup(button, self.ui, picked, none, more)

    def _outline_menu(self, button):
        """Shape Outline: weight, dashes and arrowheads."""
        items = [("Màu đường viền...", lambda: self._shape_color(button, "line")), None]
        for w in (0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6):
            items.append((f"Độ dày {w:g} pt", lambda w=w: self._set_line_w(w)))
        items.append(None)
        for key, name in shapes.DASH_NAMES.items():
            items.append((f"Nét {name.lower()}", lambda k=key: self._change_shapes("Kiểu nét", dash=k)))
        menu = self._make_menu(items)
        menu.tk_popup(button.winfo_rootx(), button.winfo_rooty() + button.winfo_height())
        menu.grab_release()

    def _set_line_w(self, w):
        self.diagram_style["line_w"] = w
        self.diagram_line["line_w"] = w
        self._change_shapes("Độ dày đường", line_w=w)

    def _arrow_menu(self, button):
        sets = [("Không mũi tên", None, None), ("Mũi tên ở cuối", None, "triangle"), ("Mũi tên ở đầu", "triangle", None),
                ("Mũi tên hai đầu", "triangle", "triangle"), ("Mũi hở ở cuối", None, "arrow"),
                ("Mũi nhọn ở cuối", None, "stealth"), ("Tròn ở đầu, mũi tên ở cuối", "oval", "triangle"),
                ("Hình thoi ở cuối", None, "diamond")]

        def set_arrows(head, tail):
            self.diagram_line.update(head=head, tail=tail)
            self._change_shapes("Mũi tên", lines=True, head=head, tail=tail)
        menu = self._make_menu([(name, lambda h=h, t=t: set_arrows(h, t)) for name, h, t in sets])
        menu.tk_popup(button.winfo_rootx(), button.winfo_rooty() + button.winfo_height())
        menu.grab_release()

    def _arrange_menu(self, button, kind):
        if kind == "order":
            items = [("Đưa lên trên cùng", lambda: self.order_objects("front")),
                     ("Đưa lên một lớp", lambda: self.order_objects("forward")),
                     ("Đưa xuống một lớp", lambda: self.order_objects("backward")),
                     ("Đưa xuống dưới cùng", lambda: self.order_objects("back"))]
        elif kind == "align":
            items = [("Căn trái", lambda: self.align_objects("left")), ("Căn giữa", lambda: self.align_objects("center")),
                     ("Căn phải", lambda: self.align_objects("right")), None,
                     ("Căn trên", lambda: self.align_objects("top")), ("Căn giữa dọc", lambda: self.align_objects("middle")),
                     ("Căn dưới", lambda: self.align_objects("bottom")), None,
                     ("Phân phối theo chiều ngang", lambda: self.align_objects("hspace")),
                     ("Phân phối theo chiều dọc", lambda: self.align_objects("vspace"))]
        else:
            items = [("Xoay phải 90°", lambda: self.rotate_shapes(90)), ("Xoay trái 90°", lambda: self.rotate_shapes(-90)),
                     ("Lật ngang", lambda: self.flip_shapes("h")), ("Lật dọc", lambda: self.flip_shapes("v"))]
        menu = self._make_menu(items)
        menu.tk_popup(button.winfo_rootx(), button.winfo_rooty() + button.winfo_height())
        menu.grab_release()

    def _shape_text_menu(self, button, kind):
        if kind == "size":
            items = [(f"{s} pt", lambda s=s: self._shape_font_size(s)) for s in (8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36)]
        else:
            items = [("Trên", lambda: self._change_shapes("Căn chữ", lines=False, valign="top")),
                     ("Giữa", lambda: self._change_shapes("Căn chữ", lines=False, valign="middle")),
                     ("Dưới", lambda: self._change_shapes("Căn chữ", lines=False, valign="bottom"))]
        menu = self._make_menu(items)
        menu.tk_popup(button.winfo_rootx(), button.winfo_rooty() + button.winfo_height())
        menu.grab_release()

    def _shape_font_size(self, size):
        self.diagram_style["font"] = dict(self.diagram_style.get("font") or {}, size=size)
        self._change_shapes("Cỡ chữ", lines=False, font={"size": size})

    def shape_context_menu(self, e):
        """Right-click on a shape."""
        sel = self.selected_shapes()
        line = bool(sel) and all(shapes.is_line(self.sh.charts[i]) for i in sel)
        items = [("Cắt", lambda: self.copy_objects(cut=True), "Ctrl+X"), ("Sao chép", self.copy_objects, "Ctrl+C"),
                 ("Dán", self.paste_objects if self.obj_clip else None, "Ctrl+V"),
                 ("Nhân bản", self.duplicate_objects, "Ctrl+D"), None]
        if not line:
            items += [("Sửa văn bản", self.edit_shape_text, "F2"),
                      ("Đổi hình", [(shapes.name_of(k), lambda k=k: self.change_shape_kind(k))
                                    for k in FLOWCHART_MAIN + ("rect", "roundRect", "ellipse", "diamond", "textBox")]),
                      None]
        else:
            items += [("Kiểu đường nối", [("Đường thẳng", lambda: self.change_shape_kind("line")),
                                          ("Đường gấp khúc", lambda: self.change_shape_kind("elbow")),
                                          ("Đường cong", lambda: self.change_shape_kind("curve"))]), None]
        items += [("Đưa lên trên cùng", lambda: self.order_objects("front")),
                  ("Đưa xuống dưới cùng", lambda: self.order_objects("back")), None,
                  ("Xóa", self.delete_objects, "Delete")]
        menu = self._make_menu(items)
        try:
            menu.tk_popup(e.x_root, e.y_root)
        finally:
            menu.grab_release()

    # ------------------------------------------------------------------ the ribbon

    def _register_shape_icons(self):
        import ribbon
        for kind in shapes.KINDS:
            ribbon.ICONS.setdefault(f"shape_{kind}", shape_icon(kind))
        ribbon.ICONS.setdefault("shape_line_arrow", shape_icon("line", {"tail": "triangle"}))
        ribbon.ICONS.setdefault("shape_line_double", shape_icon("line", {"head": "triangle", "tail": "triangle"}))
        ribbon.ICONS.setdefault("shape_elbow_arrow", shape_icon("elbow", {"tail": "triangle"}))
        ribbon.ICONS.setdefault("shape_curve_arrow", shape_icon("curve", {"tail": "triangle"}))
        ribbon.ICONS.setdefault("shapes", _icon_shapes)
        ribbon.ICONS.setdefault("select_objects", _icon_select)
        ribbon.ICONS.setdefault("snap", _icon_snap)
        ribbon.ICONS.setdefault("diagram_sheet", _icon_diagram_sheet)
        ribbon.ICONS.setdefault("quick_style", _icon_quick_style)
        ribbon.ICONS.setdefault("shape_fill", _icon_shape_fill)
        ribbon.ICONS.setdefault("shape_outline", _icon_shape_outline)
        ribbon.ICONS.setdefault("arrows", shape_icon("line", {"head": "oval", "tail": "triangle"}))
        ribbon.ICONS.setdefault("bring_front", lambda *a: _icon_order(*a, True))
        ribbon.ICONS.setdefault("send_back", lambda *a: _icon_order(*a, False))
        ribbon.ICONS.setdefault("align_objects", _icon_align_objects)
        ribbon.ICONS.setdefault("rotate", _icon_rotate)
        for d in ("down", "right", "up", "left"):
            ribbon.ICONS.setdefault(f"add_{d}", (lambda d: lambda *a: _icon_add(*a, d))(d))

    def _build_diagram_tab(self, page):
        """The Diagram tab: drawing logic diagrams and flowcharts - shapes, connectors that stick to them,
        logic gates, quick add, alignment and styles."""
        self._register_shape_icons()
        rb = self.ribbon
        self._tool_buttons = {}
        self._diagram_compact, self._diagram_full_w = False, 0
        g = rb.group(page, "Công cụ")
        b = self._btn(g, large=True, icon="select_objects", text="Chọn\nđối tượng", label="Chọn đối tượng",
                      tooltip="Chọn đối tượng\nKéo trên trang tính để chọn nhiều hình; nhấp vào ô không chọn ô")
        b.command = lambda b=b: self.toggle_object_mode(b)
        b.set_checked(self.object_mode)
        self._tool_buttons["object_mode"] = b
        col = g.column()
        b = self._btn(g, col, side="top", icon="snap", text="Bắt dính lưới",
                      tooltip="Bắt dính vào lưới ô\nHình và đường nối bám vào cạnh các ô khi kéo")
        b.command = lambda b=b: self.toggle_snap(b)
        b.set_checked(self.snap_grid)
        self._tool_buttons["snap"] = b
        self._btn(g, col, side="top", icon="shape_textBox", text="Văn bản", command=lambda: self.diagram_tool("textBox"),
                  tooltip="Hộp văn bản\nNhãn chữ không viền, ví dụ \"Có\" / \"Không\" cạnh mũi tên")
        self._btn(g, col, side="top", icon="diagram_sheet", text="Trang sơ đồ mới", command=self.new_diagram_sheet,
                  tooltip="Trang sơ đồ mới\nThêm một trang tính riêng để vẽ sơ đồ")

        g = rb.group(page, "Lưu đồ")
        self._shape_tiles(g, FLOWCHART_MAIN, 4)
        col = g.column()
        self._btn(g, col, side="top", icon="shapes", text="Thêm", label="Thêm hình lưu đồ",
                  menu=lambda b: self.shape_gallery(b, self.diagram_tool, ("Lưu đồ", "Hình chữ nhật", "Hình cơ bản",
                                                                             "Mũi tên khối", "Chú thích", "Sao"),
                                                    style={}),
                  tooltip="Các hình khác: mọi hình lưu đồ, hình cơ bản, mũi tên khối, chú thích")

        g = rb.group(page, "Đường nối")
        rows = g.rows(3)
        for n, (key, icon, name, tip) in enumerate((
                ("elbow", "shape_elbow_arrow", "Gấp khúc", "Mũi tên gấp khúc\nNối hai hình bằng các đoạn ngang/dọc, tự đi vòng"),
                ("line", "shape_line_arrow", "Thẳng", "Mũi tên thẳng"),
                ("curve", "shape_curve_arrow", "Cong", "Mũi tên cong"),
                ("line_plain", "shape_line", "Đường", "Đường thẳng không mũi tên"),
                ("line_double", "shape_line_double", "Hai đầu", "Mũi tên hai đầu"))):
            b = self._btn(g, rows[n % 3], icon=icon, text=name, label=f"Đường nối {name.lower()}",
                          command=lambda k=key: self.diagram_tool(k),
                          tooltip=tip + "\nKéo từ một điểm nối (chấm xanh) tới hình khác. Esc để thôi")
            self._tool_buttons[(key, True)] = b
        self._btn(g, rows[2], icon="arrows", text="Mũi tên", menu=lambda b: self._arrow_menu(b),
                  tooltip="Kiểu mũi tên của đường đang chọn")

        g = rb.group(page, "Cổng logic")
        self._shape_tiles(g, GATES, 4)

        g = rb.group(page, "Thêm nối tiếp")
        rows = g.rows(2)
        for n, (d, tip) in enumerate((("up", "lên trên"), ("down", "xuống dưới"), ("left", "sang trái"),
                                      ("right", "sang phải"))):
            self._btn(g, rows[n // 2], icon=f"add_{d}", label=f"Thêm hình {tip}",
                      command=lambda d=d: self.add_connected_shape(d),
                      tooltip=f"Thêm hình {tip}\nThêm một hình giống hình đang chọn, nối sẵn bằng mũi tên")

        g = rb.group(page, "Kiểu hình")
        self._btn(g, large=True, icon="quick_style", text="Kiểu\nnhanh", label="Kiểu nhanh",
                  menu=lambda b: self.quick_style_gallery(b), tooltip="Kiểu nhanh\nMàu tô, viền và chữ có sẵn")
        col = g.column()
        self._btn(g, col, side="top", icon="shape_fill", text="Màu tô", menu=lambda b: self._shape_color(b, "fill"),
                  tooltip="Màu tô hình")
        self._btn(g, col, side="top", icon="shape_outline", text="Đường viền", menu=lambda b: self._outline_menu(b),
                  tooltip="Màu, độ dày và kiểu nét của viền / đường nối")
        self._btn(g, col, side="top", icon="font_color", text="Màu chữ", menu=lambda b: self._shape_color(b, "text"),
                  tooltip="Màu chữ trong hình")

        g = rb.group(page, "Sắp xếp")
        col = g.column()
        self._btn(g, col, side="top", icon="bring_front", text="Đưa lên trước",
                  command=lambda: self.order_objects("front"), menu=lambda b: self._arrange_menu(b, "order"), split=True)
        self._btn(g, col, side="top", icon="send_back", text="Đưa ra sau", command=lambda: self.order_objects("back"),
                  menu=lambda b: self._arrange_menu(b, "order"), split=True)
        self._btn(g, col, side="top", icon="delete_x", text="Xóa", command=self.delete_objects)
        col = g.column()
        self._btn(g, col, side="top", icon="align_objects", text="Căn chỉnh", menu=lambda b: self._arrange_menu(b, "align"),
                  tooltip="Căn chỉnh và phân phối đều các hình đang chọn")
        self._btn(g, col, side="top", icon="rotate", text="Xoay", menu=lambda b: self._arrange_menu(b, "rotate"),
                  tooltip="Xoay hoặc lật hình")
        self._btn(g, col, side="top", icon="valign_center", text="Vị trí chữ",
                  menu=lambda b: self._shape_text_menu(b, "valign"), tooltip="Chữ ở trên / giữa / dưới hình")
        self._diagram_labels = [(w, w.text) for w in _walk(page)
                                if hasattr(w, "set_text") and not getattr(w, "large", True) and w.text]


def _walk(widget):
    for child in widget.winfo_children():
        yield child
        yield from _walk(child)


class _Tile(tk.Canvas):
    """A shape tool in the ribbon: the shape drawn in a small box; checked while it is the tool."""

    def __init__(self, master, ui, kind, w, h, command):
        super().__init__(master, width=w, height=h, bg=ui.bg, highlightthickness=1, highlightbackground=ui.bg,
                         bd=0, cursor="hand2")
        self.ui, self.checked = ui, False
        dw, dh = shapes.default_size(kind)
        f = min((w - 8) / max(dw, 1), (h - 8) / max(dh, 1))
        dw, dh = dw * f, dh * f
        shape = shapes.new_shape(kind, (0, 0, 0, 0), (dw, dh), ICON_STYLE)
        shapes.draw(self, shape, ((w - dw) / 2, (h - dh) / 2, (w + dw) / 2, (h + dh) / 2), max(0.6, ui.scale * 0.8))
        self.bind("<Enter>", lambda e: self.config(highlightbackground="#f29436"))
        self.bind("<Leave>", lambda e: self.config(highlightbackground="#9ebde6" if self.checked else ui.bg))
        self.bind("<ButtonRelease-1>", lambda e: command())
        Tooltip(self, shapes.name_of(kind), ui.small_font, 500)

    def set_checked(self, value):
        self.checked = bool(value)
        self.config(bg="#dbe8f8" if value else self.ui.bg,
                    highlightbackground="#9ebde6" if value else self.ui.bg)


# ---------------------------------------------------------------------- icons of the Diagram tab
# (cv, P, k, color) as ribbon.ICONS: P maps the 16 x 16 icon grid to the canvas

def _icon_shapes(cv, P, k, c):
    cv.create_oval(*P(1, 1, 9, 9), fill="#ffffff", outline="#404040")
    cv.create_rectangle(*P(6, 6, 15, 15), fill="#ffffff", outline="#404040")


def _icon_select(cv, P, k, c):
    cv.create_rectangle(*P(1, 1, 11, 11), outline="#4a7ebb", dash=(2, 2))
    cv.create_polygon(*P(6, 5, 6, 15, 8.5, 12.5, 10.5, 16, 12, 15.2, 10, 11.8, 13.5, 11.5), fill="#ffffff",
                      outline="#262626")


def _icon_snap(cv, P, k, c):
    for v in (2, 6, 10, 14):
        cv.create_line(*P(v, 1, v, 15), fill="#bfbfbf")
        cv.create_line(*P(1, v, 15, v), fill="#bfbfbf")
    cv.create_rectangle(*P(6, 6, 14, 10), fill="#4472c4", outline="#2f528f")


def _icon_diagram_sheet(cv, P, k, c):
    cv.create_rectangle(*P(1, 2, 15, 14), fill="#ffffff", outline="#7f7f7f")
    cv.create_rectangle(*P(3, 4, 8, 7), fill="#ffffff", outline="#404040")
    cv.create_rectangle(*P(8, 9, 13, 12), fill="#ffffff", outline="#404040")
    cv.create_line(*P(5.5, 7, 5.5, 10.5, 8, 10.5), fill="#404040", arrow="last", arrowshape=(3, 3, 1))


def _icon_quick_style(cv, P, k, c):
    cv.create_rectangle(*P(1, 2, 7, 7), fill="#4472c4", outline="#2f528f")
    cv.create_rectangle(*P(9, 2, 15, 7), fill="#ed7d31", outline="#ae5a21")
    cv.create_rectangle(*P(1, 9, 7, 14), fill="#ffffff", outline="#70ad47")
    cv.create_rectangle(*P(9, 9, 15, 14), fill="#fff2cc", outline="#ffc000")


def _icon_shape_fill(cv, P, k, c):
    cv.create_polygon(*P(3, 8, 8, 3, 13, 8, 8, 13), fill="#ffffff", outline="#404040")
    cv.create_rectangle(*P(1, 14, 15, 16), fill="#ffc000", outline="")


def _icon_shape_outline(cv, P, k, c):
    cv.create_rectangle(*P(3, 3, 13, 11), outline="#404040", width=max(1, round(1.5 * k)))
    cv.create_rectangle(*P(1, 14, 15, 16), fill="#4472c4", outline="")


def _icon_order(cv, P, k, c, front):
    back, top = ("#bfbfbf", "#4472c4") if front else ("#4472c4", "#bfbfbf")
    if front:
        cv.create_rectangle(*P(1, 1, 9, 9), fill=back, outline="#7f7f7f")
        cv.create_rectangle(*P(6, 6, 15, 15), fill=top, outline="#2f528f")
    else:
        cv.create_rectangle(*P(6, 6, 15, 15), fill=top, outline="#7f7f7f")
        cv.create_rectangle(*P(1, 1, 9, 9), fill=back, outline="#2f528f")


def _icon_align_objects(cv, P, k, c):
    cv.create_line(*P(2, 1, 2, 15), fill="#262626")
    cv.create_rectangle(*P(3, 3, 13, 6), fill="#4472c4", outline="")
    cv.create_rectangle(*P(3, 9, 9, 12), fill="#4472c4", outline="")


def _icon_rotate(cv, P, k, c):
    cv.create_arc(*P(2, 2, 14, 14), start=60, extent=270, style="arc", outline="#262626", width=max(1, round(1.4 * k)))
    cv.create_polygon(*P(11, 1, 14.5, 4.5, 10, 5.5), fill="#262626", outline="")


def _icon_add(cv, P, k, c, direction):
    box = {"down": (4, 1, 12, 6), "up": (4, 10, 12, 15), "right": (1, 5, 6, 11), "left": (10, 5, 15, 11)}[direction]
    cv.create_rectangle(*P(*box), fill="#ffffff", outline="#404040")
    line = {"down": (8, 6, 8, 14), "up": (8, 10, 8, 2), "right": (6, 8, 14, 8), "left": (10, 8, 2, 8)}[direction]
    cv.create_line(*P(*line), fill="#2f74d0", width=max(1, round(1.4 * k)), arrow="last",
                   arrowshape=(4 * k, 4 * k, 1.5 * k))


def ensure_shape_ids(sheet):
    """Give shapes loaded without one an id (connectors find their shapes by id)."""
    for i, obj in enumerate(sheet.charts):
        if obj.get("type") == fe.SHAPE and not obj.get("id"):
            sheet.charts[i] = dict(obj, id=fe.new_picture_id())

