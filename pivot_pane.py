"""The "PivotTable Fields" pane shown beside the grid while a cell of a pivot report is selected.

It lists the source fields with check boxes and the four areas - Filters, Columns, Rows and Values -
that fields are dragged between, as in Excel. The pane edits a copy of the pivot's areas and hands
every change to the app, which recomputes the report.
"""

import tkinter as tk
from tkinter import ttk

import pivot as pv
from ribbon import TEXT_FG

PANE_BG = "#ffffff"
DROP_BG = "#e5f1fb"
AREA_TITLES = {"filters": "Bộ lọc", "cols": "Cột", "rows": "Hàng", "values": "Σ Giá trị"}
AREA_ICONS = {"filters": "filter", "cols": "insert_col", "rows": "insert_row", "values": "autosum"}


class PivotPane(tk.Frame):
    """on_change(areas) receives {"filters": [...], "cols": [...], "rows": [...], "values": [...]} after
    each edit; on_action(name, *args) gets "close", "refresh", "options", "delete", ("field", area,
    index) for a row / column / filter field's settings and ("value", index) for a value field's."""

    def __init__(self, master, ui, icons, on_change, on_action):
        super().__init__(master, bg=PANE_BG, highlightthickness=1, highlightbackground="#d0d0d0")
        self.ui, self.icons, self.on_change, self.on_action = ui, icons, on_change, on_action
        self.fields, self.kinds, self.areas = [], {}, {area: [] for area in pv.AREAS}
        self.drag = None
        pad = self.pad = ui.px(8)
        head = tk.Frame(self, bg=PANE_BG)
        head.pack(fill="x", padx=pad, pady=(pad, 2))
        tk.Label(head, text="Các trường PivotTable", font=(ui.family, -ui.px(15)), bg=PANE_BG,
                 fg="#217346").pack(side="left")
        close = tk.Label(head, text="✕", bg=PANE_BG, fg="#606060", cursor="hand2", font=ui.font)
        close.pack(side="right")
        close.bind("<ButtonRelease-1>", lambda e: on_action("close"))
        self._hint("Chọn trường để thêm vào báo cáo:")
        self.search = tk.StringVar()
        self.search.trace_add("write", lambda *a: self._fill_fields())
        ttk.Entry(self, textvariable=self.search, font=ui.font).pack(fill="x", padx=pad, pady=2)
        box = tk.Frame(self, bg=PANE_BG, highlightthickness=1, highlightbackground="#d0d0d0")
        box.pack(fill="both", expand=True, padx=pad)
        self.field_canvas = tk.Canvas(box, bg=PANE_BG, highlightthickness=0, width=ui.px(250), height=ui.px(150))
        bar = ttk.Scrollbar(box, orient="vertical", command=self.field_canvas.yview)
        self.field_canvas.config(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        self.field_canvas.pack(side="left", fill="both", expand=True)
        self.field_frame = tk.Frame(self.field_canvas, bg=PANE_BG)
        window = self.field_canvas.create_window(0, 0, window=self.field_frame, anchor="nw")
        self.field_canvas.bind("<Configure>", lambda e: self.field_canvas.itemconfigure(window, width=e.width))
        self.field_frame.bind("<Configure>",
                              lambda e: self.field_canvas.config(scrollregion=self.field_canvas.bbox("all")))
        self._hint("Kéo trường vào các vùng bên dưới:")
        grid = tk.Frame(self, bg=PANE_BG)
        grid.pack(fill="x", padx=pad)
        self.lists, self.cells = {}, {}
        for i, area in enumerate(("filters", "cols", "rows", "values")):
            cell = tk.Frame(grid, bg=PANE_BG)
            cell.grid(row=i // 2, column=i % 2, sticky="nsew", padx=(0, 3) if i % 2 == 0 else (3, 0), pady=3)
            image = icons.get(AREA_ICONS[area]) if icons is not None else None
            tk.Label(cell, text=" " + AREA_TITLES[area], bg=PANE_BG, anchor="w", font=ui.small_font, fg=TEXT_FG,
                     **({"image": image, "compound": "left"} if image is not None else {})).pack(fill="x")
            lb = tk.Listbox(cell, height=5, width=14, activestyle="none", exportselection=False, font=ui.font,
                            relief="solid", bd=1, highlightthickness=0, selectbackground="#cce4f7",
                            selectforeground=TEXT_FG)
            lb.pack(fill="both", expand=True)
            lb.bind("<ButtonPress-1>", lambda e, a=area: self._press(e, area=a))
            lb.bind("<B1-Motion>", self._motion)
            lb.bind("<ButtonRelease-1>", self._release)
            self.lists[area], self.cells[area] = lb, cell
        grid.columnconfigure(0, weight=1, uniform="area")
        grid.columnconfigure(1, weight=1, uniform="area")
        buttons = tk.Frame(self, bg=PANE_BG)
        buttons.pack(fill="x", padx=pad, pady=pad)
        for text, action in (("Làm mới", "refresh"), ("Tùy chọn...", "options"), ("Xóa", "delete")):
            ttk.Button(buttons, text=text, command=lambda a=action: on_action(a)).pack(side="left", padx=(0, 4))

    def _hint(self, text):
        tk.Label(self, text=text, bg=PANE_BG, fg="#505050", font=self.ui.small_font, anchor="w").pack(
            fill="x", padx=self.pad, pady=(6, 0))

    # ------------------------------------------------------------------ content

    def show(self, pivot, fields, kinds):
        """Show `pivot`'s areas for a source with these field names and kinds ("number", "date", "text")."""
        self.fields, self.kinds = fields, kinds
        self.areas = {area: list(pivot[area]) if area != "values" else [dict(v) for v in pivot[area]]
                      for area in pv.AREAS}
        self._fill_fields()
        for area, lb in self.lists.items():
            lb.delete(0, "end")
            for item in self.areas[area]:
                lb.insert("end", pv.value_caption(item) if area == "values" else item)

    def _used(self, field):
        return field in self.areas["filters"] + self.areas["cols"] + self.areas["rows"] or any(
            v["field"] == field for v in self.areas["values"])

    def _fill_fields(self):
        for w in self.field_frame.winfo_children():
            w.destroy()
        query = self.search.get().strip().lower()
        for field in self.fields:
            if query and query not in field.lower():
                continue
            var = tk.BooleanVar(value=self._used(field))
            box = tk.Checkbutton(self.field_frame, text=field, variable=var, bg=PANE_BG, anchor="w", font=self.ui.font,
                                 activebackground=PANE_BG, command=lambda f=field, v=var: self._toggle(f, v.get()))
            box.var = var
            box.pack(fill="x", anchor="w")
            box.bind("<ButtonPress-1>", lambda e, f=field: self._press(e, field=f))
            box.bind("<B1-Motion>", self._motion)
            box.bind("<ButtonRelease-1>", self._release)

    def _changed(self):
        self.on_change({area: list(items) for area, items in self.areas.items()})

    def _toggle(self, field, on):
        """Ticking a field adds it where Excel would (numbers to Values, the rest to Rows); clearing
        the tick takes it out of every area."""
        if on:
            if self.kinds.get(field) == "number":
                self.areas["values"].append({"field": field, "func": "sum", "show": "normal"})
            else:
                self.areas["rows"].append(field)
        else:
            for area in ("filters", "cols", "rows"):
                self.areas[area] = [f for f in self.areas[area] if f != field]
            self.areas["values"] = [v for v in self.areas["values"] if v["field"] != field]
        self._changed()

    def move(self, field, src, src_index, dst, dst_index):
        """Move a field from area `src` (None: the field list) to position dst_index of area `dst`
        (None: out of the report)."""
        a = self.areas
        value = None
        if src is not None:
            item = a[src].pop(src_index)
            value = item if src == "values" else None
            if src == dst and src_index < dst_index:
                dst_index -= 1
        if dst == "values":
            a["values"].insert(dst_index, value or {"field": field, "func": pv.default_func(self.kinds.get(field)),
                                                    "show": "normal"})
        elif dst is not None:
            for area in ("filters", "cols", "rows"):  # a field sits in one of these at most
                if field in a[area]:
                    i = a[area].index(field)
                    a[area].remove(field)
                    if area == dst and i < dst_index:
                        dst_index -= 1
            a[dst].insert(dst_index, field)
        self._changed()

    # ------------------------------------------------------------------ dragging

    def _press(self, e, area=None, field=None):
        if area is not None:
            lb = self.lists[area]
            index = lb.nearest(e.y)
            box = lb.bbox(index) if lb.size() else None
            if box is None or not box[1] <= e.y <= box[1] + box[3]:
                self.drag = None
                return "break"
            item = self.areas[area][index]
            field = item["field"] if area == "values" else item
            lb.selection_clear(0, "end")
            lb.selection_set(index)
        else:
            index = None
        self.drag = {"field": field, "area": area, "index": index, "start": (e.x_root, e.y_root), "moved": False}
        return "break" if area is not None else None

    def _target(self, x, y):
        for area, cell in self.cells.items():
            if cell.winfo_rootx() <= x < cell.winfo_rootx() + cell.winfo_width() and \
                    cell.winfo_rooty() <= y < cell.winfo_rooty() + cell.winfo_height():
                return area
        return None

    def _motion(self, e):
        d = self.drag
        if d is None:
            return None
        if not d["moved"] and abs(e.x_root - d["start"][0]) + abs(e.y_root - d["start"][1]) > 5:
            d["moved"] = True
            self.winfo_toplevel().config(cursor="hand2")
        if d["moved"]:
            target = self._target(e.x_root, e.y_root)
            for area, lb in self.lists.items():
                lb.config(bg=DROP_BG if area == target else "#ffffff")
        return "break"

    def _release(self, e):
        d, self.drag = self.drag, None
        self.winfo_toplevel().config(cursor="")
        for lb in self.lists.values():
            lb.config(bg="#ffffff")
        if d is None:
            return None
        if not d["moved"]:
            if d["area"] is not None:
                self._item_menu(d["area"], d["index"], e.x_root, e.y_root)
                return "break"
            return None  # a plain click on a field's check box ticks it
        target = self._target(e.x_root, e.y_root)
        if target is None:
            if d["area"] is not None:  # dragged out of the areas: removed, as in Excel
                self.move(d["field"], d["area"], d["index"], None, 0)
            return "break"
        lb = self.lists[target]
        y = e.y_root - lb.winfo_rooty()
        index = lb.nearest(y) if lb.size() else 0
        box = lb.bbox(index) if lb.size() else None
        if box is not None and y > box[1] + box[3] / 2:
            index += 1
        if target == d["area"] and index in (d["index"], d["index"] + 1):
            return "break"
        self.move(d["field"], d["area"], d["index"], target, index)
        return "break"

    def _item_menu(self, area, index, x, y):
        menu = tk.Menu(self, tearoff=0)
        items = self.areas[area]
        field = items[index]["field"] if area == "values" else items[index]

        def add(label, command, icon=None, enabled=True):
            image = self.icons.get(icon) if self.icons is not None and icon else None
            menu.add_command(label=label, command=command, state="normal" if enabled else "disabled",
                             **({"image": image, "compound": "left"} if image is not None else {}))

        add("Di chuyển lên", lambda: self.move(field, area, index, area, index - 1), "fill_up", index > 0)
        add("Di chuyển xuống", lambda: self.move(field, area, index, area, index + 2), "fill_down",
            index < len(items) - 1)
        add("Chuyển lên đầu", lambda: self.move(field, area, index, area, 0), None, index > 0)
        add("Chuyển xuống cuối", lambda: self.move(field, area, index, area, len(items)), None,
            index < len(items) - 1)
        menu.add_separator()
        for dst in ("filters", "rows", "cols", "values"):
            if dst != area:
                add(f"Chuyển tới {AREA_TITLES[dst]}",
                    lambda d=dst: self.move(field, area, index, d, len(self.areas[d])), AREA_ICONS[dst])
        menu.add_separator()
        add("Xóa trường", lambda: self.move(field, area, index, None, 0), "delete_x")
        menu.add_separator()
        if area == "values":
            add("Cài đặt trường giá trị...", lambda: self.on_action("value", index), "autosum")
        else:
            add("Cài đặt trường...", lambda: self.on_action("field", area, index), "format_cells")
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()
