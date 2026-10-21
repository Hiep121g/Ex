"""PivotTables for PyExcel: summarising a block of records, independent of the GUI.

A pivot is a dict stored in the destination Sheet's `pivots` list (see new_pivot) naming its source
block, the fields in each area (filters, rows, columns, values) and display options. compute() reads
the source, keeps the records the item filters allow, aggregates the value fields for every
combination of row and column items (with subtotals and grand totals) and lays the report out the
way Excel does - compact or tabular - as cell values and formats relative to the pivot's anchor.
The app writes that output into the sheet; like Excel, the report is static until refreshed.
"""

import math
import statistics
from collections import defaultdict

import formula_engine as fe

ROW_LABELS, COL_LABELS, GRAND = "Nhãn Hàng", "Nhãn Cột", "Tổng Cuối"
BLANK_ITEM, ALL_ITEMS, MANY_ITEMS = "(trống)", "(Tất cả)", "(Nhiều mục)"
FUNCS = {"sum": "Tổng", "count": "Đếm", "average": "Trung bình", "max": "Lớn nhất", "min": "Nhỏ nhất",
         "product": "Tích", "countnums": "Đếm số", "stdev": "Độ lệch chuẩn", "stdevp": "Độ lệch chuẩn tổng thể",
         "var": "Phương sai", "varp": "Phương sai tổng thể"}
SHOW_AS = {"normal": "Không tính toán", "pct_total": "% của Tổng Cuối", "pct_col": "% của Tổng Cột",
           "pct_row": "% của Tổng Hàng"}
GROUPS = {None: "Không nhóm", "year": "Năm", "quarter": "Quý", "month": "Tháng", "day": "Ngày"}
AREAS = ("filters", "cols", "rows", "values")

HEADER_STYLE = {"bold": True, "bg": "#ddebf7"}
LINE = ["thin", "#9bc2e6"]
GRAND_STYLE = {"bold": True, "bg": "#ddebf7", "bt": ["thin", "#4472c4"]}


class PivotError(Exception):
    """A problem with the pivot's source data, worded for the user."""


def new_pivot(name, sheet, rect, anchor):
    return {"name": name, "source": {"sheet": sheet, "range": tuple(rect)}, "anchor": tuple(anchor),
            "filters": [], "cols": [], "rows": [], "values": [],
            "items": {},       # field -> labels kept (missing: all)
            "sort": {},        # field -> "desc" (labels ascending otherwise)
            "group": {},       # date field -> "year" | "quarter" | "month" | "day"
            "grand_rows": True, "grand_cols": True, "subtotals": True, "layout": "compact",
            "rect": None}      # where the report was last written, relative to nothing: sheet cells


def value_caption(v):
    """'Tổng của Doanh thu' and the like, unless the value field was renamed."""
    return v.get("name") or f"{FUNCS[v['func']]} của {v['field']}"


def default_func(kind):
    return "sum" if kind in ("number",) else "count"


# ------------------------------------------------------------------ source

class Source:
    """The source block: field names from its header row and the records below it."""

    def __init__(self, wb, pivot):
        sheet = wb.resolve_sheet(pivot["source"]["sheet"])
        if sheet is None:
            raise PivotError(f"Không còn trang tính nguồn “{pivot['source']['sheet']}”.")
        r1, c1, r2, c2 = pivot["source"]["range"]
        rows = wb.used_extent(sheet)[0]
        r2 = min(r2, max(r1, rows - 1))
        if r2 <= r1:
            raise PivotError("Vùng nguồn cần một hàng tiêu đề và ít nhất một hàng dữ liệu.")
        self.names, seen = [], set()
        for c in range(c1, c2 + 1):
            name = wb.display_text(sheet, r1, c).strip() or f"Cột{c - c1 + 1}"
            base, n = name, 2
            while name.lower() in seen:
                name, n = f"{base}{n}", n + 1
            seen.add(name.lower())
            self.names.append(name)
        self.records, self.formats, kinds = [], {}, defaultdict(lambda: defaultdict(int))
        for r in range(r1 + 1, r2 + 1):
            rec = []
            for j, c in enumerate(range(c1, c2 + 1)):
                value, numfmt = wb.cell_display(sheet, r, c)
                if isinstance(value, fe.CellError):
                    value = value.code
                rec.append((value, fe.format_value(value, numfmt)))
                if fe.is_number(value):
                    self.formats.setdefault(self.names[j], numfmt)
                    kinds[self.names[j]]["date" if numfmt and fe.is_date_format(numfmt) else "number"] += 1
                elif value not in (None, ""):
                    kinds[self.names[j]]["text"] += 1
            if any(v not in (None, "") for v, _ in rec):
                self.records.append(rec)
        self.kinds = {name: max(kinds[name], key=kinds[name].get) if kinds[name] else "text" for name in self.names}
        self.index = {name: i for i, name in enumerate(self.names)}


def _item(value, text, group):
    """(sort key, label) of a cell as a row / column item, dates grouped by year, quarter, month or day."""
    if value is None or value == "":
        return (9,), BLANK_ITEM
    if group and fe.is_number(value):
        try:
            d = fe.serial_to_datetime(value)
        except (OverflowError, ValueError):
            return fe.sort_key(value), text
        q = (d.month - 1) // 3 + 1
        return {"year": ((0, d.year), str(d.year)), "quarter": ((0, d.year, q), f"Quý {q}/{d.year}"),
                "month": ((0, d.year, d.month), f"Tháng {d.month:02d}/{d.year}"),
                "day": ((0, d.year, d.month, d.day), f"{d.day:02d}/{d.month:02d}/{d.year}")}[group]
    return fe.sort_key(value), text


def field_items(source, pivot, field):
    """Labels of a field's items in display order (for the item filter checklist)."""
    i, group = source.index[field], pivot["group"].get(field)
    keys = {}
    for rec in source.records:
        key, label = _item(*rec[i], group)
        keys.setdefault(label, key)
    return sorted(keys, key=keys.get, reverse=pivot["sort"].get(field) == "desc")


# ------------------------------------------------------------------ aggregation

def aggregate(func, values):
    """Excel's summary functions over a cell's records; None when there is nothing to show."""
    if not values:
        return None
    nums = [float(v) for v in values if fe.is_number(v)]
    if func == "count":
        return float(sum(1 for v in values if v not in (None, "")))
    if func == "countnums":
        return float(len(nums))
    if func == "sum":
        return math.fsum(nums)
    if func == "product":
        return math.prod(nums) if nums else 0.0
    if func in ("max", "min"):
        return (max if func == "max" else min)(nums) if nums else 0.0
    if func == "average":
        return math.fsum(nums) / len(nums) if nums else None
    need = 2 if func in ("stdev", "var") else 1
    if len(nums) < need:
        return None
    return {"stdev": statistics.stdev, "stdevp": statistics.pstdev, "var": statistics.variance,
            "varp": statistics.pvariance}[func](nums)


def _order(paths, keys, desc):
    """Every prefix of the label tuples `paths`, parents before their children, each level sorted by
    the items' keys (descending where desc[level])."""
    out = []

    def walk(prefix, members, level):
        if level == len(desc):
            return
        groups = defaultdict(list)
        for p in members:
            groups[p[level]].append(p)
        for label in sorted(groups, key=lambda lb: keys[level][lb], reverse=desc[level]):
            out.append(prefix + (label,))
            walk(prefix + (label,), groups[label], level + 1)
    walk((), paths, 0)
    return out


class Report:
    """compute() result: cells {(dr, dc): value} and formats {(dr, dc): fmt} relative to the anchor,
    the report's size, and for each value cell the (row items, column items) it sums up."""

    def __init__(self):
        self.cells, self.formats, self.sources = {}, {}, {}
        self.height = self.width = 0

    def put(self, r, c, value, fmt=None):
        if value is not None and value != "":
            self.cells[(r, c)] = value
        if fmt:
            self.formats[(r, c)] = dict(self.formats.get((r, c), {}), **fmt)
        self.height, self.width = max(self.height, r + 1), max(self.width, c + 1)


def compute(wb, pivot, source=None):
    """Lay out the pivot report (see Report). Raises PivotError when the source cannot be read."""
    source = source or Source(wb, pivot)
    names = source.index
    for area in AREAS:  # fields renamed or removed in the source drop out
        pivot[area] = [f for f in pivot[area] if (f["field"] if area == "values" else f) in names]
    report = Report()
    rows, cols, values = pivot["rows"], pivot["cols"], pivot["values"]
    if not (rows or cols or values or pivot["filters"]):
        return _placeholder(report, pivot)
    items = {f: set(v) for f, v in pivot["items"].items() if f in names}

    def labels(rec, fields):
        return tuple(_item(*rec[names[f]], pivot["group"].get(f))[1] for f in fields)

    keys = defaultdict(dict)  # field -> label -> sort key
    data = defaultdict(list)  # (row prefix, column prefix, value index) -> cell values
    paths_r, paths_c = set(), set()
    for rec in source.records:
        if any(_item(*rec[names[f]], pivot["group"].get(f))[1] not in allowed for f, allowed in items.items()):
            continue
        for f in rows + cols:
            key, label = _item(*rec[names[f]], pivot["group"].get(f))
            keys[f].setdefault(label, key)
        rl, cl = labels(rec, rows), labels(rec, cols)
        paths_r.add(rl)
        paths_c.add(cl)
        for i in range(len(rows) + 1):
            for j in range(len(cols) + 1):
                for vi, v in enumerate(values):
                    data[(rl[:i], cl[:j], vi)].append(rec[names[v["field"]]][0])

    def agg(rp, cp, vi):
        return aggregate(values[vi]["func"], data.get((rp, cp, vi)))

    def shown(rp, cp, vi):
        x = agg(rp, cp, vi)
        mode = values[vi].get("show", "normal")
        if x is None or mode == "normal":
            return x
        base = {"pct_total": agg((), (), vi), "pct_col": agg((), cp, vi), "pct_row": agg(rp, (), vi)}[mode]
        return x / base if base else None

    def numfmt(v):
        if v.get("num"):
            return v["num"]
        if v.get("show", "normal") != "normal":
            return "0.00%"
        if v["func"] in ("count", "countnums"):
            return None
        fmt = source.formats.get(v["field"])
        if fmt and fe.is_date_format(fmt) and v["func"] not in ("max", "min"):
            return None  # a sum of dates is not a date
        return fmt

    # report filters, above the table
    top = 0
    for f in pivot["filters"]:
        allowed = items.get(f)
        shown_items = ALL_ITEMS if allowed is None else next(iter(allowed)) if len(allowed) == 1 else MANY_ITEMS
        report.put(top, 0, f, {"bold": True})
        report.put(top, 1, shown_items, {"bg": "#f2f2f2"})
        top += 1
    if pivot["filters"]:
        top += 1

    def descending(fields):
        return [pivot["sort"].get(f) == "desc" for f in fields]

    row_order = _order(paths_r, [keys[f] for f in rows], descending(rows)) if rows else []
    col_order = _order(paths_c, [keys[f] for f in cols], descending(cols)) if cols else []

    # column entries, left to right: (column items, value index, "item" | "subtotal" | "grand")
    multi = len(values) > 1
    entries = []
    if cols:
        for idx, cp in enumerate(col_order):
            if len(cp) < len(cols):
                continue
            entries += [(cp, vi, "item") for vi in range(len(values))] or [(cp, None, "item")]
            nxt = next((p for p in col_order[idx + 1:] if len(p) == len(cols)), ())
            for level in range(len(cp) - 1, 0, -1):  # outer groups ending here get their subtotal
                if pivot["subtotals"] and nxt[:level] != cp[:level]:
                    entries += [(cp[:level], vi, "subtotal") for vi in range(len(values))]
        if pivot["grand_cols"] and values:
            entries += [((), vi, "grand") for vi in range(len(values))]
    else:
        entries = [((), vi, "item") for vi in range(len(values))]

    tabular = pivot["layout"] == "tabular" and rows
    label_cols = len(rows) if tabular else 1
    nhead = (1 + len(cols) + (1 if multi else 0)) if cols else 1
    last = top + nhead - 1
    width = label_cols + len(entries)
    for r in range(top, last + 1):
        for c in range(width):
            report.put(r, c, None, HEADER_STYLE)
    for c in range(width):
        report.put(last, c, None, {"bb": LINE})
    if cols:
        report.put(top, 0, value_caption(values[0]) if len(values) == 1 else None)
        report.put(top, label_cols, COL_LABELS)
    if tabular:
        for j, f in enumerate(rows):
            report.put(last, j, f)
    else:
        report.put(last, 0, ROW_LABELS if rows else (None if cols else ""))
    # column item headers
    prev = None
    for k, (cp, vi, kind) in enumerate(entries):
        c = label_cols + k
        if kind == "grand":
            if prev is None or prev[2] != "grand":
                report.put(top + 1, c, GRAND)
            if multi:
                report.put(last, c, value_caption(values[vi]))
        elif kind == "subtotal":
            report.put(top + len(cp), c, f"{cp[-1]} Tổng")
            if multi:
                report.put(last, c, value_caption(values[vi]))
        else:
            for level in range(len(cp)):
                if prev is None or prev[2] != "item" or prev[0][:level + 1] != cp[:level + 1]:
                    report.put(top + 1 + level, c, cp[level])
            if multi or not cols:
                report.put(last, c, value_caption(values[vi]) if vi is not None else None)
        prev = (cp, vi, kind)

    # data lines
    r = last + 1
    fmts = [numfmt(v) for v in values]

    def put_values(r, rp, style=None):
        for k, (cp, vi, _) in enumerate(entries):
            if vi is None:
                continue
            x = shown(rp, cp, vi)
            report.put(r, label_cols + k, x, dict(style or {}, **({"num": fmts[vi]} if fmts[vi] else {})))
            report.sources[(r, label_cols + k)] = (rp, cp)

    if not rows:
        report.put(r, 0, value_caption(values[0]) if len(values) == 1 else "Tổng" if values else "")
        put_values(r, ())
        r += 1
    elif not tabular:
        for rp in row_order:
            parent = len(rp) < len(rows)
            report.put(r, 0, rp[-1], {"indent": len(rp) - 1, "bold": True} if parent else
                       ({"indent": len(rp) - 1} if len(rp) > 1 else None))
            if not parent or pivot["subtotals"]:
                put_values(r, rp, {"bold": True} if parent else None)
            r += 1
    else:
        leaves = [rp for rp in row_order if len(rp) == len(rows)]
        for idx, rp in enumerate(leaves):
            prev_leaf = leaves[idx - 1] if idx else None
            for level, label in enumerate(rp):  # an outer label shows on its group's first line only
                if prev_leaf is None or prev_leaf[:level + 1] != rp[:level + 1]:
                    report.put(r, level, label)
            put_values(r, rp)
            r += 1
            nxt = leaves[idx + 1] if idx + 1 < len(leaves) else ()
            for level in range(len(rp) - 1, 0, -1):  # groups ending here get a subtotal line
                if pivot["subtotals"] and nxt[:level] != rp[:level]:
                    report.put(r, level - 1, f"{rp[level - 1]} Tổng")
                    put_values(r, rp[:level], {"bold": True, "bg": "#f2f2f2"})
                    for c in range(level - 1, label_cols):
                        report.put(r, c, None, {"bold": True, "bg": "#f2f2f2"})
                    r += 1
    if rows and pivot["grand_rows"]:
        report.put(r, 0, GRAND)
        put_values(r, ())
        for c in range(width):
            report.put(r, c, None, GRAND_STYLE)
        r += 1
    return report


def _placeholder(report, pivot):
    """What an empty pivot shows: its name and a hint, in a light box (Excel's drop-zone picture)."""
    for r in range(0, 18):
        for c in range(3):
            report.put(r, c, None, {"bg": "#f7f7f7", **({"bt": LINE} if r == 0 else {}),
                                    **({"bb": LINE} if r == 17 else {}), **({"bl": LINE} if c == 0 else {}),
                                    **({"br": LINE} if c == 2 else {})})
    report.put(1, 0, pivot["name"], {"bold": True, "fg": "#2b579a"})
    report.put(3, 0, "Để tạo báo cáo, hãy chọn")
    report.put(4, 0, "các trường trong ngăn")
    report.put(5, 0, "Các trường PivotTable.")
    return report


def drill_down(wb, pivot, row_items, col_items, source=None):
    """(field names, records) behind one value cell: the source rows with those row and column items
    that the item filters let through (Excel's "Show Details")."""
    source = source or Source(wb, pivot)
    names = source.index
    items = {f: set(v) for f, v in pivot["items"].items() if f in names}
    out = []
    for rec in source.records:
        lab = {f: _item(*rec[names[f]], pivot["group"].get(f))[1] for f in set(items) | set(pivot["rows"]) |
               set(pivot["cols"])}
        if any(lab[f] not in allowed for f, allowed in items.items()):
            continue
        if any(lab[f] != item for f, item in zip(pivot["rows"], row_items)):
            continue
        if any(lab[f] != item for f, item in zip(pivot["cols"], col_items)):
            continue
        out.append(rec)
    return source.names, out


# ------------------------------------------------------------------ writing the report

def report_rect(pivot, report):
    r, c = pivot["anchor"]
    rect = (r, c, r + max(report.height, 1) - 1, c + max(report.width, 1) - 1)
    if rect[2] >= fe.MAX_ROWS or rect[3] >= fe.MAX_COLS:
        raise PivotError("Báo cáo PivotTable vượt ra ngoài trang tính; hãy đặt nó ở chỗ khác.")
    return rect


def overlaps(wb, sheet, pivot, report):
    """Whether the report would cover cells holding something that is not its own previous output."""
    r1, c1, r2, c2 = report_rect(pivot, report)
    old = pivot["rect"] or (0, 0, -1, -1)
    return any(r1 <= r <= r2 and c1 <= c <= c2 and not (old[0] <= r <= old[2] and old[1] <= c <= old[3])
               for r, c in wb.sheets[sheet].cells)


def place(wb, sheet, pivot, report):
    """Replace the pivot's previous output on `sheet` by `report` (values, never formulas)."""
    sh = wb.sheets[sheet]
    rect = report_rect(pivot, report)
    if pivot["rect"]:
        a1, b1, a2, b2 = pivot["rect"]
        for key in [k for k in set(sh.cells) | set(sh.formats) if a1 <= k[0] <= a2 and b1 <= k[1] <= b2]:
            sh.cells.pop(key, None)
            sh.formats.pop(key, None)
    r0, c0 = pivot["anchor"]
    for (dr, dc), value in report.cells.items():
        sh.cells[(r0 + dr, c0 + dc)] = fe.value_to_raw(value) if not fe.is_number(value) else fe.number_text(value)
    for (dr, dc), fmt in report.formats.items():
        sh.formats[(r0 + dr, c0 + dc)] = dict(fmt)
    pivot["rect"] = rect
    wb._ast.clear()
    wb.invalidate()
    return rect


def refresh(wb, sheet, pivot):
    """Recompute and rewrite a pivot; returns the Report (raises PivotError)."""
    report = compute(wb, pivot)
    place(wb, sheet, pivot, report)
    return report


def pivot_at(wb, sheet, r, c):
    """The pivot whose report covers (r, c) on `sheet`, or None."""
    for p in wb.sheets[sheet].pivots:
        if p["rect"] and p["rect"][0] <= r <= p["rect"][2] and p["rect"][1] <= c <= p["rect"][3]:
            return p
    return None
