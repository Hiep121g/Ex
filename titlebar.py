"""PyExcel's title bar, laid out like Excel's: the Quick Access Toolbar on the left, the file name in the
middle, then the user's name, the ribbon display options and the window buttons.

With custom=True (the default on Windows) PyExcel draws the whole window frame itself: the system title
bar is removed and this bar moves the window (drag it; drag to the top of the screen to maximise, to a
side to fill half of it), maximises it to the screen's work area (double-click), minimises and closes
it, and thin edges around the window resize it. Otherwise the system title bar stays and this bar only
holds the toolbar and the buttons next to it.
"""

import getpass
import sys
import tkinter as tk

from ribbon import RIBBON_BG, TAB_BAR_BG, TAB_HOVER_BG, Tooltip
from screens import user32, work_area

BAR_HEIGHT = 30   # px at 100 %, like Excel's
EDGE = 4          # px at 100 %: width of the window edges that resize it
TITLE_FG = "#ffffff"
CLOSE_HOVER, CLOSE_PRESSED = "#e81123", "#f1707a"
CAPTION_PRESSED = "#1a5c38"
SNAP_ZONE = 2     # px from a screen edge where letting go of a dragged window snaps it

if sys.platform == "win32":
    CURSORS = {"n": "size_ns", "s": "size_ns", "e": "size_we", "w": "size_we",
               "nw": "size_nw_se", "se": "size_nw_se", "ne": "size_ne_sw", "sw": "size_ne_sw"}
else:
    CURSORS = {"n": "top_side", "s": "bottom_side", "e": "right_side", "w": "left_side",
               "nw": "top_left_corner", "se": "bottom_right_corner", "ne": "top_right_corner",
               "sw": "bottom_left_corner"}


def user_name():
    """The signed-in user's name as Windows shows it (Excel shows the account here), else the login name."""
    if sys.platform == "win32":
        try:
            import ctypes
            size = ctypes.c_ulong(256)
            buf = ctypes.create_unicode_buffer(size.value)
            if ctypes.WinDLL("secur32").GetUserNameExW(3, buf, ctypes.byref(size)) and buf.value:  # 3: NameDisplay
                return buf.value
        except (AttributeError, OSError):
            pass
    try:
        return getpass.getuser()
    except (ImportError, KeyError, OSError):
        return ""


# ------------------------------------------------------------------ Windows

def show_in_taskbar(root):
    """A borderless (override-redirect) Tk window is a tool window to Windows: no taskbar button and no
    Alt+Tab. Make it an application window again, which can also be minimised from the taskbar."""
    GWL_STYLE, GWL_EXSTYLE = -16, -20
    WS_EX_TOOLWINDOW, WS_EX_APPWINDOW = 0x80, 0x40000
    WS_MINIMIZEBOX, WS_SYSMENU = 0x20000, 0x80000
    u = user32()
    hwnd = u.GetParent(root.winfo_id())
    u.SetWindowLongW(hwnd, GWL_EXSTYLE, (u.GetWindowLongW(hwnd, GWL_EXSTYLE) & ~WS_EX_TOOLWINDOW) | WS_EX_APPWINDOW)
    u.SetWindowLongW(hwnd, GWL_STYLE, u.GetWindowLongW(hwnd, GWL_STYLE) | WS_MINIMIZEBOX | WS_SYSMENU)


# ------------------------------------------------------------------ widgets

class CaptionButton(tk.Canvas):
    """Minimise, maximise/restore or close, drawn like Windows' caption buttons."""

    def __init__(self, master, ui, kind, command, tooltip):
        self.ui, self.kind, self.command = ui, kind, command
        self.hover = self.pressed = False
        super().__init__(master, width=ui.px(46), height=ui.px(BAR_HEIGHT), bg=TAB_BAR_BG, highlightthickness=0,
                         bd=0, takefocus=0)
        self.bind("<Enter>", lambda e: self._set(hover=True))
        self.bind("<Leave>", lambda e: self._set(hover=False, pressed=False))
        self.bind("<ButtonPress-1>", lambda e: self._set(pressed=True))
        self.bind("<ButtonRelease-1>", self._release)
        self.tip = Tooltip(self, tooltip, ui.small_font)
        self.draw()

    def set_kind(self, kind, tooltip):
        if kind != self.kind:
            self.kind, self.tip.text = kind, tooltip
            self.draw()

    def _set(self, **state):
        for key, value in state.items():
            setattr(self, key, value)
        self.draw()

    def _release(self, e):
        run = self.pressed and 0 <= e.x < self.winfo_width() and 0 <= e.y < self.winfo_height()
        self._set(pressed=False)
        if run:
            self.command()

    def draw(self):
        self.delete("all")
        w, h = int(self["width"]), int(self["height"])
        if self.kind == "close":
            bg = CLOSE_PRESSED if self.pressed else CLOSE_HOVER if self.hover else TAB_BAR_BG
        else:
            bg = CAPTION_PRESSED if self.pressed else TAB_HOVER_BG if self.hover else TAB_BAR_BG
        self.config(bg=bg)
        s = self.ui.scale
        d, lw = round(5 * s), max(1, round(s))
        cx, cy = w // 2, h // 2
        if self.kind == "minimize":
            self.create_line(cx - d, cy, cx + d + 1, cy, fill=TITLE_FG, width=lw)
        elif self.kind == "maximize":
            self.create_rectangle(cx - d, cy - d, cx + d, cy + d, outline=TITLE_FG, width=lw)
        elif self.kind == "restore":
            g = round(2 * s)
            self.create_rectangle(cx - d, cy - d + g, cx + d - g, cy + d, outline=TITLE_FG, width=lw)
            self.create_line(cx - d + g, cy - d + g, cx - d + g, cy - d, cx + d, cy - d, cx + d, cy + d - g,
                             cx + d - g, cy + d - g, fill=TITLE_FG, width=lw)
        else:
            self.create_line(cx - d, cy - d, cx + d + 1, cy + d + 1, fill=TITLE_FG, width=lw)
            self.create_line(cx - d, cy + d, cx + d + 1, cy - d - 1, fill=TITLE_FG, width=lw)


class GlyphButton(tk.Canvas):
    """A small white glyph button for the green bar (Customize toolbar, ribbon display options)."""

    def __init__(self, master, ui, glyph, command, tooltip, width=None):
        self.ui, self.glyph, self.command = ui, glyph, command
        self.hover = False
        super().__init__(master, width=width or ui.px(22), height=ui.px(BAR_HEIGHT), bg=TAB_BAR_BG,
                         highlightthickness=0, bd=0, takefocus=0)
        self.bind("<Enter>", lambda e: self._set_hover(True))
        self.bind("<Leave>", lambda e: self._set_hover(False))
        self.bind("<ButtonRelease-1>", lambda e: self.command(self))
        Tooltip(self, tooltip, ui.small_font)
        self.draw()

    def _set_hover(self, value):
        self.hover = value
        self.draw()

    def draw(self):
        self.delete("all")
        w, h, s = int(self["width"]), int(self["height"]), self.ui.scale
        self.config(bg=TAB_HOVER_BG if self.hover else TAB_BAR_BG)
        cx, cy = w // 2, h // 2
        lw = max(1, round(s))
        if self.glyph == "customize":  # a bar over a small triangle
            d = round(3 * s)
            self.create_line(cx - d, cy - d, cx + d + 1, cy - d, fill=TITLE_FG, width=lw)
            self.create_polygon(cx - d, cy - 1, cx + d + 1, cy - 1, cx, cy + d - 1, fill=TITLE_FG, outline="")
        elif self.glyph == "ribbon":  # a window with an arrow pointing up at its top edge
            dx, dy = round(7 * s), round(5 * s)
            self.create_rectangle(cx - dx, cy - dy, cx + dx, cy + dy, outline=TITLE_FG, width=lw)
            self.create_line(cx - dx, cy - dy + round(3 * s), cx + dx, cy - dy + round(3 * s), fill=TITLE_FG,
                             width=lw)
            a = round(2.5 * s)
            self.create_polygon(cx - a, cy + round(2 * s), cx + a + 1, cy + round(2 * s), cx, cy - round(0.5 * s),
                                fill=TITLE_FG, outline="")
        else:  # "more": the three dots of a hidden ribbon
            for i in (-1, 0, 1):
                x = cx + i * round(5 * s)
                r = max(1, round(1.2 * s))
                self.create_oval(x - r, cy - r, x + r, cy + r, fill=TITLE_FG, outline="")


class TitleBar(tk.Frame):
    """The green bar at the top of the window. `qat` and `tools` are frames for the caller's buttons."""

    def __init__(self, root, ui, custom, on_close):
        super().__init__(root, bg=TAB_BAR_BG, height=ui.px(BAR_HEIGHT))
        self.root, self.ui, self.custom, self.on_close = root, ui, custom, on_close
        self.pack_propagate(False)
        self.pack(side="top", fill="x")
        self.maximized = False
        self._normal = None   # geometry to go back to from maximised or snapped
        self._drag = self._resize = self._snap = None  # _snap: geometry of a window snapped to half the screen
        self._geometry_job = self._pending_geometry = None
        self.band = 0         # height of the green tab strip the side edges continue
        self.text = ""

        self.qat = tk.Frame(self, bg=TAB_BAR_BG)
        self.qat.pack(side="left", fill="y", padx=(ui.px(4), 0))
        self.caption = {}
        if custom:
            for kind, command, tip in (("close", self.on_close, "Đóng"),
                                       ("maximize", self.toggle_maximize, "Phóng to"),
                                       ("minimize", self.minimize, "Thu nhỏ")):
                button = self.caption[kind] = CaptionButton(self, ui, kind, command, tip)
                button.pack(side="right")
        self.tools = tk.Frame(self, bg=TAB_BAR_BG)
        self.tools.pack(side="right", fill="y", padx=(0, ui.px(10)))
        self.user = tk.Label(self, text=user_name(), bg=TAB_BAR_BG, fg=TITLE_FG, font=ui.font)
        self.user.pack(side="right", padx=(0, ui.px(14)))
        self.title = tk.Label(self, bg=TAB_BAR_BG, fg=TITLE_FG, font=ui.font)
        for widget in (self, self.qat, self.user):
            widget.bind("<Configure>", lambda e: self._place_title(), add="+")

        self.edges = {}
        if custom:
            root.overrideredirect(True)
            for widget in (self, self.title, self.user):
                widget.bind("<ButtonPress-1>", self._press)
                widget.bind("<B1-Motion>", self._move)
                widget.bind("<ButtonRelease-1>", self._release)
                widget.bind("<Motion>", self._hover)
                widget.bind("<Double-Button-1>", lambda e: self._edge_at(e) or self.toggle_maximize())
                widget.bind("<Button-3>", self._system_menu)
            for side in ("bottom", "left", "right"):
                edge = self.edges[side] = tk.Canvas(root, bg=RIBBON_BG, highlightthickness=0, bd=0,
                                                    width=self._edge_px(), height=self._edge_px())
                edge.bind("<Configure>", lambda e, s=side: self._draw_edge(s))
                edge.bind("<Motion>", lambda e, s=side: e.widget.config(cursor=CURSORS[self._edge_kind(s, e)]))
                edge.bind("<ButtonPress-1>", lambda e, s=side: self._start_resize(e, self._edge_kind(s, e)))
                edge.bind("<B1-Motion>", self._do_resize)
                edge.bind("<ButtonRelease-1>", lambda e: setattr(self, "_resize", None))
            self._show_edges()
            if sys.platform == "win32":
                root.after(10, self._fix_taskbar)

    # -------------------------------------------------------------- title

    def set_title(self, text):
        self.text = text
        self._place_title()

    def _place_title(self):
        """The title in the middle of the window, moved aside or shortened when the toolbar is in the way."""
        if not self.custom:
            return  # the system title bar shows it
        font, pad = self.ui.font, self.ui.px(12)
        left = self.qat.winfo_x() + self.qat.winfo_width() + pad
        right = (self.user.winfo_x() if self.user.winfo_ismapped() else self.tools.winfo_x()) - pad
        text = self.text
        while text and font.measure(text) > right - left:
            text = text[:-2] + "…" if len(text) > 2 else ""
        if not text:
            self.title.place_forget()
            return
        width = font.measure(text)
        x = max(left, min((self.winfo_width() - width) // 2, right - width))
        self.title.config(text=text)
        self.title.place(x=x, rely=0.5, anchor="w")

    # -------------------------------------------------------------- window state

    def start(self, width, height, maximized):
        """Size the window at start-up: `width` x `height` in the middle of the screen, maximised if asked."""
        if not self.custom:
            self.root.geometry(f"{width}x{height}")
            if maximized:
                self.root.state("zoomed")
            return
        x, y, w, h = work_area(self.root)
        self._normal = f"{width}x{height}+{x + max(0, (w - width) // 2)}+{y + max(0, (h - height) // 2)}"
        if maximized:
            self.maximize()
        else:
            self.root.geometry(self._normal)

    def maximize(self, point=None):
        if not self.maximized and self._resize is None:
            if self._normal is None or not self._snapped():
                self._normal = self.root.geometry()
            x, y, w, h = work_area(self.root, point)
            self._set_maximized(True)
            self.root.geometry(f"{w}x{h}+{x}+{y}")

    def restore(self):
        if self.maximized:
            self._set_maximized(False)
            self.root.geometry(self._normal or "")

    def toggle_maximize(self):
        self.restore() if self.maximized else self.maximize()

    def minimize(self):
        if sys.platform == "win32":
            try:
                user32().ShowWindow(user32().GetParent(self.root.winfo_id()), 6)  # SW_MINIMIZE
                return
            except (AttributeError, OSError):
                pass
        # elsewhere a borderless window can't be iconified: let it have its frame while it is iconic
        self.root.overrideredirect(False)
        self.root.iconify()
        self.root.bind("<Map>", self._remap, add="+")

    def _remap(self, e):
        if e.widget is self.root and self.custom:
            self.root.unbind("<Map>")
            self.root.overrideredirect(True)

    def _set_maximized(self, value):
        self.maximized = value
        self.caption["maximize"].set_kind(*(("restore", "Khôi phục") if value else ("maximize", "Phóng to")))
        if value:
            for edge in self.edges.values():
                edge.pack_forget()
        else:
            self._show_edges()

    def _snapped(self):
        """True when the window was put in half of the screen by dragging it to a side."""
        return self._snap is not None and self._snap == self.root.geometry()

    def _fix_taskbar(self):
        if not self.root.winfo_ismapped():  # the window's frame (whose styles change) exists once it is shown
            self.root.after(20, self._fix_taskbar)
            return
        try:
            show_in_taskbar(self.root)
        except (AttributeError, OSError, ValueError):
            return
        # Windows applies the new styles when the window is shown again
        self.root.withdraw()
        self.root.after(10, self._reshow)

    def _reshow(self):
        self.root.deiconify()
        self.root.event_generate("<<WindowShown>>")

    def _system_menu(self, e):
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Khôi phục", command=self.restore, state="normal" if self.maximized else "disabled")
        menu.add_command(label="Thu nhỏ", command=self.minimize)
        menu.add_command(label="Phóng to", command=self.maximize, state="disabled" if self.maximized else "normal")
        menu.add_separator()
        menu.add_command(label="Đóng", accelerator="Alt+F4", command=self.on_close)
        menu.tk_popup(e.x_root, e.y_root)
        menu.grab_release()

    # -------------------------------------------------------------- moving

    def _press(self, e):
        edge = self._edge_at(e)
        if edge:
            self._start_resize(e, edge)
            return
        root = self.root
        self._drag = (e.x_root - root.winfo_rootx(), e.y_root - root.winfo_rooty(), e.x_root, e.y_root, False)

    def _move(self, e):
        if self._resize:
            self._do_resize(e)
            return
        if not self._drag:
            return
        dx, dy, x0, y0, moved = self._drag
        if not moved:
            if abs(e.x_root - x0) + abs(e.y_root - y0) < self.ui.px(5):
                return
            if self.maximized or self._snapped():  # dragging a maximised window brings back its own size
                width = self.root.winfo_width()
                self._set_maximized(False)
                normal = self._normal or self.root.geometry()
                size = normal.split("+")[0]
                w, h = (int(v) for v in size.split("x"))
                dx = round(dx * w / max(1, width))
                self.root.geometry(f"{w}x{h}+{e.x_root - dx}+{e.y_root - dy}")
                self._snap = None
            self._drag = (dx, dy, x0, y0, True)
        self._set_geometry(f"+{e.x_root - dx}+{e.y_root - dy}")

    def _release(self, e):
        if self._resize:
            self._resize = None
            return
        drag, self._drag = self._drag, None
        if not drag or not drag[4]:
            return
        x, y, w, h = work_area(self.root, (e.x_root, e.y_root))
        if e.y_root <= y + SNAP_ZONE:
            self.maximize((e.x_root, e.y_root))
        elif e.x_root <= x + SNAP_ZONE or e.x_root >= x + w - 1 - SNAP_ZONE:
            if not self._snapped():
                self._normal = self.root.geometry()
            half = w // 2
            self._snap = f"{half}x{h}+{x if e.x_root <= x + SNAP_ZONE else x + w - half}+{y}"
            self.root.geometry(self._snap)

    def _set_geometry(self, geometry):
        """Move or resize at most once per idle moment (each step redraws the sheet)."""
        self._pending_geometry = geometry
        if self._geometry_job is None:
            self._geometry_job = self.after_idle(self._apply_geometry)

    def _apply_geometry(self):
        self._geometry_job = None
        self.root.geometry(self._pending_geometry)

    # -------------------------------------------------------------- resizing

    def _edge_px(self):
        return max(EDGE, self.ui.px(EDGE))

    def _show_edges(self):
        if not self.edges:
            return
        after = self
        for side, fill in (("bottom", "x"), ("left", "y"), ("right", "y")):
            self.edges[side].pack(side=side, fill=fill, after=after)
            after = self.edges[side]

    def set_band(self, height):
        """The side edges continue the green tab strip below this bar, `height` pixels tall."""
        if height != self.band:
            self.band = height
            for side in ("left", "right"):
                if side in self.edges:
                    self._draw_edge(side)

    def _draw_edge(self, side):
        edge = self.edges[side]
        edge.delete("all")
        w, h = edge.winfo_width(), edge.winfo_height()
        if side != "bottom" and self.band:
            edge.create_rectangle(0, 0, w, self.band, fill=TAB_BAR_BG, outline="")
        # a one-pixel line of the title bar's green around the window, as Excel's frame has
        x0, y0, x1, y1 = {"left": (0, 0, 0, h), "right": (w - 1, 0, w - 1, h), "bottom": (0, h - 1, w, h - 1)}[side]
        edge.create_line(x0, y0, x1 + (side == "bottom"), y1, fill=TAB_BAR_BG)

    def _edge_kind(self, side, e):
        corner = self.ui.px(16)
        if side == "bottom":
            return "sw" if e.x < corner else "se" if e.x >= e.widget.winfo_width() - corner else "s"
        if e.y >= e.widget.winfo_height() - corner:
            return "sw" if side == "left" else "se"
        return "w" if side == "left" else "e"

    def _edge_at(self, e):
        """'n', 'nw' or 'ne' when the pointer is on the top edge of the window (the top of this bar)."""
        if self.maximized or e.y_root - self.winfo_rooty() >= self._edge_px():
            return None
        x = e.x_root - self.winfo_rootx()
        corner = self.ui.px(16)
        return "nw" if x < corner else "ne" if x >= self.winfo_width() - corner else "n"

    def _hover(self, e):
        if self._drag or self._resize:
            return
        edge = self._edge_at(e)
        e.widget.config(cursor=CURSORS[edge] if edge else "")

    def _start_resize(self, e, edge):
        root = self.root
        self._resize = (edge, e.x_root, e.y_root, root.winfo_rootx(), root.winfo_rooty(),
                        root.winfo_width(), root.winfo_height())

    def _do_resize(self, e):
        if not self._resize:
            return
        edge, x0, y0, x, y, w, h = self._resize
        min_w, min_h = self.root.minsize()
        dx, dy = e.x_root - x0, e.y_root - y0
        nx, ny, nw, nh = x, y, w, h
        if "e" in edge:
            nw = max(min_w, w + dx)
        if "w" in edge:
            nw = max(min_w, w - dx)
            nx = x + w - nw
        if "s" in edge:
            nh = max(min_h, h + dy)
        if "n" in edge:
            nh = max(min_h, h - dy)
            ny = y + h - nh
        self._snap = None
        self._set_geometry(f"{nw}x{nh}+{nx}+{ny}")
