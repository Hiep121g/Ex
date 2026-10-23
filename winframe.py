"""A real Windows frame behind PyExcel's own title bar (Windows only).

PyExcel draws its title bar itself (titlebar.py), in a window without a system frame (Tk's overrideredirect). Such a
window misses what Windows does for a window's frame: the maximise, restore and minimise animations, Aero Snap
(drag the window to the top of the screen to maximise it, to a side for half of it, with Windows 11's snap bar),
Snap Layouts over the maximise button (Windows 11), Win+arrow keys and the system menu.

Here the window keeps Tk's borderless setup - so Tk keeps laying it out as a window that is all client area - while
Windows is told it has a caption and a sizing frame, and the frame is given no room (WM_NCCALCSIZE): the window
looks as before and Windows handles it as a normal window. Which points are its caption, its maximise button and
its sizing edges is answered in WM_NCHITTEST; the title bar's own windows let those points through to the frame
(HTTRANSPARENT), so pressing there starts Windows' own moving, sizing and snapping.

The window procedures run inside Tk's event loop through ctypes and must not call Tcl through tkinter (tkinter
would lose its record of the thread state). What the title bar has to redraw is asked for with Tcl's "after idle"
on a command registered beforehand, through Tcl's C API.
"""

import ctypes
import tkinter as tk
from ctypes import wintypes

WM_WINDOWPOSCHANGED = 0x0047
WM_NCDESTROY, WM_NCCALCSIZE, WM_NCHITTEST, WM_NCACTIVATE = 0x0082, 0x0083, 0x0084, 0x0086
WM_NCMOUSEMOVE, WM_NCLBUTTONDOWN, WM_NCLBUTTONUP, WM_NCLBUTTONDBLCLK = 0x00A0, 0x00A1, 0x00A2, 0x00A3
WM_ENTERSIZEMOVE, WM_EXITSIZEMOVE, WM_NCMOUSELEAVE = 0x0231, 0x0232, 0x02A2

HTTRANSPARENT, HTCAPTION, HTMAXBUTTON = -1, 2, 9
HTLEFT, HTRIGHT, HTTOP, HTTOPLEFT, HTTOPRIGHT, HTBOTTOM, HTBOTTOMLEFT, HTBOTTOMRIGHT = range(10, 18)
RESIZE = frozenset(range(10, 18))  # the sizing edges and corners

SW_MAXIMIZE, SW_MINIMIZE, SW_RESTORE = 3, 6, 9
GWL_STYLE, GWL_EXSTYLE, GWLP_WNDPROC = -16, -20, -4
WS_POPUP, WS_CAPTION, WS_SYSMENU, WS_THICKFRAME = 0x80000000, 0x00C00000, 0x00080000, 0x00040000
WS_MINIMIZEBOX, WS_MAXIMIZEBOX = 0x00020000, 0x00010000
WS_EX_TOOLWINDOW, WS_EX_APPWINDOW = 0x80, 0x40000
SWP_NOSIZE, SWP_NOMOVE, SWP_NOZORDER, SWP_NOACTIVATE, SWP_FRAMECHANGED = 0x1, 0x2, 0x4, 0x10, 0x20
TME_LEAVE, TME_NONCLIENT = 0x2, 0x10
ABM_GETAUTOHIDEBAREX = 0xB
AUTOHIDE_GAP = 2  # px left free on the edge of an auto-hidden taskbar, so it can still be brought up

LRESULT = ctypes.c_ssize_t
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)


class WINDOWPOS(ctypes.Structure):
    _fields_ = [("hwnd", wintypes.HWND), ("hwndInsertAfter", wintypes.HWND), ("x", ctypes.c_int),
                ("y", ctypes.c_int), ("cx", ctypes.c_int), ("cy", ctypes.c_int), ("flags", wintypes.UINT)]


class NCCALCSIZE_PARAMS(ctypes.Structure):
    _fields_ = [("rgrc", wintypes.RECT * 3), ("lppos", ctypes.POINTER(WINDOWPOS))]


class MONITORINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", wintypes.RECT), ("rcWork", wintypes.RECT),
                ("dwFlags", wintypes.DWORD)]


class TRACKMOUSEEVENT(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD), ("dwFlags", wintypes.DWORD), ("hwndTrack", wintypes.HWND),
                ("dwHoverTime", wintypes.DWORD)]


class APPBARDATA(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD), ("hWnd", wintypes.HWND), ("uCallbackMessage", wintypes.UINT),
                ("uEdge", wintypes.UINT), ("rc", wintypes.RECT), ("lParam", wintypes.LPARAM)]


_api = None


def _user32():
    """user32 with the prototypes used here (a private copy, as in screens.user32)."""
    global _api
    if _api is None:
        u = ctypes.WinDLL("user32")
        HWND, UINT, WPARAM, LPARAM = wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM
        get = getattr(u, "GetWindowLongPtrW", u.GetWindowLongW)  # (32-bit Windows has only the Long ones)
        put = getattr(u, "SetWindowLongPtrW", u.SetWindowLongW)
        get.argtypes, get.restype = [HWND, ctypes.c_int], ctypes.c_void_p
        put.argtypes, put.restype = [HWND, ctypes.c_int, ctypes.c_void_p], ctypes.c_void_p
        u.get_proc, u.set_proc = get, put
        u.GetWindowLongW.argtypes, u.GetWindowLongW.restype = [HWND, ctypes.c_int], ctypes.c_long
        u.SetWindowLongW.argtypes, u.SetWindowLongW.restype = [HWND, ctypes.c_int, ctypes.c_long], ctypes.c_long
        u.CallWindowProcW.argtypes, u.CallWindowProcW.restype = [ctypes.c_void_p, HWND, UINT, WPARAM, LPARAM], LRESULT
        u.SetWindowPos.argtypes = [HWND, HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, UINT]
        u.SetWindowPos.restype = wintypes.BOOL
        u.ShowWindow.argtypes, u.ShowWindow.restype = [HWND, ctypes.c_int], wintypes.BOOL
        u.IsZoomed.argtypes, u.IsZoomed.restype = [HWND], wintypes.BOOL
        u.IsIconic.argtypes, u.IsIconic.restype = [HWND], wintypes.BOOL
        u.GetWindowRect.argtypes, u.GetWindowRect.restype = [HWND, ctypes.POINTER(wintypes.RECT)], wintypes.BOOL
        u.GetClientRect.argtypes, u.GetClientRect.restype = [HWND, ctypes.POINTER(wintypes.RECT)], wintypes.BOOL
        u.MapWindowPoints.argtypes = [HWND, HWND, ctypes.POINTER(wintypes.RECT), UINT]
        u.MapWindowPoints.restype = ctypes.c_int
        u.MonitorFromRect.argtypes = [ctypes.POINTER(wintypes.RECT), wintypes.DWORD]
        u.MonitorFromRect.restype = wintypes.HMONITOR
        u.GetMonitorInfoW.argtypes, u.GetMonitorInfoW.restype = [wintypes.HMONITOR, ctypes.c_void_p], wintypes.BOOL
        u.TrackMouseEvent.argtypes, u.TrackMouseEvent.restype = [ctypes.POINTER(TRACKMOUSEEVENT)], wintypes.BOOL
        u.GetMessageTime.argtypes, u.GetMessageTime.restype = [], ctypes.c_long
        u.GetDoubleClickTime.argtypes, u.GetDoubleClickTime.restype = [], UINT
        u.GetSystemMetrics.argtypes, u.GetSystemMetrics.restype = [ctypes.c_int], ctypes.c_int
        _api = u
    return _api


def _tcl_eval(root):
    """A function running a Tcl script in root's interpreter through Tcl's C API, or None."""
    k = ctypes.WinDLL("kernel32")
    k.GetModuleHandleW.argtypes, k.GetModuleHandleW.restype = [wintypes.LPCWSTR], wintypes.HMODULE
    version = str(tk.TclVersion).replace(".", "")
    for name in (f"tcl{version}t.dll", f"tcl{version}.dll", f"tcl9tcl{version}.dll"):
        handle = k.GetModuleHandleW(name)  # already loaded by tkinter
        if handle:
            tcl = ctypes.CDLL(name, handle=handle)
            run = tcl.Tcl_EvalEx
            run.argtypes, run.restype = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_int, ctypes.c_int], ctypes.c_int
            interp = root.tk.interpaddr()
            return lambda script: run(interp, script, -1, 0x20000)  # TCL_EVAL_GLOBAL
    return None


def _inside(rect, x, y):
    return rect.left <= x < rect.right and rect.top <= y < rect.bottom


class Frame:
    """The system frame of a Tk root window made borderless with overrideredirect (see the module's notes).

    hit(x, y) - set by the title bar - names the part of the window at a screen point (HTCAPTION, HTMAXBUTTON,
    HTLEFT...) or None for the window's contents; it runs in a window procedure (no Tk calls). on_change() is
    called through Tk when the window was maximised or restored, or the maximise button's hover or pressed state
    changed (self.hover, self.pressed)."""

    def __init__(self, root):
        self.root = root
        self.hwnd = int(root.wm_frame(), 16)  # Tk's wrapper: the top-level window Windows knows
        self.hit = self.on_change = None
        self.hover = self.pressed = self.moving = False
        self._procs, self._graveyard = {}, []
        self._rect = wintypes.RECT()
        self._last_press = (0, 0, 0)
        self._tracking = self._notify_pending = False
        self._was_zoomed = None
        self._notify = None
        try:
            run = _tcl_eval(root)
        except (AttributeError, OSError, tk.TclError):
            run = None
        if run:
            command = root.register(self._changed)
            script = f"after idle {command}".encode()
            self._notify = lambda: run(script)

    # -------------------------------------------------------------- setting up

    def install(self):
        """Give the window its frame. Done once it has been mapped (Tk makes its top-level window then)."""
        u = _user32()
        self._subclass(self.hwnd, self._frame_proc)
        style = u.GetWindowLongW(self.hwnd, GWL_STYLE) & 0xFFFFFFFF
        style = (style & ~WS_POPUP) | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX
        u.SetWindowLongW(self.hwnd, GWL_STYLE, ctypes.c_long(style & 0xFFFFFFFF).value)
        ex = u.GetWindowLongW(self.hwnd, GWL_EXSTYLE)
        u.SetWindowLongW(self.hwnd, GWL_EXSTYLE, (ex & ~WS_EX_TOOLWINDOW) | WS_EX_APPWINDOW)  # a taskbar button
        u.SetWindowPos(self.hwnd, None, 0, 0, 0, 0,
                       SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE)
        try:  # Windows 11: round corners like other windows (no such attribute before)
            dwm = ctypes.WinDLL("dwmapi")
            dwm.DwmSetWindowAttribute(wintypes.HWND(self.hwnd), 33, ctypes.byref(ctypes.c_int(2)), 4)
        except (AttributeError, OSError):
            pass

    def let_through(self, widget, codes=None, deep=False):
        """Have widget's window pass the frame's parts (those in `codes`, all when None) to the frame, so a press
        there moves or sizes the window instead of reaching the widget; deep: its descendants too."""
        hwnd = widget.winfo_id()
        if hwnd and hwnd not in self._procs:
            self._subclass(hwnd, self._through_proc(codes))
        if deep:
            for child in widget.winfo_children():
                self.let_through(child, codes, deep)

    def forget_old(self):
        """Drop the procedures of windows destroyed since (they may not be dropped while they run)."""
        del self._graveyard[:]

    def _subclass(self, hwnd, handler):
        u = _user32()
        old = u.get_proc(hwnd, GWLP_WNDPROC)

        def proc(h, msg, wp, lp):
            try:
                result = handler(h, msg, wp, lp, old)
            except Exception:  # (an error must not leave the window without its procedure)
                result = None
            if result is None:
                result = u.CallWindowProcW(old, h, msg, wp, lp)
            if msg == WM_NCDESTROY:
                self._graveyard.append(self._procs.pop(h, None))
            return result
        callback = WNDPROC(proc)
        self._procs[hwnd] = callback
        u.set_proc(hwnd, GWLP_WNDPROC, ctypes.cast(callback, ctypes.c_void_p))

    # -------------------------------------------------------------- state

    def zoomed(self):
        return bool(_user32().IsZoomed(self.hwnd))

    def show(self, command):
        """ShowWindow: SW_MAXIMIZE, SW_RESTORE or SW_MINIMIZE, with Windows' animation."""
        _user32().ShowWindow(self.hwnd, command)

    def toggle(self):
        self.show(SW_RESTORE if self.zoomed() else SW_MAXIMIZE)

    def rect_of(self, hwnd):
        """The screen rectangle of a window (a RECT reused for each call)."""
        _user32().GetWindowRect(hwnd, self._rect)
        return self._rect

    def _changed(self):
        self._notify_pending = False
        if self.on_change:
            self.on_change()

    def _tell(self):
        if self._notify and not self._notify_pending:
            self._notify_pending = True
            self._notify()

    def _set_button(self, hover, pressed):
        if (hover, pressed) != (self.hover, self.pressed):
            self.hover, self.pressed = hover, pressed
            self._tell()

    # -------------------------------------------------------------- window procedures

    def _hit_code(self, lp):
        x, y = ctypes.c_short(lp & 0xFFFF).value, ctypes.c_short((lp >> 16) & 0xFFFF).value
        return self.hit(x, y) if self.hit else None

    def _through_proc(self, codes):
        def handler(h, msg, wp, lp, old):
            if msg == WM_NCHITTEST:
                code = self._hit_code(lp)
                if code is not None and (codes is None or code in codes):
                    return HTTRANSPARENT
            return None
        return handler

    def _frame_proc(self, h, msg, wp, lp, old):
        u = _user32()
        if msg == WM_NCHITTEST:
            return self._hit_code(lp)  # (None: the default, the client area)
        if msg == WM_NCCALCSIZE:
            # no room for a frame: the client area is the whole window - except when maximised, where Windows
            # lets the frame hang over the monitor's edges: then it is the monitor's work area
            rect = NCCALCSIZE_PARAMS.from_address(lp).rgrc[0] if wp else wintypes.RECT.from_address(lp)
            if u.IsZoomed(h):
                self._fit_work_area(rect)
            return 0
        if msg == WM_WINDOWPOSCHANGED:
            pos = WINDOWPOS.from_address(lp)
            zoomed = bool(u.IsZoomed(h))
            if zoomed != self._was_zoomed:
                self._was_zoomed = zoomed
                self._tell()
            if zoomed and not pos.flags & SWP_NOSIZE:
                # Tk takes the window's size for its contents' (it sees no frame): give it the client area's
                client = wintypes.RECT()
                u.GetClientRect(h, client)
                u.MapWindowPoints(h, None, client, 2)
                tk_pos = WINDOWPOS.from_buffer_copy(pos)
                tk_pos.x, tk_pos.y = client.left, client.top
                tk_pos.cx, tk_pos.cy = client.right - client.left, client.bottom - client.top
                return u.CallWindowProcW(old, h, msg, wp, ctypes.addressof(tk_pos))
            return None
        if msg == WM_NCACTIVATE:
            return u.CallWindowProcW(old, h, msg, wp, -1)  # -1: don't paint a frame (there is none to see)
        if msg in (WM_NCLBUTTONDOWN, WM_NCLBUTTONDBLCLK):
            if wp == HTMAXBUTTON:
                self._set_button(True, True)
                return 0
            if wp == HTCAPTION and msg == WM_NCLBUTTONDOWN and self._double_click(lp):
                self.toggle()  # (Tk's window class takes no double-clicks: they come as two presses)
                return 0
            return None
        if msg == WM_NCLBUTTONUP and wp == HTMAXBUTTON:
            if self.pressed:
                self._set_button(True, False)
                self.toggle()
            return 0
        if msg == WM_NCMOUSEMOVE:
            over = wp == HTMAXBUTTON
            self._set_button(over, self.pressed and over)
            if not self._tracking:
                self._tracking = True
                track = TRACKMOUSEEVENT(ctypes.sizeof(TRACKMOUSEEVENT), TME_LEAVE | TME_NONCLIENT, h, 0)
                u.TrackMouseEvent(track)
        elif msg == WM_NCMOUSELEAVE:
            self._tracking = False
            self._set_button(False, False)
        elif msg == WM_ENTERSIZEMOVE:
            self.moving = True
        elif msg == WM_EXITSIZEMOVE:
            self.moving = False
        return None

    def _double_click(self, lp):
        """True when a press on the caption is the second one of a double-click."""
        u = _user32()
        x, y = ctypes.c_short(lp & 0xFFFF).value, ctypes.c_short((lp >> 16) & 0xFFFF).value
        t = u.GetMessageTime() & 0xFFFFFFFF
        t0, x0, y0 = self._last_press
        double = (0 <= t - t0 <= u.GetDoubleClickTime() and abs(x - x0) <= u.GetSystemMetrics(36)  # SM_CXDOUBLECLK
                  and abs(y - y0) <= u.GetSystemMetrics(37))
        self._last_press = (0, 0, 0) if double else (t, x, y)
        return double

    def _fit_work_area(self, rect):
        """Clip the maximised window's client rectangle to its monitor's work area (leaving the edge of an
        auto-hidden taskbar free)."""
        u = _user32()
        info = MONITORINFO()
        info.cbSize = ctypes.sizeof(info)
        if not u.GetMonitorInfoW(u.MonitorFromRect(rect, 2), ctypes.byref(info)):  # 2: nearest
            return
        work = info.rcWork
        rect.left, rect.top = max(rect.left, work.left), max(rect.top, work.top)
        rect.right, rect.bottom = min(rect.right, work.right), min(rect.bottom, work.bottom)
        try:
            shell = ctypes.WinDLL("shell32")
            shell.SHAppBarMessage.argtypes = [wintypes.DWORD, ctypes.POINTER(APPBARDATA)]
            shell.SHAppBarMessage.restype = ctypes.c_size_t
            for edge in range(4):  # ABE_LEFT, ABE_TOP, ABE_RIGHT, ABE_BOTTOM
                bar = APPBARDATA(ctypes.sizeof(APPBARDATA), None, 0, edge, info.rcMonitor, 0)
                if shell.SHAppBarMessage(ABM_GETAUTOHIDEBAREX, ctypes.byref(bar)):
                    if edge == 0:
                        rect.left += AUTOHIDE_GAP
                    elif edge == 1:
                        rect.top += AUTOHIDE_GAP
                    elif edge == 2:
                        rect.right -= AUTOHIDE_GAP
                    else:
                        rect.bottom -= AUTOHIDE_GAP
        except (AttributeError, OSError):
            pass


def install(root):
    """A Frame for Tk's root window, installed; None when Windows refuses."""
    try:
        frame = Frame(root)
        if not frame.hwnd:
            return None
        frame.install()
        return frame
    except (AttributeError, OSError, ValueError, tk.TclError):
        return None
