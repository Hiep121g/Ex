"""Where PyExcel's windows may go with several monitors.

Tk only knows the primary screen's size on Windows (winfo_screenwidth), so popups and dialogs that were
kept "on screen" with it jumped to the primary monitor when PyExcel was on another one. These functions
give the work area (the monitor less its taskbar) of the monitor holding a point or a window instead.
"""

import sys

_user32 = None


def user32():
    """user32 with the prototypes PyExcel uses (a private copy: other modules declare their own)."""
    global _user32
    if _user32 is None:
        import ctypes
        from ctypes import wintypes
        u = ctypes.WinDLL("user32")
        u.GetParent.argtypes, u.GetParent.restype = [wintypes.HWND], wintypes.HWND
        u.GetWindowLongW.argtypes, u.GetWindowLongW.restype = [wintypes.HWND, ctypes.c_int], ctypes.c_long
        u.SetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_long]
        u.SetWindowLongW.restype = ctypes.c_long
        u.ShowWindow.argtypes, u.ShowWindow.restype = [wintypes.HWND, ctypes.c_int], wintypes.BOOL
        u.MonitorFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
        u.MonitorFromWindow.restype = wintypes.HMONITOR
        u.MonitorFromPoint.argtypes = [wintypes.POINT, wintypes.DWORD]
        u.MonitorFromPoint.restype = wintypes.HMONITOR
        u.GetMonitorInfoW.argtypes, u.GetMonitorInfoW.restype = [wintypes.HMONITOR, ctypes.c_void_p], wintypes.BOOL
        _user32 = u
    return _user32


def _monitor_work_area(monitor):
    import ctypes
    from ctypes import wintypes

    class MONITORINFO(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", wintypes.RECT), ("rcWork", wintypes.RECT),
                    ("dwFlags", wintypes.DWORD)]
    info = MONITORINFO()
    info.cbSize = ctypes.sizeof(info)
    if not monitor or not user32().GetMonitorInfoW(monitor, ctypes.byref(info)):
        return None
    r = info.rcWork
    return r.left, r.top, r.right - r.left, r.bottom - r.top


def work_area(widget, point=None):
    """(x, y, width, height) of the work area of the monitor holding `point` (x, y) - or else the
    monitor of widget's window. Elsewhere than Windows Tk's screen already spans all monitors."""
    if sys.platform == "win32":
        try:
            from ctypes import wintypes
            u = user32()
            monitor = (u.MonitorFromPoint(wintypes.POINT(round(point[0]), round(point[1])), 2) if point
                       else u.MonitorFromWindow(u.GetParent(widget.winfo_toplevel().winfo_id()), 2))  # 2: nearest
            area = _monitor_work_area(monitor)
            if area:
                return area
        except (AttributeError, OSError, ValueError):
            pass
    return 0, 0, widget.winfo_screenwidth(), widget.winfo_screenheight()


def keep_on_screen(widget, x, y, width, height, flip=None):
    """(x, y) for a width x height window at (x, y), moved to stay on the monitor (x, y) is on. With
    `flip`, a window that would go below the bottom opens `flip` pixels higher instead (a dropdown
    then opens above what it belongs to)."""
    ax, ay, aw, ah = work_area(widget, (x, y))
    x = max(ax, min(x, ax + aw - width - 4))
    if flip is not None and y + height > ay + ah - 4:
        y -= height + flip
    return x, max(ay, min(y, ay + ah - height)) if flip is None else max(ay, y)


def center_on(window, parent, width, height, rise=2):
    """(x, y) placing a width x height window over the middle of parent's window (a third of the way
    down with rise=3), on parent's monitor."""
    top = parent.winfo_toplevel()
    x = top.winfo_rootx() + (top.winfo_width() - width) // 2
    y = top.winfo_rooty() + (top.winfo_height() - height) // rise
    ax, ay, aw, ah = work_area(top, (top.winfo_rootx() + top.winfo_width() // 2,
                                     top.winfo_rooty() + top.winfo_height() // 2))
    return max(ax, min(x, ax + aw - width)), max(ay, min(y, ay + ah - height))


def monitor_scale(widget):
    """Scale (1.0 = 96 DPI, 1.5 = 150 %) of the monitor showing widget's window; None when not known
    (only Windows tells it per monitor)."""
    if sys.platform != "win32":
        return None
    try:
        import ctypes
        from ctypes import wintypes
        u = user32()
        monitor = u.MonitorFromWindow(u.GetParent(widget.winfo_toplevel().winfo_id()), 2)  # 2: nearest
        shcore = ctypes.WinDLL("shcore")
        shcore.GetDpiForMonitor.argtypes = [wintypes.HMONITOR, ctypes.c_int, ctypes.POINTER(wintypes.UINT),
                                            ctypes.POINTER(wintypes.UINT)]
        dpi_x, dpi_y = wintypes.UINT(), wintypes.UINT()
        if monitor and shcore.GetDpiForMonitor(monitor, 0, ctypes.byref(dpi_x), ctypes.byref(dpi_y)) == 0:
            return dpi_x.value / 96  # 0: the effective DPI, which includes the user's scale setting
    except (AttributeError, OSError, ValueError):
        pass
    return None


def follow_monitor_dpi(per_monitor=True):
    """Tell Windows that PyExcel draws itself at each monitor's DPI (before the first window is made). A
    program that doesn't is drawn at the main monitor's scale and stretched by Windows on a monitor with
    another scale: blurred, and with many more pixels to draw each time (slow scrolling there). Returns
    whether it worked; otherwise (or with per_monitor=False) the main monitor's DPI is used."""
    if sys.platform != "win32":
        return False
    import ctypes
    if per_monitor:
        try:
            if ctypes.windll.shcore.SetProcessDpiAwareness(2) == 0:  # 2: PROCESS_PER_MONITOR_DPI_AWARE
                return True
        except (AttributeError, OSError):
            pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI: crisp on the main monitor
    except (AttributeError, OSError):
        pass
    return False


def grab_screen(x, y, width, height):
    """What the screen shows in a rectangle, as PPM image data Tk can show without Pillow; None when it
    can't be read (only Windows here). Used to cross-fade the window while it is drawn again."""
    if sys.platform != "win32" or width <= 0 or height <= 0:
        return None
    try:
        import ctypes
        from ctypes import wintypes
        u, gdi = user32(), ctypes.WinDLL("gdi32")
        u.GetDC.argtypes, u.GetDC.restype = [wintypes.HWND], wintypes.HDC
        u.ReleaseDC.argtypes, u.ReleaseDC.restype = [wintypes.HWND, wintypes.HDC], ctypes.c_int
        gdi.CreateCompatibleDC.argtypes, gdi.CreateCompatibleDC.restype = [wintypes.HDC], wintypes.HDC
        gdi.CreateCompatibleBitmap.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int]
        gdi.CreateCompatibleBitmap.restype = wintypes.HBITMAP
        gdi.SelectObject.argtypes, gdi.SelectObject.restype = [wintypes.HDC, wintypes.HGDIOBJ], wintypes.HGDIOBJ
        gdi.BitBlt.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                               wintypes.HDC, ctypes.c_int, ctypes.c_int, wintypes.DWORD]
        gdi.BitBlt.restype = wintypes.BOOL
        gdi.GetDIBits.argtypes = [wintypes.HDC, wintypes.HBITMAP, wintypes.UINT, wintypes.UINT, ctypes.c_void_p,
                                  ctypes.c_void_p, wintypes.UINT]
        gdi.GetDIBits.restype = ctypes.c_int
        gdi.DeleteObject.argtypes, gdi.DeleteObject.restype = [wintypes.HGDIOBJ], wintypes.BOOL
        gdi.DeleteDC.argtypes, gdi.DeleteDC.restype = [wintypes.HDC], wintypes.BOOL

        class BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [("biSize", wintypes.DWORD), ("biWidth", ctypes.c_long), ("biHeight", ctypes.c_long),
                        ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD),
                        ("biCompression", wintypes.DWORD), ("biSizeImage", wintypes.DWORD),
                        ("biXPelsPerMeter", ctypes.c_long), ("biYPelsPerMeter", ctypes.c_long),
                        ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)]
        screen = u.GetDC(None)
        if not screen:
            return None
        memory = gdi.CreateCompatibleDC(screen)
        bitmap = gdi.CreateCompatibleBitmap(screen, width, height)
        try:
            old = gdi.SelectObject(memory, bitmap)
            ok = gdi.BitBlt(memory, 0, 0, width, height, screen, x, y, 0x00CC0020)  # SRCCOPY
            gdi.SelectObject(memory, old)
            header = BITMAPINFOHEADER(ctypes.sizeof(BITMAPINFOHEADER), width, -height, 1, 32, 0)  # top-down
            pixels = ctypes.create_string_buffer(width * height * 4)
            if not ok or gdi.GetDIBits(memory, bitmap, 0, height, pixels, ctypes.byref(header), 0) != height:
                return None
        finally:
            gdi.DeleteObject(bitmap)
            gdi.DeleteDC(memory)
            u.ReleaseDC(None, screen)
        bgra = pixels.raw
        rgb = bytearray(width * height * 3)
        rgb[0::3], rgb[1::3], rgb[2::3] = bgra[2::4], bgra[1::4], bgra[0::4]
        return b"P6 %d %d 255\n" % (width, height) + bytes(rgb)
    except (AttributeError, OSError, ValueError, MemoryError):
        return None
