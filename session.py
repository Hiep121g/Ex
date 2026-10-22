"""PyExcel's own files outside workbooks: settings, the recent-files list and AutoRecover copies.

They live in %APPDATA%\\PyExcel on Windows (~/.config/PyExcel elsewhere). Every running PyExcel writes a
small heartbeat file each minute and, while there are unsaved changes, a recovery copy of the workbook
every few minutes. A recovery copy whose heartbeat has stopped belongs to a PyExcel that closed without
saving (a crash or power cut) and is offered when PyExcel starts again.
"""

import json
import os
import sys
import threading
import time

import formula_engine as fe

MAX_RECENT = 10
QAT_DEFAULT = ["save", "undo", "redo"]  # Excel's Quick Access Toolbar out of the box
RIBBON_MODES = ("full", "tabs", "auto")  # tabs and commands / tabs only / auto-hide
DEFAULTS = {"recent": [], "autosave_minutes": 5, "qat": QAT_DEFAULT, "ribbon": "full", "custom_title_bar": None}


def app_dir():
    if sys.platform == "win32":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "PyExcel")


def _read_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return default


def _write_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


class Settings:
    """settings.json: the recent files, the AutoRecover interval, the Quick Access Toolbar's commands, how
    the ribbon shows and whether PyExcel draws the window's title bar."""

    def __init__(self, folder=None):
        self.folder = folder or app_dir()
        self.path = os.path.join(self.folder, "settings.json")
        data = _read_json(self.path, {})
        self.data = dict(DEFAULTS, **(data if isinstance(data, dict) else {}))

    def save(self):
        try:
            os.makedirs(self.folder, exist_ok=True)
            _write_json(self.path, self.data)
        except OSError:
            pass  # settings are a convenience: a read-only profile must not stop the program

    def recent(self):
        """Recent files that still exist, newest first."""
        return [p for p in self.data.get("recent", []) if isinstance(p, str) and os.path.exists(p)]

    def add_recent(self, path):
        path = os.path.abspath(path)
        key = os.path.normcase(path)
        items = [p for p in self.data.get("recent", []) if os.path.normcase(p) != key]
        self.data["recent"] = [path] + items[:MAX_RECENT - 1]
        self.save()

    def clear_recent(self):
        self.data["recent"] = []
        self.save()

    @property
    def autosave_minutes(self):
        try:
            return max(0, int(self.data.get("autosave_minutes", 5)))
        except (TypeError, ValueError):
            return 5

    @autosave_minutes.setter
    def autosave_minutes(self, minutes):
        self.data["autosave_minutes"] = max(0, int(minutes))
        self.save()

    @property
    def qat(self):
        items = self.data.get("qat")
        return [k for k in items if isinstance(k, str)] if isinstance(items, list) else list(QAT_DEFAULT)

    @qat.setter
    def qat(self, keys):
        self.data["qat"] = list(keys)
        self.save()

    @property
    def ribbon(self):
        mode = self.data.get("ribbon")
        return mode if mode in RIBBON_MODES else "full"

    @ribbon.setter
    def ribbon(self, mode):
        self.data["ribbon"] = mode
        self.save()

    @property
    def custom_title_bar(self):
        """PyExcel's own title bar instead of the system's: by default on Windows (as Excel)."""
        value = self.data.get("custom_title_bar")
        return sys.platform == "win32" if value is None else bool(value)

    @custom_title_bar.setter
    def custom_title_bar(self, value):
        self.data["custom_title_bar"] = bool(value)
        self.save()


class Recovery:
    """This session's AutoRecover copy and heartbeat, and the copies left by sessions that died."""

    STALE = 150  # seconds without a heartbeat after which a session counts as gone

    def __init__(self, folder=None):
        self.dir = os.path.join(folder or app_dir(), "AutoRecover")
        self.id = f"{os.getpid()}-{int(time.time())}-{os.urandom(3).hex()}"
        self.file = os.path.join(self.dir, self.id + ".pyxl")
        self.meta = os.path.join(self.dir, self.id + ".json")
        self.info = {}
        self.saved_at = None

    def beat(self, **info):
        """Say this session is alive (and what it is editing)."""
        self.info.update(info)
        try:
            os.makedirs(self.dir, exist_ok=True)
            _write_json(self.meta, dict(self.info, alive=time.time(), saved=self.saved_at,
                                        has_copy=os.path.exists(self.file)))
        except OSError:
            pass

    def save(self, wb, active=None):
        """Write the recovery copy of the workbook; False when it could not be written."""
        try:
            os.makedirs(self.dir, exist_ok=True)
            tmp = self.file + ".tmp.pyxl"
            fe.save_workbook(wb, tmp, active=active)
            os.replace(tmp, self.file)
        except (OSError, ValueError):
            return False
        self.saved_at = time.time()
        self.beat()
        return True

    def save_in_background(self, snapshot, active=None):
        """Write the recovery copy from a workbook snapshot (Workbook.snapshot()) on another thread, so a big
        workbook doesn't stop the program while it is written. Returns the thread; saved_at is set when done."""
        def work():
            copy = fe.Workbook(())
            copy.restore(snapshot)
            self.save(copy, active)
        thread = threading.Thread(target=work, name="PyExcel AutoRecover", daemon=True)
        thread.start()
        return thread

    def discard(self):
        """The work is saved (or deliberately dropped): no recovery copy is needed any more."""
        for path in (self.file, self.file + ".tmp.pyxl"):
            try:
                os.remove(path)
            except OSError:
                pass
        self.saved_at = None
        self.beat()

    def close(self):
        for path in (self.file, self.meta):
            try:
                os.remove(path)
            except OSError:
                pass

    def orphans(self, now=None):
        """Recovery copies of sessions that stopped without closing: [{"file", "meta", "path", "saved"}],
        newest first. Heartbeats of dead sessions without a copy are cleaned up."""
        now = now or time.time()
        found, alive = [], {self.id}
        try:
            names = os.listdir(self.dir)
        except OSError:
            return found
        for name in names:
            if not name.endswith(".json") or name[:-5] == self.id:
                continue
            meta = os.path.join(self.dir, name)
            info = _read_json(meta, {})
            if not isinstance(info, dict) or now - float(info.get("alive") or 0) < self.STALE:
                alive.add(name[:-5])
                continue
            copy = meta[:-5] + ".pyxl"
            if os.path.exists(copy):
                found.append({"file": copy, "meta": meta, "path": info.get("path"),
                              "saved": info.get("saved") or os.path.getmtime(copy)})
            else:
                self.forget({"meta": meta})
        for name in names:  # copies a dead session was still writing are of no use
            if name.endswith(".tmp.pyxl") and name.split(".", 1)[0] not in alive:
                self.forget({"file": os.path.join(self.dir, name)})
        return sorted(found, key=lambda e: e["saved"], reverse=True)

    @staticmethod
    def forget(entry):
        for key in ("file", "meta"):
            if entry.get(key):
                try:
                    os.remove(entry[key])
                except OSError:
                    pass
