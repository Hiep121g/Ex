"""TrueType fonts for PDF export: finding installed fonts, metrics, character maps and subsetting.

Only fonts with TrueType outlines (glyf table) are used, single .ttf files or members of .ttc
collections - the kind Windows ships. A subset keeps the glyph ids of the original font (unused
glyphs become empty), so a PDF can address glyphs by their original ids.
"""

import os
import struct
import sys

_TT_VERSIONS = (0x00010000, 0x74727565)  # version 1.0 / 'true'; 'OTTO' (CFF outlines) is not supported
BOLD, ITALIC = 1, 2  # head.macStyle bits


def _u16(data, off):
    return struct.unpack_from(">H", data, off)[0]


def _i16(data, off):
    return struct.unpack_from(">h", data, off)[0]


def _u32(data, off):
    return struct.unpack_from(">I", data, off)[0]


def _directory(read, base):
    """{tag: (offset, length)} of the font starting at `base`, or None when it is not TrueType."""
    head = read(base, 12)
    if len(head) < 12 or _u32(head, 0) not in _TT_VERSIONS:
        return None
    count = _u16(head, 4)
    entries = read(base + 12, 16 * count)
    tables = {}
    for i in range(count):
        tag, _, off, length = struct.unpack_from(">4sIII", entries, 16 * i)
        tables[tag.decode("latin-1")] = (off, length)
    return tables


def _members(read):
    """Offsets of the fonts in a file (several for a .ttc collection)."""
    head = read(0, 12)
    if head[:4] == b"ttcf":
        n = _u32(head, 8)
        return [_u32(read(12 + 4 * i, 4), 0) for i in range(n)]
    return [0]


def _names(data):
    """{name id: string} from a 'name' table, preferring Windows Unicode English records."""
    count, strings = _u16(data, 2), _u16(data, 4)
    found = {}
    for i in range(count):
        platform, encoding, language, name_id, length, off = struct.unpack_from(">6H", data, 6 + 12 * i)
        raw = data[strings + off:strings + off + length]
        if platform == 3 and encoding in (0, 1, 10):
            text, score = raw.decode("utf-16-be", "replace"), 3 if language == 0x409 else 2
        elif platform == 1 and encoding == 0:
            text, score = raw.decode("mac_roman", "replace"), 1
        elif platform == 0:
            text, score = raw.decode("utf-16-be", "replace"), 2
        else:
            continue
        if score > found.get(name_id, (0, ""))[0]:
            found[name_id] = (score, text)
    return {k: v[1] for k, v in found.items()}


def describe(path):
    """[(family, style bits, member index)] for each TrueType font in a file; [] if none can be read."""
    out = []
    try:
        with open(path, "rb") as f:
            def read(off, n):
                f.seek(off)
                return f.read(n)
            for index, base in enumerate(_members(read)):
                tables = _directory(read, base)
                if not tables or "glyf" not in tables or "name" not in tables or "head" not in tables:
                    continue
                names = _names(read(*tables["name"]))
                family = names.get(16) or names.get(1)  # typographic family, else legacy family
                style = _u16(read(tables["head"][0] + 44, 2), 0) & 3
                sub = (names.get(17) or names.get(2) or "").lower()
                style |= (BOLD if "bold" in sub else 0) | (ITALIC if "italic" in sub or "oblique" in sub else 0)
                if family:
                    out.append((family, style, index))
    except (OSError, struct.error, IndexError, ValueError):
        return []
    return out


def font_dirs():
    if sys.platform == "win32":
        dirs = [os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")]
        local = os.environ.get("LOCALAPPDATA")
        if local:
            dirs.append(os.path.join(local, "Microsoft", "Windows", "Fonts"))
        return dirs
    if sys.platform == "darwin":
        return ["/System/Library/Fonts", "/Library/Fonts", os.path.expanduser("~/Library/Fonts")]
    return ["/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.fonts"),
            os.path.expanduser("~/.local/share/fonts")]


# fonts standing in for common Windows fonts that are not installed (metric-compatible where possible)
ALIASES = {"calibri": ["carlito"], "cambria": ["caladea"], "arial": ["liberation sans", "arimo"],
           "times new roman": ["liberation serif", "tinos"], "courier new": ["liberation mono", "cousine"],
           "segoe ui": ["arial", "liberation sans"], "tahoma": ["dejavu sans"], "verdana": ["dejavu sans"]}
# tried in turn for characters the chosen font lacks (e.g. Japanese in Calibri)
FALLBACKS = ["Arial", "Segoe UI", "Tahoma", "Arial Unicode MS", "Yu Gothic", "Meiryo", "MS Gothic", "MS Mincho",
             "Microsoft YaHei", "SimSun", "Malgun Gothic", "Segoe UI Symbol", "Noto Sans CJK JP", "Noto Sans",
             "IPAGothic", "TakaoGothic", "WenQuanYi Zen Hei", "DejaVu Sans", "Liberation Sans", "FreeSans"]


class FontLibrary:
    """Installed TrueType fonts by family and style, loaded on demand."""

    def __init__(self, dirs=None):
        self.dirs = dirs or font_dirs()
        self._index = None
        self._fonts = {}

    def index(self):
        if self._index is None:
            self._index = {}
            for folder in self.dirs:
                for root, _, files in os.walk(folder):
                    for name in sorted(files):
                        if name.lower().endswith((".ttf", ".ttc")):
                            path = os.path.join(root, name)
                            for family, style, member in describe(path):
                                self._index.setdefault((family.lower(), style), (path, member))
        return self._index

    def families(self):
        return sorted({family for family, _ in self.index()})

    def _load(self, key):
        font = self._fonts.get(key)
        if font is None:
            try:
                font = self._fonts[key] = TTFont(*key)
            except (OSError, ValueError, struct.error, IndexError):
                font = self._fonts[key] = False
        return font or None

    def find(self, family, bold=False, italic=False):
        """The installed font closest to family + style (aliases, then any regular font), or None."""
        index = self.index()
        want = (BOLD if bold else 0) | (ITALIC if italic else 0)
        for name in [family.lower()] + ALIASES.get(family.lower(), []) + [f.lower() for f in FALLBACKS]:
            for style in (want, want & BOLD, want & ITALIC, 0, BOLD, ITALIC, 3):
                key = index.get((name, style))
                if key:
                    font = self._load(key)
                    if font:
                        return font
        return None

    def fallback_for(self, char):
        """A font that has `char`, or None."""
        for family in FALLBACKS:
            font = self.find(family)
            if font and font.has(char):
                return font
        for key in self.index().values():
            font = self._load(key)
            if font and font.has(char):
                return font
        return None


class TTFont:
    def __init__(self, path, member=0):
        with open(path, "rb") as f:
            self.data = data = f.read()
        self.path, self.member = path, member
        base = _members(lambda off, n: data[off:off + n])[member]
        self.tables = _directory(lambda off, n: data[off:off + n], base)
        if not self.tables or any(t not in self.tables for t in ("head", "hhea", "hmtx", "maxp", "cmap", "loca", "glyf")):
            raise ValueError("not a TrueType font")
        head, hhea = self.table("head"), self.table("hhea")
        self.units = _u16(head, 18)
        self.bbox = [_i16(head, 36 + 2 * i) * 1000 // self.units for i in range(4)]
        self.long_loca = _i16(head, 50) == 1
        self.ascent = _i16(hhea, 4) * 1000 // self.units
        self.descent = _i16(hhea, 6) * 1000 // self.units
        self.num_glyphs = _u16(self.table("maxp"), 4)
        n_metrics = _u16(hhea, 34)
        hmtx = self.table("hmtx")
        widths = [_u16(hmtx, 4 * i) for i in range(n_metrics)]
        self.widths = widths + [widths[-1]] * (self.num_glyphs - n_metrics)
        names = _names(self.table("name")) if "name" in self.tables else {}
        self.family = names.get(16) or names.get(1) or os.path.basename(path)
        self.postscript = "".join(ch for ch in (names.get(6) or self.family) if ch.isalnum() or ch in "-_") or "Font"
        self.style = _u16(head, 44) & 3
        os2 = self.table("OS/2") if "OS/2" in self.tables else b""
        self.cap_height = (_i16(os2, 88) * 1000 // self.units) if len(os2) >= 90 else self.ascent
        post = self.table("post") if "post" in self.tables else b""
        self.italic_angle = struct.unpack_from(">i", post, 4)[0] / 65536 if len(post) >= 8 else 0.0
        self.cmap = self._read_cmap()

    def table(self, tag):
        off, length = self.tables[tag]
        return self.data[off:off + length]

    def _read_cmap(self):
        data = self.table("cmap")
        best = None
        for i in range(_u16(data, 2)):
            platform, encoding, off = struct.unpack_from(">HHI", data, 4 + 8 * i)
            fmt = _u16(data, off)
            score = {(3, 10, 12): 4, (0, 4, 12): 4, (3, 1, 4): 3, (0, 3, 4): 2, (3, 0, 4): 1}.get(
                (platform, encoding, fmt), 2 if platform == 0 and fmt in (4, 12) else 0)
            if score and (best is None or score > best[0]):
                best = (score, off, fmt)
        cmap = {}
        if best is None:
            return cmap
        _, off, fmt = best
        if fmt == 4:
            seg2 = _u16(data, off + 6)
            ends, starts = off + 14, off + 16 + seg2
            deltas, ranges = starts + seg2, starts + 2 * seg2
            for s in range(seg2 // 2):
                end, start = _u16(data, ends + 2 * s), _u16(data, starts + 2 * s)
                delta, rng = _i16(data, deltas + 2 * s), _u16(data, ranges + 2 * s)
                for c in range(start, min(end, 0xFFFE) + 1):
                    if rng == 0:
                        gid = (c + delta) & 0xFFFF
                    else:
                        g = _u16(data, ranges + 2 * s + rng + 2 * (c - start))
                        gid = (g + delta) & 0xFFFF if g else 0
                    if gid:
                        cmap[c] = gid
        else:
            for g in range(_u32(data, off + 12)):
                start, end, first = struct.unpack_from(">III", data, off + 16 + 12 * g)
                for c in range(start, min(end, 0x10FFFF) + 1):
                    cmap[c] = first + c - start
        return cmap

    def has(self, char):
        return ord(char) in self.cmap

    def glyph(self, char):
        return self.cmap.get(ord(char), 0)

    def advance(self, gid):
        """Advance width of a glyph in thousandths of the font size."""
        return self.widths[gid] * 1000 / self.units if gid < len(self.widths) else 0

    # subsetting
    def _loca(self):
        data = self.table("loca")
        if self.long_loca:
            return [_u32(data, 4 * i) for i in range(self.num_glyphs + 1)]
        return [2 * _u16(data, 2 * i) for i in range(self.num_glyphs + 1)]

    def subset(self, gids):
        """A standalone TrueType font keeping only glyphs `gids` (plus .notdef and composite parts)."""
        loca, glyf = self._loca(), self.table("glyf")
        keep, todo = set(), {0} | set(gids)
        while todo:
            gid = todo.pop()
            if gid in keep or gid >= self.num_glyphs:
                continue
            keep.add(gid)
            start, end = loca[gid], loca[gid + 1]
            if end - start > 10 and _i16(glyf, start) < 0:  # composite glyph: keep its components
                pos = start + 10
                while True:
                    flags, component = _u16(glyf, pos), _u16(glyf, pos + 2)
                    todo.add(component)
                    pos += 4 + (4 if flags & 0x0001 else 2)
                    pos += 8 if flags & 0x0080 else 4 if flags & 0x0040 else 2 if flags & 0x0008 else 0
                    if not flags & 0x0020:
                        break
        count = max(keep) + 1  # glyphs after the last one used are dropped altogether
        new_glyf, new_loca = bytearray(), []
        for gid in range(count):
            new_loca.append(len(new_glyf))
            if gid in keep:
                new_glyf += glyf[loca[gid]:loca[gid + 1]]
                new_glyf += b"\0" * (-len(new_glyf) % 4)
        new_loca.append(len(new_glyf))
        head = bytearray(self.table("head"))
        struct.pack_into(">I", head, 8, 0)    # checkSumAdjustment, set below
        struct.pack_into(">h", head, 50, 1)   # long loca offsets
        hhea, maxp, hmtx = bytearray(self.table("hhea")), bytearray(self.table("maxp")), self.table("hmtx")
        metrics = _u16(hhea, 34)
        if count <= metrics:
            hmtx, metrics = hmtx[:4 * count], count
        else:
            hmtx = hmtx[:4 * metrics + 2 * (count - metrics)]
        struct.pack_into(">H", hhea, 34, metrics)
        struct.pack_into(">H", maxp, 4, count)
        tables = {"head": bytes(head), "hhea": bytes(hhea), "maxp": bytes(maxp), "hmtx": hmtx,
                  "loca": struct.pack(f">{len(new_loca)}I", *new_loca), "glyf": bytes(new_glyf)}
        for tag in ("cvt ", "fpgm", "prep", "OS/2"):  # PDF needs no cmap or name; hinting tables stay
            if tag in self.tables:
                tables[tag] = self.table(tag)
        if "post" in self.tables:  # version 3: no glyph names
            tables["post"] = struct.pack(">I", 0x00030000) + self.table("post")[4:32]
        return _build_font(tables)


def _checksum(data):
    data += b"\0" * (-len(data) % 4)
    return sum(struct.unpack(f">{len(data) // 4}I", data)) & 0xFFFFFFFF


def _build_font(tables):
    tags = sorted(tables)
    n = len(tags)
    power = 1
    while power * 2 <= n:
        power *= 2
    header = struct.pack(">IHHHH", 0x00010000, n, power * 16, power.bit_length() - 1, n * 16 - power * 16)
    offset = 12 + 16 * n
    directory, body = bytearray(), bytearray()
    for tag in tags:
        data = tables[tag]
        directory += struct.pack(">4sIII", tag.encode("latin-1"), _checksum(data), offset + len(body), len(data))
        body += data + b"\0" * (-len(data) % 4)
    font = bytearray(header + directory + body)
    head_at = offset + sum(len(tables[t]) + (-len(tables[t]) % 4) for t in tags[:tags.index("head")])
    struct.pack_into(">I", font, head_at + 8, (0xB1B0AFBA - _checksum(bytes(font))) & 0xFFFFFFFF)
    return bytes(font)
