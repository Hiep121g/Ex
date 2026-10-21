"""Pictures for PyExcel (Insert > Pictures): what a picture file is, Tk images of it at any size and the
pixel data a PDF needs.

Pillow is used when it is installed (any format, smooth resizing). Without it PNG and GIF are still
shown - Tk reads them itself and scales by whole-number steps - and pictures go into PDFs as they are
(JPEG) or decoded here in pure Python (PNG); JPEG pictures on screen need Pillow.
"""

import base64
import io
import struct
import zlib
from fractions import Fraction

try:
    from PIL import Image as PILImage
except ImportError:  # Pillow is optional
    PILImage = None

FORMATS = ("png", "gif", "jpeg")


def image_info(data):
    """(format, width, height) read from the header of PNG, GIF, JPEG or BMP data; None if unknown."""
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
        return ("png",) + struct.unpack(">II", data[16:24])
    if data[:6] in (b"GIF87a", b"GIF89a") and len(data) >= 10:
        return ("gif",) + struct.unpack("<HH", data[6:10])
    if data[:2] == b"\xff\xd8":
        sof = _jpeg_sof(data)
        return ("jpeg", sof[1], sof[0]) if sof else None
    if data[:2] == b"BM" and len(data) >= 26:
        w, h = struct.unpack("<ii", data[18:26])
        return "bmp", w, abs(h)
    return None


def _jpeg_sof(data):
    """(height, width, components) from a JPEG's start-of-frame segment."""
    i = 2
    while i + 9 < len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7 or marker == 0xFF:
            i += 1 if marker == 0xFF else 2
            continue
        if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
            h, w = struct.unpack(">HH", data[i + 5:i + 9])
            return h, w, data[i + 9]
        i += 2 + struct.unpack(">H", data[i + 2:i + 4])[0]
    return None


def load_picture(path):
    """(base64 text, format, (width, height)) of a picture file, converting other formats to PNG with
    Pillow. Raises ValueError for files that are not pictures PyExcel can keep."""
    with open(path, "rb") as f:
        data = f.read()
    info = image_info(data)
    if info and info[0] in FORMATS:
        return base64.b64encode(data).decode("ascii"), info[0], info[1:]
    if PILImage is not None:
        try:
            img = PILImage.open(io.BytesIO(data))
            out = io.BytesIO()
            img.save(out, format="PNG")
            return base64.b64encode(out.getvalue()).decode("ascii"), "png", img.size
        except (OSError, ValueError) as err:
            raise ValueError(str(err)) from None
    raise ValueError("Chỉ chèn được ảnh PNG, GIF hoặc JPEG (cài thêm Pillow để mở các định dạng khác).")


# ------------------------------------------------------------------ PNG decoding (for PDFs without Pillow)

def decode_png(data):
    """(width, height, rgb bytes, alpha bytes or None) of an 8-bit (or paletted) non-interlaced PNG;
    None for kinds this small decoder does not handle."""
    pos, chunks, idat = 8, {}, []
    while pos + 8 <= len(data):
        length, kind = struct.unpack(">I4s", data[pos:pos + 8])
        body = data[pos + 8:pos + 8 + length]
        if kind == b"IDAT":
            idat.append(body)
        else:
            chunks.setdefault(kind, body)
        pos += 12 + length
        if kind == b"IEND":
            break
    if b"IHDR" not in chunks:
        return None
    w, h, depth, ctype, _, _, interlace = struct.unpack(">IIBBBBB", chunks[b"IHDR"])
    if interlace or ctype not in (0, 2, 3, 4, 6) or (depth != 8 and ctype != 3) or depth > 8:
        return None
    channels = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[ctype]
    bpp = max(1, channels * depth // 8)
    stride = (w * channels * depth + 7) // 8
    raw = zlib.decompress(b"".join(idat))
    rows, prev = [], bytearray(stride)
    for y in range(h):
        start = y * (stride + 1)
        ftype, line = raw[start], bytearray(raw[start + 1:start + 1 + stride])
        if ftype == 1:
            for i in range(bpp, stride):
                line[i] = (line[i] + line[i - bpp]) & 255
        elif ftype == 2:
            line = bytearray((a + b) & 255 for a, b in zip(line, prev))
        elif ftype == 3:
            for i in range(stride):
                left = line[i - bpp] if i >= bpp else 0
                line[i] = (line[i] + ((left + prev[i]) >> 1)) & 255
        elif ftype == 4:
            for i in range(stride):
                a = line[i - bpp] if i >= bpp else 0
                b, c = prev[i], prev[i - bpp] if i >= bpp else 0
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                line[i] = (line[i] + (a if pa <= pb and pa <= pc else b if pb <= pc else c)) & 255
        rows.append(line)
        prev = line
    pixels = b"".join(rows)
    if ctype == 3:  # palette: expand indices to RGB (and alpha from tRNS)
        if depth < 8:
            per = 8 // depth
            mask = (1 << depth) - 1
            idx = bytearray()
            for line in rows:
                row = [(byte >> (8 - depth * (k + 1))) & mask for byte in line for k in range(per)]
                idx += bytes(row[:w])
        else:
            idx = pixels
        pal, trns = chunks.get(b"PLTE", b""), chunks.get(b"tRNS", b"")
        rgb = b"".join(pal[3 * i:3 * i + 3].ljust(3, b"\0") for i in range(256))
        rgb_px = b"".join(rgb[3 * i:3 * i + 3] for i in idx)
        alpha = bytes(trns[i] if i < len(trns) else 255 for i in idx) if trns else None
        return w, h, rgb_px, alpha
    if ctype == 0:
        return w, h, bytes(v for v in pixels for _ in range(3)), None
    if ctype == 4:
        gray, alpha = pixels[0::2], pixels[1::2]
        return w, h, bytes(v for v in gray for _ in range(3)), alpha
    if ctype == 6:
        rgb = bytearray(w * h * 3)
        rgb[0::3], rgb[1::3], rgb[2::3] = pixels[0::4], pixels[1::4], pixels[2::4]
        return w, h, bytes(rgb), pixels[3::4]
    return w, h, pixels, None


def pdf_image(data, fmt):
    """What a PDF image XObject needs: {"w", "h", "space", "filter", "data", "alpha"} (alpha: deflated
    8-bit soft mask or None), or None when the picture can't be put in a PDF here."""
    if fmt == "jpeg":
        sof = _jpeg_sof(data)
        if sof:
            space = {1: "DeviceGray", 3: "DeviceRGB", 4: "DeviceCMYK"}.get(sof[2])
            if space:
                return {"w": sof[1], "h": sof[0], "space": space, "filter": "DCTDecode", "data": data,
                        "alpha": None, "invert": sof[2] == 4}
    decoded = None
    if PILImage is not None:
        try:
            img = PILImage.open(io.BytesIO(data))
            img = img.convert("RGBA") if "A" in img.getbands() or "transparency" in img.info else img.convert("RGB")
            alpha = img.getchannel("A").tobytes() if img.mode == "RGBA" else None
            decoded = img.width, img.height, img.convert("RGB").tobytes(), alpha
        except (OSError, ValueError):
            decoded = None
    elif fmt == "png":
        decoded = decode_png(data)
    if decoded is None:
        return None
    w, h, rgb, alpha = decoded
    if alpha is not None and min(alpha) == 255:
        alpha = None
    return {"w": w, "h": h, "space": "DeviceRGB", "filter": "FlateDecode", "data": zlib.compress(rgb, 6),
            "alpha": zlib.compress(alpha, 6) if alpha else None, "invert": False}


# ------------------------------------------------------------------ Tk images

class PictureCache:
    """Tk images of pictures at given pixel sizes, kept while they are drawn."""

    def __init__(self, master):
        self.master = master
        self._source, self._sized = {}, {}

    def photo(self, key, data64, fmt, width, height):
        """A Tk image of picture `key` about width x height pixels, or None when it can't be shown."""
        width, height = max(1, round(width)), max(1, round(height))
        img = self._sized.get((key, width, height))
        if img is not None:
            return img
        if len(self._sized) > 48:
            self._sized.clear()
        try:
            img = self._make(key, data64, fmt, width, height)
        except Exception:  # a damaged picture is drawn as an empty frame, not an error
            img = None
        self._sized[(key, width, height)] = img
        return img

    def _make(self, key, data64, fmt, width, height):
        import tkinter as tk
        if PILImage is not None:
            try:
                from PIL import ImageTk
                pil = PILImage.open(io.BytesIO(base64.b64decode(data64)))
                pil = pil.convert("RGBA").resize((width, height), getattr(PILImage, "LANCZOS", 1))
                return ImageTk.PhotoImage(pil, master=self.master)
            except ImportError:  # Pillow without its Tk part: fall back to Tk's own PNG/GIF reader
                pass
        if fmt not in ("png", "gif"):
            return None
        src = self._source.get(key)
        if src is None:
            src = self._source[key] = tk.PhotoImage(master=self.master, data=data64)
        img = src
        # Tk scales by whole numbers only: step down to just above the size, then a small fraction (n/d with
        # d <= 4) when the zoomed-in step stays a reasonable size
        sx = max(1, int(src.width() / width)) if width < src.width() else 1
        sy = max(1, int(src.height() / height)) if height < src.height() else 1
        if sx > 1 or sy > 1:
            img = img.subsample(sx, sy)
        fx = Fraction(width / img.width()).limit_denominator(4)
        fy = Fraction(height / img.height()).limit_denominator(4)
        if (fx != 1 or fy != 1) and fx.numerator and fy.numerator and \
                img.width() * fx.numerator * img.height() * fy.numerator <= 4_000_000:
            img = img.zoom(fx.numerator, fy.numerator)
            if fx.denominator > 1 or fy.denominator > 1:
                img = img.subsample(fx.denominator, fy.denominator)
        return img
