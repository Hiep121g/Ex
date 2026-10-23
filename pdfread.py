"""Reading PDF files (Tiện ích > Nhập PDF): each page's text with its position, size, font and colour, the
lines and shapes drawn on it and its pictures - in pure Python (zlib only), nothing to install.

It covers what PDFs made by programs use: classic and compressed cross-reference tables (and rebuilding a
damaged one), object streams, the Flate / LZW / ASCII85 / ASCIIHex / RunLength filters with predictors,
the standard security handler with an empty password (RC4, AES-128, AES-256), simple and composite (CID)
fonts with ToUnicode maps, the usual encodings and glyph names, and the CJK CMaps (Shift-JIS, GBK, Big5,
UHC, UCS-2). A scanned page is only a picture: it gives no text (reading it would take OCR).

Coordinates are PDF points with y going up, as in the file; pdf_import turns pages into cells.
"""

import hashlib
import math
import re
import struct
import unicodedata
import zlib
from collections import namedtuple


class PDFError(Exception):
    pass


class Name(str):
    """A PDF name (/Type); strings are bytes."""
    __slots__ = ()


Ref = namedtuple("Ref", "num gen")


class Stream:
    __slots__ = ("dict", "raw", "_data", "doc", "num", "gen")

    def __init__(self, d, raw, doc=None, num=0, gen=0):
        self.dict, self.raw, self._data, self.doc, self.num, self.gen = d, raw, None, doc, num, gen

    def get(self, key, default=None):
        return self.dict.get(key, default)

    def data(self):
        """The decoded bytes (stopping before an image codec such as DCT, whose data stays encoded)."""
        if self._data is None:
            raw = self.raw
            if self.doc is not None and self.doc.crypt is not None and self.dict.get("Type") != "XRef":
                raw = self.doc.crypt.decrypt(raw, self.num, self.gen, stream=True)
            filters, parms = self.filters()
            self._data = decode(raw, filters, parms)
        return self._data

    def filters(self):
        r = self.doc.resolve if self.doc is not None else (lambda x: x)
        f = r(self.dict.get("Filter") or self.dict.get("F"))
        p = r(self.dict.get("DecodeParms") or self.dict.get("DP"))
        f = [r(x) for x in f] if isinstance(f, list) else ([f] if f else [])
        p = [r(x) for x in p] if isinstance(p, list) else [p] * len(f)
        p = [x if isinstance(x, dict) else {} for x in p] + [{}] * (len(f) - len(p))
        return f, p


# ------------------------------------------------------------------ lexing

_WS = b"\x00\t\n\x0c\r "
_DELIM = b"()<>[]{}/%"
_SKIP = re.compile(rb"(?:[\x00\t\n\x0c\r ]+|%[^\r\n]*)*")
_NUMBER = re.compile(rb"[+-]?(?:\d+\.?\d*|\.\d+)")
_REGULAR = re.compile(rb"[^\x00\t\n\x0c\r ()<>\[\]{}/%]*")
_REF = re.compile(rb"(\d+)\s+(\d+)\s+R(?![^\x00\t\n\x0c\r ()<>\[\]{}/%])")
_ESCAPES = {ord("n"): b"\n", ord("r"): b"\r", ord("t"): b"\t", ord("b"): b"\b", ord("f"): b"\f",
            ord("("): b"(", ord(")"): b")", ord("\\"): b"\\"}


def _name(raw):
    if b"#" in raw:
        raw = re.sub(rb"#([0-9A-Fa-f]{2})", lambda m: bytes([int(m.group(1), 16)]), raw)
    return Name(raw.decode("latin-1"))


def _literal(data, pos):
    """A (string) starting after its "(": (bytes, position after the closing parenthesis)."""
    out, depth, n = bytearray(), 1, len(data)
    while pos < n:
        ch = data[pos]
        if ch == 0x5C:  # backslash
            pos += 1
            if pos >= n:
                break
            ch = data[pos]
            if ch in _ESCAPES:
                out += _ESCAPES[ch]
                pos += 1
            elif 0x30 <= ch <= 0x37:
                m = re.compile(rb"[0-7]{1,3}").match(data, pos)
                out.append(int(m.group(0), 8) & 0xFF)
                pos = m.end()
            elif ch == 0x0D:  # an escaped line end continues the string
                pos += 2 if data[pos + 1:pos + 2] == b"\n" else 1
            elif ch == 0x0A:
                pos += 1
            else:
                out.append(ch)
                pos += 1
            continue
        if ch == 0x28:
            depth += 1
        elif ch == 0x29:
            depth -= 1
            if depth == 0:
                return bytes(out), pos + 1
        out.append(ch)
        pos += 1
    return bytes(out), pos


def _hexstring(data, pos):
    end = data.find(b">", pos)
    if end < 0:
        end = len(data)
    digits = re.sub(rb"[^0-9A-Fa-f]", b"", data[pos:end])
    if len(digits) % 2:
        digits += b"0"
    return bytes.fromhex(digits.decode("ascii")), end + 1


class Lexer:
    """PDF objects from bytes; with refs=True "1 0 R" gives a Ref (file objects; content streams have none)."""

    def __init__(self, data, pos=0, refs=True):
        self.data, self.pos, self.refs = data, pos, refs

    def skip(self):
        self.pos = _SKIP.match(self.data, self.pos).end()

    def token(self):
        """The next object or operator (as a Keyword); None at the end."""
        self.skip()
        data, pos = self.data, self.pos
        if pos >= len(data):
            return None
        ch = data[pos]
        if ch == 0x2F:  # /
            m = _REGULAR.match(data, pos + 1)
            self.pos = m.end()
            return _name(m.group(0))
        if ch == 0x28:
            s, self.pos = _literal(data, pos + 1)
            return s
        if ch == 0x3C:
            if data[pos + 1:pos + 2] == b"<":
                self.pos = pos + 2
                return self._dict()
            s, self.pos = _hexstring(data, pos + 1)
            return s
        if ch == 0x5B:
            self.pos = pos + 1
            return self._array()
        if ch in (0x5D, 0x3E, 0x7B, 0x7D, 0x29):
            self.pos = pos + (2 if data[pos:pos + 2] == b">>" else 1)
            return Keyword(chr(ch) if data[pos:pos + 2] != b">>" else ">>")
        if self.refs and 0x30 <= ch <= 0x39:
            m = _REF.match(data, pos)
            if m:
                self.pos = m.end()
                return Ref(int(m.group(1)), int(m.group(2)))
        m = _NUMBER.match(data, pos)
        if m and (m.end() >= len(data) or data[m.end()] in _WS or data[m.end()] in _DELIM):
            self.pos = m.end()
            text = m.group(0)
            return int(text) if b"." not in text else float(text)
        m = _REGULAR.match(data, pos)
        end = m.end() if m.end() > pos else pos + 1
        self.pos = end
        word = data[pos:end]
        if word == b"true":
            return True
        if word == b"false":
            return False
        if word == b"null":
            return None
        return Keyword(word.decode("latin-1"))

    def _array(self):
        out = []
        while True:
            t = self.token()
            if t is None or (isinstance(t, Keyword) and t == "]"):
                return out
            out.append(t)

    def _dict(self):
        out = {}
        key = None
        while True:
            t = self.token()
            if t is None or (isinstance(t, Keyword) and t == ">>"):
                return out
            if key is None:
                if isinstance(t, Name):
                    key = t
            else:
                out[key] = t
                key = None


class Keyword(str):
    __slots__ = ()


# ------------------------------------------------------------------ filters

def _inflate(data):
    d = zlib.decompressobj()
    try:
        return d.decompress(data) + d.flush()
    except zlib.error:
        pass
    out = bytearray()  # damaged: keep what decompresses
    d = zlib.decompressobj()
    for i in range(0, len(data), 1024):
        try:
            out += d.decompress(data[i:i + 1024])
        except zlib.error:
            break
    if not out:
        try:
            return zlib.decompressobj(-15).decompress(data[2:])
        except zlib.error:
            return b""
    return bytes(out)


def _lzw(data, early=1):
    out = bytearray()
    table = [bytes([i]) for i in range(256)] + [b"", b""]
    bits, buf, nbits, prev = 9, 0, 0, None
    for byte in data:
        buf = (buf << 8) | byte
        nbits += 8
        while nbits >= bits:
            nbits -= bits
            code = (buf >> nbits) & ((1 << bits) - 1)
            if code == 256:
                table = table[:258]
                bits, prev = 9, None
                continue
            if code == 257:
                return bytes(out)
            if prev is None:
                entry = table[code] if code < len(table) else b""
            elif code < len(table):
                entry = table[code]
                table.append(prev + entry[:1])
            else:
                entry = prev + prev[:1]
                table.append(entry)
            out += entry
            prev = entry
            if len(table) + early >= (1 << bits) and bits < 12:
                bits += 1
    return bytes(out)


def _ascii85(data):
    data = re.sub(rb"\s", b"", data)
    if data.startswith(b"<~"):
        data = data[2:]
    end = data.find(b"~>")
    if end >= 0:
        data = data[:end]
    out, group = bytearray(), []
    for ch in data:
        if ch == ord("z") and not group:
            out += b"\0\0\0\0"
            continue
        if not 33 <= ch <= 117:
            continue
        group.append(ch - 33)
        if len(group) == 5:
            n = 0
            for v in group:
                n = n * 85 + v
            out += struct.pack(">I", n & 0xFFFFFFFF)
            group = []
    if group:
        k = len(group)
        group += [84] * (5 - k)
        n = 0
        for v in group:
            n = n * 85 + v
        out += struct.pack(">I", n & 0xFFFFFFFF)[:k - 1]
    return bytes(out)


def _runlength(data):
    out, i = bytearray(), 0
    while i < len(data):
        n = data[i]
        if n == 128:
            break
        if n < 128:
            out += data[i + 1:i + 2 + n]
            i += n + 2
        else:
            out += data[i + 1:i + 2] * (257 - n)
            i += 2
    return bytes(out)


def _unpredict(data, parms):
    predictor = parms.get("Predictor", 1)
    if not isinstance(predictor, int) or predictor < 2:
        return data
    colors = parms.get("Colors", 1) or 1
    bpc = parms.get("BitsPerComponent", 8) or 8
    columns = parms.get("Columns", 1) or 1
    bpp = max(1, colors * bpc // 8)
    row_len = (colors * bpc * columns + 7) // 8
    if predictor == 2:  # TIFF: each byte adds the one bpp before it
        out = bytearray(data)
        if bpc == 8:
            for r in range(0, len(out), row_len):
                for i in range(r + bpp, min(r + row_len, len(out))):
                    out[i] = (out[i] + out[i - bpp]) & 0xFF
        return bytes(out)
    out = bytearray()
    prev = bytearray(row_len)
    stride = row_len + 1
    for r in range(0, len(data) - row_len, stride):
        kind = data[r]
        row = bytearray(data[r + 1:r + stride])
        if len(row) < row_len:
            row += bytes(row_len - len(row))
        if kind == 1:
            for i in range(bpp, row_len):
                row[i] = (row[i] + row[i - bpp]) & 0xFF
        elif kind == 2:
            for i in range(row_len):
                row[i] = (row[i] + prev[i]) & 0xFF
        elif kind == 3:
            for i in range(row_len):
                left = row[i - bpp] if i >= bpp else 0
                row[i] = (row[i] + ((left + prev[i]) >> 1)) & 0xFF
        elif kind == 4:
            for i in range(row_len):
                a = row[i - bpp] if i >= bpp else 0
                b, c = prev[i], prev[i - bpp] if i >= bpp else 0
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                row[i] = (row[i] + (a if pa <= pb and pa <= pc else b if pb <= pc else c)) & 0xFF
        out += row
        prev = row
    return bytes(out)


IMAGE_CODECS = {"DCTDecode", "DCT", "JPXDecode", "CCITTFaxDecode", "CCF", "JBIG2Decode"}


def decode(data, filters, parms):
    for f, p in zip(filters, parms):
        if f in ("FlateDecode", "Fl"):
            data = _unpredict(_inflate(data), p)
        elif f in ("LZWDecode", "LZW"):
            data = _unpredict(_lzw(data, p.get("EarlyChange", 1)), p)
        elif f in ("ASCII85Decode", "A85"):
            data = _ascii85(data)
        elif f in ("ASCIIHexDecode", "AHx"):
            data, _ = _hexstring(data, 0)
        elif f in ("RunLengthDecode", "RL"):
            data = _runlength(data)
        elif f in IMAGE_CODECS:
            break
    return data


# ------------------------------------------------------------------ encryption (standard handler)

_PAD = bytes.fromhex("28BF4E5E4E758A4164004E56FFFA01082E2E00B6D0683E802F0CA9FE6453697A")


def _rc4(key, data):
    s = list(range(256))
    j = 0
    for i in range(256):
        j = (j + s[i] + key[i % len(key)]) & 0xFF
        s[i], s[j] = s[j], s[i]
    out = bytearray(len(data))
    i = j = 0
    for k, b in enumerate(data):
        i = (i + 1) & 0xFF
        j = (j + s[i]) & 0xFF
        s[i], s[j] = s[j], s[i]
        out[k] = b ^ s[(s[i] + s[j]) & 0xFF]
    return bytes(out)


def _xtime(a):
    return ((a << 1) ^ 0x1B) & 0xFF if a & 0x80 else a << 1


def _aes_tables():
    sbox = [0] * 256
    p = q = 1
    while True:
        p = p ^ ((p << 1) & 0xFF) ^ (0x1B if p & 0x80 else 0)
        q ^= q << 1
        q ^= q << 2
        q ^= q << 4
        q &= 0xFF
        if q & 0x80:
            q ^= 0x09
        x = q ^ (q << 1 | q >> 7) ^ (q << 2 | q >> 6) ^ (q << 3 | q >> 5) ^ (q << 4 | q >> 4)
        sbox[p] = (x ^ 0x63) & 0xFF
        if p == 1:
            break
    sbox[0] = 0x63
    inv = [0] * 256
    for i, v in enumerate(sbox):
        inv[v] = i
    return sbox, inv


_SBOX, _INV_SBOX = _aes_tables()


def _mul(a, b):
    r = 0
    while b:
        if b & 1:
            r ^= a
        a = _xtime(a)
        b >>= 1
    return r


_MUL = {k: [_mul(i, k) for i in range(256)] for k in (9, 11, 13, 14)}


def _expand(key):
    nk = len(key) // 4
    rounds = nk + 6
    w = [list(key[4 * i:4 * i + 4]) for i in range(nk)]
    rcon = 1
    for i in range(nk, 4 * (rounds + 1)):
        t = list(w[i - 1])
        if i % nk == 0:
            t = t[1:] + t[:1]
            t = [_SBOX[b] for b in t]
            t[0] ^= rcon
            rcon = _xtime(rcon)
        elif nk > 6 and i % nk == 4:
            t = [_SBOX[b] for b in t]
        w.append([a ^ b for a, b in zip(w[i - nk], t)])
    return [sum(w[4 * r:4 * r + 4], []) for r in range(rounds + 1)]


def _aes_decrypt_block(block, keys):
    s = [b ^ k for b, k in zip(block, keys[-1])]
    m9, m11, m13, m14 = _MUL[9], _MUL[11], _MUL[13], _MUL[14]
    for rnd in range(len(keys) - 2, -1, -1):
        s = [s[0], s[13], s[10], s[7], s[4], s[1], s[14], s[11], s[8], s[5], s[2], s[15], s[12], s[9], s[6], s[3]]
        s = [_INV_SBOX[b] for b in s]
        s = [b ^ k for b, k in zip(s, keys[rnd])]
        if rnd:
            out = []
            for c in range(4):
                a0, a1, a2, a3 = s[4 * c:4 * c + 4]
                out += [m14[a0] ^ m11[a1] ^ m13[a2] ^ m9[a3], m9[a0] ^ m14[a1] ^ m11[a2] ^ m13[a3],
                        m13[a0] ^ m9[a1] ^ m14[a2] ^ m11[a3], m11[a0] ^ m13[a1] ^ m9[a2] ^ m14[a3]]
            s = out
    return bytes(s)


def aes_cbc_decrypt(key, data, iv=None, unpad=True):
    keys = _expand(key)
    if iv is None:
        iv, data = data[:16], data[16:]
    out = bytearray()
    prev = iv
    for i in range(0, len(data) - len(data) % 16, 16):
        block = data[i:i + 16]
        out += bytes(a ^ b for a, b in zip(_aes_decrypt_block(block, keys), prev))
        prev = block
    if unpad and out and 1 <= out[-1] <= 16:
        out = out[:-out[-1]]
    return bytes(out)


def _aes_cbc_encrypt_nopad(key, data, iv):
    """AES encryption (for the R6 password hash only): uses the forward cipher."""
    keys = _expand(key)
    out = bytearray()
    prev = iv
    for i in range(0, len(data), 16):
        block = bytes(a ^ b for a, b in zip(data[i:i + 16], prev))
        enc = _aes_encrypt_block(block, keys)
        out += enc
        prev = enc
    return bytes(out)


def _aes_encrypt_block(block, keys):
    s = [b ^ k for b, k in zip(block, keys[0])]
    for rnd in range(1, len(keys)):
        s = [_SBOX[b] for b in s]
        s = [s[0], s[5], s[10], s[15], s[4], s[9], s[14], s[3], s[8], s[13], s[2], s[7], s[12], s[1], s[6], s[11]]
        if rnd != len(keys) - 1:
            out = []
            for c in range(4):
                a0, a1, a2, a3 = s[4 * c:4 * c + 4]
                out += [_xtime(a0) ^ _xtime(a1) ^ a1 ^ a2 ^ a3, a0 ^ _xtime(a1) ^ _xtime(a2) ^ a2 ^ a3,
                        a0 ^ a1 ^ _xtime(a2) ^ _xtime(a3) ^ a3, _xtime(a0) ^ a0 ^ a1 ^ a2 ^ _xtime(a3)]
            s = out
        s = [b ^ k for b, k in zip(s, keys[rnd])]
    return bytes(s)


class Security:
    """The standard security handler, opened with an (empty) user password."""

    def __init__(self, doc, enc, ids, password=b""):
        r = doc.resolve
        self.v, self.rev = r(enc.get("V", 0)), r(enc.get("R", 2))
        o, u = r(enc.get("O", b"")), r(enc.get("U", b""))
        p = r(enc.get("P", 0))
        self.aes, self.identity = False, False
        length = r(enc.get("Length", 40)) or 40
        if self.v >= 4:
            cf = r(enc.get("CF", {})) or {}
            stmf = r(enc.get("StmF", "Identity"))
            filt = r(cf.get(stmf, {})) or {}
            cfm = r(filt.get("CFM", "None"))
            self.identity = stmf == "Identity" or cfm == "None"
            self.aes = cfm in ("AESV2", "AESV3")
            length = (r(filt.get("Length", length // 8 if length > 16 else length)) or 16)
            length = length * 8 if length <= 32 else length
        if self.rev >= 5:
            self.key = self._key_r6(password, o, u, r(enc.get("UE", b"")), r(enc.get("OE", b"")))
            self.aes = True
            return
        n = 5 if self.rev == 2 else max(5, length // 8)
        first_id = ids[0] if ids else b""
        h = hashlib.md5((password + _PAD)[:32] + o[:32] + struct.pack("<i", p if p < 2 ** 31 else p - 2 ** 32)
                        + first_id)
        if self.rev >= 4 and r(enc.get("EncryptMetadata", True)) is False:
            h.update(b"\xff\xff\xff\xff")
        key = h.digest()[:n]
        if self.rev >= 3:
            for _ in range(50):
                key = hashlib.md5(key[:n]).digest()[:n]
        self.key = key

    def _key_r6(self, password, o, u, ue, oe):
        password = password[:127]

        def hash_r6(data, salt, udata=b""):
            k = hashlib.sha256(data + salt + udata).digest()
            if self.rev < 6:
                return k
            i = 0
            while True:
                k1 = (data + k + udata) * 64
                e = _aes_cbc_encrypt_nopad(k[:16], k1, k[16:32])
                total = sum(e[:16]) % 3
                k = (hashlib.sha256, hashlib.sha384, hashlib.sha512)[total](e).digest()
                i += 1
                if i >= 64 and e[-1] <= i - 32:
                    return k[:32]
        inter = hash_r6(password, u[40:48])
        return aes_cbc_decrypt(inter, ue[:32], iv=bytes(16), unpad=False)

    def decrypt(self, data, num, gen, stream=False):
        if self.identity or not data:
            return data
        if self.rev >= 5:
            return aes_cbc_decrypt(self.key, data)
        key = self.key + struct.pack("<I", num)[:3] + struct.pack("<I", gen)[:2] + (b"sAlT" if self.aes else b"")
        key = hashlib.md5(key).digest()[:min(16, len(self.key) + 5)]
        return aes_cbc_decrypt(key, data) if self.aes else _rc4(key, data)


# ------------------------------------------------------------------ the document

class Document:
    """A PDF file's objects and pages."""

    def __init__(self, data, password=b""):
        if isinstance(data, str):
            with open(data, "rb") as f:
                data = f.read()
        start = data.find(b"%PDF-")
        if start < 0:
            raise PDFError("Không phải tệp PDF.")
        self.data = data[start:] if start else data
        self.xref, self.trailer, self.cache, self.crypt = {}, {}, {}, None
        try:
            self._read_xref()
        except (PDFError, ValueError, IndexError, KeyError, TypeError, zlib.error):
            self.xref, self.trailer = {}, {}
        if "Root" not in self.trailer or not isinstance(self.resolve(self.trailer.get("Root")), dict):
            self._rebuild()
        enc = self.trailer.get("Encrypt")
        if enc is not None:
            enc = self.resolve(enc)
            ids = self.resolve(self.trailer.get("ID")) or [b""]
            if not isinstance(enc, dict) or self.resolve(enc.get("Filter")) != "Standard":
                raise PDFError("PDF được mã hóa bằng cách PyExcel chưa đọc được.")
            self.crypt = Security(self, enc, [self.resolve(i) for i in ids])
            self.cache.clear()

    # cross-reference
    def _read_xref(self):
        data = self.data
        at = data.rfind(b"startxref")
        if at < 0:
            raise PDFError("no startxref")
        offset = int(re.match(rb"startxref\s+(\d+)", data[at:at + 40]).group(1))
        seen = set()
        while offset is not None and offset not in seen and 0 <= offset < len(data):
            seen.add(offset)
            pos = _SKIP.match(data, offset).end()
            if data.startswith(b"xref", pos):
                trailer = self._xref_table(pos + 4)
            else:
                trailer = self._xref_stream(pos)
            for k, v in trailer.items():
                self.trailer.setdefault(k, v)
            if isinstance(trailer.get("XRefStm"), int):
                try:
                    self._xref_stream(trailer["XRefStm"])
                except (PDFError, ValueError, IndexError):
                    pass
            prev = trailer.get("Prev")
            offset = prev if isinstance(prev, int) else None

    def _xref_table(self, pos):
        data = self.data
        sub = re.compile(rb"\s*(\d+)\s+(\d+)\s*[\r\n]")
        entry = re.compile(rb"\s*(\d{1,10})\s+(\d{1,5})\s+([nf])")
        while True:
            m = sub.match(data, pos)
            if not m:
                break
            first, count = int(m.group(1)), int(m.group(2))
            pos = m.end()
            for i in range(count):
                e = entry.match(data, pos)
                if not e:
                    break
                pos = e.end()
                if e.group(3) == b"n":
                    self.xref.setdefault(first + i, ("at", int(e.group(1)), int(e.group(2))))
                else:
                    self.xref.setdefault(first + i, None)
        t = data.find(b"trailer", pos)
        if t < 0:
            return {}
        lex = Lexer(data, t + 7)
        d = lex.token()
        return d if isinstance(d, dict) else {}

    def _xref_stream(self, pos):
        num, gen, obj = self._object_at(pos)
        if not isinstance(obj, Stream):
            raise PDFError("bad xref stream")
        d = obj.dict
        widths = d.get("W", [1, 2, 1])
        size = d.get("Size", 0)
        index = d.get("Index", [0, size])
        raw = decode(obj.raw, *obj.filters())
        step = sum(widths)
        k = 0
        for start, count in zip(index[0::2], index[1::2]):
            for i in range(count):
                row = raw[k:k + step]
                k += step
                if len(row) < step:
                    break
                vals, p = [], 0
                for w in widths:
                    vals.append(int.from_bytes(row[p:p + w], "big") if w else None)
                    p += w
                kind = vals[0] if widths[0] else 1
                if kind == 1:
                    self.xref.setdefault(start + i, ("at", vals[1], vals[2] or 0))
                elif kind == 2:
                    self.xref.setdefault(start + i, ("in", vals[1], vals[2]))
                elif kind == 0:
                    self.xref.setdefault(start + i, None)
        return d

    def _rebuild(self):
        """A damaged cross-reference table: find the objects by scanning the file."""
        self.xref, self.cache = {}, {}
        for m in re.finditer(rb"(?<![0-9])(\d+)\s+(\d+)\s+obj\b", self.data):
            self.xref[int(m.group(1))] = ("at", m.start(), int(m.group(2)))
        trailer = {}
        for m in re.finditer(rb"trailer\s*<<", self.data):
            d = Lexer(self.data, m.end() - 2).token()
            if isinstance(d, dict) and "Root" in d:
                trailer = d
        if "Root" not in trailer:
            for num in list(self.xref):
                obj = self.get(num)
                if isinstance(obj, dict) and obj.get("Type") == "Catalog":
                    trailer["Root"] = Ref(num, 0)
                    break
                if isinstance(obj, Stream) and obj.dict.get("Type") == "XRef" and "Root" in obj.dict:
                    trailer = dict(obj.dict)
        self.trailer = trailer
        if "Root" not in trailer:
            raise PDFError("Tệp PDF bị hỏng: không tìm thấy danh sách trang.")

    def _object_at(self, pos):
        lex = Lexer(self.data, pos)
        num, gen, kw = lex.token(), lex.token(), lex.token()
        if not (isinstance(num, int) and isinstance(gen, int) and kw == "obj"):
            raise PDFError(f"no object at {pos}")
        obj = lex.token()
        lex.skip()
        if isinstance(obj, dict) and self.data.startswith(b"stream", lex.pos):
            start = lex.pos + 6
            if self.data[start:start + 2] == b"\r\n":
                start += 2
            elif self.data[start:start + 1] in (b"\n", b"\r"):
                start += 1
            length = obj.get("Length")
            if isinstance(length, Ref):
                try:
                    length = self.resolve(length)
                except (PDFError, RecursionError):
                    length = None
            end = start + length if isinstance(length, int) and length >= 0 else -1
            if end < 0 or not re.match(rb"\s*endstream", self.data[end:end + 20]):
                found = self.data.find(b"endstream", start)
                end = found if found >= 0 else len(self.data)
                while end > start and self.data[end - 1] in b"\r\n":
                    end -= 1
            obj = Stream(obj, self.data[start:end], self, num, gen)
        return num, gen, obj

    def get(self, num):
        if num in self.cache:
            return self.cache[num]
        self.cache[num] = None  # (a loop of references ends here)
        entry = self.xref.get(num)
        obj = None
        try:
            if entry and entry[0] == "at":
                n, gen, obj = self._object_at(entry[1])
                if n != num:
                    obj = self._search(num)
            elif entry and entry[0] == "in":
                obj = self._from_object_stream(entry[1], entry[2], num)
        except (PDFError, ValueError, IndexError, TypeError):
            obj = self._search(num)
        self.cache[num] = obj
        return obj

    def _search(self, num):
        m = re.search(rb"(?<![0-9])%d\s+\d+\s+obj\b" % num, self.data)
        if not m:
            return None
        try:
            return self._object_at(m.start())[2]
        except PDFError:
            return None

    def _from_object_stream(self, stm_num, index, num):
        stm = self.get(stm_num)
        if not isinstance(stm, Stream):
            return None
        key = ("objstm", stm_num)
        parsed = self.cache.get(key)
        if parsed is None:
            data = stm.data()
            first = stm.dict.get("First", 0)
            n = stm.dict.get("N", 0)
            lex = Lexer(data, 0)
            pairs = [(lex.token(), lex.token()) for _ in range(n)]
            parsed = self.cache[key] = (data, first, pairs)
        data, first, pairs = parsed
        for objnum, off in pairs:
            if objnum == num and isinstance(off, int):
                return Lexer(data, first + off).token()
        if index < len(pairs) and isinstance(pairs[index][1], int):
            return Lexer(data, first + pairs[index][1]).token()
        return None

    def resolve(self, obj, depth=0):
        while isinstance(obj, Ref) and depth < 32:
            obj = self.get(obj.num)
            depth += 1
        return obj

    # pages
    def pages(self):
        root = self.resolve(self.trailer.get("Root"))
        out, seen = [], set()

        def walk(node, inherited):
            if isinstance(node, Ref):
                if node.num in seen:
                    return
                seen.add(node.num)
            node = self.resolve(node)
            if not isinstance(node, dict):
                return
            inh = dict(inherited)
            for k in ("Resources", "MediaBox", "CropBox", "Rotate"):
                if k in node:
                    inh[k] = node[k]
            kids = self.resolve(node.get("Kids"))
            if node.get("Type") == "Pages" or isinstance(kids, list):
                for kid in kids or []:
                    walk(kid, inh)
            else:
                out.append(Page(self, node, inh))
        walk(root.get("Pages") if isinstance(root, dict) else None, {})
        return out


# ------------------------------------------------------------------ fonts

_ACCENTS = {"acute": "́", "grave": "̀", "circumflex": "̂", "tilde": "̃", "dieresis": "̈",
            "ring": "̊", "cedilla": "̧", "caron": "̌", "macron": "̄", "breve": "̆",
            "ogonek": "̨", "dotaccent": "̇", "hungarumlaut": "̋", "hookabove": "̉",
            "dotbelow": "̣", "horn": "̛"}
_GLYPHS = {
    "space": " ", "exclam": "!", "quotedbl": '"', "numbersign": "#", "dollar": "$", "percent": "%", "ampersand": "&",
    "quotesingle": "'", "quoteright": "’", "parenleft": "(", "parenright": ")", "asterisk": "*", "plus": "+",
    "comma": ",", "hyphen": "-", "period": ".", "slash": "/", "zero": "0", "one": "1", "two": "2", "three": "3",
    "four": "4", "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9", "colon": ":", "semicolon": ";",
    "less": "<", "equal": "=", "greater": ">", "question": "?", "at": "@", "bracketleft": "[", "backslash": "\\",
    "bracketright": "]", "asciicircum": "^", "underscore": "_", "grave": "`", "quoteleft": "‘",
    "braceleft": "{", "bar": "|", "braceright": "}", "asciitilde": "~", "bullet": "•", "endash": "–",
    "emdash": "—", "quotedblleft": "“", "quotedblright": "”", "quotesinglbase": "‚",
    "quotedblbase": "„", "ellipsis": "…", "dagger": "†", "daggerdbl": "‡",
    "perthousand": "‰", "trademark": "™", "copyright": "©", "registered": "®",
    "degree": "°", "plusminus": "±", "multiply": "×", "divide": "÷", "section": "§",
    "paragraph": "¶", "periodcentered": "·", "cent": "¢", "sterling": "£", "yen": "¥",
    "Euro": "€", "fi": "fi", "fl": "fl", "ff": "ff", "ffi": "ffi", "ffl": "ffl", "dotlessi": "ı",
    "germandbls": "ß", "ae": "æ", "AE": "Æ", "oe": "œ", "OE": "Œ", "oslash": "ø",
    "Oslash": "Ø", "eth": "ð", "Eth": "Ð", "thorn": "þ", "Thorn": "Þ", "minus": "−",
    "fraction": "⁄", "exclamdown": "¡", "questiondown": "¿", "guillemotleft": "«",
    "guillemotright": "»", "guilsinglleft": "‹", "guilsinglright": "›", "mu": "µ",
    "onehalf": "½", "onequarter": "¼", "threequarters": "¾", "logicalnot": "¬",
    "brokenbar": "¦", "nbspace": " ", "dcroat": "đ", "Dcroat": "Đ", "dmacron": "đ",
    "Dmacron": "Đ", "florin": "ƒ", "lslash": "ł", "Lslash": "Ł", "ordfeminine": "ª",
    "ordmasculine": "º", "onesuperior": "¹", "twosuperior": "²", "threesuperior": "³",
    "currency": "¤", "dieresis": "¨", "acute": "´", "cedilla": "¸", "macron": "¯",
}


def glyph_text(name):
    """The text of a glyph name (Adobe's names, uniXXXX, accented letters such as "ohornhookabove")."""
    if name in _GLYPHS:
        return _GLYPHS[name]
    base = name.split(".")[0].split("_")[0]
    if base in _GLYPHS:
        return _GLYPHS[base]
    m = re.fullmatch(r"uni((?:[0-9A-Fa-f]{4})+)", base)
    if m:
        hexs = m.group(1)
        return "".join(chr(int(hexs[i:i + 4], 16)) for i in range(0, len(hexs), 4))
    m = re.fullmatch(r"u([0-9A-Fa-f]{4,6})", base)
    if m:
        return chr(int(m.group(1), 16))
    if len(base) == 1:
        return base
    m = re.fullmatch(r"([A-Za-z])((?:" + "|".join(sorted(_ACCENTS, key=len, reverse=True)) + r")+)", base)
    if m:
        marks = re.findall("|".join(sorted(_ACCENTS, key=len, reverse=True)), m.group(2))
        return unicodedata.normalize("NFC", m.group(1) + "".join(_ACCENTS[a] for a in marks))
    return ""


def _standard_encoding():
    enc = {i: chr(i) for i in range(32, 127)}
    enc[0x27], enc[0x60] = "’", "‘"
    extra = {0xA1: "¡", 0xA2: "¢", 0xA3: "£", 0xA4: "⁄", 0xA5: "¥", 0xA6: "ƒ", 0xA7: "§", 0xA8: "¤",
             0xA9: "'", 0xAA: "“", 0xAB: "«", 0xAC: "‹", 0xAD: "›", 0xAE: "fi", 0xAF: "fl",
             0xB1: "–", 0xB2: "†", 0xB3: "‡", 0xB4: "·", 0xB6: "¶", 0xB7: "•", 0xB8: "‚",
             0xB9: "„", 0xBA: "”", 0xBB: "»", 0xBC: "…", 0xBD: "‰", 0xBF: "¿", 0xD0: "—",
             0xE1: "Æ", 0xE8: "Ł", 0xE9: "Ø", 0xEA: "Œ", 0xF1: "æ", 0xF5: "ı", 0xF8: "ł", 0xF9: "ø", 0xFA: "œ",
             0xFB: "ß"}
    enc.update(extra)
    return enc


_ENCODINGS = {
    "WinAnsiEncoding": {i: bytes([i]).decode("cp1252", "replace") for i in range(32, 256)},
    "MacRomanEncoding": {i: bytes([i]).decode("mac_roman", "replace") for i in range(32, 256)},
    "StandardEncoding": _standard_encoding(),
}
_ENCODINGS["WinAnsiEncoding"].update({0x7F: "•", 0x81: "•", 0x8D: "•", 0x8F: "•",
                                      0x90: "•", 0x9D: "•", 0xA0: " ", 0xAD: "-"})
_CJK_CODECS = (("RKSJ", "cp932"), ("EUC-H", None), ("GBK", "gbk"), ("GB-EUC", "gb2312"), ("B5", "cp950"),
               ("ETen", "cp950"), ("UHC", "cp949"), ("KSC", "cp949"))
_STD_WIDTHS = {"Courier": 600, "Helvetica": 556, "Arial": 556, "Times": 500, "Symbol": 500, "ZapfDingbats": 700}


def _codec_for(cmap_name):
    if "UCS2" in cmap_name or "UTF16" in cmap_name:
        return "utf-16-be"
    for key, codec in _CJK_CODECS:
        if key in cmap_name:
            if codec is None:
                return "euc_jp" if "JIS" in cmap_name or cmap_name.startswith(("H", "EUC")) else "gb2312"
            return codec
    return None


class CMap:
    """A CMap: the byte lengths of codes (codespace ranges), and codes to Unicode text (ToUnicode) or to
    CIDs (an encoding CMap)."""

    def __init__(self, data=b""):
        self.ranges = []  # [(length, low bytes, high bytes)]
        self.text, self.cids = {}, {}
        if data:
            self._parse(data)

    def _parse(self, data):
        lex = Lexer(data, refs=False)
        stack, mode = [], None
        while True:
            t = lex.token()
            if t is None:
                break
            if isinstance(t, Keyword):
                if t in ("begincodespacerange", "beginbfchar", "beginbfrange", "begincidrange", "begincidchar",
                         "beginnotdefrange"):
                    mode, stack = t[5:], []
                elif t.startswith("end") and mode is not None:
                    self._apply(mode, stack)
                    mode = None
                continue
            if mode is not None:
                stack.append(t)

    def _apply(self, mode, items):
        if mode == "codespacerange":
            for lo, hi in zip(items[0::2], items[1::2]):
                if isinstance(lo, bytes) and isinstance(hi, bytes) and lo:
                    self.ranges.append((len(lo), lo, hi))
        elif mode == "bfchar":
            for src, dst in zip(items[0::2], items[1::2]):
                if isinstance(src, bytes):
                    self.text[int.from_bytes(src, "big")] = _target(dst)
        elif mode == "bfrange":
            for i in range(0, len(items) - 2, 3):
                lo, hi, dst = items[i:i + 3]
                if not (isinstance(lo, bytes) and isinstance(hi, bytes)):
                    continue
                a, b = int.from_bytes(lo, "big"), int.from_bytes(hi, "big")
                if b - a > 65535:
                    continue
                if isinstance(dst, list):
                    for k, d in enumerate(dst[:b - a + 1]):
                        self.text[a + k] = _target(d)
                elif isinstance(dst, bytes) and dst:
                    head, last = dst[:-1], dst[-1]
                    for k in range(b - a + 1):
                        if last + k > 255:  # (carries into the byte before)
                            v = int.from_bytes(dst, "big") + k
                            self.text[a + k] = v.to_bytes(len(dst), "big").decode("utf-16-be", "replace")
                        else:
                            self.text[a + k] = (head + bytes([last + k])).decode("utf-16-be", "replace")
        elif mode in ("cidrange", "notdefrange"):
            for i in range(0, len(items) - 2, 3):
                lo, hi, cid = items[i:i + 3]
                if isinstance(lo, bytes) and isinstance(hi, bytes) and isinstance(cid, int):
                    a, b = int.from_bytes(lo, "big"), int.from_bytes(hi, "big")
                    if b - a <= 65535:
                        for k in range(b - a + 1):
                            self.cids[a + k] = cid + k
        elif mode == "cidchar":
            for src, cid in zip(items[0::2], items[1::2]):
                if isinstance(src, bytes) and isinstance(cid, int):
                    self.cids[int.from_bytes(src, "big")] = cid

    def split(self, data, default=1):
        """The codes in a string's bytes: [(code, length)]."""
        out, i, n = [], 0, len(data)
        ranges = self.ranges
        while i < n:
            length = 0
            for size, lo, hi in ranges:
                part = data[i:i + size]
                if len(part) == size and all(lo[k] <= part[k] <= hi[k] for k in range(size)):
                    length = size
                    break
            if not length:
                length = default if not ranges else min(r[0] for r in ranges)
            part = data[i:i + length]
            out.append((int.from_bytes(part, "big"), part))
            i += max(1, length)
        return out


def _target(dst):
    if isinstance(dst, bytes):
        return dst.decode("utf-16-be", "replace") if len(dst) >= 2 else dst.decode("latin-1")
    if isinstance(dst, Name):
        return glyph_text(dst)
    return ""


class Font:
    """What a page needs from a font: the text of its codes and their widths (in 1/1000 of the size)."""

    def __init__(self, doc, d):
        r = doc.resolve
        self.subtype = r(d.get("Subtype"))
        base = str(r(d.get("BaseFont")) or "")
        self.base = base.split("+", 1)[1] if re.match(r"[A-Z]{6}\+", base) else base
        self.composite = self.subtype == "Type0"
        desc = {}
        self.tounicode = None
        tu = r(d.get("ToUnicode"))
        if isinstance(tu, Stream):
            try:
                self.tounicode = CMap(tu.data())
            except (PDFError, ValueError, IndexError, TypeError):
                self.tounicode = None
        self.widths, self.default_width = {}, 1000 if self.composite else 500
        self.vertical = False
        self.codec, self.cmap = None, None
        self.scale = 1.0
        if self.composite:
            enc = r(d.get("Encoding"))
            if isinstance(enc, Stream):
                self.cmap = CMap(enc.data())
            else:
                name = str(enc or "Identity-H")
                self.vertical = name.endswith("-V")
                self.codec = _codec_for(name)
                self.cmap = CMap()
                if name.startswith("Identity") or "UCS2" in name or "UTF16" in name:
                    self.cmap.ranges = [(2, b"\x00\x00", b"\xff\xff")]
                elif self.codec:
                    self.cmap.ranges = _CODEC_RANGES.get(self.codec, [(1, b"\x00", b"\x7f"), (2, b"\x80\x00", b"\xff\xff")])
            desc_fonts = r(d.get("DescendantFonts")) or []
            cid = r(desc_fonts[0]) if desc_fonts else {}
            cid = cid if isinstance(cid, dict) else {}
            desc = r(cid.get("FontDescriptor")) or {}
            self.default_width = r(cid.get("DW", 1000)) or 1000
            w = r(cid.get("W")) or []
            i = 0
            while i < len(w) - 1:
                first = r(w[i])
                nxt = r(w[i + 1])
                if isinstance(nxt, list):
                    for k, v in enumerate(nxt):
                        self.widths[first + k] = r(v)
                    i += 2
                elif i + 2 < len(w):
                    last, v = nxt, r(w[i + 2])
                    if isinstance(last, int) and isinstance(first, int) and last - first < 65536:
                        for c in range(first, last + 1):
                            self.widths[c] = v
                    i += 3
                else:
                    break
        else:
            desc = r(d.get("FontDescriptor")) or {}
            first = r(d.get("FirstChar", 0)) or 0
            for k, v in enumerate(r(d.get("Widths")) or []):
                v = r(v)
                if isinstance(v, (int, float)):
                    self.widths[first + k] = v
            missing = r(desc.get("MissingWidth")) if isinstance(desc, dict) else None
            std = next((w for k, w in _STD_WIDTHS.items() if k.lower() in self.base.lower()), 500)
            self.default_width = missing if isinstance(missing, (int, float)) and missing else std
            if self.subtype == "Type3":
                matrix = r(d.get("FontMatrix")) or [0.001]
                self.scale = (matrix[0] if matrix else 0.001) * 1000
            self.encoding = self._simple_encoding(doc, d, desc)
        desc = desc if isinstance(desc, dict) else {}
        flags = r(desc.get("Flags", 0)) or 0
        weight = r(desc.get("FontWeight", 400)) or 400
        name = self.base.lower()
        self.bold = "bold" in name or "black" in name or "heavy" in name or weight >= 600 or bool(flags & (1 << 18))
        self.italic = "italic" in name or "oblique" in name or bool(r(desc.get("ItalicAngle", 0)))
        self.symbolic = bool(flags & 4) and not (flags & 32)

    def _simple_encoding(self, doc, d, desc):
        r = doc.resolve
        enc = r(d.get("Encoding"))
        base = "StandardEncoding" if self.subtype in ("Type1", "MMType1") else "WinAnsiEncoding"
        diffs = []
        if isinstance(enc, Name):
            base = enc if enc in _ENCODINGS else base
        elif isinstance(enc, dict):
            b = r(enc.get("BaseEncoding"))
            if b in _ENCODINGS:
                base = b
            diffs = r(enc.get("Differences")) or []
        table = dict(_ENCODINGS[base])
        code = 0
        for item in diffs:
            item = r(item)
            if isinstance(item, int):
                code = item
            elif isinstance(item, Name):
                text = glyph_text(item)
                if text:
                    table[code] = text
                else:
                    table.pop(code, None)
                code += 1
        return table

    def codes(self, data):
        """[(code, text, width in 1/1000, is a one-byte space)] of a shown string."""
        out = []
        if self.composite:
            for code, raw in self.cmap.split(data, 2):
                text = self.tounicode.text.get(code) if self.tounicode else None
                if text is None:
                    if self.codec:
                        try:
                            text = raw.decode(self.codec)
                        except UnicodeDecodeError:
                            text = ""
                    else:
                        text = ""
                cid = self.cmap.cids.get(code, code) if self.cmap.cids else code
                width = self.widths.get(cid)
                if width is None:
                    width = 500 if (self.codec and len(raw) == 1) else self.default_width
                out.append((code, text, width, False))
            return out
        split = self.tounicode.split(data, 1) if self.tounicode and self.tounicode.ranges else \
            [(b, bytes([b])) for b in data]
        for code, raw in split:
            text = self.tounicode.text.get(code) if self.tounicode else None
            if text is None:
                text = self.encoding.get(code, "")
                if not text and self.symbolic and 32 <= (code & 0xFF) < 127:
                    text = chr(code & 0xFF)
            width = self.widths.get(code, self.default_width) * self.scale
            out.append((code, text, width, len(raw) == 1 and code == 32))
        return out


_CODEC_RANGES = {
    "cp932": [(1, b"\x00", b"\x80"), (2, b"\x81\x40", b"\x9f\xfc"), (1, b"\xa0", b"\xdf"), (2, b"\xe0\x40", b"\xfc\xfc")],
    "gbk": [(1, b"\x00", b"\x80"), (2, b"\x81\x40", b"\xfe\xfe")],
    "gb2312": [(1, b"\x00", b"\x80"), (2, b"\xa1\xa1", b"\xfe\xfe")],
    "cp950": [(1, b"\x00", b"\x80"), (2, b"\x81\x40", b"\xfe\xfe")],
    "cp949": [(1, b"\x00", b"\x80"), (2, b"\x81\x41", b"\xfe\xfe")],
    "euc_jp": [(1, b"\x00", b"\x80"), (2, b"\x8e\xa0", b"\x8e\xdf"), (2, b"\xa1\xa1", b"\xfe\xfe")],
}


# ------------------------------------------------------------------ pages

Span = namedtuple("Span", "x0 y0 x1 y1 size text font bold italic color angle")
# a run of text: from (x0, y0) to (x1, y1) along its baseline (page points, y up), size in points
Path = namedtuple("Path", "subpaths stroke fill width evenodd")
# subpaths: [(points, closed)] in page points; stroke / fill: "#rrggbb" or None; width: points
Image = namedtuple("Image", "box stream")  # box: (x0, y0, x1, y1) in page points


def _mult(m, n):
    a, b, c, d, e, f = m
    a2, b2, c2, d2, e2, f2 = n
    return (a * a2 + b * c2, a * b2 + b * d2, c * a2 + d * c2, c * b2 + d * d2, e * a2 + f * c2 + e2, e * b2 + f * d2 + f2)


def _apply(m, x, y):
    return m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]


def _hex(r, g, b):
    return "#%02x%02x%02x" % tuple(max(0, min(255, round(v * 255))) for v in (r, g, b))


def _color(space, values):
    vals = [v for v in values if isinstance(v, (int, float))]
    n = space[1] if isinstance(space, tuple) else {"DeviceGray": 1, "G": 1, "CalGray": 1, "DeviceRGB": 3, "RGB": 3,
                                                   "CalRGB": 3, "Lab": 3, "DeviceCMYK": 4, "CMYK": 4}.get(space, len(vals))
    kind = space[0] if isinstance(space, tuple) else space
    if kind in ("Separation", "DeviceN") and vals:  # a spot colour's tint: shades of grey
        g = 1 - min(1.0, max(vals))
        return _hex(g, g, g)
    if kind == "Indexed":
        return "#000000"
    if n == 1 and vals:
        return _hex(vals[0], vals[0], vals[0])
    if n == 3 and len(vals) >= 3:
        if kind == "Lab":
            g = vals[0] / 100
            return _hex(g, g, g)
        return _hex(*vals[:3])
    if n == 4 and len(vals) >= 4:
        c, m, y, k = vals[:4]
        return _hex((1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))
    return "#000000"


class Page:
    def __init__(self, doc, d, inherited):
        self.doc, self.dict = doc, d
        r = doc.resolve
        box = r(inherited.get("CropBox")) or r(inherited.get("MediaBox")) or [0, 0, 612, 792]
        box = [r(v) for v in box][:4] if isinstance(box, list) and len(box) >= 4 else [0, 0, 612, 792]
        self.box = (min(box[0], box[2]), min(box[1], box[3]), max(box[0], box[2]), max(box[1], box[3]))
        self.rotate = (r(inherited.get("Rotate", 0)) or 0) % 360
        self.resources = r(inherited.get("Resources")) or {}

    @property
    def width(self):
        return self.box[2] - self.box[0] if self.rotate in (0, 180) else self.box[3] - self.box[1]

    @property
    def height(self):
        return self.box[3] - self.box[1] if self.rotate in (0, 180) else self.box[2] - self.box[0]

    def contents(self):
        c = self.doc.resolve(self.dict.get("Contents"))
        parts = c if isinstance(c, list) else [c]
        out = []
        for part in parts:
            part = self.doc.resolve(part)
            if isinstance(part, Stream):
                out.append(part.data())
        return b"\n".join(out)

    def base_matrix(self):
        """From PDF user space to the page as shown, y going up from its bottom-left corner (Rotate applied)."""
        x0, y0, x1, y1 = self.box
        w, h = x1 - x0, y1 - y0
        m = (1, 0, 0, 1, -x0, -y0)
        if self.rotate == 90:
            m = _mult(m, (0, -1, 1, 0, 0, w))
        elif self.rotate == 180:
            m = _mult(m, (-1, 0, 0, -1, w, h))
        elif self.rotate == 270:
            m = _mult(m, (0, 1, -1, 0, h, 0))
        return m

    def read(self, limit=400000):
        """(spans, paths, images) drawn on the page; `limit` caps the operators read (huge drawings)."""
        reader = _Reader(self.doc, limit)
        reader.run(self.contents(), self.resources, self.base_matrix())
        return reader.spans, reader.paths, reader.images


class _State:
    __slots__ = ("ctm", "stroke", "fill", "stroke_space", "fill_space", "width", "font", "size", "tc", "tw", "th",
                 "tl", "rise", "mode")

    def __init__(self):
        self.ctm = (1, 0, 0, 1, 0, 0)
        self.stroke = self.fill = "#000000"
        self.stroke_space = self.fill_space = "DeviceGray"
        self.width = 1.0
        self.font, self.size = None, 0.0
        self.tc = self.tw = 0.0
        self.th = 1.0
        self.tl = self.rise = 0.0
        self.mode = 0

    def copy(self):
        s = _State.__new__(_State)
        for k in _State.__slots__:
            setattr(s, k, getattr(self, k))
        return s


_CONTENT = re.compile(
    rb"(?P<ws>(?:[\x00\t\n\x0c\r ]+|%[^\r\n]*)+)|(?P<num>[+-]?(?:\d+\.?\d*|\.\d+))(?![^\x00\t\n\x0c\r ()<>\[\]{}/%])"
    rb"|/(?P<name>[^\x00\t\n\x0c\r ()<>\[\]{}/%]*)|(?P<open>\()|(?P<dict><<)|(?P<hex><)|(?P<lb>\[)|(?P<rb>\])"
    rb"|(?P<dend>>>)|(?P<op>[^\x00\t\n\x0c\r ()<>\[\]{}/%]+)|(?P<other>.)", re.S)


class _Reader:
    """Runs a content stream, collecting text spans, painted paths and pictures."""

    def __init__(self, doc, limit):
        self.doc, self.limit = doc, limit
        self.spans, self.paths, self.images = [], [], []
        self.fonts = {}
        self.ops = 0
        self.depth = 0

    def font(self, resources, name):
        fonts = self.doc.resolve(resources.get("Font")) or {}
        ref = fonts.get(name)
        key = ref if isinstance(ref, Ref) else (id(fonts), name)
        f = self.fonts.get(key)
        if f is None:
            d = self.doc.resolve(ref)
            try:
                f = Font(self.doc, d) if isinstance(d, dict) else None
            except (PDFError, ValueError, TypeError, IndexError, KeyError, AttributeError):
                f = None
            self.fonts[key] = f
        return f

    def run(self, data, resources, ctm):
        st = _State()
        st.ctm = ctm
        self._run(data, resources, st)

    def _run(self, data, resources, st):
        r = self.doc.resolve
        resources = r(resources) or {}
        stack = []
        operands = []
        path, current, start = [], [], None
        tm = tlm = (1, 0, 0, 1, 0, 0)
        pos, n = 0, len(data)
        arrays = []
        clip_pending = False
        while pos < n:
            m = _CONTENT.match(data, pos)
            if not m:
                break
            pos = m.end()
            kind = m.lastgroup
            if kind == "ws" or kind == "other":
                continue
            if kind == "num":
                text = m.group("num")
                operands.append(float(text) if b"." in text else int(text))
            elif kind == "name":
                operands.append(_name(m.group("name")))
            elif kind == "open":
                v, pos = _literal(data, pos)
                operands.append(v)
            elif kind == "hex":
                v, pos = _hexstring(data, pos)
                operands.append(v)
            elif kind == "dict":
                lex = Lexer(data, pos, refs=False)
                operands.append(lex._dict())
                pos = lex.pos
            elif kind == "lb":
                arrays.append(operands)
                operands = []
                continue
            elif kind == "rb":
                if arrays:
                    inner = operands
                    operands = arrays.pop()
                    operands.append(inner)
                continue
            elif kind == "dend":
                continue
            else:
                op = m.group("op").decode("latin-1")
                if op in ("true", "false", "null"):
                    operands.append({"true": True, "false": False, "null": None}[op])
                    continue
                self.ops += 1
                if self.ops > self.limit:
                    return
                args = operands
                operands = []
                try:
                    if op == "BI":  # an inline image: skipped
                        end = re.compile(rb"\sEI(?=[\x00\t\n\x0c\r /]|$)").search(data, pos)
                        pos = end.end() if end else n
                        continue
                    # graphics state
                    if op == "q":
                        stack.append(st.copy())
                    elif op == "Q":
                        if stack:
                            st = stack.pop()
                    elif op == "cm" and len(args) == 6:
                        st.ctm = _mult(tuple(args), st.ctm)
                    elif op == "w" and args:
                        st.width = float(args[0])
                    elif op == "gs" and args:
                        gs = r((r(resources.get("ExtGState")) or {}).get(args[0])) or {}
                        if "LW" in gs:
                            st.width = float(r(gs["LW"]))
                    # colours
                    elif op in ("g", "G", "rg", "RG", "k", "K"):
                        space = {"g": "DeviceGray", "rg": "DeviceRGB", "k": "DeviceCMYK"}[op.lower()]
                        color = _color(space, args)
                        if op.islower():
                            st.fill, st.fill_space = color, space
                        else:
                            st.stroke, st.stroke_space = color, space
                    elif op in ("cs", "CS") and args:
                        space = self._space(resources, args[0])
                        if op == "cs":
                            st.fill_space, st.fill = space, "#000000"
                        else:
                            st.stroke_space, st.stroke = space, "#000000"
                    elif op in ("sc", "scn"):
                        if args and not isinstance(args[-1], Name):
                            st.fill = _color(st.fill_space, args)
                    elif op in ("SC", "SCN"):
                        if args and not isinstance(args[-1], Name):
                            st.stroke = _color(st.stroke_space, args)
                    # paths
                    elif op == "m" and len(args) >= 2:
                        if len(current) > 1:
                            path.append((current, False))
                        start = _apply(st.ctm, args[0], args[1])
                        current = [start]
                    elif op == "l" and len(args) >= 2:
                        p = _apply(st.ctm, args[0], args[1])
                        if not current:
                            current = [p]
                        else:
                            current.append(p)
                    elif op in ("c", "v", "y") and current:
                        p0 = current[-1]
                        ctm = st.ctm
                        if op == "c" and len(args) >= 6:
                            c1, c2, p3 = _apply(ctm, *args[0:2]), _apply(ctm, *args[2:4]), _apply(ctm, *args[4:6])
                        elif op == "v" and len(args) >= 4:
                            c1, c2, p3 = p0, _apply(ctm, *args[0:2]), _apply(ctm, *args[2:4])
                        elif op == "y" and len(args) >= 4:
                            c1, c2, p3 = _apply(ctm, *args[0:2]), _apply(ctm, *args[2:4]), _apply(ctm, *args[2:4])
                        else:
                            continue
                        steps = 8
                        for i in range(1, steps + 1):
                            t = i / steps
                            u = 1 - t
                            current.append((u * u * u * p0[0] + 3 * u * u * t * c1[0] + 3 * u * t * t * c2[0] + t * t * t * p3[0],
                                            u * u * u * p0[1] + 3 * u * u * t * c1[1] + 3 * u * t * t * c2[1] + t * t * t * p3[1]))
                    elif op == "h":
                        if len(current) > 1:
                            path.append((current, True))
                            current = [current[0]]
                    elif op == "re" and len(args) >= 4:
                        if len(current) > 1:
                            path.append((current, False))
                        x, y, w, h = args[:4]
                        ctm = st.ctm
                        path.append(([_apply(ctm, x, y), _apply(ctm, x + w, y), _apply(ctm, x + w, y + h),
                                      _apply(ctm, x, y + h)], True))
                        current = []
                    elif op in ("S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "n"):
                        if len(current) > 1:
                            path.append((current, op in ("s", "b", "b*")))
                        if op != "n" and path and not clip_pending:
                            stroke = st.stroke if op in ("S", "s", "B", "B*", "b", "b*") else None
                            fill = st.fill if op in ("f", "F", "f*", "B", "B*", "b", "b*") else None
                            a, b, c, d = st.ctm[:4]
                            scale = math.sqrt(abs(a * d - b * c)) or 1.0
                            subpaths = [(pts, closed or (fill is not None)) for pts, closed in path]
                            self.paths.append(Path(subpaths, stroke, fill, st.width * scale, op.endswith("*")))
                        path, current, clip_pending = [], [], False
                    elif op in ("W", "W*"):
                        clip_pending = False  # (clipping is not followed; the path still paints if asked)
                    # text
                    elif op == "BT":
                        tm = tlm = (1, 0, 0, 1, 0, 0)
                    elif op == "Tf" and len(args) >= 2:
                        st.font, st.size = self.font(resources, args[0]), float(args[1])
                    elif op == "Tc" and args:
                        st.tc = float(args[0])
                    elif op == "Tw" and args:
                        st.tw = float(args[0])
                    elif op == "Tz" and args:
                        st.th = float(args[0]) / 100
                    elif op == "TL" and args:
                        st.tl = float(args[0])
                    elif op == "Ts" and args:
                        st.rise = float(args[0])
                    elif op == "Tr" and args:
                        st.mode = int(args[0])
                    elif op in ("Td", "TD") and len(args) >= 2:
                        if op == "TD":
                            st.tl = -float(args[1])
                        tlm = _mult((1, 0, 0, 1, float(args[0]), float(args[1])), tlm)
                        tm = tlm
                    elif op == "Tm" and len(args) >= 6:
                        tlm = tm = tuple(float(v) for v in args[:6])
                    elif op == "T*":
                        tlm = _mult((1, 0, 0, 1, 0, -st.tl), tlm)
                        tm = tlm
                    elif op in ("Tj", "'", '"', "TJ") and args:
                        if op == "'":
                            tlm = _mult((1, 0, 0, 1, 0, -st.tl), tlm)
                            tm = tlm
                        elif op == '"' and len(args) >= 3:
                            st.tw, st.tc = float(args[0]), float(args[1])
                            tlm = _mult((1, 0, 0, 1, 0, -st.tl), tlm)
                            tm = tlm
                        items = args[-1] if op == "TJ" else [args[-1]]
                        if isinstance(items, list):
                            tm = self._show(st, tm, items)
                    # pictures and forms
                    elif op == "Do" and args:
                        self._do(resources, st, args[0])
                except (TypeError, ValueError, ZeroDivisionError, IndexError, AttributeError, OverflowError):
                    pass

    def _space(self, resources, name):
        r = self.doc.resolve
        if name in ("DeviceGray", "DeviceRGB", "DeviceCMYK", "Pattern", "G", "RGB", "CMYK"):
            return name
        space = r((r(resources.get("ColorSpace")) or {}).get(name)) if isinstance(name, Name) else r(name)
        if isinstance(space, list) and space:
            kind = r(space[0])
            if kind == "ICCBased" and len(space) > 1:
                s = r(space[1])
                n = r(s.dict.get("N", 3)) if isinstance(s, Stream) else 3
                return ("ICC", n)
            if kind in ("Separation", "DeviceN", "Indexed"):
                return (kind, 1)
            if kind in ("CalRGB", "Lab"):
                return (kind, 3)
            if kind == "CalGray":
                return (kind, 1)
            return (kind, 3)
        return space if isinstance(space, str) else "DeviceGray"

    def _show(self, st, tm, items):
        """Show strings (TJ numbers move along); spans break where a big move leaves a gap."""
        font = st.font
        if font is None or st.size == 0:
            return tm
        fs, th = st.size, st.th
        texts, start, last_end = [], None, None
        for item in items:
            if isinstance(item, (int, float)):
                shift = -item / 1000 * fs * th
                if font.vertical:
                    tm = _mult((1, 0, 0, 1, 0, -item / 1000 * fs), tm)
                else:
                    tm = _mult((1, 0, 0, 1, shift, 0), tm)
                if texts and item < 0 and -item > 250:  # a gap of more than a quarter of the size: a new span
                    self._emit(st, start, last_end, "".join(texts))
                    texts, start = [], None
                elif texts and item < -120:
                    texts.append(" ")
                continue
            if not isinstance(item, bytes):
                continue
            for code, text, width, space in font.codes(item):
                trm = _mult((fs * th, 0, 0, fs, 0, st.rise), _mult(tm, st.ctm))
                if start is None:
                    start = trm
                if font.vertical:
                    ty = -(width / 1000 * fs + st.tc + (st.tw if space else 0))
                    tm = _mult((1, 0, 0, 1, 0, ty), tm)
                else:
                    tx = (width / 1000 * fs + st.tc + (st.tw if space else 0)) * th
                    tm = _mult((1, 0, 0, 1, tx, 0), tm)
                last_end = _mult((fs * th, 0, 0, fs, 0, st.rise), _mult(tm, st.ctm))
                texts.append(text)
        if texts:
            self._emit(st, start, last_end, "".join(texts))
        return tm

    def _emit(self, st, start, end, text):
        if start is None or not text.strip():
            return
        a, b, c, d = start[:4]
        size = math.hypot(c, d)
        angle = math.degrees(math.atan2(b, a))
        font = st.font
        self.spans.append(Span(start[4], start[5], end[4], end[5], size, text, font.base, font.bold, font.italic,
                               st.fill if st.mode != 3 else "#000000", round(angle, 1)))

    def _do(self, resources, st, name):
        r = self.doc.resolve
        xobj = r((r(resources.get("XObject")) or {}).get(name))
        if not isinstance(xobj, Stream):
            return
        sub = r(xobj.dict.get("Subtype"))
        if sub == "Image":
            ctm = st.ctm
            pts = [_apply(ctm, x, y) for x, y in ((0, 0), (1, 0), (1, 1), (0, 1))]
            xs, ys = [p[0] for p in pts], [p[1] for p in pts]
            self.images.append(Image((min(xs), min(ys), max(xs), max(ys)), xobj))
        elif sub == "Form" and self.depth < 12:
            matrix = r(xobj.dict.get("Matrix")) or [1, 0, 0, 1, 0, 0]
            inner = st.copy()
            inner.ctm = _mult(tuple(r(v) for v in matrix[:6]), st.ctm)
            self.depth += 1
            try:
                self._run(xobj.data(), r(xobj.dict.get("Resources")) or resources, inner)
            finally:
                self.depth -= 1


# ------------------------------------------------------------------ pictures

def image_data(doc, stream):
    """(format, data, (width, height)) of an image XObject as PyExcel keeps pictures - JPEG as it is, other
    8-bit images as PNG - or None for kinds it can't show (JPEG 2000, fax, JBIG2...)."""
    import images
    r = doc.resolve
    d = stream.dict
    w, h = r(d.get("Width", 0)) or 0, r(d.get("Height", 0)) or 0
    if not (w and h) or w * h > 40_000_000:
        return None
    filters, _ = stream.filters()
    last = filters[-1] if filters else None
    if last in ("DCTDecode", "DCT"):
        data = stream.data()
        info = images.image_info(data)
        if info and info[0] == "jpeg":
            return "jpeg", data, (w, h)
        return None
    if last in IMAGE_CODECS:
        return None
    bpc = r(d.get("BitsPerComponent", 8)) or 8
    space = r(d.get("ColorSpace"))
    raw = stream.data()
    palette, comps = None, 1
    if isinstance(space, list) and space and r(space[0]) == "Indexed":
        base = r(space[1])
        base_n = 3
        if isinstance(base, list) and base and r(base[0]) == "ICCBased":
            s = r(base[1])
            base_n = r(s.dict.get("N", 3)) if isinstance(s, Stream) else 3
        elif base in ("DeviceGray", "CalGray"):
            base_n = 1
        elif base == "DeviceCMYK":
            base_n = 4
        lookup = r(space[3]) if len(space) > 3 else b""
        lookup = lookup.data() if isinstance(lookup, Stream) else (lookup or b"")
        palette = []
        for i in range(0, len(lookup) - base_n + 1, base_n):
            vals = [v / 255 for v in lookup[i:i + base_n]]
            palette.append(bytes(int(c[1:][k:k + 2], 16) for c in [_color({1: "DeviceGray", 3: "DeviceRGB",
                                                                             4: "DeviceCMYK"}[base_n], vals)]
                                 for k in (0, 2, 4)))
    elif isinstance(space, list) and space and r(space[0]) == "ICCBased":
        s = r(space[1])
        comps = r(s.dict.get("N", 3)) if isinstance(s, Stream) else 3
    elif space in ("DeviceRGB", "CalRGB") or (isinstance(space, list) and space and r(space[0]) == "CalRGB"):
        comps = 3
    elif space == "DeviceCMYK":
        comps = 4
    elif d.get("ImageMask") or space is None:
        comps = 1
    row_bits = w * bpc * (1 if palette is not None else comps)
    stride = (row_bits + 7) // 8
    if len(raw) < stride * h:
        return None
    rows = []
    for y in range(h):
        line = raw[y * stride:(y + 1) * stride]
        if palette is not None:
            idx = _unpack(line, bpc, w)
            px = bytearray()
            for i in idx:
                px += palette[i] if i < len(palette) else b"\0\0\0"
            rows.append(px)
        elif comps == 1:
            vals = _unpack(line, bpc, w)
            top = (1 << bpc) - 1
            g = bytes(round(v * 255 / top) for v in vals) if bpc != 8 else bytes(vals)
            if d.get("ImageMask") or (d.get("Decode") and r(d.get("Decode"))[:1] == [1]):
                g = bytes(255 - v for v in g)
            px = bytearray(w * 3)
            px[0::3] = px[1::3] = px[2::3] = g
            rows.append(px)
        elif comps == 3 and bpc == 8:
            rows.append(bytearray(line[:w * 3]))
        elif comps == 4 and bpc == 8:
            px = bytearray(w * 3)
            for i in range(w):
                c, m, y2, k = (v / 255 for v in line[4 * i:4 * i + 4])
                px[3 * i], px[3 * i + 1], px[3 * i + 2] = (round(255 * (1 - c) * (1 - k)), round(255 * (1 - m) * (1 - k)),
                                                           round(255 * (1 - y2) * (1 - k)))
            rows.append(px)
        else:
            return None
    return "png", images.encode_png(w, h, rows), (w, h)


def _unpack(line, bpc, count):
    if bpc == 8:
        return list(line[:count])
    out = []
    mask = (1 << bpc) - 1
    for byte in line:
        for shift in range(8 - bpc, -1, -bpc):
            out.append((byte >> shift) & mask)
    return out[:count]
