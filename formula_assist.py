"""Typing help for formulas, as in Excel: Formula AutoComplete (the list of functions that start with
what is typed), the argument ScreenTip of the function the cursor is in, and F4 to cycle a reference
through $A$1, A$1, $A1 and A1.

The text functions at the top work on a formula string and a cursor position; FormulaAssist wires them to
a Tk entry (the in-cell editor or the formula bar).
"""

import re

try:
    import tkinter as tk
except ImportError:  # the text functions are usable (and tested) without Tk
    tk = None

_NAME_CHARS = "._"
_REF_RE = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)")


def _string_spans(text):
    """[(start, end)] of the "..." literals in a formula (an open one runs to the end)."""
    spans, start = [], None
    for i, ch in enumerate(text):
        if ch == '"':
            if start is None:
                start = i
            else:
                spans.append((start, i + 1))
                start = None
    if start is not None:
        spans.append((start, len(text)))
    return spans


def _in_string(text, pos):
    return text[:pos].count('"') % 2 == 1


def completion_prefix(text, pos):
    """(start, prefix) of the function name being typed just before the cursor, or None."""
    if not text.startswith("=") or _in_string(text, pos):
        return None
    start = pos
    while start > 1 and (text[start - 1].isalnum() or text[start - 1] in _NAME_CHARS):
        start -= 1
    prefix = text[start:pos]
    if not prefix or not prefix[0].isalpha() or text[start - 1] in "!$'\"" or text[pos:pos + 1] == "(":
        return None
    return start, prefix


def matching_functions(prefix, names):
    prefix = prefix.upper()
    return [n for n in names if n.upper().startswith(prefix)]


def call_context(text, pos):
    """(function, argument index) of the innermost function call the cursor is in, or None."""
    if not text.startswith("="):
        return None
    stack, in_str = [], False
    for i in range(1, min(pos, len(text))):
        ch = text[i]
        if in_str:
            in_str = ch != '"'
        elif ch == '"':
            in_str = True
        elif ch in "({":
            k = i
            while ch == "(" and k > 1 and (text[k - 1].isalnum() or text[k - 1] in _NAME_CHARS):
                k -= 1
            stack.append([text[k:i].upper() if ch == "(" else "{", 0])
        elif ch in ")}":
            if stack:
                stack.pop()
        elif ch in ",;" and stack:
            stack[-1][1] += 1
    for name, index in reversed(stack):
        if name == "{":
            return None  # inside an array constant
        if name:
            return name, index
    return None


def signature_args(syntax):
    """'VLOOKUP(a, b, [c])' -> ('VLOOKUP', ['a', 'b', '[c]'])."""
    name, _, rest = syntax.partition("(")
    rest = rest.rsplit(")", 1)[0].strip()
    return name, [a.strip() for a in rest.split(",")] if rest else []


def current_arg(args, index):
    """Which of the signature's arguments argument `index` is (repeating ones map to '...'), or None."""
    if index < len(args):
        return index
    if args and args[-1] == "...":
        return len(args) - 1
    return None


def _cycle(col_abs, row_abs):
    # A1 -> $A$1 -> A$1 -> $A1 -> A1, as Excel's F4
    return {(False, False): (True, True), (True, True): (False, True), (False, True): (True, False),
            (True, False): (False, False)}[(col_abs, row_abs)]


def _refs(text):
    """Cell references in a formula, outside strings, not part of a name or function call."""
    strings = _string_spans(text)
    out = []
    for m in _REF_RE.finditer(text):
        before, after = text[m.start() - 1:m.start()], text[m.end():m.end() + 1]
        if (before and (before.isalnum() or before in "_.$")) or (after and (after.isalnum() or after in "_.(")):
            continue
        if any(a <= m.start() < b for a, b in strings):
            continue
        out.append(m)
    return out


def toggle_reference(text, pos):
    """F4: cycle the reference at or left of the cursor (both ends of a range together).
    Returns (text, cursor) or None when there is no reference there."""
    refs = _refs(text)
    hit = next((i for i, m in enumerate(refs) if m.start() <= pos <= m.end()), None)
    if hit is None:
        hit = next((i for i in range(len(refs) - 1, -1, -1) if refs[i].end() <= pos), None)
        if hit is None:
            return None
    group = [hit]
    if hit + 1 < len(refs) and refs[hit + 1].start() == refs[hit].end() + 1 and text[refs[hit].end()] == ":":
        group.append(hit + 1)
    elif hit > 0 and refs[hit - 1].end() + 1 == refs[hit].start() and text[refs[hit - 1].end()] == ":":
        group.insert(0, hit - 1)
    first = refs[group[0]]
    col_abs, row_abs = _cycle(bool(first.group(1)), bool(first.group(3)))
    out, last = [], 0
    for i in group:
        m = refs[i]
        out.append(text[last:m.start()])
        out.append(("$" if col_abs else "") + m.group(2) + ("$" if row_abs else "") + m.group(4))
        last = m.end()
    new_text = "".join(out) + text[last:]
    end = refs[group[-1]].end() + len(new_text) - len(text)
    return new_text, end


class FormulaAssist:
    """AutoComplete list, argument tip and F4 for one entry widget.

    names() gives the function and defined names to offer and info(name) its (syntax, description) or None;
    a third item False in the info marks a defined name, inserted without "(".
    The assist puts its own bind tag in front of the entry's, so while the list is open Up/Down/Tab/
    Escape work on the list before the entry's own bindings see them."""

    IGNORE = {"Up", "Down", "Tab", "ISO_Left_Tab", "Escape", "Return", "KP_Enter", "F4", "Shift_L", "Shift_R",
              "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock"}

    def __init__(self, entry, font, small_font, names, info, on_change=None):
        self.entry, self.font, self.small_font = entry, font, small_font
        self.names, self.info, self.on_change = names, info, on_change or (lambda: None)
        self.popup = self.tip = self.tip_key = None
        self.matches, self.span = [], None
        tag = f"FormulaAssist{id(self)}"
        entry.bindtags((tag,) + entry.bindtags())
        for seq, func in (("<Down>", lambda e: self._move(1)), ("<Up>", lambda e: self._move(-1)),
                          ("<Next>", lambda e: self._move(7)), ("<Prior>", lambda e: self._move(-7)),
                          ("<Tab>", lambda e: self._accept()), ("<Escape>", lambda e: self._escape()),
                          ("<F4>", lambda e: self._toggle_ref())):
            entry.bind_class(tag, seq, func)
        entry.bind_class(tag, "<KeyRelease>", self._on_key)
        entry.bind_class(tag, "<ButtonRelease-1>", lambda e: entry.after_idle(self.update))
        entry.bind_class(tag, "<FocusOut>", lambda e: entry.after(150, self._check_focus))
        entry.bind_class(tag, "<Destroy>", lambda e: self.hide())

    # ------------------------------------------------------------------ state

    def _cursor(self):
        return self.entry.index("insert")

    def update(self):
        """Show the list or the tip for the text and cursor as they are now."""
        if not self.entry.winfo_exists() or self.entry.focus_get() is not self.entry:
            self.hide()
            return
        text, pos = self.entry.get(), self._cursor()
        found = completion_prefix(text, pos)
        matches = matching_functions(found[1], self.names()) if found else []
        if matches and not (len(matches) == 1 and matches[0].upper() == found[1].upper()):
            self.span = (found[0], pos)
            self._show_list(matches)
            self._hide_tip()
            return
        self._hide_list()
        context = call_context(text, pos)
        if context and self.info(context[0]):
            self._show_tip(*context)
        else:
            self._hide_tip()

    def hide(self):
        self._hide_list()
        self._hide_tip()

    def _check_focus(self):
        try:
            focus = self.entry.focus_get()
        except (KeyError, tk.TclError):
            focus = None
        if focus is not self.entry and (self.popup is None or focus is None or
                                        not str(focus).startswith(str(self.popup))):
            self.hide()

    def _on_key(self, e):
        if e.keysym not in self.IGNORE:
            self.update()

    # ------------------------------------------------------------------ the list

    def _show_list(self, matches):
        if self.popup is None:
            top = self.popup = tk.Toplevel(self.entry)
            top.overrideredirect(True)
            top.attributes("-topmost", True)
            frame = tk.Frame(top, bg="#ffffff", bd=1, relief="solid")
            frame.pack(fill="both", expand=True)
            self.listbox = tk.Listbox(frame, font=self.font, activestyle="none", bd=0, highlightthickness=0,
                                      selectbackground="#cde8ff", selectforeground="#000000", exportselection=False,
                                      width=18)
            self.listbox.pack(side="left", fill="y")
            self.desc = tk.Label(frame, bg="#ffffff", fg="#444444", font=self.small_font, justify="left",
                                 anchor="nw", wraplength=self.font.measure("M") * 16)
            self.desc.pack(side="left", fill="both", padx=(6, 8), pady=4)
            self.listbox.bind("<<ListboxSelect>>", lambda e: self._describe())
            self.listbox.bind("<Double-Button-1>", lambda e: self._accept())
        lb = self.listbox
        if matches != self.matches:
            self.matches = matches
            lb.delete(0, "end")
            for name in matches:
                lb.insert("end", name)
            lb.configure(height=min(8, len(matches)))
            lb.selection_clear(0, "end")
            lb.selection_set(0)
            lb.see(0)
        self._describe()
        self._place(self.popup, self.span[0])

    def _place(self, win, index):
        win.update_idletasks()
        box = self.entry.bbox(index)
        x = self.entry.winfo_rootx() + (box[0] if box else 0)
        y = self.entry.winfo_rooty() + self.entry.winfo_height() + 1
        x = max(0, min(x, win.winfo_screenwidth() - win.winfo_reqwidth() - 4))
        win.geometry(f"+{x}+{y}")
        win.deiconify()

    def _describe(self):
        sel = self.listbox.curselection()
        info = self.info(self.matches[sel[0]]) if sel else None
        self.desc.configure(text=info[1] if info else "")

    def _hide_list(self):
        if self.popup is not None:
            self.popup.destroy()
            self.popup, self.matches = None, []

    def _move(self, step):
        if self.popup is None:
            return None
        lb = self.listbox
        sel = lb.curselection()
        i = max(0, min(len(self.matches) - 1, (sel[0] if sel else -1) + step))
        lb.selection_clear(0, "end")
        lb.selection_set(i)
        lb.see(i)
        self._describe()
        return "break"

    def _accept(self):
        if self.popup is None:
            return None
        sel = self.listbox.curselection()
        name = self.matches[sel[0] if sel else 0]
        info = self.info(name)
        text = name if info and len(info) > 2 and not info[2] else name + "("  # defined names take no "("
        start, end = self.span
        e = self.entry
        e.delete(start, end)
        e.insert(start, text)
        e.icursor(start + len(text))
        e.focus_set()
        self.on_change()
        self._hide_list()
        self.update()
        return "break"

    def _escape(self):
        if self.popup is None:
            self._hide_tip()
            return None
        self._hide_list()
        return "break"

    def _toggle_ref(self):
        text = self.entry.get()
        result = toggle_reference(text, self._cursor()) if text.startswith("=") else None
        if result is not None:
            e = self.entry
            e.delete(0, "end")
            e.insert(0, result[0])
            e.icursor(result[1])
            self.on_change()
            self.update()
        return "break"

    # ------------------------------------------------------------------ argument tip

    def _show_tip(self, name, index):
        syntax = self.info(name)[0]
        fname, args = signature_args(syntax)
        bold = current_arg(args, index)
        key = (syntax, bold)
        if self.tip is not None and self.tip_key == key:
            self._place(self.tip, 0)
            return
        self._hide_tip()
        tip = self.tip = tk.Toplevel(self.entry)
        self.tip_key = key
        tip.overrideredirect(True)
        tip.attributes("-topmost", True)
        row = tk.Frame(tip, bg="#ffffe1", bd=1, relief="solid", padx=5, pady=2)
        row.pack()
        bold_font = self.small_font.copy()
        bold_font.configure(weight="bold")
        self._bold_font = bold_font
        parts = [(fname + "(", None)]
        for i, arg in enumerate(args):
            parts.append((arg, i))
            parts.append((", " if i < len(args) - 1 else ")", None))
        if not args:
            parts.append((")", None))
        for text, i in parts:
            tk.Label(row, text=text, bg="#ffffe1", fg="#000000", bd=0, padx=0,
                     font=bold_font if i is not None and i == bold else self.small_font).pack(side="left")
        self._place(tip, 0)

    def _hide_tip(self):
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None
