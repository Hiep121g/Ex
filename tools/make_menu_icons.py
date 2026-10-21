"""Render the ribbon's vector icons into the PNG sprite sheets that menus use.

tk.Menu entries can only show images, so the icons drawn by ribbon.ICONS are rasterised once, here,
into icons/menu_<size>.png (sizes 16, 20, 24 and 32 pixels, for 100% to 200% display scaling) with
icons/menu_icons.txt listing the icon in each slot, row by row, MenuIcons.COLUMNS per row.

Each icon is drawn four times larger on a white and on a black canvas and grabbed from the screen; the
difference between the two grabs gives its transparency, and scaling down smooths its edges.

Run from the repository root on a machine with a display and Pillow:  python tools/make_menu_icons.py
(on a headless Linux box: xvfb-run -s "-screen 0 1600x1200x24" python tools/make_menu_icons.py).
"""

import os
import sys
import tkinter as tk

from PIL import Image, ImageChops, ImageGrab

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import excel_app  # noqa: E402  (for the colours of the data-bar and colour-scale menus)
from ribbon import ICONS, MenuIcons, draw_icon  # noqa: E402

BIG = 64     # drawing size: 4 x 16
CELL = 80    # spacing on the canvas, so icons that overshoot their square do not touch


def icon_names():
    """Every icon, plus coloured variants "name@colour" used by menus."""
    names = sorted(ICONS)
    names += [f"cf_databar@{color}" for _, color in excel_app.DATA_BARS]
    names += [f"cf_scale@{','.join(colors)}" for _, colors in excel_app.COLOR_SCALES]
    return names


def grab(root, canvas, names, background):
    canvas.delete("all")
    canvas.configure(bg=background)
    cols = canvas.winfo_width() // CELL
    for i, name in enumerate(names):
        base, _, color = name.partition("@")
        x, y = (i % cols) * CELL + (CELL - BIG) // 2, (i // cols) * CELL + (CELL - BIG) // 2
        draw_icon(canvas, base, x, y, BIG, color or None)
    root.update()
    root.after(300)
    root.update()
    x0, y0 = canvas.winfo_rootx(), canvas.winfo_rooty()
    shot = ImageGrab.grab(bbox=(x0, y0, x0 + canvas.winfo_width(), y0 + canvas.winfo_height())).convert("RGB")
    out = []
    for i in range(len(names)):
        x, y = (i % cols) * CELL + (CELL - BIG) // 2, (i // cols) * CELL + (CELL - BIG) // 2
        out.append(shot.crop((x, y, x + BIG, y + BIG)))
    return out


def with_alpha(on_white, on_black):
    """RGBA icon from grabs over white and black: alpha = 1 - (white - black), colour = black / alpha."""
    alpha = ImageChops.invert(ImageChops.subtract(on_white, on_black).convert("L"))
    return Image.merge("RGBa", (*on_black.split(), alpha)).convert("RGBA")


def main():
    names = icon_names()
    root = tk.Tk()
    root.geometry("+0+0")
    width = min(root.winfo_screenwidth() - 40, 12 * CELL)
    per_page = (width // CELL) * max(1, (root.winfo_screenheight() - 120) // CELL)
    canvas = tk.Canvas(root, width=width, height=CELL * -(-min(per_page, len(names)) // (width // CELL)),
                       highlightthickness=0, bd=0)
    canvas.pack()
    root.update()
    icons = []
    for start in range(0, len(names), per_page):
        page = names[start:start + per_page]
        white, black = grab(root, canvas, page, "#ffffff"), grab(root, canvas, page, "#000000")
        icons += [with_alpha(w, b) for w, b in zip(white, black)]
    root.destroy()
    folder = os.path.join(ROOT, "icons")
    os.makedirs(folder, exist_ok=True)
    cols = MenuIcons.COLUMNS
    rows = -(-len(icons) // cols)
    for size in MenuIcons.SIZES:
        sheet = Image.new("RGBA", (cols * size, rows * size), (0, 0, 0, 0))
        for i, icon in enumerate(icons):
            small = icon.convert("RGBa").resize((size, size), Image.BOX).convert("RGBA")
            sheet.paste(small, ((i % cols) * size, (i // cols) * size))
        sheet.save(os.path.join(folder, f"menu_{size}.png"), optimize=True)
    with open(os.path.join(folder, "menu_icons.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(names) + "\n")
    print(f"{len(names)} icons -> icons/menu_<{'|'.join(map(str, MenuIcons.SIZES))}>.png")


if __name__ == "__main__":
    main()
