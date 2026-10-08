"""Draws the app icon and writes every format the builds need.

    python3 assets/make-icon.py

Outputs into src/OpenSheet.App/Assets/:
  icon.png       1024px, used as the window icon
  app.ico        Windows exe icon (16..256)
  AppIcon.icns   macOS bundle icon (built with iconutil)

Design: emerald rounded square (macOS icon grid), a white sheet with a header
row, and one selected cell with a fill handle, like the app's own grid.
"""
import os
import shutil
import subprocess
import sys
import tempfile

from PIL import Image, ImageChops, ImageDraw, ImageFilter

S = 4  # supersampling: draw at 4096px, scale down for clean edges
N = 1024 * S
OUT = os.path.join(os.path.dirname(__file__), "..", "src", "OpenSheet.App", "Assets")


def px(v):
    return int(v * S)


def rounded_mask(box, radius):
    m = Image.new("L", (N, N), 0)
    ImageDraw.Draw(m).rounded_rectangle([px(c) for c in box], radius=px(radius), fill=255)
    return m


def vertical_gradient(top, bottom):
    g = Image.new("RGBA", (1, 256))
    for y in range(256):
        t = y / 255
        g.putpixel((0, y), tuple(round(a + (b - a) * t) for a, b in zip(top, bottom)) + (255,))
    return g.resize((N, N))


def shadow(mask, offset, blur, alpha):
    sh = Image.new("RGBA", (N, N), (0, 0, 0, 0))
    a = mask.point(lambda v: v * alpha // 255).filter(ImageFilter.GaussianBlur(px(blur)))
    sh.putalpha(ImageChops.offset(a, 0, px(offset)))
    return sh


def draw():
    img = Image.new("RGBA", (N, N), (0, 0, 0, 0))

    # Base: rounded square on Apple's 1024 grid (824px body, 100px margin), soft drop shadow.
    body = (100, 100, 924, 924)
    body_mask = rounded_mask(body, 185)
    img.alpha_composite(shadow(body_mask, 14, 18, 90))
    img.paste(vertical_gradient((52, 199, 123), (8, 112, 62)), (0, 0), body_mask)

    # Gentle highlight across the top half.
    glow = vertical_gradient((255, 255, 255), (255, 255, 255))
    glow_alpha = Image.new("L", (1, 256))
    for y in range(256):
        glow_alpha.putpixel((0, y), max(0, round(46 * (1 - y / 140))))
    glow.putalpha(ImageChops.multiply(glow_alpha.resize((N, N)), body_mask))
    img.alpha_composite(glow)

    # The sheet: white card with its own shadow.
    card = (238, 262, 786, 778)
    card_mask = rounded_mask(card, 44)
    img.alpha_composite(shadow(card_mask, 12, 16, 110))
    sheet = Image.new("RGBA", (N, N), (0, 0, 0, 0))
    d = ImageDraw.Draw(sheet)
    d.rectangle([px(c) for c in card], fill=(255, 255, 255, 255))

    x0, y0, x1, y1 = card
    cols, header_h = 3, 104
    col_w = (x1 - x0) / cols
    row_h = (y1 - y0 - header_h) / 3

    # Header row, light green.
    d.rectangle([px(x0), px(y0), px(x1), px(y0 + header_h)], fill=(214, 240, 225, 255))

    # Grid lines.
    line = (200, 222, 209, 255)
    for i in range(1, cols):
        x = x0 + col_w * i
        d.line([px(x), px(y0), px(x), px(y1)], fill=line, width=px(5))
    for i in range(0, 3):
        y = y0 + header_h + row_h * i
        d.line([px(x0), px(y), px(x1), px(y)], fill=line, width=px(5))

    # "Text" in some cells: rounded bars.
    def bar(col, row, frac, color):
        cx = x0 + col_w * col + 30
        cy = (y0 + header_h / 2) if row < 0 else (y0 + header_h + row_h * row + row_h / 2)
        h = 22 if row < 0 else 26
        d.rounded_rectangle([px(cx), px(cy - h / 2), px(cx + (col_w - 60) * frac), px(cy + h / 2)],
                            radius=px(h / 2), fill=color)

    head, text = (120, 176, 143, 255), (188, 204, 195, 255)
    for c, f in ((0, 0.7), (1, 0.55), (2, 0.65)):
        bar(c, -1, f, head)
    for c, r, f in ((0, 0, 0.8), (2, 0, 0.5), (0, 1, 0.6), (1, 2, 0.75), (2, 2, 0.45), (0, 2, 0.5)):
        bar(c, r, f, text)

    sheet.putalpha(ImageChops.multiply(sheet.getchannel("A"), card_mask))
    img.alpha_composite(sheet)

    # Selected cell (middle) with Excel-style border and fill handle; drawn over the card edge-safe.
    sel = ImageDraw.Draw(img)
    sx0, sy0 = x0 + col_w, y0 + header_h + row_h
    sx1, sy1 = sx0 + col_w, sy0 + row_h
    green = (14, 138, 74, 255)
    sel.rectangle([px(sx0 + 4), px(sy0 + 4), px(sx1 - 4), px(sy1 - 4)], fill=(232, 247, 238, 255))
    cx = sx0 + 30
    cy = sy0 + row_h / 2
    sel.rounded_rectangle([px(cx), px(cy - 13), px(cx + (col_w - 60) * 0.7), px(cy + 13)], radius=px(13), fill=green)
    sel.rectangle([px(sx0), px(sy0), px(sx1), px(sy1)], outline=green, width=px(14))
    hs = 34
    sel.rectangle([px(sx1 - hs / 2), px(sy1 - hs / 2), px(sx1 + hs / 2), px(sy1 + hs / 2)],
                  fill=green, outline=(255, 255, 255, 255), width=px(7))

    return img.resize((1024, 1024), Image.LANCZOS)


def main():
    os.makedirs(OUT, exist_ok=True)
    icon = draw()
    icon.save(os.path.join(OUT, "icon.png"))
    icon.save(os.path.join(OUT, "app.ico"), sizes=[(s, s) for s in (16, 24, 32, 48, 64, 128, 256)])

    if shutil.which("iconutil"):  # macOS only
        with tempfile.TemporaryDirectory() as tmp:
            iconset = os.path.join(tmp, "AppIcon.iconset")
            os.makedirs(iconset)
            for s in (16, 32, 128, 256, 512):
                icon.resize((s, s), Image.LANCZOS).save(os.path.join(iconset, f"icon_{s}x{s}.png"))
                icon.resize((s * 2, s * 2), Image.LANCZOS).save(os.path.join(iconset, f"icon_{s}x{s}@2x.png"))
            subprocess.run(["iconutil", "-c", "icns", iconset, "-o", os.path.join(OUT, "AppIcon.icns")], check=True)
    else:
        print("iconutil not found (not macOS): skipped AppIcon.icns", file=sys.stderr)
    print("wrote", os.path.abspath(OUT))


if __name__ == "__main__":
    main()
