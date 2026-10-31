#!/usr/bin/env python3
"""Rasterize Material 3-styled extension icons (16/32/48/128) without any SVG/font deps.
Drawn at 8x then downsampled for crisp anti-aliased edges."""
from PIL import Image, ImageDraw
import os

BRAND = (147, 51, 234, 255)      # purple-600 (brand continuity with assets/icon.svg)
SHEET = (255, 255, 255, 255)
LINE = (216, 180, 254, 255)      # purple-300 bars on the sheet
CHECK_BG = (34, 197, 94, 255)    # green check badge
OUT = os.path.join(os.path.dirname(__file__), "..", "public", "icons")


def rounded(draw, box, radius, fill):
    draw.rounded_rectangle(box, radius=radius, fill=fill)


def draw_icon(SS):  # SS = supersample size (square)
    img = Image.new("RGBA", (SS, SS), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # rounded purple background
    rounded(d, (0, 0, SS - 1, SS - 1), radius=int(SS * 0.24), fill=BRAND)
    clip = Image.new("RGBA", (SS, SS), (0, 0, 0, 0))
    cd = ImageDraw.Draw(clip)
    rounded(cd, (0, 0, SS - 1, SS - 1), radius=int(SS * 0.24), fill=(255, 255, 255, 255))

    # white "assignment sheet" slightly rotated
    sheet = Image.new("RGBA", (SS, SS), (0, 0, 0, 0))
    sd = ImageDraw.Draw(sheet)
    x0, y0, x1, y1 = int(SS * 0.24), int(SS * 0.18), int(SS * 0.76), int(SS * 0.82)
    r = int(SS * 0.07)
    rounded(sd, (x0, y0, x1, y1), radius=r, fill=SHEET)
    # two purple text bars
    rounded(sd, (int(SS * 0.33), int(SS * 0.36), int(SS * 0.67), int(SS * 0.42)), radius=int(SS * 0.03), fill=LINE)
    rounded(sd, (int(SS * 0.33), int(SS * 0.50), int(SS * 0.67), int(SS * 0.56)), radius=int(SS * 0.03), fill=LINE)
    rounded(sd, (int(SS * 0.33), int(SS * 0.64), int(SS * 0.55), int(SS * 0.70)), radius=int(SS * 0.03), fill=LINE)
    img.alpha_composite(sheet.rotate(-6, center=(SS / 2, SS / 2)))
    img.alpha_composite(clip)

    # green check badge, bottom-right
    badge = Image.new("RGBA", (SS, SS), (0, 0, 0, 0))
    bd = ImageDraw.Draw(badge)
    cx, cy = int(SS * 0.70), int(SS * 0.70)
    r = int(SS * 0.17)
    bd.ellipse((cx - r, cy - r, cx + r, cy + r), fill=CHECK_BG)
    s = SS / 128.0
    check = [(cx - 0.075 * SS, cy), (cx - 0.02 * SS, cy + 0.055 * SS), (cx + 0.075 * SS, cy - 0.06 * SS)]
    bd.line(check, fill=(255, 255, 255, 255), width=max(3, int(9 * s / 2.0) * 2))
    bd.ellipse((cx - r, cy - r, cx + r, cy + r), outline=(255, 255, 255, 60), width=max(1, int(2 * s)))
    img.alpha_composite(badge)
    return img


def main():
    os.makedirs(OUT, exist_ok=True)
    for size in (16, 32, 48, 128):
        img = draw_icon(size * 8)
        img = img.resize((size, size), Image.LANCZOS)
        img.save(os.path.join(OUT, f"icon-{size}.png"))
    print("icons written to", os.path.abspath(OUT))


if __name__ == "__main__":
    main()
