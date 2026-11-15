#!/usr/bin/env python3
"""Rasterize the Assignments+ backpack icon (16/32/48/128) without any SVG/font deps.

Geometry mirrors assets/icon.svg on a 64x64 design grid; the purple rounded tile
is the only background. Drawn at 8x then downsampled for crisp anti-aliased edges."""
from PIL import Image, ImageDraw
import math
import os

BRAND = (92, 95, 192, 255)   # #5C5FC0 indigo tile
INK = (255, 255, 255, 255)   # white backpack
OUT = os.path.join(os.path.dirname(__file__), "..", "public", "icons")

# --- design grid (64 x 64), symmetric about x = 32 -------------------------
TILE_RADIUS = 18.22

BODY_L, BODY_R = 16.2, 47.8          # body sides
BODY_TOP, DOME_CY, DOME_R = 17.0, 32.8, 15.8   # domed top = half-width arc
BODY_BOTTOM, BODY_R_BOTTOM = 52.4, 8.44        # rounded bottom corners

ARCH_CY, ARCH_OUTER, ARCH_INNER = 19.0, 6.98, 4.18   # top carry handle

BAND_TOP, BAND_BOTTOM = 35.78, 39.33  # seam between flap and body
STRAP_L, STRAP_R, STRAP_BOTTOM, STRAP_RADIUS = 24.2, 26.9, 43.78, 1.33
POCKET = (24.2, 23.56, 39.8, 32.0, 4.2)   # purple zip pocket, stadium shape
ZIP = (27.35, 26.67, 36.65, 29.11, 1.22)  # white zip pull inside the pocket


def body_outline(k):
    """Single closed polygon for the backpack body, scaled by k px per unit."""
    pts = []
    # left edge, dome -> bottom corner
    pts.append((BODY_L, DOME_CY))
    # domed top: half circle, left -> top -> right
    for i in range(0, 181, 2):
        a = math.radians(180 - i * 1.0)
        pts.append((32 + DOME_R * math.cos(a), DOME_CY - DOME_R * math.sin(a)))
    # right edge down to the bottom-right corner
    pts.append((BODY_R, BODY_BOTTOM - BODY_R_BOTTOM))
    # bottom-right corner
    cbx, cby = BODY_R - BODY_R_BOTTOM, BODY_BOTTOM - BODY_R_BOTTOM
    for i in range(0, 91, 2):
        a = math.radians(i)
        pts.append((cbx + BODY_R_BOTTOM * math.cos(a), cby + BODY_R_BOTTOM * math.sin(a)))
    # bottom edge / bottom-left corner
    cbx = BODY_L + BODY_R_BOTTOM
    for i in range(0, 91, 2):
        a = math.radians(90 + i)
        pts.append((cbx + BODY_R_BOTTOM * math.cos(a), cby + BODY_R_BOTTOM * math.sin(a)))
    return [(x * k, y * k) for x, y in pts]


def draw_icon(size):
    ss = size * 8
    k = ss / 64.0
    img = Image.new("RGBA", (ss, ss), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # purple rounded tile
    d.rounded_rectangle((0, 0, ss - 1, ss - 1), radius=TILE_RADIUS * k, fill=BRAND)

    # carry handle: white ring, lower half later hidden by the body
    cx, cy = 32 * k, ARCH_CY * k
    d.ellipse((cx - ARCH_OUTER * k, cy - ARCH_OUTER * k,
               cx + ARCH_OUTER * k, cy + ARCH_OUTER * k), fill=INK)
    d.ellipse((cx - ARCH_INNER * k, cy - ARCH_INNER * k,
               cx + ARCH_INNER * k, cy + ARCH_INNER * k), fill=BRAND)

    # backpack body (drawn over the handle's lower half)
    d.polygon(body_outline(k), fill=INK)

    # seam across the body
    d.rectangle((0, BAND_TOP * k, ss - 1, BAND_BOTTOM * k), fill=BRAND)

    # small strap tab hanging from the seam
    d.rounded_rectangle((STRAP_L * k, BAND_TOP * k, STRAP_R * k, STRAP_BOTTOM * k),
                        radius=STRAP_RADIUS * k, fill=BRAND)

    # zip pocket + pull
    x0, y0, x1, y1, r = POCKET
    d.rounded_rectangle((x0 * k, y0 * k, x1 * k, y1 * k), radius=r * k, fill=BRAND)
    x0, y0, x1, y1, r = ZIP
    d.rounded_rectangle((x0 * k, y0 * k, x1 * k, y1 * k), radius=r * k, fill=INK)

    return img.resize((size, size), Image.LANCZOS)


def main():
    os.makedirs(OUT, exist_ok=True)
    for size in (16, 32, 48, 128):
        draw_icon(size).save(os.path.join(OUT, f"icon-{size}.png"))
    print("icons written to", os.path.abspath(OUT))


if __name__ == "__main__":
    main()