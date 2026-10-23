"""The MarginKite "Margin bars" mark and wordmark, defined once.

Every script that draws the logo imports from here:
  products/brand/build_logo.py, products/seller-profit-tracker/previews/build_previews.py
  (square and Creative Market previews), products/seller-profit-tracker/build_guide.py (via
  build_previews) and marketing/*/build_pins.py.

Approved spec (owner, Concept C): viewBox 0 0 100 100, three rounded bars rising left to right
at 50% / 75% / 100% opacity, and a diamond top right.
  dark version  (on navy #1E3A5F): bars #3FC1AE, diamond #FFFFFF
  light version (on white/cream):  bars #0F8B7D, diamond #1E3A5F
Wordmark: "MarginKite" in Sora 700, letter-spacing -0.02em, navy on light, white on dark.
Lockup: mark left, gap about 0.25x the mark size.
"""

from __future__ import annotations

NAVY = "#1E3A5F"
BAR_LIGHT = "#0F8B7D"
BAR_DARK = "#3FC1AE"
WHITE = "#FFFFFF"

BARS = [  # x, y, width, height, opacity  (rx = 4 on all)
    (16, 58, 16, 26, 0.5),
    (40, 44, 16, 40, 0.75),
    (64, 30, 16, 54, 1),
]
DIAMOND = "80,4 88,12 80,20 72,12"
GAP_RATIO = 0.25

SORA_LINK = ('<link href="https://fonts.googleapis.com/css2?family=Sora:wght@700&display=swap" '
             'rel="stylesheet">')
WORDMARK_FONT = "'Sora', 'Segoe UI', Arial, sans-serif"


def mark_svg(size: int | None = None, variant: str = "light", viewbox: str = "0 0 100 100") -> str:
    """The mark as inline SVG. `size` sets width/height in px (None leaves sizing to CSS)."""
    bar = BAR_LIGHT if variant == "light" else BAR_DARK
    diamond = NAVY if variant == "light" else WHITE
    dims = f' width="{size}" height="{size}"' if size else ""
    rects = "".join(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="4" fill="{bar}" fill-opacity="{o}"/>'
                    for x, y, w, h, o in BARS)
    return (f'<svg class="mk-mark" xmlns="http://www.w3.org/2000/svg"{dims} viewBox="{viewbox}">'
            f'{rects}<polygon points="{DIAMOND}" fill="{diamond}"/></svg>')


def wordmark_css(selector: str = ".wm", variant: str = "light") -> str:
    color = NAVY if variant == "light" else WHITE
    return (f"{selector} {{ font-family: {WORDMARK_FONT}; font-weight: 700; letter-spacing: -0.02em; "
            f"color: {color}; }}")


def lockup_html(mark_px: int, font_px: int, variant: str = "light") -> str:
    """Mark + wordmark in a flex row with the approved gap."""
    return (f'<div style="display:flex;align-items:center;gap:{round(mark_px * GAP_RATIO)}px">'
            f'{mark_svg(mark_px, variant)}<span class="wm" style="font-size:{font_px}px">MarginKite</span></div>')
