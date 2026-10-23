r"""Export the MarginKite "Margin bars" logo files (mark defined in mark.py).

Usage (from the repo root):
    .venv\Scripts\python.exe products\brand\build_logo.py

Writes next to this file:
  marginkite-logo-1024.png              avatar: navy square, dark-version mark at ~63% width, centred
  marginkite-logo-transparent-1024.png  light-version mark on transparent, same framing
  marginkite-lockup-light.png           mark + wordmark for light backgrounds (transparent), 1600 px wide
  marginkite-lockup-dark.png            mark + wordmark on navy, 1600 px wide
and 40 px circle-crop check copies in _check/ (git-ignored).
Rendered with headless Edge so the Sora wordmark comes from Google Fonts.
"""

from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from mark import NAVY, SORA_LINK, lockup_html, mark_svg, wordmark_css  # noqa: E402

EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
]
# Content box of the mark inside its 100x100 viewBox (bars x16-80, diamond to x88; y4-84).
CONTENT_VIEWBOX = "16 4 72 80"
MARK_SHARE = 0.63  # mark width as a share of the avatar width


def page(body: str, w: int, h: int, bg: str) -> str:
    return (f'<!doctype html><html><head><meta charset="utf-8">{SORA_LINK}<style>'
            f"* {{ margin: 0; padding: 0; }} html, body {{ width: {w}px; height: {h}px; overflow: hidden; "
            f"background: {bg}; }} body {{ display: flex; align-items: center; justify-content: center; }} "
            f"{wordmark_css('.wm', 'light')} .dark .wm {{ color: #FFFFFF; }}"
            f"</style></head><body>{body}</body></html>")


def shoot(browser, doc: str, out: Path, w: int, h: int, tmp: Path, transparent: bool):
    hp = tmp / (out.stem + ".html")
    hp.write_text(doc, encoding="utf-8")
    args = [browser, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run",
            "--force-device-scale-factor=1", f"--user-data-dir={tmp / 'p'}", f"--window-size={w},{h}",
            "--virtual-time-budget=10000", f"--screenshot={out}", hp.as_uri()]
    if transparent:
        args.insert(1, "--default-background-color=00000000")
    subprocess.run(args, check=True, capture_output=True, timeout=120)
    img = Image.open(out)
    assert img.size == (w, h), img.size
    img = img.convert("RGBA") if transparent else img.convert("RGB")
    img.save(out, optimize=True)


def framed(variant: str, w: int, h: int) -> str:
    """The mark cropped to its content box, so it sits optically centred in the square."""
    return mark_svg(None, variant, CONTENT_VIEWBOX).replace("<svg ", f'<svg width="{w}" height="{h}" ', 1)


def circle_check(src: Path, dst: Path, px: int = 40):
    img = Image.open(src).convert("RGBA").resize((px, px), Image.LANCZOS)
    mask = Image.new("L", (px * 4, px * 4), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, px * 4 - 1, px * 4 - 1), fill=255)
    out = Image.new("RGBA", (px, px), (255, 255, 255, 0))
    out.paste(img, (0, 0), mask.resize((px, px), Image.LANCZOS))
    out.save(dst)


def main():
    browser = next((p for p in EDGE_CANDIDATES if Path(p).exists()), None) or sys.exit("Edge/Chrome not found")
    check = HERE / "_check"
    check.mkdir(exist_ok=True)
    w = round(1024 * MARK_SHARE)
    h = round(w * 80 / 72)
    with tempfile.TemporaryDirectory() as t:
        tmp = Path(t)
        jobs = [
            ("marginkite-logo-1024.png", page(framed("dark", w, h), 1024, 1024, NAVY), 1024, 1024, False),
            ("marginkite-logo-transparent-1024.png",
             page(framed("light", w, h), 1024, 1024, "transparent"), 1024, 1024, True),
            ("marginkite-lockup-light.png", page(lockup_html(320, 200, "light"), 1600, 420, "transparent"),
             1600, 420, True),
            ("marginkite-lockup-dark.png", page(f'<div class="dark">{lockup_html(320, 200, "dark")}</div>',
                                                1600, 420, NAVY), 1600, 420, False),
        ]
        for name, doc, ww, hh, transparent in jobs:
            shoot(browser, doc, HERE / name, ww, hh, tmp, transparent)
            print(f"Built {HERE / name}")
    for name in ("marginkite-logo-1024.png", "marginkite-logo-transparent-1024.png"):
        circle_check(HERE / name, check / name.replace("1024", "40-circle"))


if __name__ == "__main__":
    main()
