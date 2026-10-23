"""Render the buyer guide (listing/buyer-guide.md) to dist/MarginKite-Buyer-Guide.pdf.

Usage (from the repo root):
    .venv\\Scripts\\python.exe products\\seller-profit-tracker\\build_guide.py

Uses the MarginKite style from previews/build_previews.py and headless Edge --print-to-pdf.
[[PLACEHOLDER: ...]] markers stay visible and highlighted until the owner fills them in;
then re-run this script. The listing has no lite guide, so only the main PDF is built.
"""

from __future__ import annotations

import html
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import markdown

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "previews"))
from build_previews import BRAND, FONT_LINK, KITE, find_browser  # noqa: E402

SOURCE = HERE / "listing" / "buyer-guide.md"
OUT = HERE / "dist" / "MarginKite-Buyer-Guide.pdf"

CSS = f"""
@page {{ size: Letter; margin: 0.5in 0.6in; }}
* {{ box-sizing: border-box; }}
body {{ font-family: {BRAND['font']}; color: {BRAND['ink']}; font-size: 10.5pt; line-height: 1.42; margin: 0; }}
.top {{ display: flex; align-items: center; gap: 10px; font-weight: 800; font-size: 15pt; color: {BRAND['navy']};
       border-bottom: 3px solid {BRAND['teal']}; padding-bottom: 8px; margin-bottom: 12px; }}
.top svg {{ width: 26px; height: 26px; }}
.top .tag {{ margin-left: auto; font-size: 9pt; color: #fff; background: {BRAND['teal']}; padding: 3px 10px;
            border-radius: 12px; font-weight: 700; }}
h1 {{ font-size: 20pt; font-weight: 900; color: {BRAND['navy']}; margin: 0 0 6px; line-height: 1.15; }}
h2 {{ font-size: 12.5pt; font-weight: 800; color: {BRAND['teal']}; margin: 14px 0 4px; }}
p {{ margin: 4px 0; }}
ol, ul {{ margin: 4px 0; padding-left: 22px; }}
li {{ margin: 2px 0; }}
strong {{ color: {BRAND['navy']}; }}
code {{ font-family: Consolas, monospace; font-size: 9.5pt; background: #EEF1F5; padding: 1px 4px; border-radius: 3px; }}
mark.ph {{ background: #FFE08A; color: #7A2E00; font-weight: 800; padding: 1px 4px; border-radius: 3px;
          outline: 2px dashed {BRAND['coral']}; }}
em {{ color: {BRAND['muted']}; }}
"""


def render_html(md_text: str) -> str:
    body = markdown.markdown(md_text, extensions=["sane_lists"])
    body = re.sub(r"\[\[PLACEHOLDER:(.*?)\]\]",
                  lambda m: f'<mark class="ph">[[PLACEHOLDER:{m.group(1)}]]</mark>', body)
    return (f'<!doctype html><html><head><meta charset="utf-8"><title>MarginKite Buyer Guide</title>{FONT_LINK}'
            f"<style>{CSS}</style></head><body>"
            f'<div class="top">{KITE}MarginKite<span class="tag">{html.escape("Google Sheets + Excel")}</span></div>'
            f"{body}</body></html>")


def main():
    doc = render_html(SOURCE.read_text(encoding="utf-8"))
    OUT.parent.mkdir(parents=True, exist_ok=True)
    browser = find_browser()
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        page = tmp / "guide.html"
        page.write_text(doc, encoding="utf-8")
        subprocess.run([browser, "--headless=new", "--disable-gpu", "--no-first-run", f"--user-data-dir={tmp / 'p'}",
                        "--no-pdf-header-footer", "--virtual-time-budget=15000",
                        f"--print-to-pdf={OUT}", page.as_uri()], check=True, capture_output=True, timeout=120)
    pages = len(re.findall(rb"/Type\s*/Page(?!s)", OUT.read_bytes()))
    print(f"Built {OUT} ({pages} page{'s' if pages != 1 else ''})")


if __name__ == "__main__":
    main()
