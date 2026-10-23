"""Build the listing preview images for the MarginKite Seller Profit Tracker and the lite calculator.

Usage (from the repo root, after build.py):
    .venv\\Scripts\\python.exe products\\seller-profit-tracker\\previews\\build_previews.py

How it works:
  1. Evaluates the built workbooks (dist/*.xlsx) with the `formulas` engine, so every
     number shown is the workbook's own result for its sample data. Nothing is typed by hand.
  2. Recreates cell ranges as HTML using each cell's real value, number format, fill
     colour and font, read with openpyxl. Charts are drawn as SVG from the evaluated
     Dashboard cells.
  3. Composes each 2000x2000 listing image in HTML and captures it with headless Edge.
  4. Saves 300 px thumbnails to previews/_thumbs/ (git-ignored) for the readability check.
"""

from __future__ import annotations

import datetime as dt
import html
import logging
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import formulas
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from PIL import Image

logging.disable(logging.WARNING)

HERE = Path(__file__).resolve().parent
PRODUCT = HERE.parent
sys.path.insert(0, str(PRODUCT))
from build import FEES_AS_OF_TEXT, LITE_NAME, TRACKER_NAME  # noqa: E402

DIST = PRODUCT / "dist"
THUMBS = HERE / "_thumbs"
SIZE = 2000
EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
]

# ------------------------------------------------------------------ brand style (defined once)
BRAND = {
    "navy": "#1F3A5F",      # same navy as the workbook headers
    "teal": "#0F8B7D",      # same teal as the workbook section labels
    "coral": "#E8735A",     # callout accent
    "cream": "#F6F3EC",     # background
    "ink": "#1E2430",
    "muted": "#5E6675",
    "font": "'Inter', 'Segoe UI', Arial, sans-serif",
    "sheet_font": "Arial, 'Segoe UI', sans-serif",
}
FONT_LINK = ('<link rel="preconnect" href="https://fonts.googleapis.com">'
             '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap" '
             'rel="stylesheet">')
SAMPLE_NOTE = "Illustration of the template; numbers are sample data."


# ------------------------------------------------------------------ workbook values
class Book:
    """A built workbook: openpyxl for values/styles, formulas for computed results."""

    def __init__(self, path: Path):
        self.path = path
        self.wb = load_workbook(path)
        sol = formulas.ExcelModel().loads(str(path)).finish().calculate()
        prefix = f"'[{path.name}]"
        self.calc = {}
        for k, v in sol.items():
            if k.startswith(prefix) and ":" not in k:
                sheet, cell = k[len(prefix):].split("'!")
                val = v.value[0, 0] if hasattr(v, "value") else v
                if hasattr(val, "item"):
                    val = val.item()
                self.calc[(sheet, cell)] = val

    def value(self, sheet, coord):
        cell = self.wb[sheet][coord]
        v = cell.value
        if isinstance(v, str) and v.startswith("="):
            v = self.calc.get((sheet.upper(), coord))
            if v is not None and str(v) == "empty":
                v = ""
            if isinstance(v, str):
                try:
                    v = float(v)
                except ValueError:
                    pass
        return v

    def text(self, sheet, coord):
        cell = self.wb[sheet][coord]
        return fmt_value(self.value(sheet, coord), cell.number_format)


def serial_to_date(v):
    return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))


def fmt_value(v, nf):
    if v is None or v == "":
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (dt.datetime, dt.date)) or (isinstance(v, (int, float)) and nf in ("yyyy-mm-dd", "mmm yyyy")):
        d = v if isinstance(v, (dt.datetime, dt.date)) else serial_to_date(v)
        return d.strftime("%b %Y") if nf == "mmm yyyy" else d.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)):
        if '"$"' in nf:
            s = f"${abs(v):,.2f}"
            return "-" + s if v < -0.004 else s
        if "%" in nf:
            dec = len(nf.split(".")[1].rstrip("%")) if "." in nf else 0
            return f"{v * 100:.{dec}f}%"
        if nf == "0":
            return f"{v:,.0f}"
        if nf == "0.00":
            return f"{v:.2f}"
        return f"{v:g}"
    return str(v)


def rgb(color):
    if color is None or color.type != "rgb" or not color.rgb or color.rgb in ("00000000",):
        return None
    return "#" + color.rgb[-6:]


# ------------------------------------------------------------------ sheet renderer
def render_range(book: Book, sheet: str, rows, cols, marks=None, width_px=7.6, headers=True):
    """Recreate a block of cells as an HTML table.

    rows/cols: lists of row numbers / column letters (may skip; the frame shows real letters
    and numbers so a skipped range is visible). marks: {"C25": 1} draws a numbered callout.
    """
    ws = book.wb[sheet]
    marks = marks or {}
    merged = {}
    covered = set()
    for mr in ws.merged_cells.ranges:
        tl = (mr.min_row, mr.min_col)
        span_cols = [c for c in cols if mr.min_col <= column_index_from_string(c) <= mr.max_col]
        merged[tl] = len(span_cols)
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                if (r, c) != tl:
                    covered.add((r, c))
    colw = []
    for c in cols:
        w = ws.column_dimensions[c].width or 8.43
        colw.append(max(30, int(w * width_px)))
    out = ['<table class="sheet"><colgroup>']
    if headers:
        out.append(f'<col style="width:{110 if marks else 46}px">')
    out += [f'<col style="width:{w}px">' for w in colw]
    out.append("</colgroup>")
    if headers:
        out.append('<tr class="hdr"><td class="corner"></td>' + "".join(f"<td>{c}</td>" for c in cols) + "</tr>")
    prev = None
    for r in rows:
        if prev is not None and r != prev + 1:
            out.append(f'<tr class="gap"><td class="rn">&#8942;</td><td colspan="{len(cols)}"></td></tr>')
        prev = r
        out.append("<tr>")
        if headers:
            nums = sorted({n for k, n in marks.items() if int(k.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ")) == r})
            badges = "".join(f'<span class="badge">{n}</span>' for n in nums)
            out.append(f'<td class="rn">{badges}{r}</td>')
        for c in cols:
            ci = column_index_from_string(c)
            if (r, ci) in covered:
                continue
            cell = ws[f"{c}{r}"]
            raw = book.value(sheet, f"{c}{r}")
            txt = html.escape(fmt_value(raw, cell.number_format))
            style = []
            if cell.fill is not None and cell.fill.fill_type == "solid":
                bg = rgb(cell.fill.fgColor)
                if bg:
                    style.append(f"background:{bg}")
            f = cell.font
            if f is not None:
                if f.b:
                    style.append("font-weight:700")
                if f.i:
                    style.append("font-style:italic")
                col = rgb(f.color) if f.color is not None else None
                if col and col != "#000000":
                    style.append(f"color:{col}")
                if f.sz and f.sz >= 14:
                    style.append(f"font-size:{f.sz / 11:.2f}em")
            numeric = isinstance(raw, (int, float)) and not isinstance(raw, bool)
            h = cell.alignment.horizontal if cell.alignment is not None else None
            if numeric or h == "right":
                style.append("text-align:right")
            if cell.alignment is not None and cell.alignment.wrap_text:
                style.append("white-space:normal")
            if cell.border is not None and cell.border.left is not None and cell.border.left.style:
                style.append("box-shadow:inset 0 0 0 1px #BFBFBF")
            span = merged.get((r, ci), 1)
            attrs = f' colspan="{span}"' if span > 1 else ""
            if f"{c}{r}" in marks:
                style.append(f"outline:4px solid {BRAND['coral']};outline-offset:-2px")
            out.append(f'<td{attrs} style="{";".join(style)}">{txt}</td>')
        out.append("</tr>")
    out.append("</table>")
    return "".join(out)


def tabbar(book: Book, active: str):
    tabs = "".join(f'<span class="tab{" on" if n == active else ""}">{html.escape(n)}</span>'
                   for n in book.wb.sheetnames)
    return f'<div class="tabbar">{tabs}</div>'


def sheet_card(inner, book=None, active=None, max_h=1150, max_w=1800):
    bar = tabbar(book, active) if book else ""
    return (f'<div class="card" style="max-width:{max_w}px"><div class="fit" data-maxh="{max_h}" '
            f'data-maxw="{max_w - 40}">{inner}</div>{bar}</div>')


# ------------------------------------------------------------------ charts (SVG from evaluated cells)
def column_chart(labels, values, w=900, h=520, title="Net profit by month"):
    pad_l, pad_b, pad_t = 110, 70, 70
    vmax = max(max(values), 0) or 1
    vmin = min(min(values), 0)
    span = vmax - vmin or 1
    plot_h = h - pad_b - pad_t
    zero_y = pad_t + plot_h * vmax / span
    bw = (w - pad_l - 30) / len(values)
    parts = [f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" class="chart">',
             f'<text x="{w / 2}" y="40" text-anchor="middle" class="ct">{html.escape(title)}</text>']
    for frac in (0, 0.5, 1):
        val = vmin + span * frac
        y = pad_t + plot_h * (1 - frac)
        parts.append(f'<line x1="{pad_l}" x2="{w - 20}" y1="{y:.1f}" y2="{y:.1f}" stroke="#E2E2E2"/>'
                     f'<text x="{pad_l - 12}" y="{y + 8:.1f}" text-anchor="end" class="ax">${val:,.0f}</text>')
    for i, (lab, v) in enumerate(zip(labels, values)):
        x = pad_l + i * bw + bw * 0.18
        bh = abs(v) / span * plot_h
        y = zero_y - bh if v >= 0 else zero_y
        color = BRAND["teal"] if v >= 0 else BRAND["coral"]
        parts.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bw * 0.64:.1f}" height="{max(bh, 1):.1f}" fill="{color}"/>')
        parts.append(f'<text x="{x + bw * 0.32:.1f}" y="{h - pad_b + 34}" text-anchor="middle" class="ax">{html.escape(lab)}</text>')
    parts.append(f'<line x1="{pad_l}" x2="{w - 20}" y1="{zero_y:.1f}" y2="{zero_y:.1f}" stroke="#888"/></svg>')
    return "".join(parts)


def bar_chart(labels, values, w=900, h=520, title="Where your fees went"):
    pad_l, pad_t, pad_r = 330, 70, 120
    vmax = max(abs(v) for v in values) or 1
    vmin = min(min(values), 0)
    span = max(values) - vmin or 1
    plot_w = w - pad_l - pad_r
    zero_x = pad_l + plot_w * (-vmin) / span
    bh = (h - pad_t - 20) / len(values)
    parts = [f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" class="chart">',
             f'<text x="{w / 2}" y="40" text-anchor="middle" class="ct">{html.escape(title)}</text>']
    for i, (lab, v) in enumerate(zip(labels, values)):
        y = pad_t + i * bh + bh * 0.18
        bw = abs(v) / span * plot_w
        x = zero_x if v >= 0 else zero_x - bw
        color = BRAND["teal"] if v >= 0 else BRAND["coral"]
        parts.append(f'<text x="{pad_l - 16}" y="{y + bh * 0.42:.1f}" text-anchor="end" class="ax">{html.escape(lab)}</text>')
        parts.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{max(bw, 1):.1f}" height="{bh * 0.64:.1f}" fill="{color}"/>')
        sv = f"-${abs(v):,.2f}" if v < 0 else f"${v:,.2f}"
        tx = (x + bw + 10) if v >= 0 else (zero_x + 10)
        parts.append(f'<text x="{tx:.1f}" y="{y + bh * 0.42:.1f}" class="ax">{sv}</text>')
    parts.append(f'<line x1="{zero_x:.1f}" x2="{zero_x:.1f}" y1="{pad_t - 10}" y2="{h - 10}" stroke="#888"/></svg>')
    return "".join(parts)


# ------------------------------------------------------------------ page shell
CSS = f"""
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{ width: {SIZE}px; height: {SIZE}px; overflow: hidden; }}
body {{ background: {BRAND['cream']}; font-family: {BRAND['font']}; color: {BRAND['ink']};
       display: flex; flex-direction: column; padding: 90px 100px 60px; position: relative; }}
.brand {{ display: flex; align-items: center; gap: 22px; font-weight: 800; font-size: 46px; color: {BRAND['navy']}; }}
.brand .tag {{ margin-left: auto; font-size: 34px; font-weight: 700; color: #fff; background: {BRAND['teal']};
              padding: 12px 26px; border-radius: 40px; }}
h1 {{ font-size: 128px; line-height: 1.02; font-weight: 900; color: {BRAND['navy']}; margin-top: 56px;
      letter-spacing: -2px; }}
h1 em {{ font-style: normal; color: {BRAND['teal']}; }}
.sub {{ font-size: 50px; color: {BRAND['muted']}; margin-top: 26px; font-weight: 600; line-height: 1.25; }}
.stage {{ flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center;
         margin-top: 40px; gap: 34px; min-height: 0; }}
.card {{ background: #fff; border-radius: 26px; box-shadow: 0 20px 60px rgba(31,58,95,.18); padding: 22px 20px 0;
        overflow: hidden; display: flex; flex-direction: column; align-items: flex-start; }}
.fit {{ transform-origin: top left; }}
table.sheet {{ border-collapse: collapse; table-layout: fixed; font-family: {BRAND['sheet_font']}; font-size: 15px;
              color: #000; }}
table.sheet td {{ height: 24px; padding: 2px 5px; white-space: nowrap; overflow: visible; position: relative;
                 border: 1px solid #ECECEC; vertical-align: middle; }}
table.sheet tr.hdr td, table.sheet td.rn, table.sheet td.corner {{ background: #F1F3F4; color: #5F6368;
                 text-align: center; font-size: 13px; border-color: #D9D9D9; }}
table.sheet tr.gap td {{ height: 14px; background: #FAFAFA; font-size: 12px; color: #999; }}
.badge {{ display: inline-block; margin-right: 6px; vertical-align: middle; width: 30px; height: 30px;
         border-radius: 50%; background: {BRAND['coral']}; color: #fff; font: 800 18px/30px {BRAND['font']};
         text-align: center; z-index: 3; box-shadow: 0 2px 6px rgba(0,0,0,.25); }}
.tabbar {{ display: flex; gap: 4px; margin-top: 18px; width: calc(100% + 40px); margin-left: -20px;
          background: #F1F3F4; padding: 10px 16px; border-top: 1px solid #D9D9D9; }}
.tab {{ font: 600 22px {BRAND['sheet_font']}; color: #5F6368; padding: 8px 18px; border-radius: 6px; }}
.tab.on {{ background: #fff; color: {BRAND['navy']}; box-shadow: 0 1px 3px rgba(0,0,0,.2); }}
.callouts {{ display: flex; gap: 26px; flex-wrap: wrap; justify-content: center; }}
.co {{ display: flex; align-items: center; gap: 16px; font-size: 40px; font-weight: 700; color: {BRAND['ink']};
      background: #fff; padding: 16px 30px 16px 16px; border-radius: 60px; box-shadow: 0 6px 20px rgba(31,58,95,.12); }}
.co b {{ width: 54px; height: 54px; border-radius: 50%; background: {BRAND['coral']}; color: #fff; text-align: center;
        font: 800 32px/54px {BRAND['font']}; flex: none; }}
.note {{ text-align: center; font-size: 24px; color: #8A8F99; }}
body.with-note {{ padding-bottom: 150px; }}  /* keeps the bottom strip clear of gallery dots */
.chart .ct {{ font: 700 30px {BRAND['font']}; fill: {BRAND['navy']}; }}
.chart .ax {{ font: 500 22px {BRAND['font']}; fill: #4A4F59; }}
.row {{ display: flex; gap: 34px; align-items: stretch; }}
.panel {{ background: #fff; border-radius: 26px; box-shadow: 0 20px 60px rgba(31,58,95,.14); padding: 46px 50px; }}
.big {{ font-size: 150px; font-weight: 900; color: {BRAND['teal']}; line-height: 1; }}
.list {{ list-style: none; font-size: 44px; line-height: 1.3; }}
.list li {{ display: flex; gap: 24px; align-items: flex-start; padding: 16px 0; border-bottom: 2px solid #EEE9DF; }}
.list li:last-child {{ border-bottom: none; }}
.list .k {{ font-weight: 800; color: {BRAND['navy']}; min-width: 380px; }}
.list .v {{ color: {BRAND['muted']}; font-weight: 500; }}
.pills {{ display: flex; flex-wrap: wrap; gap: 22px; justify-content: center; }}
.pill {{ font-size: 42px; font-weight: 800; color: #fff; background: {BRAND['navy']}; padding: 20px 36px; border-radius: 60px; }}
.pill.t {{ background: {BRAND['teal']}; }}
.step {{ display: flex; gap: 40px; align-items: center; background: #fff; border-radius: 26px; padding: 40px 50px;
        box-shadow: 0 12px 40px rgba(31,58,95,.12); width: 100%; }}
.step .n {{ width: 130px; height: 130px; border-radius: 50%; background: {BRAND['teal']}; color: #fff; flex: none;
           font: 900 72px/130px {BRAND['font']}; text-align: center; }}
.step h3 {{ font-size: 60px; font-weight: 800; color: {BRAND['navy']}; }}
.step p {{ font-size: 38px; color: {BRAND['muted']}; margin-top: 8px; font-weight: 500; }}
.col h2 {{ font-size: 62px; font-weight: 900; color: {BRAND['navy']}; margin-bottom: 18px; }}
.col.full h2 {{ color: {BRAND['teal']}; }}
.col ul {{ list-style: none; font-size: 38px; line-height: 1.3; }}
.col li {{ padding: 12px 0 12px 50px; position: relative; }}
.col li::before {{ content: "\\2713"; position: absolute; left: 0; color: {BRAND['teal']}; font-weight: 900; }}
.col li.plus::before {{ content: "+"; color: {BRAND['coral']}; }}
"""

FIT_JS = """
<script>
document.fonts.ready.then(() => {
  document.querySelectorAll('.fit').forEach(el => {
    const mw = +el.dataset.maxw, mh = +el.dataset.maxh;
    const w = el.scrollWidth, h = el.scrollHeight;
    const s = Math.min(mw / w, mh / h);
    el.style.transform = 'scale(' + s + ')';
    el.style.width = w + 'px'; el.style.height = h + 'px';
    el.parentElement.style.width = (w * s + 40) + 'px';
    el.style.marginBottom = (h * s - h) + 'px';
    el.style.marginRight = (w * s - w) + 'px';
  });
});
</script>
"""

KITE = (f'<svg width="64" height="64" viewBox="0 0 64 64"><path d="M32 4 L56 26 L32 60 L8 26 Z" fill="{BRAND["teal"]}"/>'
        f'<path d="M32 4 L32 60 M8 26 L56 26" stroke="#fff" stroke-width="3"/></svg>')


def page(headline, sub, stage, tag="Google Sheets + Excel", note=True):
    return (f'<!doctype html><html><head><meta charset="utf-8">{FONT_LINK}<style>{CSS}</style></head>'
            + ('<body class="with-note">' if note else "<body>")
            + f'<div class="brand">{KITE}MarginKite<span class="tag">{html.escape(tag)}</span></div>'
            f"<h1>{headline}</h1>" + (f'<div class="sub">{sub}</div>' if sub else "")
            + f'<div class="stage">{stage}'
            + (f'<div class="note">{SAMPLE_NOTE}</div>' if note else "") + "</div>" + FIT_JS + "</body></html>")


def callouts(items):
    return '<div class="callouts">' + "".join(f'<div class="co"><b>{i}</b>{html.escape(t)}</div>'
                                              for i, t in enumerate(items, 1)) + "</div>"


def cols(a, b):
    return [get_column_letter(i) for i in range(column_index_from_string(a), column_index_from_string(b) + 1)]


# ------------------------------------------------------------------ the images
def images(t: Book, lite: Book):
    out = {}
    D, PC, OL = "Dashboard", "Price Calculator", "Orders Log"

    # month table rows that hold data, plus the total row
    month_rows = list(range(6, 12)) + [19]
    labels = [t.value(D, f"B{r}") for r in range(7, 19)]
    profits = [float(t.value(D, f"J{r}") or 0) for r in range(7, 19)]
    n_data = max(i for i, v in enumerate(profits) if v) + 1 if any(profits) else 3
    chart = column_chart(labels[:max(n_data, 3) + 1], profits[:max(n_data, 3) + 1], w=820, h=470)
    fee_labels = [t.value(D, f"A{r}") for r in range(31, 37)]
    fee_vals = [float(t.value(D, f"D{r}") or 0) for r in range(31, 37)]

    # 1 hero
    dash = render_range(t, D, month_rows, ["A"] + cols("C", "K"), marks={"J19": 1})
    stage = (sheet_card(dash, t, D, max_h=430, max_w=1800)
             + f'<div class="row"><div class="panel" style="padding:24px">{chart}</div>'
               f'<div class="panel" style="padding:24px">{bar_chart(fee_labels, fee_vals, w=820, h=470)}</div></div>')
    out["01-hero.png"] = page("See your real profit after every <em>Etsy &amp; Shopify</em> fee",
                              "MarginKite Seller Profit Tracker", stage)

    # 2 price calculator
    pc = render_range(t, PC, list(range(3, 28)), cols("B", "D"),
                      marks={"C22": 1, "D22": 1, "C25": 2, "D25": 2, "D21": 3})
    out["02-price-calculator.png"] = page(
        "Etsy vs Shopify, <em>side by side</em>",
        f"One product, your price and shipping: every fee line for both platforms.",
        callouts(["Every fee line", "Net profit & margin", "Shopify plan cost per order"])
        + sheet_card(pc, t, PC, max_h=1050, max_w=1300))

    # 3 reverse margin
    target = f"{float(t.value(PC, 'C11')) * 100:g}%"
    rv = render_range(t, PC, list(range(3, 12)) + list(range(29, 32)), cols("B", "D"),
                      marks={"C11": 1, "C31": 2, "D31": 2})
    out["03-reverse-margin.png"] = page(
        f"What price do I need for <em>{target} margin?</em>",
        "Type a target margin. Reverse mode gives the item price for Etsy and for Shopify.",
        callouts(["Your target margin", "Price to charge, per platform"])
        + sheet_card(rv, t, PC, max_h=900, max_w=1500))

    # 4 orders log
    ol_cols = ["A", "B", "C", "I", "J", "K", "L"] + cols("M", "S") + ["V", "W"]
    ol = render_range(t, OL, list(range(5, 16)) + [26], ol_cols,
                      marks={"I8": 1, "Q8": 1, "J26": 2, "K26": 2, "R26": 2, "V6": 3})
    out["04-orders-log.png"] = page(
        "Fees &amp; profit <em>fill in automatically</em>",
        "One row per order, 500 rows ready. Offsite Ads, refunds and fee credits included.",
        callouts(["Offsite Ads fee with the $100 cap", "Refunds & fee credits", "Profit per order"])
        + sheet_card(ol, t, OL, max_h=1000, max_w=1800))

    # 5 break-even
    be = render_range(t, PC, list(range(33, 41)), cols("B", "E"), marks={"C38": 1})
    n = t.text(PC, "C38")
    price = t.text(PC, "C6")
    out["05-shopify-break-even.png"] = page(
        "When does <em>Shopify pay off?</em>",
        "Compares Shopify's plan cost with the Etsy fees you'd save on each order.",
        f'<div class="panel" style="text-align:center"><div class="big">{n} orders / month</div>'
        f'<div class="sub" style="margin-top:20px">break-even for the sample product at {price}, '
        f'Shopify {t.text("Settings", "C7")} billed {t.text("Settings", "C8").lower()}</div></div>'
        + sheet_card(be, t, PC, max_h=560, max_w=1700))

    # 6 what's included
    blurbs = {
        "Start Here": "Setup steps, colour key, how to clear the sample data",
        "Settings": "Dated fee table with sources; Shopify plan & billing picker",
        "Products": "Materials, packaging & labour into a unit cost",
        "Price Calculator": "Etsy vs Shopify, target-margin price, break-even",
        "Orders Log": "500 rows: Offsite Ads, discounts, refunds",
        "Expenses": "Monthly overhead, Etsy Ads spend, Shopify plan",
        "Dashboard": "Monthly profit, top & bottom products, 2 charts",
    }
    items = "".join(f'<li><span class="k">{html.escape(s)}</span><span class="v">{html.escape(blurbs.get(s, ""))}</span></li>'
                    for s in t.wb.sheetnames)
    out["06-whats-included.png"] = page(
        f"What's included: <em>{len(t.wb.sheetnames)} tabs</em>", None,
        f'<div class="panel" style="width:100%"><ul class="list">{items}</ul></div>'
        '<div class="pills"><span class="pill">Google Sheets + Excel 2016+</span>'
        '<span class="pill">Instant download</span>'
        f'<span class="pill t">US fees dated {FEES_AS_OF_TEXT}</span>'
        '<span class="pill">No macros, no password</span></div>', note=False)

    # 7 how it works
    steps = [("Make your copy", "Open the link in your PDF guide and click Make a copy. Or open the .xlsx in Excel."),
             ("Add your products", "Materials, packaging and labour minutes give a unit cost."),
             ("Log your orders", "Pick platform and product; fees fill in for you."),
             ("See your profit", "Monthly profit, margin and your best and worst products.")]
    st = "".join(f'<div class="step"><div class="n">{i}</div><div><h3>{h}</h3><p>{p}</p></div></div>'
                 for i, (h, p) in enumerate(steps, 1))
    out["07-how-it-works.png"] = page("How it works", "Four steps to your first real profit number.",
                                      st, note=False)

    # lite 1 hero
    LS = lite.wb.sheetnames[0]
    lt = render_range(lite, LS, list(range(4, 22)), cols("B", "C"), marks={"C18": 1, "C20": 2, "C21": 2})
    out["lite-01-hero.png"] = page(
        "Free Etsy <em>fee calculator</em>",
        "One item: every Etsy fee, your profit and margin.",
        callouts(["Listing, transaction, processing & Offsite Ads", "Profit & margin"])
        + sheet_card(lt, None, None, max_h=1000, max_w=1100),
        tag="Free · Google Sheets + Excel")

    # lite 2 what you get / upgrade
    free = ["One item at a time", "Listing, transaction & processing fees",
            "Offsite Ads at 15% or 12%, with the $100 cap", "Net profit and margin",
            f"Dated fee table ({FEES_AS_OF_TEXT})"]
    full = ["Etsy and Shopify side by side", "Price for your target margin", "When Shopify pays off",
            "Orders Log with refunds & fee credits (500 rows)", "Products, expenses & monthly Dashboard"]
    out["lite-02-upgrade.png"] = page(
        "Free lite, or the <em>full tracker</em>", None,
        '<div class="row" style="width:100%">'
        '<div class="panel col" style="flex:1"><h2>Free Lite</h2><ul>'
        + "".join(f"<li>{html.escape(x)}</li>" for x in free) + "</ul></div>"
        '<div class="panel col full" style="flex:1"><h2>Full Tracker</h2><ul><li>Everything in Lite</li>'
        + "".join(f'<li class="plus">{html.escape(x)}</li>' for x in full) + "</ul></div></div>",
        tag="Free · Google Sheets + Excel", note=False)
    return out


# ------------------------------------------------------------------ Creative Market variant (3:2 landscape)
CM_W, CM_H = 1820, 1214
CM_DIR = HERE / "creative-market"
CSS_WIDE = CSS.replace(f"html, body {{ width: {SIZE}px; height: {SIZE}px;",
                       f"html, body {{ width: {CM_W}px; height: {CM_H}px;") + f"""
body.wide {{ flex-direction: row; padding: 70px 80px; gap: 64px; }}
body.wide.stacked {{ flex-direction: column; gap: 0; }}
.wide .left {{ width: 600px; flex: none; display: flex; flex-direction: column; }}
.wide .right {{ flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center;
               gap: 30px; min-width: 0; }}
.wide .brand {{ font-size: 36px; gap: 18px; }}
.wide .right .row {{ gap: 22px; }}
.wide .chart .ax {{ font-size: 19px; }}
.wide .brand .tag {{ font-size: 20px; padding: 7px 14px; white-space: nowrap; }}
.wide .tabbar {{ gap: 2px; padding: 8px 12px; overflow: hidden; }}
.wide .tab {{ font-size: 18px; padding: 6px 10px; white-space: nowrap; }}
.wide h1 {{ font-size: 92px; margin-top: 60px; letter-spacing: -1.5px; }}
.wide .sub {{ font-size: 38px; margin-top: 24px; }}
.wide .left .callouts {{ flex-direction: column; align-items: flex-start; margin-top: 44px; gap: 18px; }}
.wide .co {{ font-size: 32px; padding: 12px 26px 12px 12px; }}
.wide .co b {{ width: 46px; height: 46px; font: 800 28px/46px {BRAND['font']}; }}
.wide .left .note {{ margin-top: auto; text-align: left; font-size: 22px; }}
.wide .right .note {{ font-size: 22px; }}
.wide.stacked h1 {{ margin-top: 34px; }}
.wide.stacked .stage {{ margin-top: 30px; gap: 26px; }}
.wide .big {{ font-size: 92px; white-space: nowrap; }}
.wide .list {{ font-size: 34px; }}
.wide .list li {{ padding: 11px 0; }}
.wide .list .k {{ min-width: 300px; }}
.wide .left .pills {{ flex-direction: column; align-items: flex-start; margin-top: 44px; gap: 16px; }}
.wide .pill {{ font-size: 30px; padding: 14px 28px; }}
.wide .step {{ padding: 26px 40px; gap: 32px; }}
.wide .step .n {{ width: 100px; height: 100px; font: 900 56px/100px {BRAND['font']}; }}
.wide .step h3 {{ font-size: 46px; }}
.wide .step p {{ font-size: 30px; }}
"""


def brand_bar(tag):
    return f'<div class="brand">{KITE}MarginKite<span class="tag">{html.escape(tag)}</span></div>'


def page_wide(headline, sub, right, left_extra="", note=True, tag="Google Sheets + Excel"):
    """Two columns: text on the left (headline, sub, callouts, note), visual on the right."""
    left = (brand_bar(tag) + f"<h1>{headline}</h1>" + (f'<div class="sub">{sub}</div>' if sub else "")
            + left_extra + (f'<div class="note">{SAMPLE_NOTE}</div>' if note else ""))
    return (f'<!doctype html><html><head><meta charset="utf-8">{FONT_LINK}<style>{CSS_WIDE}</style></head>'
            f'<body class="wide"><div class="left">{left}</div><div class="right">{right}</div>'
            + FIT_JS + "</body></html>")


def page_wide_stacked(headline, sub, stage, note=True, tag="Google Sheets + Excel"):
    """Full-width layout for very wide sheets: headline row on top, visual below, note under it."""
    return (f'<!doctype html><html><head><meta charset="utf-8">{FONT_LINK}<style>{CSS_WIDE}</style></head>'
            f'<body class="wide stacked">{brand_bar(tag)}<h1>{headline}</h1>'
            + (f'<div class="sub">{sub}</div>' if sub else "")
            + f'<div class="stage">{stage}' + (f'<div class="note">{SAMPLE_NOTE}</div>' if note else "")
            + "</div>" + FIT_JS + "</body></html>")


def images_cm(t: Book):
    """The 7 tracker images for Creative Market at 3:2. Same content as images(), landscape layout."""
    out = {}
    D, PC, OL = "Dashboard", "Price Calculator", "Orders Log"
    labels = [t.value(D, f"B{r}") for r in range(7, 19)]
    profits = [float(t.value(D, f"J{r}") or 0) for r in range(7, 19)]
    n_data = max(i for i, v in enumerate(profits) if v) + 1 if any(profits) else 3
    k = max(n_data, 3) + 1
    fee_labels = [t.value(D, f"A{r}") for r in range(31, 37)]
    fee_vals = [float(t.value(D, f"D{r}") or 0) for r in range(31, 37)]

    dash = render_range(t, D, list(range(6, 12)) + [19], ["A"] + cols("C", "K"), marks={"J19": 1})
    out["cm-01-hero.png"] = page_wide(
        "See your real profit after every <em>Etsy &amp; Shopify</em> fee", "MarginKite Seller Profit Tracker",
        sheet_card(dash, t, D, max_h=330, max_w=990)
        + f'<div class="row"><div class="panel" style="padding:12px">{column_chart(labels[:k], profits[:k], w=460, h=380)}</div>'
          f'<div class="panel" style="padding:12px">{bar_chart(fee_labels, fee_vals, w=470, h=380)}</div></div>')

    pc = render_range(t, PC, list(range(3, 28)), cols("B", "D"),
                      marks={"C22": 1, "D22": 1, "C25": 2, "D25": 2, "D21": 3})
    out["cm-02-price-calculator.png"] = page_wide(
        "Etsy vs Shopify, <em>side by side</em>",
        "One product, your price and shipping: every fee line for both platforms.",
        sheet_card(pc, t, PC, max_h=1000, max_w=990),
        callouts(["Every fee line", "Net profit & margin", "Shopify plan cost per order"]))

    target = f"{float(t.value(PC, 'C11')) * 100:g}%"
    rv = render_range(t, PC, list(range(3, 12)) + list(range(29, 32)), cols("B", "D"),
                      marks={"C11": 1, "C31": 2, "D31": 2})
    out["cm-03-reverse-margin.png"] = page_wide(
        f"What price do I need for <em>{target} margin?</em>",
        "Type a target margin. Reverse mode gives the item price for Etsy and for Shopify.",
        sheet_card(rv, t, PC, max_h=1000, max_w=990),
        callouts(["Your target margin", "Price to charge, per platform"]))

    ol_cols = ["A", "B", "C", "I", "J", "K", "L"] + cols("M", "S") + ["V", "W"]
    ol = render_range(t, OL, list(range(5, 16)) + [26], ol_cols,
                      marks={"I8": 1, "Q8": 1, "J26": 2, "K26": 2, "R26": 2, "V6": 3})
    out["cm-04-orders-log.png"] = page_wide_stacked(
        "Fees &amp; profit <em>fill in automatically</em>",
        "One row per order, 500 rows ready. Offsite Ads, refunds and fee credits included.",
        callouts(["Offsite Ads fee with the $100 cap", "Refunds & fee credits", "Profit per order"])
        + sheet_card(ol, t, OL, max_h=600, max_w=1660))

    be = render_range(t, PC, list(range(33, 41)), cols("B", "E"), marks={"C38": 1})
    n = t.text(PC, "C38")
    price = t.text(PC, "C6")
    out["cm-05-shopify-break-even.png"] = page_wide(
        "When does <em>Shopify pay off?</em>",
        "Compares Shopify's plan cost with the Etsy fees you'd save on each order.",
        f'<div class="panel" style="text-align:center;width:100%"><div class="big">{n} orders / month</div>'
        f'<div class="sub" style="margin-top:16px;font-size:32px">break-even for the sample product at {price}, '
        f'Shopify {t.text("Settings", "C7")} billed {t.text("Settings", "C8").lower()}</div></div>'
        + sheet_card(be, t, PC, max_h=480, max_w=990))

    blurbs = {
        "Start Here": "Setup steps, colour key, how to clear the sample data",
        "Settings": "Dated fee table with sources; Shopify plan & billing picker",
        "Products": "Materials, packaging & labour into a unit cost",
        "Price Calculator": "Etsy vs Shopify, target-margin price, break-even",
        "Orders Log": "500 rows: Offsite Ads, discounts, refunds",
        "Expenses": "Monthly overhead, Etsy Ads spend, Shopify plan",
        "Dashboard": "Monthly profit, top & bottom products, 2 charts",
    }
    items = "".join(f'<li><span class="k">{html.escape(x)}</span><span class="v">{html.escape(blurbs.get(x, ""))}</span></li>'
                    for x in t.wb.sheetnames)
    out["cm-06-whats-included.png"] = page_wide(
        f"What's included: <em>{len(t.wb.sheetnames)} tabs</em>", None,
        f'<div class="panel" style="width:100%"><ul class="list">{items}</ul></div>',
        '<div class="pills"><span class="pill">Google Sheets + Excel 2016+</span>'
        '<span class="pill">Instant download</span>'
        f'<span class="pill t">US fees dated {FEES_AS_OF_TEXT}</span>'
        '<span class="pill">No macros, no password</span></div>', note=False)

    steps = [("Open your copy", "Open the .xlsx in Excel, or make a Google Sheets copy (links in the PDF guide)."),
             ("Add your products", "Materials, packaging and labour minutes give a unit cost."),
             ("Log your orders", "Pick platform and product; fees fill in for you."),
             ("See your profit", "Monthly profit, margin and your best and worst products.")]
    st = "".join(f'<div class="step"><div class="n">{i}</div><div><h3>{h}</h3><p>{p}</p></div></div>'
                 for i, (h, p) in enumerate(steps, 1))
    out["cm-07-how-it-works.png"] = page_wide("How it works", "Four steps to your first real profit number.",
                                              st, note=False)
    return out


# ------------------------------------------------------------------ capture
def find_browser():
    for p in EDGE_CANDIDATES:
        if Path(p).exists():
            return p
    sys.exit("Edge or Chrome not found.")


def capture(browser, html_path: Path, png_path: Path, profile: Path, w=SIZE, h=SIZE):
    subprocess.run([browser, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run",
                    "--force-device-scale-factor=1", f"--user-data-dir={profile}",
                    f"--window-size={w},{h}", "--virtual-time-budget=15000",
                    f"--screenshot={png_path}", html_path.as_uri()],
                   check=True, capture_output=True, timeout=120)
    img = Image.open(png_path)
    if img.size != (w, h):
        sys.exit(f"{png_path.name}: unexpected size {img.size}")
    img.convert("RGB").save(png_path, optimize=True)


def main():
    t = Book(DIST / TRACKER_NAME)
    lite = Book(DIST / LITE_NAME)
    pages = images(t, lite)
    browser = find_browser()
    if THUMBS.exists():
        shutil.rmtree(THUMBS)
    THUMBS.mkdir()
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        for name, doc in pages.items():
            hp = tmp / (name[:-4] + ".html")
            hp.write_text(doc, encoding="utf-8")
            out = HERE / name
            capture(browser, hp, out, tmp / "profile")
            Image.open(out).resize((300, 300), Image.LANCZOS).save(THUMBS / name)
            print(f"{name}  {os.path.getsize(out) // 1024} KB")
        # Creative Market: 3:2 landscape, tracker only (the lite isn't sold there)
        CM_DIR.mkdir(exist_ok=True)
        for name, doc in images_cm(t).items():
            hp = tmp / (name[:-4] + ".html")
            hp.write_text(doc, encoding="utf-8")
            out = CM_DIR / name
            capture(browser, hp, out, tmp / "profile", CM_W, CM_H)
            if os.path.getsize(out) >= 5 * 1024 * 1024:
                sys.exit(f"{name}: over 5 MB")
            Image.open(out).resize((400, round(400 * CM_H / CM_W)), Image.LANCZOS).save(THUMBS / name)
            print(f"creative-market/{name}  {os.path.getsize(out) // 1024} KB")


if __name__ == "__main__":
    main()
