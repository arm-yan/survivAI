"""Build the Pinterest pin images for MarginKite batch 01 (week of 2026-10-26).

Usage (from the repo root):
    .venv\\Scripts\\python.exe marketing\\2026-10-26-batch-01\\build_pins.py

How it works:
  1. Every fee rate is read from FEE_ROWS / the Settings defaults in
     products/seller-profit-tracker/build.py. Nothing is typed by hand.
  2. Each worked example is computed below with Decimal arithmetic, using the same
     formulas as the workbooks (lite tab rows 13-21; Price Calculator rows 14-38).
     Every computed number is printed with its arithmetic to worked-examples.txt so the
     drafts file (posts.md) can show it, and asserted against the figures quoted there.
  3. Each 1000x1500 (2:3) pin is laid out in HTML in the MarginKite style (palette and
     fonts copied from products/seller-profit-tracker/previews/build_previews.py) and
     captured with headless Edge.
"""

from __future__ import annotations

import html
import subprocess
import sys
import tempfile
from decimal import ROUND_CEILING, ROUND_HALF_UP, Decimal as D
from pathlib import Path

from PIL import Image

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
PRODUCT = ROOT / "products" / "seller-profit-tracker"
sys.path.insert(0, str(PRODUCT))
sys.path.insert(0, str(PRODUCT.parent / "brand"))
from build import FEE_ROWS, FEES_AS_OF_TEXT  # noqa: E402
from mark import GAP_RATIO, SORA_LINK, mark_svg, wordmark_css  # noqa: E402  (logo defined once)

OUT = HERE / "pins"
W, H = 1000, 1500
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

# ------------------------------------------------------------------ rates from build.py
RATE = {row[0]: D(str(row[2])) for row in FEE_ROWS if isinstance(row[2], (int, float))}
LISTING = RATE["ETSY_LISTING"]          # 0.20 per order (lite: one unit)
TXN = RATE["ETSY_TXN"]                  # 6.5% of item + shipping
PROC_PCT = RATE["ETSY_PROC_PCT"]        # 3% of order total
PROC_FIXED = RATE["ETSY_PROC_FIXED"]    # 0.25 per order
OFF15 = RATE["ETSY_OFFSITE_15"]         # 15%
OFF12 = RATE["ETSY_OFFSITE_12"]         # 12%
OFF_CAP = RATE["ETSY_OFFSITE_CAP"]      # 100
TAX_BASE = RATE["RULE_TAX_BASE"]        # 0 by default
SHOP_PLAN = RATE["Shopify|Basic|Monthly"]   # Settings default: Basic, Monthly
SHOP_PCT = RATE["Shopify|Basic|Card %"]     # Settings default provider: Shopify Payments
SHOP_FIXED = RATE["Shopify|Basic|Card fixed"]

CENT = D("0.01")


def c(x: D) -> D:
    """Round to cents, half up (how the workbook displays $ cells)."""
    return x.quantize(CENT, rounding=ROUND_HALF_UP)


def usd(x: D) -> str:
    return f"${c(x):,.2f}"


def pct(x: D, dp=1) -> str:
    return f"{(x * 100).quantize(D(1).scaleb(-dp), rounding=ROUND_HALF_UP)}%"


def pr(x: D) -> str:
    return f"{(x * 100).normalize():f}%"


LOG: list[str] = []


def log(s=""):
    LOG.append(s)


def etsy_fees(item: D, ship: D, offsite: D = D(0)):
    """Lite-tab formulas: C13..C18."""
    total = item + ship
    f = {
        "Order total (buyer pays)": total,
        "Listing fee": LISTING,
        "Transaction fee": TXN * total,
        "Payment processing": PROC_PCT * total * (1 + TAX_BASE) + PROC_FIXED,
        "Offsite Ads fee": min(offsite * total, OFF_CAP),
    }
    f["Total Etsy fees"] = f["Listing fee"] + f["Transaction fee"] + f["Payment processing"] + f["Offsite Ads fee"]
    return f


def log_fees(name, item, ship, offsite=D(0)):
    f = etsy_fees(item, ship, offsite)
    t = f["Order total (buyer pays)"]
    log(f"## {name}")
    log(f"Order total = item {usd(item)} + shipping charged {usd(ship)} = {usd(t)}")
    log(f"Listing fee = {usd(LISTING)}")
    log(f"Transaction fee = {pr(TXN)} x {usd(t)} = {f['Transaction fee']:f} -> {usd(f['Transaction fee'])}")
    log(f"Payment processing = {pr(PROC_PCT)} x {usd(t)} + {usd(PROC_FIXED)} = {f['Payment processing']:f} -> {usd(f['Payment processing'])}")
    if offsite:
        raw = offsite * t
        capped = " (under the $100 cap)" if raw <= OFF_CAP else f" -> capped at {usd(OFF_CAP)}"
        log(f"Offsite Ads fee = {pr(offsite)} x {usd(t)} = {raw:f}{capped} -> {usd(f['Offsite Ads fee'])}")
    log(f"Total Etsy fees = {f['Total Etsy fees']:f} -> {usd(f['Total Etsy fees'])} "
        f"= {pct(f['Total Etsy fees'] / t, 2)} of the order total")
    log()
    return f


def check(label, got: str, want: str):
    if got != want:
        sys.exit(f"MISMATCH {label}: computed {got}, drafts say {want}")


# ------------------------------------------------------------------ worked examples
def examples():
    ex = {}
    log(f"MarginKite batch 01 worked examples. Rates from build.py FEE_ROWS, as of {FEES_AS_OF_TEXT}.")
    log(f"Etsy: listing {usd(LISTING)}, transaction {pr(TXN)} of item + shipping, processing {pr(PROC_PCT)} + "
        f"{usd(PROC_FIXED)}, Offsite Ads {pr(OFF15)} / {pr(OFF12)} capped at {usd(OFF_CAP)}. "
        f"Shopify (Settings default): Basic monthly {usd(SHOP_PLAN)}, card {pr(SHOP_PCT)} + {usd(SHOP_FIXED)}.")
    log("Rounding: each line is computed exactly, then shown to the cent (half up).")
    log()

    # E1: $25 item + $5 shipping
    e1 = log_fees("E1 (pins P01, P09; Reddit R1): $25 item + $5 shipping, no Offsite Ad", D(25), D(5))
    keep = e1["Order total (buyer pays)"] - e1["Total Etsy fees"]
    log(f"Left after Etsy fees, before your costs = {usd(e1['Order total (buyer pays)'])} - {usd(e1['Total Etsy fees'])} = {usd(keep)}")
    log()
    ex["E1"] = e1
    for k, want in [("Transaction fee", "$1.95"), ("Payment processing", "$1.15"), ("Total Etsy fees", "$3.30")]:
        check(f"E1 {k}", usd(e1[k]), want)
    check("E1 keep", usd(keep), "$26.70")

    # E1b: $30 item + free shipping (same total)
    e1b = log_fees("E1b (pin P09): $30 item + free shipping, no Offsite Ad", D(30), D(0))
    check("E1b total", usd(e1b["Total Etsy fees"]), "$3.30")

    # E2: same order through an Offsite Ad at 15%
    e2 = log_fees("E2 (pin P02; Reddit R2): $25 item + $5 shipping, sold through an Offsite Ad (15%)", D(25), D(5), OFF15)
    check("E2 offsite", usd(e2["Offsite Ads fee"]), "$4.50")
    check("E2 total", usd(e2["Total Etsy fees"]), "$7.80")
    check("E2 pct", pct(e2["Total Etsy fees"] / D(30), 0), "26%")
    ex["E2"] = e2

    # E3: fees on shipping alone ($8 shipping)
    s = D(8)
    t3, p3 = TXN * s, PROC_PCT * s
    log("## E3 (pin P03; Reddit R3): the part of the fees that comes from $8 shipping charged")
    log(f"Transaction fee on shipping = {pr(TXN)} x {usd(s)} = {usd(t3)}")
    log(f"Processing % on shipping = {pr(PROC_PCT)} x {usd(s)} = {usd(p3)}")
    log(f"Fees on the shipping line = {usd(t3)} + {usd(p3)} = {usd(t3 + p3)} (the $0.25 fixed fee is per order, so not counted here)")
    log()
    check("E3", usd(t3 + p3), "$0.76")
    ex["E3"] = (t3, p3)

    # E4: small vs larger order
    e4a = log_fees("E4a (pin P04; Reddit R4): $4 item, no shipping charged", D(4), D(0))
    e4b = log_fees("E4b (pin P04; Reddit R4): $40 item, no shipping charged", D(40), D(0))
    check("E4a", usd(e4a["Total Etsy fees"]), "$0.83")
    check("E4a pct", pct(e4a["Total Etsy fees"] / D(4), 2), "20.75%")
    check("E4b", usd(e4b["Total Etsy fees"]), "$4.25")
    check("E4b pct", pct(e4b["Total Etsy fees"] / D(40), 1), "10.6%")
    ex["E4"] = (e4a, e4b)

    # E5: lite calculator's own example item (build_lite rows 5-10)
    item, ship, ship_cost, make = D("32.00"), D("6.95"), D("7.20"), D("14.20")
    e5 = log_fees("E5 (pin P05): the free calculator's example item ($32 + $6.95 shipping)", item, ship)
    costs = ship_cost + make
    profit = e5["Order total (buyer pays)"] - e5["Total Etsy fees"] - costs
    margin = profit / e5["Order total (buyer pays)"]
    log(f"Your costs = shipping cost {usd(ship_cost)} + cost to make {usd(make)} = {usd(costs)}")
    log(f"Net profit = {usd(e5['Order total (buyer pays)'])} - {e5['Total Etsy fees']:f} - {usd(costs)} = {profit:f} -> {usd(profit)}")
    log(f"Margin = {profit:f} / {usd(e5['Order total (buyer pays)'])} = {pct(margin, 1)}")
    log()
    check("E5 fees", usd(e5["Total Etsy fees"]), "$4.15")
    check("E5 profit", usd(profit), "$13.40")
    check("E5 margin", pct(margin, 1), "34.4%")
    ex["E5"] = (e5, costs, profit, margin)

    # E6: Offsite Ads cap
    e6_raw = OFF15 * D(800)
    log("## E6 (pin P06; Reddit R5): Offsite Ads cap on an $800 order")
    log(f"{pr(OFF15)} x $800.00 = {usd(e6_raw)}; cap is {usd(OFF_CAP)}, so the fee is {usd(min(e6_raw, OFF_CAP))} "
        f"(saves {usd(e6_raw - OFF_CAP)} vs no cap)")
    log(f"Cap kicks in above {usd(OFF_CAP / OFF15)} at 15% and above {usd(OFF_CAP / OFF12)} at 12%")
    log()
    check("E6", usd(min(e6_raw, OFF_CAP)), "$100.00")
    check("E6 thr15", usd(OFF_CAP / OFF15), "$666.67")
    check("E6 thr12", usd(OFF_CAP / OFF12), "$833.33")

    # E7: 15% vs 12% on $50
    a, b = OFF15 * 50, OFF12 * 50
    log("## E7 (pin P07): Offsite Ads 15% vs 12% on a $50 order")
    log(f"15% x $50.00 = {usd(a)}; 12% x $50.00 = {usd(b)}; difference {usd(a - b)}")
    log()
    check("E7", f"{usd(a)} {usd(b)}", "$7.50 $6.00")

    # E8: reverse mode, Price Calculator C31 formula (no Offsite Ad)
    unit, sc, sch, tgt = D(10), D(5), D(5), D("0.30")
    fixed = LISTING + PROC_FIXED + unit + sc
    var = TXN + PROC_PCT * (1 + TAX_BASE)
    total_needed = fixed / (1 - var - tgt)
    price = (total_needed - sch).quantize(CENT, rounding=ROUND_CEILING)
    log("## E8 (pin P08; Reddit R6): price for a 30% margin (Price Calculator reverse mode, C31)")
    log(f"Inputs: unit cost {usd(unit)}, your shipping cost {usd(sc)}, shipping charged {usd(sch)}, target margin 30%, no Offsite Ad")
    log(f"Fixed per order = listing {usd(LISTING)} + processing fixed {usd(PROC_FIXED)} + unit cost {usd(unit)} + shipping cost {usd(sc)} = {usd(fixed)}")
    log(f"Variable share = transaction {pr(TXN)} + processing {pr(PROC_PCT)} = {pr(var)}")
    log(f"Order total needed = {usd(fixed)} / (1 - {pr(var)} - 30%) = {usd(fixed)} / {1 - var - tgt:f} = {total_needed:.4f}")
    log(f"Item price = {total_needed:.4f} - shipping charged {usd(sch)} = {total_needed - sch:.4f} -> rounded up {usd(price)}")
    chk = etsy_fees(price, sch)
    pf = chk["Order total (buyer pays)"] - chk["Total Etsy fees"] - unit - sc
    log(f"Check at {usd(price)}: order total {usd(price + sch)}; fees {usd(LISTING)} + {usd(chk['Transaction fee'])} + "
        f"{usd(chk['Payment processing'])} = {usd(chk['Total Etsy fees'])}; profit {usd(price + sch)} - {usd(chk['Total Etsy fees'])} "
        f"- {usd(unit + sc)} = {usd(pf)}; margin {pct(pf / (price + sch), 1)}")
    log()
    check("E8 price", usd(price), "$20.54")
    check("E8 fees", usd(chk["Total Etsy fees"]), "$2.88")
    check("E8 profit", usd(pf), "$7.66")
    check("E8 margin", pct(pf / (price + sch), 1), "30.0%")
    ex["E8"] = (price, chk, pf)

    # E9: When does Shopify pay off (Price Calculator rows 34-38), $25 + $5 order
    etsy_f = e1["Total Etsy fees"]
    shop_f = SHOP_PCT * D(30) + SHOP_FIXED
    save = etsy_f - shop_f
    orders = (SHOP_PLAN / save).quantize(D(1), rounding=ROUND_CEILING)
    log("## E9 (pin P10; Reddit R7): when does a Shopify Basic plan pay for itself? ($25 item + $5 shipping)")
    log(f"Etsy fees per order (E1) = {usd(etsy_f)}")
    log(f"Shopify fees per order before plan = {pr(SHOP_PCT)} x $30.00 + {usd(SHOP_FIXED)} = {usd(shop_f)}")
    log(f"Saved per order on Shopify = {usd(etsy_f)} - {usd(shop_f)} = {usd(save)}")
    log(f"Orders a month to cover Basic (monthly billing) = {usd(SHOP_PLAN)} / {usd(save)} = {SHOP_PLAN / save:.2f} -> rounded up {orders}")
    log("Leaves out the cost of bringing your own traffic to Shopify (ads, apps, theme). A floor, not a promise.")
    log()
    check("E9 shop", usd(shop_f), "$1.17")
    check("E9 save", usd(save), "$2.13")
    check("E9 orders", str(orders), "13")
    ex["E9"] = (etsy_f, shop_f, save, orders)
    return ex


# ------------------------------------------------------------------ style (from build_previews.py BRAND)
BRAND = {
    "navy": "#1F3A5F", "teal": "#0F8B7D", "coral": "#E8735A", "cream": "#F6F3EC",
    "ink": "#1E2430", "muted": "#5E6675",
    "font": "'Inter', 'Segoe UI', Arial, sans-serif",
}
FONT_LINK = ('<link rel="preconnect" href="https://fonts.googleapis.com">'
             '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap" '
             'rel="stylesheet">') + SORA_LINK
MARK_PX = 54
MARK = mark_svg(MARK_PX, "light")  # the Margin bars mark (products/brand/mark.py)
FOOT = (f"Current published US Etsy rates as of {FEES_AS_OF_TEXT} (from fee guides; check etsy.com/legal/fees). "
        "Estimates, not tax advice.")

CSS = f"""
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{ width: {W}px; height: {H}px; overflow: hidden; }}
body {{ background: {BRAND['cream']}; font-family: {BRAND['font']}; color: {BRAND['ink']};
       display: flex; flex-direction: column; padding: 60px 64px 0; position: relative; }}
.brand {{ display: flex; align-items: center; gap: {round(MARK_PX * GAP_RATIO)}px; font-weight: 800; font-size: 36px; color: {BRAND['navy']}; }}
{wordmark_css('.brand .wm', 'light')}
.brand .tag {{ margin-left: auto; font-size: 24px; font-weight: 700; color: #fff; background: {BRAND['teal']};
              padding: 10px 22px; border-radius: 40px; }}
.kicker {{ margin-top: 52px; font-size: 30px; font-weight: 800; letter-spacing: 3px; text-transform: uppercase; color: {BRAND['coral']}; }}
h1 {{ font-size: 82px; line-height: 1.04; font-weight: 900; color: {BRAND['navy']}; margin-top: 14px; letter-spacing: -1.5px; }}
h1 em {{ font-style: normal; color: {BRAND['teal']}; }}
.sub {{ font-size: 34px; color: {BRAND['muted']}; margin-top: 22px; font-weight: 600; line-height: 1.3; }}
.stage {{ flex: 1; display: flex; flex-direction: column; justify-content: center; gap: 28px; padding-bottom: 150px; }}
.card {{ background: #fff; border-radius: 28px; box-shadow: 0 18px 50px rgba(31,58,95,.16); padding: 34px 42px; }}
.line {{ display: flex; justify-content: space-between; align-items: baseline; font-size: 36px; padding: 13px 0;
        border-bottom: 2px solid #EEE9DF; font-weight: 600; }}
.line:last-child {{ border-bottom: none; }}
.line .v {{ font-weight: 800; color: {BRAND['navy']}; font-variant-numeric: tabular-nums; }}
.line .k small {{ display: block; font-size: 24px; color: {BRAND['muted']}; font-weight: 500; margin-top: 2px; }}
.line.tot {{ border-top: 4px solid {BRAND['navy']}; border-bottom: none; margin-top: 6px; padding-top: 20px; font-weight: 900; }}
.line.tot .v {{ color: {BRAND['coral']}; font-size: 46px; }}
.line.good .v {{ color: {BRAND['teal']}; }}
.big {{ font-size: 140px; font-weight: 900; color: {BRAND['teal']}; line-height: 1; }}
.big.coral {{ color: {BRAND['coral']}; }}
.bigcap {{ font-size: 34px; font-weight: 700; color: {BRAND['navy']}; margin-top: 10px; line-height: 1.25; }}
.two {{ display: flex; gap: 26px; }}
.two .card {{ flex: 1; text-align: center; padding: 34px 20px; }}
.two h3 {{ font-size: 32px; color: {BRAND['muted']}; font-weight: 700; }}
.two .n {{ font-size: 92px; font-weight: 900; color: {BRAND['navy']}; margin-top: 8px; }}
.two .n.coral {{ color: {BRAND['coral']}; }}
.two .n.teal {{ color: {BRAND['teal']}; }}
.two p {{ font-size: 28px; color: {BRAND['muted']}; margin-top: 8px; font-weight: 600; line-height: 1.25; }}
.list {{ list-style: none; }}
.list li {{ font-size: 36px; font-weight: 700; padding: 16px 0 16px 64px; position: relative; border-bottom: 2px solid #EEE9DF; line-height: 1.2; }}
.list li:last-child {{ border-bottom: none; }}
.list li::before {{ content: attr(data-n); position: absolute; left: 0; top: 14px; width: 46px; height: 46px; border-radius: 50%;
                   background: {BRAND['coral']}; color: #fff; font: 800 26px/46px {BRAND['font']}; text-align: center; }}
.list li.check::before {{ content: "\\2713"; background: {BRAND['teal']}; }}
.list li small {{ display: block; font-size: 27px; font-weight: 500; color: {BRAND['muted']}; margin-top: 4px; }}
.cta {{ position: absolute; left: 64px; right: 64px; bottom: 74px; background: {BRAND['navy']}; color: #fff; border-radius: 22px;
       padding: 24px 30px; font-size: 32px; font-weight: 800; text-align: center; }}
.cta.teal {{ background: {BRAND['teal']}; }}
.foot {{ position: absolute; left: 40px; right: 40px; bottom: 22px; text-align: center; font-size: 19px; color: #7D828C; line-height: 1.3; }}
.note {{ font-size: 26px; color: {BRAND['muted']}; font-weight: 600; line-height: 1.3; }}
"""

LITE_CTA = "Free Etsy fee calculator (Google Sheets + Excel)"
FULL_CTA = "MarginKite Seller Profit Tracker - $12"


def page(kicker, headline, sub, stage, cta, cta_teal=True, tag="US Etsy sellers"):
    return (f'<!doctype html><html><head><meta charset="utf-8">{FONT_LINK}<style>{CSS}</style></head><body>'
            f'<div class="brand">{MARK}<span class="wm">MarginKite</span><span class="tag">{html.escape(tag)}</span></div>'
            f'<div class="kicker">{kicker}</div><h1>{headline}</h1>'
            + (f'<div class="sub">{sub}</div>' if sub else "")
            + f'<div class="stage">{stage}</div>'
            f'<div class="cta{" teal" if cta_teal else ""}">{html.escape(cta)}</div>'
            f'<div class="foot">{html.escape(FOOT)}</div></body></html>')


def line(k, v, cls="", small=""):
    s = f"<small>{small}</small>" if small else ""
    return f'<div class="line {cls}"><span class="k">{k}{s}</span><span class="v">{v}</span></div>'


def pins(ex):
    e1, e2 = ex["E1"], ex["E2"]
    t3, p3 = ex["E3"]
    e4a, e4b = ex["E4"]
    e5, costs, profit, margin = ex["E5"]
    price, chk, pf = ex["E8"]
    etsy_f, shop_f, save, orders = ex["E9"]
    out = {}

    out["P01-where-your-30-etsy-order-goes.png"] = page(
        "Etsy fees, worked out", "Where your <em>$30</em> Etsy order goes",
        "$25 item + $5 shipping, no Offsite Ad.",
        '<div class="card">'
        + line("Listing fee", usd(LISTING), small="per item sold")
        + line("Transaction fee", usd(e1["Transaction fee"]), small=f"{pr(TXN)} of item + shipping")
        + line("Payment processing", usd(e1["Payment processing"]), small=f"{pr(PROC_PCT)} of $30 + {usd(PROC_FIXED)}")
        + line("Etsy fees", usd(e1["Total Etsy fees"]), "tot")
        + line("Left before your costs", usd(D(30) - e1["Total Etsy fees"]), "good")
        + '</div><div class="note">That is 11% of the order, before materials, packaging and postage.</div>',
        LITE_CTA)

    out["P02-offsite-ads-same-order.png"] = page(
        "Offsite Ads", "Same $30 order. <em>One Offsite Ad.</em>",
        "If the sale came through an Etsy Offsite Ad (15% rate):",
        '<div class="two">'
        f'<div class="card"><h3>No Offsite Ad</h3><div class="n">{usd(e1["Total Etsy fees"])}</div><p>11% of the order</p></div>'
        f'<div class="card"><h3>With Offsite Ad</h3><div class="n coral">{usd(e2["Total Etsy fees"])}</div><p>26% of the order</p></div>'
        '</div><div class="card">'
        + line("Offsite Ads fee", usd(e2["Offsite Ads fee"]), small="15% x $30 order total")
        + line("Everything else", usd(e1["Total Etsy fees"]), small="listing + transaction + processing")
        + line("Total Etsy fees", usd(e2["Total Etsy fees"]), "tot")
        + "</div>",
        LITE_CTA)

    out["P03-fees-on-shipping.png"] = page(
        "Pricing tip", "Etsy takes a cut of <em>shipping</em> too",
        "Transaction and processing fees apply to the shipping you charge.",
        f'<div class="card" style="text-align:center"><div class="big coral">{usd(t3 + p3)}</div>'
        '<div class="bigcap">in fees on $8.00 of shipping charged</div></div>'
        '<div class="card">'
        + line("Transaction fee on shipping", usd(t3), small=f"{pr(TXN)} x $8.00")
        + line("Processing on shipping", usd(p3), small=f"{pr(PROC_PCT)} x $8.00")
        + line("Fees on the shipping line", usd(t3 + p3), "tot")
        + "</div>",
        LITE_CTA)

    out["P04-small-orders-pay-more.png"] = page(
        "Pricing tip", "Why small orders pay <em>more</em> in fees",
        "The $0.20 listing and $0.25 processing fees are fixed per order.",
        '<div class="two">'
        f'<div class="card"><h3>$4 item</h3><div class="n coral">{usd(e4a["Total Etsy fees"])}</div><p>{pct(e4a["Total Etsy fees"] / 4, 2)} of the order</p></div>'
        f'<div class="card"><h3>$40 item</h3><div class="n teal">{usd(e4b["Total Etsy fees"])}</div><p>about {pct(e4b["Total Etsy fees"] / 40, 1)} of the order</p></div>'
        '</div><div class="card">'
        + line("$4: 0.20 + 0.26 + 0.12 + 0.25", usd(e4a["Total Etsy fees"]), small="listing + 6.5% + 3% + fixed")
        + line("$40: 0.20 + 2.60 + 1.20 + 0.25", usd(e4b["Total Etsy fees"]), small="listing + 6.5% + 3% + fixed")
        + '</div><div class="note">No shipping charged in either example, no Offsite Ad.</div>',
        LITE_CTA)

    out["P05-free-etsy-fee-calculator.png"] = page(
        "Free download", "Free <em>Etsy fee</em> calculator",
        "Type in one item. See every fee, your profit and margin.",
        '<div class="card">'
        + line("Order total", usd(e5["Order total (buyer pays)"]), small="example: $32 item + $6.95 shipping")
        + line("Total Etsy fees", usd(e5["Total Etsy fees"]), small="listing, transaction, processing, Offsite Ads")
        + line("Your costs", usd(costs), small="making + shipping")
        + line("Net profit", usd(profit), "good")
        + line("Margin", pct(margin, 1), "tot")
        + '</div><div class="note">One tab. Google Sheets or Excel. Made by MarginKite with help from AI. '
          'Gumroad asks for your email.</div>',
        LITE_CTA, tag="Free download")

    out["P06-offsite-ads-100-cap.png"] = page(
        "Offsite Ads", "The Offsite Ads fee stops at <em>$100</em>",
        "On big orders the cap matters. Example: an $800 order at 15%.",
        '<div class="card">'
        + line("15% x $800", "$120.00", small="what 15% would be")
        + line("Cap per order", usd(OFF_CAP))
        + line("Offsite Ads fee charged", usd(OFF_CAP), "tot")
        + '</div><div class="card">'
        + line("Cap starts to matter above", "$666.67", small="at 15%: $100 / 0.15")
        + line("At the 12% rate", "$833.33", small="$100 / 0.12")
        + "</div>",
        LITE_CTA)

    out["P07-offsite-ads-15-vs-12.png"] = page(
        "Offsite Ads", "15% or 12%? <em>The $10k line</em>",
        "The Offsite Ads rate depends on your Etsy sales in the last 12 months.",
        '<div class="two">'
        '<div class="card"><h3>Under $10k</h3><div class="n coral">15%</div><p>$7.50 on a $50 order<br>(you can opt out)</p></div>'
        '<div class="card"><h3>$10k or more</h3><div class="n teal">12%</div><p>$6.00 on a $50 order<br>(no opt-out)</p></div>'
        '</div><div class="note">Both capped at $100 per order. Only charged when a sale comes from an Offsite Ad.</div>',
        LITE_CTA)

    out["P08-price-for-30-percent-margin.png"] = page(
        "Pricing tip", "What to charge for a <em>30% margin</em>",
        "$10 to make, $5 to ship, $5 shipping charged, no Offsite Ad.",
        f'<div class="card" style="text-align:center"><div class="big">{usd(price)}</div>'
        '<div class="bigcap">item price for a 30% margin on Etsy</div></div>'
        '<div class="card">'
        + line("Order total", usd(price + 5))
        + line("Etsy fees", usd(chk["Total Etsy fees"]), small="0.20 + 1.66 + 1.02")
        + line("Your costs", "$15.00")
        + line("Profit", usd(pf), "good")
        + line("Margin", pct(pf / (price + 5), 1), "tot")
        + "</div>",
        FULL_CTA, cta_teal=False)

    out["P09-free-shipping-same-fees.png"] = page(
        "Pricing tip", "Free shipping doesn't <em>dodge</em> Etsy fees",
        "Fees follow the order total, however you split it.",
        '<div class="two">'
        f'<div class="card"><h3>$25 + $5 shipping</h3><div class="n">{usd(e1["Total Etsy fees"])}</div><p>Etsy fees</p></div>'
        f'<div class="card"><h3>$30, free shipping</h3><div class="n">{usd(e1["Total Etsy fees"])}</div><p>Etsy fees</p></div>'
        '</div><div class="card">'
        + line("Transaction fee", usd(e1["Transaction fee"]), small="6.5% of $30 either way")
        + line("Processing", usd(e1["Payment processing"]), small="3% of $30 + $0.25")
        + line("Listing fee", usd(LISTING))
        + "</div>",
        LITE_CTA)

    out["P10-when-shopify-pays-off.png"] = page(
        "Etsy vs Shopify", "When does Shopify <em>pay for itself?</em>",
        "Same $30 order ($25 + $5 shipping). Shopify Basic, billed monthly.",
        f'<div class="card" style="text-align:center"><div class="big">{orders}</div>'
        '<div class="bigcap">orders a month for the plan to pay for itself</div></div>'
        '<div class="card">'
        + line("Etsy fees per order", usd(etsy_f))
        + line("Shopify card fees per order", usd(shop_f), small="2.9% x $30 + $0.30")
        + line("Saved per order", usd(save), "good")
        + line("Plan $27 / $2.13", f"{orders} orders", "tot")
        + '</div><div class="note">Card rate: check your Shopify admin. Leaves out the cost of bringing your own traffic. A floor, not a promise.</div>',
        FULL_CTA, cta_teal=False)

    out["P11-five-etsy-fees-checklist.png"] = page(
        "Checklist", "5 Etsy costs to put <em>in your price</em>",
        None,
        '<div class="card"><ul class="list">'
        '<li data-n="1">Listing fee<small>$0.20, charged again each time an item sells</small></li>'
        '<li data-n="2">Transaction fee<small>6.5% of item price + shipping</small></li>'
        '<li data-n="3">Payment processing<small>3% + $0.25 per order (US)</small></li>'
        '<li data-n="4">Offsite Ads, when a sale comes from one<small>15% or 12%, capped at $100</small></li>'
        '<li data-n="5">Etsy Ads, if you run them<small>your own daily budget</small></li>'
        "</ul></div>",
        LITE_CTA)

    out["P12-seller-profit-tracker.png"] = page(
        "Spreadsheet template", "Your real profit on <em>every order</em>",
        "Etsy and Shopify fees, refunds and a monthly dashboard.",
        '<div class="card"><ul class="list">'
        '<li class="check" data-n="">Every Etsy fee per order<small>listing, transaction, processing, Offsite Ads</small></li>'
        '<li class="check" data-n="">Etsy vs Shopify side by side</li>'
        '<li class="check" data-n="">Price for your target margin</li>'
        '<li class="check" data-n="">500-order log with refunds</li>'
        '<li class="check" data-n="">12-month profit dashboard</li>'
        '</ul></div><div class="note">Google Sheets + Excel. One honest price: $12. Made by MarginKite with help from AI.</div>',
        FULL_CTA, cta_teal=False, tag="$12")
    return out


def capture(html_path: Path, png_path: Path, profile: Path):
    subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run",
                    "--force-device-scale-factor=1", f"--user-data-dir={profile}",
                    f"--window-size={W},{H}", "--virtual-time-budget=15000",
                    f"--screenshot={png_path}", html_path.as_uri()],
                   check=True, capture_output=True, timeout=120)
    img = Image.open(png_path)
    if img.size != (W, H):
        sys.exit(f"{png_path.name}: unexpected size {img.size}")
    img.convert("RGB").save(png_path, optimize=True)


def main():
    ex = examples()
    (HERE / "worked-examples.txt").write_text("\n".join(LOG) + "\n", encoding="utf-8")
    OUT.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        for name, doc in pins(ex).items():
            hp = tmp / (name[:-4] + ".html")
            hp.write_text(doc, encoding="utf-8")
            capture(hp, OUT / name, tmp / "profile")
            print(name)


if __name__ == "__main__":
    main()
