"""Build the MarginKite Seller Profit Tracker and the free Etsy Fee Calculator Lite.

Usage (from the repo root):
    .venv\\Scripts\\python.exe products\\seller-profit-tracker\\build.py

Writes both workbooks into ./dist next to this file, deleting the folder first.
The output is byte-for-byte deterministic: fixed document dates, fixed zip
timestamps, no randomness.

Compatibility rules (Excel 2016+ and Google Sheets):
  - only IF, AND, OR, MIN, MAX, SUM, SUMIFS, COUNTIFS, INDEX, MATCH, IFERROR,
    LARGE, SMALL, ROUNDUP, DATE, YEAR, MONTH, ROW, TEXT, ISNUMBER
  - no LET / LAMBDA / XLOOKUP / FILTER / dynamic arrays / macros / external links
  - every rate lives in the Settings fee table and is referenced by cell
"""

from __future__ import annotations

import datetime as dt
import re
import shutil
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).resolve().parent
DIST = HERE / "dist"
TRACKER_NAME = "MarginKite-Seller-Profit-Tracker.xlsx"
LITE_NAME = "MarginKite-Etsy-Fee-Calculator-Lite.xlsx"

FEES_AS_OF = dt.date(2026, 10, 21)
FEES_AS_OF_TEXT = "2026-10-21"
FIXED_STAMP = dt.datetime(2026, 10, 21, 0, 0, 0)

# ---------------------------------------------------------------- styles
NAVY = "1F3A5F"
TEAL = "0F8B7D"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")    # light yellow = you type here
FORMULA_FILL = PatternFill("solid", fgColor="DDEBF7")  # light blue = calculated, locked
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

USD = '"$"#,##0.00;-"$"#,##0.00'
PCT = "0.0%"
PCT2 = "0.00%"
DATE_FMT = "yyyy-mm-dd"
MONTH_FMT = "mmm yyyy"

UNLOCKED = Protection(locked=False)
LOCKED = Protection(locked=True)


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = Font(size=16, bold=True, color=NAVY)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(italic=True, color="595959")


def header(ws, row, col, labels, widths=None):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BOX
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[get_column_letter(col + i)].width = w


def inp(cell, value=None, fmt=None):
    """Style a cell as an input (unlocked, yellow)."""
    if value is not None:
        cell.value = value
    cell.fill = INPUT_FILL
    cell.protection = UNLOCKED
    cell.border = BOX
    if fmt:
        cell.number_format = fmt
    return cell


def fx(cell, formula, fmt=None, bold=False):
    """Style a cell as a formula (locked, blue)."""
    cell.value = formula
    cell.fill = FORMULA_FILL
    cell.protection = LOCKED
    cell.border = BOX
    if fmt:
        cell.number_format = fmt
    if bold:
        cell.font = Font(bold=True)
    return cell


def label(ws, ref, text, bold=False, wrap=False, italic=False, color=None):
    c = ws[ref]
    c.value = text
    c.font = Font(bold=bold, italic=italic, color=color)
    if wrap:
        c.alignment = Alignment(wrap_text=True, vertical="top")
    return c


def protect(ws):
    """Protect with NO password. Buyers can unprotect any time."""
    p = ws.protection
    p.sheet = True
    p.formatColumns = False
    p.formatRows = False
    p.formatCells = False
    p.selectLockedCells = False
    p.selectUnlockedCells = False


def add_list_validation(ws, formula1, ranges, prompt=None):
    dv = DataValidation(type="list", formula1=formula1, allow_blank=True)
    dv.error = "Please pick a value from the list."
    dv.errorTitle = "Pick from the list"
    if prompt:
        dv.prompt = prompt
        dv.showInputMessage = True
    ws.add_data_validation(dv)
    for r in ranges:
        dv.add(r)
    return dv


# ================================================================ Settings
# Fee table rows: (key, fee, value, number format, unit, source, status, used, notes)
ETSY_SRC_A = "https://checkoutpage.com/blog/etsy-fees ; https://www.listadum.com/etsy-fee-calculator"
ETSY_SRC_B = "https://checkoutpage.com/blog/etsy-fees ; https://craftybase.com/etsy-fee-calculator"
ETSY_SRC_OFF = "https://checkoutpage.com/blog/etsy-fees ; https://blog.marmalead.com/etsy-vs-shopify/"
ETSY_SRC_REF = "https://ecomcalctools.com/blog/fees-etsy/etsy-refund-returns/ ; Etsy Community threads"
SHOP_SRC = "https://www.shopify.com/pricing"
SHOP_CARD_SRC = "https://www.shopify.com/pricing (search index) ; https://qualimero.com/en/blog/shopify-transaction-fees"
ETSY_HELP_FEES = "https://help.etsy.com/hc/en-us/articles/115014483627 (cited by third-party guides)"
OFFICIAL_ETSY = "Official page to check: https://www.etsy.com/legal/fees/"

FEE_ROWS = [
    ("ETSY_LISTING", "Etsy listing fee", 0.20, USD, "$ per unit sold", ETSY_SRC_A, "UNVERIFIED", "Yes",
     "Renews when an item sells, again for each extra unit. ASSUMPTION (conservative): $0.20 x quantity on every Etsy order. " + OFFICIAL_ETSY),
    ("ETSY_TXN", "Etsy transaction fee", 0.065, PCT2, "of item price + shipping", ETSY_SRC_B, "UNVERIFIED", "Yes",
     "Charged on item price plus shipping (and gift wrap). " + OFFICIAL_ETSY),
    ("ETSY_PROC_PCT", "Etsy payment processing (US) - percent", 0.03, PCT2, "of order total", ETSY_SRC_B, "UNVERIFIED", "Yes",
     "Charged on the order total, including sales tax (see the sales tax rate below)."),
    ("ETSY_PROC_FIXED", "Etsy payment processing (US) - fixed", 0.25, USD, "$ per order", ETSY_SRC_B, "UNVERIFIED", "Yes", ""),
    ("ETSY_OFFSITE_15", "Etsy Offsite Ads - shop under $10k / 12 months", 0.15, PCT2, "of attributed order", ETSY_SRC_OFF, "UNVERIFIED", "Yes",
     "Used when you answer N to the $10k question above. Opt-out allowed."),
    ("ETSY_OFFSITE_12", "Etsy Offsite Ads - shop at $10k+ / 12 months", 0.12, PCT2, "of attributed order", ETSY_SRC_OFF, "UNVERIFIED", "Yes",
     "Used when you answer Y to the $10k question above. Mandatory at this level."),
    ("ETSY_OFFSITE_CAP", "Etsy Offsite Ads - cap per order", 100, USD, "$ max per order", ETSY_SRC_OFF, "UNVERIFIED", "Yes", ""),
    ("ETSY_OFFSITE_THRESHOLD", "Etsy Offsite Ads - sales threshold", 10000, USD, "$ in last 12 months", ETSY_SRC_OFF, "UNVERIFIED", "Info",
     "For reference: you answer the Y/N question above."),
    ("ETSY_REG_FEE", "Etsy Regulatory Operating Fee (US)", 0, PCT2, "none for US sellers", ETSY_SRC_B, "UNVERIFIED", "No",
     "Applies only in some other countries (e.g. UK). Not modeled in v1 (US only)."),
    ("ETSY_CURRENCY", "Etsy currency conversion", 0.025, PCT2, "only if currencies differ", "Prior knowledge, not checked this session", "UNVERIFIED", "No",
     "Not modeled: US sellers listing in USD with a USD account don't pay it."),
    ("ETSY_PLUS", "Etsy Plus subscription (optional)", 10, USD, "$ per month", "Prior knowledge, not checked this session", "UNVERIFIED", "No",
     "If you subscribe, add it on the Expenses tab."),
    ("ETSY_ADS", "Etsy Ads (onsite)", "Your spend", None, "daily budget you set", "https://www.listadum.com/etsy-fee-calculator", "UNVERIFIED", "Expenses tab",
     "Enter your monthly spend on the Expenses tab."),
    ("ETSY_SHARE_SAVE", "Etsy Share & Save credit", "Not modeled", None, "rate not confirmed", "Competitor listing snippet", "UNVERIFIED", "No",
     "A credit to you; leaving it out understates profit slightly."),
]
for plan, mo, yr, card, third in [("Basic", 27, 19, 0.029, 0.02), ("Grow", 72, 54, 0.027, 0.01),
                                  ("Advanced", 399, 299, 0.025, 0.006), ("Plus", 2300, 2300, 0.0225, 0.002)]:
    plus_note = "ASSUMPTION: Plus is 'from $2,300/mo'; same figure used for yearly. " if plan == "Plus" else ""
    FEE_ROWS += [
        (f"Shopify|{plan}|Monthly", f"Shopify {plan} plan - billed monthly", mo, USD, "$ per month", SHOP_SRC, "VERIFIED", "Yes", plus_note),
        (f"Shopify|{plan}|Yearly", f"Shopify {plan} plan - billed yearly", yr, USD, "$ per month", SHOP_SRC, "VERIFIED", "Yes", plus_note),
        (f"Shopify|{plan}|Card %", f"Shopify Payments online card rate - {plan}", card, PCT2, "of order total", SHOP_CARD_SRC, "UNVERIFIED", "Yes",
         "Check Settings > Payments in your Shopify admin." + (" Some sources say 2.6%; higher figure used (conservative)." if plan == "Grow" else "")
         + (" Some sources say 2.4%; higher figure used (conservative)." if plan == "Advanced" else "")
         + (" Shown as 'about 2.25%'." if plan == "Plus" else "")),
        (f"Shopify|{plan}|Card fixed", f"Shopify Payments fixed fee - {plan}", 0.30, USD, "$ per order", SHOP_CARD_SRC, "UNVERIFIED", "Yes", ""),
        (f"Shopify|{plan}|Third-party %", f"Shopify fee if NOT using Shopify Payments - {plan}", third, PCT2, "of order total", SHOP_SRC, "VERIFIED", "Yes",
         "Charged by Shopify on top of your provider's own fee."),
    ]
FEE_ROWS += [
    ("THIRD_PARTY_PCT", "Your third-party provider's rate (PayPal, Stripe...)", 0.029, PCT2, "of order total", "Not from research: placeholder, enter yours", "ASSUMPTION", "Yes",
     "Only used when Payment provider = Third-party provider."),
    ("THIRD_PARTY_FIXED", "Your third-party provider's fixed fee", 0.30, USD, "$ per order", "Not from research: placeholder, enter yours", "ASSUMPTION", "Yes", ""),
    # ---- rules where the research is ambiguous (conservative reading chosen)
    ("RULE_DISCOUNT_BASE", "RULE: Etsy fees use the price AFTER your discount? (Y/N)", "Y", None, "Y or N", ETSY_HELP_FEES, "UNVERIFIED", "Yes",
     "Y (default) = transaction and Offsite Ads fees are worked out on the discounted price, as Etsy's fees help article says. "
     "Alternative: N = fees on the price BEFORE your discount (higher, more cautious). Processing always uses what the buyer paid."),
    ("RULE_TAX_BASE", "Average sales tax rate on your orders (processing fee only)", 0, PCT2, "% of order total", ETSY_HELP_FEES, "UNVERIFIED", "Yes",
     "Etsy's 3% + $0.25 processing fee includes sales tax. Enter your average tax rate for a more accurate fee. Leave it at 0% if unsure."),
    ("RULE_REFUND_PCT_FEES", "RULE: refunds credit back Etsy % fees in proportion", "Info", None, "", ETSY_SRC_REF, "ASSUMPTION", "Yes",
     "When 'Etsy fees credited back?' = Y: transaction fee and % processing fee are credited in proportion to refund / order total."),
    ("RULE_REFUND_LISTING", "RULE: listing fee credited on refund", "Never", None, "", ETSY_SRC_REF, "UNVERIFIED", "Yes",
     "Per research the listing fee is not refunded."),
    ("RULE_REFUND_PROC_FIXED", "RULE: $0.25 processing fee credited on a FULL refund? (Y/N)", "N", None, "Y or N", "Research silent", "ASSUMPTION", "Yes",
     "N (conservative)."),
    ("RULE_REFUND_OFFSITE", "RULE: Offsite Ads fee credited on refund? (Y/N)", "N", None, "Y or N", ETSY_SRC_REF, "ASSUMPTION", "Yes",
     "N (conservative): sellers report this fee often does not come back."),
    ("RULE_REFUND_SHOPIFY", "RULE: Shopify payment fees credited on refund", "Never", None, "", "Research silent", "ASSUMPTION", "Yes",
     "Conservative: no Shopify fees come back on refunds."),
]

FEE_HEADER_ROW = 18
FEE_FIRST = FEE_HEADER_ROW + 1
FEE_LAST = FEE_FIRST + len(FEE_ROWS) - 1
FEE_ROW_OF = {row[0]: FEE_FIRST + i for i, row in enumerate(FEE_ROWS)}


def S(key):
    """Absolute reference to a fee value in Settings."""
    return f"Settings!$C${FEE_ROW_OF[key]}"


# "Your settings" cells
SET = {
    "hourly": "Settings!$C$5",
    "over10k": "Settings!$C$6",
    "plan": "Settings!$C$7",
    "billing": "Settings!$C$8",
    "provider": "Settings!$C$9",
    "offsite_rate": "Settings!$C$12",
    "plan_cost": "Settings!$C$13",
    "shop_pct": "Settings!$C$14",
    "shop_fixed": "Settings!$C$15",
}


def build_settings(ws):
    title(ws, "Settings", f"Fees as of {FEES_AS_OF_TEXT}. Estimates only - check your own Etsy and Shopify statements.")
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 44
    ws.column_dimensions["G"].width = 13
    ws.column_dimensions["H"].width = 12
    ws.column_dimensions["I"].width = 80

    label(ws, "A4", "YOUR SETTINGS", bold=True, color=TEAL)
    rows = [
        (5, "Your hourly rate for labour ($/hour)", 18, USD, "Used on the Products tab to cost your time."),
        (6, "Etsy sales in the last 12 months $10,000 or more? (Y/N)", "N", None, "Sets the Offsite Ads rate: N = 15%, Y = 12%."),
        (7, "Shopify plan", "Basic", None, "Basic, Grow, Advanced or Plus."),
        (8, "Shopify billing", "Monthly", None, "Monthly or Yearly (yearly is cheaper per month)."),
        (9, "Payment provider on Shopify", "Shopify Payments", None, "Third-party adds Shopify's extra fee on top of your provider's fee."),
    ]
    for r, lab, val, fmt, note in rows:
        label(ws, f"B{r}", lab)
        inp(ws[f"C{r}"], val, fmt)
        label(ws, f"D{r}", note, italic=True, color="595959")

    label(ws, "A11", "CALCULATED FROM YOUR SETTINGS", bold=True, color=TEAL)
    fee_vals = f"$C${FEE_FIRST}:$C${FEE_LAST}"
    fee_keys = f"$A${FEE_FIRST}:$A${FEE_LAST}"

    def lookup(suffix_expr):
        return f'INDEX({fee_vals},MATCH("Shopify|"&$C$7&"|"&{suffix_expr},{fee_keys},0))'

    calc = [
        (12, "Etsy Offsite Ads rate in use", f'=IF($C$6="Y",$C${FEE_ROW_OF["ETSY_OFFSITE_12"]},$C${FEE_ROW_OF["ETSY_OFFSITE_15"]})', PCT2),
        (13, "Shopify plan cost per month (plan + billing)", "=" + lookup("$C$8"), USD),
        (14, "Shopify % fee per order in use", f'=IF($C$9="Shopify Payments",{lookup(chr(34) + "Card %" + chr(34))},'
             f'$C${FEE_ROW_OF["THIRD_PARTY_PCT"]}+{lookup(chr(34) + "Third-party %" + chr(34))})', PCT2),
        (15, "Shopify fixed fee per order in use", f'=IF($C$9="Shopify Payments",{lookup(chr(34) + "Card fixed" + chr(34))},'
             f'$C${FEE_ROW_OF["THIRD_PARTY_FIXED"]})', USD),
    ]
    for r, lab, f, fmt in calc:
        label(ws, f"B{r}", lab)
        fx(ws[f"C{r}"], f, fmt)

    label(ws, "A17", "FEE TABLE - every rate the workbook uses. Edit Value / As of / Source / Status when fees change.",
          bold=True, color=TEAL)
    header(ws, FEE_HEADER_ROW, 1, ["Key (don't edit)", "Fee", "Value", "Unit", "As of", "Source", "Status",
                                   "Used in formulas?", "Notes"])
    yn_cells = []
    for i, (key, fee, val, fmt, unit, src, status, used, notes) in enumerate(FEE_ROWS):
        r = FEE_FIRST + i
        ws.cell(row=r, column=1, value=key).font = Font(color="7F7F7F", size=9)
        ws.cell(row=r, column=2, value=fee)
        if isinstance(val, str) and val in ("Info", "Never", "Your spend", "Not modeled"):
            c = ws.cell(row=r, column=3, value=val)
            c.alignment = Alignment(horizontal="right")
        else:
            inp(ws.cell(row=r, column=3), val, fmt)
            if val in ("Y", "N"):
                yn_cells.append(f"C{r}")
        ws.cell(row=r, column=4, value=unit)
        inp(ws.cell(row=r, column=5), FEES_AS_OF, DATE_FMT)
        inp(ws.cell(row=r, column=6), src)
        st = inp(ws.cell(row=r, column=7), status)
        st.font = Font(bold=True, color={"VERIFIED": "2E7D32", "UNVERIFIED": "C00000"}.get(status, "9C5700"))
        ws.cell(row=r, column=8, value=used)
        ws.cell(row=r, column=9, value=notes).alignment = Alignment(wrap_text=False)

    add_list_validation(ws, '"Y,N"', ["C6"] + yn_cells)
    add_list_validation(ws, '"Basic,Grow,Advanced,Plus"', ["C7"])
    add_list_validation(ws, '"Monthly,Yearly"', ["C8"])
    add_list_validation(ws, '"Shopify Payments,Third-party provider"', ["C9"])
    add_list_validation(ws, '"VERIFIED,UNVERIFIED,ASSUMPTION"', [f"G{FEE_FIRST}:G{FEE_LAST}"])
    ws.freeze_panes = "A4"
    protect(ws)


# ================================================================ Products
P_FIRST, P_LAST = 5, 54
P_NAMES = f"Products!$A${P_FIRST}:$A${P_LAST}"
P_COST = f"Products!$F${P_FIRST}:$F${P_LAST}"

SAMPLE_PRODUCTS = [
    # name, materials, packaging, labour minutes
    ("Ceramic Mug", 6.50, 1.20, 25),
    ("Linen Tote Bag", 5.00, 0.80, 20),
    ("Soy Candle 8oz", 3.40, 1.00, 15),
    ("Beaded Bracelet", 2.10, 0.60, 30),
    ("Art Print A4", 1.50, 0.90, 5),
]


def build_products(ws, products):
    title(ws, "Products", "One row per product. Yellow = you type. Blue = calculated. Unit cost feeds every profit figure.")
    header(ws, 4, 1, ["Product name", "Materials ($/unit)", "Packaging ($/unit)", "Labour (minutes/unit)",
                      "Labour cost ($/unit)", "Total unit cost", "Units sold (Orders Log)",
                      "Net profit (Orders Log)", "Rank helper"],
           [26, 13, 13, 13, 13, 13, 13, 15, 11])
    ws.row_dimensions[4].height = 32
    o = "'Orders Log'"
    for r in range(P_FIRST, P_LAST + 1):
        i = r - P_FIRST
        data = products[i] if i < len(products) else (None, None, None, None)
        inp(ws[f"A{r}"], data[0])
        inp(ws[f"B{r}"], data[1], USD)
        inp(ws[f"C{r}"], data[2], USD)
        inp(ws[f"D{r}"], data[3], "0")
        fx(ws[f"E{r}"], f'=IF($A{r}="","",$D{r}/60*{SET["hourly"]})', USD)
        fx(ws[f"F{r}"], f'=IF($A{r}="","",$B{r}+$C{r}+$E{r})', USD, bold=True)
        fx(ws[f"G{r}"], f'=IF($A{r}="","",SUMIFS({o}!$D${O_FIRST}:$D${O_LAST},{o}!$C${O_FIRST}:$C${O_LAST},$A{r}))', "0")
        fx(ws[f"H{r}"], f'=IF($A{r}="","",SUMIFS({o}!$V${O_FIRST}:$V${O_LAST},{o}!$C${O_FIRST}:$C${O_LAST},$A{r}))', USD)
        fx(ws[f"I{r}"], f'=IF($A{r}="","",$H{r}-ROW()/1000000000)', "0.00")
    label(ws, "K4", "Notes", bold=True)
    label(ws, "K5", "Labour cost = minutes / 60 x your hourly rate (Settings).")
    label(ws, "K6", "Rank helper breaks ties for the Dashboard top/bottom lists. Ignore it.")
    label(ws, "K7", "Product names must be unique; the Orders Log looks them up by name.")
    ws.column_dimensions["K"].width = 60
    ws.freeze_panes = "B5"
    protect(ws)


# ================================================================ Orders Log
O_FIRST, O_LAST = 6, 505
O_COLS = ["Date", "Platform", "Product", "Qty", "Item price (each)", "Shipping charged", "Your shipping cost",
          "Discount ($ off order)", "Sold via Offsite Ad? (Etsy)", "Refunded?", "Partial refund $ (blank = full)",
          "Etsy fees credited back?",
          "Order total (buyer paid)", "Listing fee", "Transaction fee", "Payment processing", "Offsite Ads fee",
          "Fee credit on refund", "Total fees", "Refund amount", "Product cost", "Net profit", "Margin", "Month"]
O_WIDTHS = [11, 10, 20, 6, 10, 10, 10, 10, 10, 10, 12, 10, 11, 9, 11, 11, 10, 10, 10, 10, 10, 11, 8, 10]


def sample_orders(products):
    """30 deterministic sample orders, Aug-Oct 2026."""
    prices = {"Ceramic Mug": (32.00, 6.95, 7.20), "Linen Tote Bag": (28.00, 5.50, 5.00),
              "Soy Candle 8oz": (24.00, 6.00, 6.50), "Beaded Bracelet": (18.00, 0.00, 4.00),
              "Art Print A4": (15.00, 0.00, 1.80)}
    names = [p[0] for p in products]
    out = []
    for i in range(30):
        month = 8 + i // 10
        day = 1 + (i % 10) * 2
        name = names[(i * 3) % 5]
        price, ship, shipcost = prices[name]
        platform = "Shopify" if i % 4 == 3 else "Etsy"
        qty = 2 if i % 7 == 0 else 1
        offsite = "Y" if (platform == "Etsy" and i % 6 == 2) else "N"
        discount = 3.00 if i % 9 == 4 else None
        refunded, partial, credited = "N", None, None
        if i == 11:
            refunded, credited = "Y", "Y"
        elif i == 20:
            refunded, partial, credited = "Y", 8.00, "Y"
        elif i == 25:
            refunded, partial, credited = "Y", 5.00, "N"
        out.append([dt.date(2026, month, day), platform, name, qty, price, ship, shipcost, discount,
                    offsite, refunded, partial, credited])
    return out


def order_formulas(r):
    qty = f'IF($D{r}="",1,$D{r})'
    base = f'($M{r}+IF({S("RULE_DISCOUNT_BASE")}="Y",0,$H{r}))'  # transaction / Offsite Ads fee base
    proc_pct = f'{S("ETSY_PROC_PCT")}*$M{r}*(1+{S("RULE_TAX_BASE")})'
    blank = f'$A{r}=""'
    return {
        "M": (f'=IF({blank},"",{qty}*$E{r}+$F{r}-$H{r})', USD),
        "N": (f'=IF({blank},"",IF($B{r}="Etsy",{qty}*{S("ETSY_LISTING")},0))', USD),
        "O": (f'=IF({blank},"",IF($B{r}="Etsy",{S("ETSY_TXN")}*{base},0))', USD),
        "P": (f'=IF({blank},"",IF($B{r}="Etsy",{proc_pct}+{S("ETSY_PROC_FIXED")},'
              f'IF($B{r}="Shopify",{SET["shop_pct"]}*$M{r}+{SET["shop_fixed"]},0)))', USD),
        "Q": (f'=IF({blank},"",IF(AND($B{r}="Etsy",$I{r}="Y"),MIN({SET["offsite_rate"]}*{base},{S("ETSY_OFFSITE_CAP")}),0))', USD),
        "R": (f'=IF({blank},"",IF(AND($B{r}="Etsy",$J{r}="Y",$L{r}="Y",$M{r}>0),'
              f'($O{r}+{proc_pct}+IF({S("RULE_REFUND_OFFSITE")}="Y",$Q{r},0))*$T{r}/$M{r}'
              f'+IF(AND({S("RULE_REFUND_PROC_FIXED")}="Y",$T{r}>=$M{r}),{S("ETSY_PROC_FIXED")},0),0))', USD),
        "S": (f'=IF({blank},"",$N{r}+$O{r}+$P{r}+$Q{r}-$R{r})', USD),
        "T": (f'=IF({blank},"",IF($J{r}="Y",IF($K{r}="",$M{r},MIN($K{r},$M{r})),0))', USD),
        "U": (f'=IF({blank},"",{qty}*IFERROR(INDEX({P_COST},MATCH($C{r},{P_NAMES},0)),0))', USD),
        "V": (f'=IF({blank},"",$M{r}-$T{r}-$S{r}-$U{r}-$G{r})', USD),
        "W": (f'=IF({blank},"",IF($M{r}-$T{r}>0,$V{r}/($M{r}-$T{r}),""))', PCT),
        "X": (f'=IF({blank},"",DATE(YEAR($A{r}),MONTH($A{r}),1))', MONTH_FMT),
    }


def build_orders(ws, orders):
    title(ws, "Orders Log", "One row per order. Fill the yellow columns A-L; everything blue is worked out for you. "
                            "Sales tax collected by Etsy is NOT revenue - leave it out.")
    label(ws, "A3", "Product cost uses today's unit cost from the Products tab. Pre-sized for 500 orders "
                    f"(rows {O_FIRST}-{O_LAST}).", italic=True, color="595959")
    header(ws, 5, 1, O_COLS, O_WIDTHS)
    ws.row_dimensions[5].height = 58
    fmts = [DATE_FMT, None, None, "0", USD, USD, USD, USD, None, None, USD, None]
    for r in range(O_FIRST, O_LAST + 1):
        i = r - O_FIRST
        data = orders[i] if i < len(orders) else [None] * 12
        for c in range(12):
            inp(ws.cell(row=r, column=c + 1), data[c], fmts[c])
        for col, (f, fmt) in order_formulas(r).items():
            fx(ws[f"{col}{r}"], f, fmt, bold=(col == "V"))
    rng = lambda col: f"{col}{O_FIRST}:{col}{O_LAST}"  # noqa: E731
    add_list_validation(ws, '"Etsy,Shopify"', [rng("B")])
    add_list_validation(ws, P_NAMES.replace("Products!", "Products!"), [rng("C")])
    add_list_validation(ws, '"Y,N"', [rng("I"), rng("J"), rng("L")])
    dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1", allow_blank=True)
    dv.error = "Quantity must be a whole number, 1 or more."
    ws.add_data_validation(dv)
    dv.add(rng("D"))
    dv2 = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)
    dv2.error = "Enter a date, e.g. 2026-10-21."
    ws.add_data_validation(dv2)
    dv2.add(rng("A"))
    ws.freeze_panes = f"D{O_FIRST}"
    protect(ws)


# ================================================================ Expenses
E_FIRST, E_LAST = 6, 65
SAMPLE_EXPENSES = [
    (dt.date(2026, 8, 1), 12.99, 27.00, 8.40, 15.00, 0),
    (dt.date(2026, 9, 1), 12.99, 27.00, 6.15, 15.00, 4.99),
    (dt.date(2026, 10, 1), 12.99, 27.00, 5.80, 20.00, 0),
]


def build_expenses(ws, expenses):
    title(ws, "Expenses", "Monthly overhead that isn't tied to one product. One row per month (or several rows per month - they add up).")
    label(ws, "A3", "Any date inside the month works. Put your Shopify plan bill here; Settings' plan picker only drives the "
                    "Price Calculator.", italic=True, color="595959")
    header(ws, 5, 1, ["Month (any date in it)", "Subscriptions & tools", "Shopify plan & apps", "Supplies (not per product)",
                      "Etsy Ads spend", "Other", "Total", "Month key"], [14, 14, 14, 15, 13, 12, 13, 11])
    ws.row_dimensions[5].height = 45
    for r in range(E_FIRST, E_LAST + 1):
        i = r - E_FIRST
        data = expenses[i] if i < len(expenses) else [None] * 6
        inp(ws[f"A{r}"], data[0], DATE_FMT)
        for c in range(1, 6):
            inp(ws.cell(row=r, column=c + 1), data[c], USD)
        fx(ws[f"G{r}"], f'=IF($A{r}="","",SUM($B{r}:$F{r}))', USD, bold=True)
        fx(ws[f"H{r}"], f'=IF($A{r}="","",DATE(YEAR($A{r}),MONTH($A{r}),1))', MONTH_FMT)
    ws.freeze_panes = f"A{E_FIRST}"
    protect(ws)


# ================================================================ Price Calculator
def build_calculator(ws):
    title(ws, "Price Calculator", "One item, Etsy vs Shopify side by side. Fill the yellow cells.")
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 60

    label(ws, "B3", "INPUTS", bold=True, color=TEAL)
    inputs = [
        (4, "Product", "Ceramic Mug", None, "Pick from your Products tab."),
        (6, "Item price", 32.00, USD, ""),
        (7, "Shipping charged to the buyer", 6.95, USD, "0 for free shipping."),
        (8, "Your shipping cost (label + postage)", 7.20, USD, ""),
        (9, "Sold via Etsy Offsite Ad? (Y/N)", "N", None, "Etsy only. Rate comes from Settings (15% or 12%)."),
        (10, "Expected Shopify orders per month", 30, "0", "Spreads the Shopify plan cost over your orders."),
        (11, "Target margin (for reverse mode)", 0.30, PCT, "Profit as % of what the buyer pays."),
    ]
    for r, lab, val, fmt, note in inputs:
        label(ws, f"B{r}", lab)
        inp(ws[f"C{r}"], val, fmt)
        label(ws, f"E{r}", note, italic=True, color="595959")
    label(ws, "B5", "Your unit cost (from Products)")
    fx(ws["C5"], f'=IFERROR(INDEX({P_COST},MATCH($C$4,{P_NAMES},0)),0)', USD)
    add_list_validation(ws, P_NAMES, ["C4"])
    add_list_validation(ws, '"Y,N"', ["C9"])

    header(ws, 13, 2, ["Per order", "Etsy", "Shopify"])
    ws["E13"] = "Shopify plan & provider are set on the Settings tab."
    ws["E13"].font = Font(italic=True, color="595959")
    off = f'IF($C$9="Y",{SET["offsite_rate"]},0)'
    plan_per = f'IF($C$10>0,{SET["plan_cost"]}/$C$10,0)'
    proc_pct = f'{S("ETSY_PROC_PCT")}*(1+{S("RULE_TAX_BASE")})'
    lines = [
        (14, "Item price", "=$C$6", "=$C$6"),
        (15, "Shipping charged", "=$C$7", "=$C$7"),
        (16, "Order total (buyer pays)", "=C14+C15", "=D14+D15"),
        (17, "Listing fee", f"={S('ETSY_LISTING')}", "=0"),
        (18, "Transaction fee", f"={S('ETSY_TXN')}*C16", "=0"),
        (19, "Payment processing", f"={proc_pct}*C16+{S('ETSY_PROC_FIXED')}",
         f"={SET['shop_pct']}*D16+{SET['shop_fixed']}"),
        (20, "Offsite Ads fee", f"=MIN({off}*C16,{S('ETSY_OFFSITE_CAP')})", "=0"),
        (21, "Shopify plan cost per order", "=0", f"={plan_per}"),
        (22, "Total fees", "=SUM(C17:C21)", "=SUM(D17:D21)"),
        (23, "Product cost", "=$C$5", "=$C$5"),
        (24, "Shipping cost", "=$C$8", "=$C$8"),
        (25, "NET PROFIT", "=C16-C22-C23-C24", "=D16-D22-D23-D24"),
        (26, "Margin (profit / order total)", '=IF(C16>0,C25/C16,"")', '=IF(D16>0,D25/D16,"")'),
    ]
    for r, lab, fe, fs in lines:
        bold = r in (16, 22, 25)
        label(ws, f"B{r}", lab, bold=bold)
        fmt = PCT if r == 26 else USD
        fx(ws[f"C{r}"], fe, fmt, bold=bold)
        fx(ws[f"D{r}"], fs, fmt, bold=bold)
    label(ws, "B27", "Shopify minus Etsy profit per order")
    fx(ws["D27"], "=D25-C25", USD, bold=True)

    # ---- reverse mode
    label(ws, "B29", "REVERSE MODE: price to charge for your target margin", bold=True, color=TEAL)
    header(ws, 30, 2, ["", "Etsy", "Shopify"])
    fixed_e = f"({S('ETSY_LISTING')}+{S('ETSY_PROC_FIXED')}+$C$5+$C$8)"
    var_e = f"({S('ETSY_TXN')}+{proc_pct})"
    t1 = f"{fixed_e}/(1-{var_e}-{off}-$C$11)"
    t2 = f"({fixed_e}+{S('ETSY_OFFSITE_CAP')})/(1-{var_e}-$C$11)"
    etsy_total = f"IF({off}*{t1}>{S('ETSY_OFFSITE_CAP')},{t2},{t1})"
    etsy_price = (f'=IF(1-{var_e}-{off}-$C$11<=0,"Not reachable",'
                  f'MAX(0,ROUNDUP({etsy_total}-$C$7,2)))')
    fixed_s = f"({SET['shop_fixed']}+{plan_per}+$C$5+$C$8)"
    shop_price = (f'=IF(1-{SET["shop_pct"]}-$C$11<=0,"Not reachable",'
                  f'MAX(0,ROUNDUP({fixed_s}/(1-{SET["shop_pct"]}-$C$11)-$C$7,2)))')
    label(ws, "B31", "Item price needed (shipping charged stays as entered)", bold=True)
    fx(ws["C31"], etsy_price, USD, bold=True)
    fx(ws["D31"], shop_price, USD, bold=True)
    label(ws, "E31", "Rounded up to the cent. Uses the same fees, costs and Offsite Ad setting as above.",
          italic=True, color="595959")

    # ---- break-even
    label(ws, "B33", "WHEN DOES SHOPIFY PAY OFF?", bold=True, color=TEAL)
    rows = [
        (34, "Etsy fees per order at this price", "=C22", USD),
        (35, "Shopify fees per order, before the plan cost", "=D22-D21", USD),
        (36, "You save per order on Shopify", "=C34-C35", USD),
        (37, "Shopify plan cost per month (Settings)", f"={SET['plan_cost']}", USD),
        (38, "Orders per month where Shopify's plan pays for itself",
         '=IF(C36<=0,"Never at this price",ROUNDUP(C37/C36,0))', "0"),
    ]
    for r, lab, f, fmt in rows:
        label(ws, f"B{r}", lab, bold=(r == 38))
        fx(ws[f"C{r}"], f, fmt, bold=(r == 38))
    fx(ws["B39"], '=IF(ISNUMBER(C38),"From "&C38&" orders a month at this price, Shopify\'s lower per-order fees cover its plan cost.",'
                  '"At this price Etsy\'s per-order fees are not higher than Shopify\'s, so the plan never pays for itself.")')
    ws.merge_cells("B39:E39")
    label(ws, "B40", "Leaves out the cost of bringing your own traffic to Shopify (ads, apps, theme) and Etsy Ads. "
                     "Treat it as a floor, not a promise.", italic=True, color="595959")
    ws.merge_cells("B40:E40")
    protect(ws)


# ================================================================ Dashboard
D_FIRST, D_LAST = 7, 18


def build_dashboard(ws):
    title(ws, "Dashboard", "Everything here is calculated from the Orders Log and Expenses tabs.")
    label(ws, "A4", "First month to show")
    inp(ws["C4"], dt.date(2026, 8, 1), MONTH_FMT)
    label(ws, "D4", "Shows 12 months from this one.", italic=True, color="595959")
    header(ws, 6, 1, ["Month", "Label", "Orders", "Revenue (order totals)", "Refunds", "Fees", "Product cost",
                      "Shipping cost", "Overhead (Expenses)", "Net profit", "Margin"],
           [11, 10, 8, 13, 11, 11, 12, 12, 13, 13, 9])
    ws.row_dimensions[6].height = 32
    o = "'Orders Log'"
    X = f"{o}!$X${O_FIRST}:$X${O_LAST}"

    def oc(col):
        return f"{o}!${col}${O_FIRST}:${col}${O_LAST}"

    for r in range(D_FIRST, D_LAST + 1):
        if r == D_FIRST:
            fx(ws[f"A{r}"], "=DATE(YEAR($C$4),MONTH($C$4),1)", MONTH_FMT)
        else:
            fx(ws[f"A{r}"], f"=DATE(YEAR(A{r-1}),MONTH(A{r-1})+1,1)", MONTH_FMT)
        fx(ws[f"B{r}"], f'=TEXT(A{r},"mmm yy")')
        fx(ws[f"C{r}"], f"=COUNTIFS({X},$A{r})", "0")
        for col, src in (("D", "M"), ("E", "T"), ("F", "S"), ("G", "U"), ("H", "G")):
            fx(ws[f"{col}{r}"], f"=SUMIFS({oc(src)},{X},$A{r})", USD)
        fx(ws[f"I{r}"], f"=SUMIFS(Expenses!$G${E_FIRST}:$G${E_LAST},Expenses!$H${E_FIRST}:$H${E_LAST},$A{r})", USD)
        fx(ws[f"J{r}"], f"=D{r}-E{r}-F{r}-G{r}-H{r}-I{r}", USD, bold=True)
        fx(ws[f"K{r}"], f'=IF(D{r}-E{r}>0,J{r}/(D{r}-E{r}),"")', PCT)
    tr = D_LAST + 1
    label(ws, f"A{tr}", "12-month total", bold=True)
    for col in "CDEFGHIJ":
        fx(ws[f"{col}{tr}"], f"=SUM({col}{D_FIRST}:{col}{D_LAST})", "0" if col == "C" else USD, bold=True)
    fx(ws[f"K{tr}"], f'=IF(D{tr}-E{tr}>0,J{tr}/(D{tr}-E{tr}),"")', PCT, bold=True)

    # ---- top / bottom products
    label(ws, "A22", "TOP 3 PRODUCTS BY NET PROFIT (all orders in the log)", bold=True, color=TEAL)
    label(ws, "G22", "BOTTOM 3 PRODUCTS BY NET PROFIT", bold=True, color=TEAL)
    header(ws, 23, 1, ["#", "Product", "", "Net profit", "Units"])
    header(ws, 23, 7, ["#", "Product", "", "Net profit", "Units"])
    key = f"Products!$I${P_FIRST}:$I${P_LAST}"
    for k in range(1, 4):
        r = 23 + k
        for c0, fn in ((1, "LARGE"), (7, "SMALL")):
            L = lambda off: get_column_letter(c0 + off)  # noqa: E731
            ws[f"{L(0)}{r}"] = k
            m = f"MATCH({fn}({key},{k}),{key},0)"
            fx(ws[f"{L(1)}{r}"], f'=IFERROR(INDEX({P_NAMES},{m}),"")')
            ws.merge_cells(f"{L(1)}{r}:{L(2)}{r}")
            fx(ws[f"{L(3)}{r}"], f'=IFERROR(INDEX(Products!$H${P_FIRST}:$H${P_LAST},{m}),"")', USD)
            fx(ws[f"{L(4)}{r}"], f'=IFERROR(INDEX(Products!$G${P_FIRST}:$G${P_LAST},{m}),"")', "0")

    # ---- fee breakdown (same 12-month window)
    label(ws, "A29", "WHERE YOUR FEES WENT (the 12 months above)", bold=True, color=TEAL)
    header(ws, 30, 1, ["Fee type", "", "", "Amount"])
    win = f'{X},">="&$A${D_FIRST},{X},"<="&$A${D_LAST}'
    fee_lines = [
        ("Etsy listing fees", f"=SUMIFS({oc('N')},{win})"),
        ("Etsy transaction fees", f"=SUMIFS({oc('O')},{win})"),
        ("Etsy payment processing", f'=SUMIFS({oc("P")},{win},{oc("B")},"Etsy")'),
        ("Etsy Offsite Ads", f"=SUMIFS({oc('Q')},{win})"),
        ("Shopify payment fees", f'=SUMIFS({oc("P")},{win},{oc("B")},"Shopify")'),
        ("Fee credits on refunds", f"=-SUMIFS({oc('R')},{win})"),
    ]
    for i, (lab, f) in enumerate(fee_lines):
        r = 31 + i
        ws[f"A{r}"] = lab
        ws.merge_cells(f"A{r}:C{r}")
        fx(ws[f"D{r}"], f, USD)
    label(ws, "A37", "Total", bold=True)
    fx(ws["D37"], "=SUM(D31:D36)", USD, bold=True)
    label(ws, "A38", "Matches the Fees column above. Etsy Ads and the Shopify plan are overhead (Expenses tab).",
          italic=True, color="595959")

    ch = BarChart()
    ch.type = "bar"
    ch.title = "Where your fees went"
    ch.style = 10
    ch.legend = None
    ch.add_data(Reference(ws, min_col=4, min_row=30, max_row=36), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=31, max_row=36))
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.numFmt = '"$"#,##0'
    ch.height, ch.width = 7.5, 16
    ws.add_chart(ch, "M22")

    ch2 = BarChart()
    ch2.type = "col"
    ch2.title = "Net profit by month"
    ch2.style = 10
    ch2.legend = None
    ch2.add_data(Reference(ws, min_col=10, min_row=6, max_row=D_LAST), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=2, min_row=D_FIRST, max_row=D_LAST))
    ch2.x_axis.delete = False
    ch2.y_axis.delete = False
    ch2.y_axis.numFmt = '"$"#,##0'
    ch2.height, ch2.width = 7.5, 16
    ws.add_chart(ch2, "M4")
    protect(ws)


# ================================================================ Start Here
def build_start(ws):
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 95
    ws["B1"] = "MarginKite Seller Profit Tracker"
    ws["B1"].font = Font(size=20, bold=True, color=NAVY)
    ws["B2"] = f"Fees last checked: {FEES_AS_OF_TEXT}  |  US sellers  |  Etsy + Shopify"
    ws["B2"].font = Font(bold=True, color="FFFFFF")
    ws["B2"].fill = PatternFill("solid", fgColor=TEAL)
    ws.merge_cells("B2:C2")

    r = 4

    def section(text):
        nonlocal r
        r += 1
        c = ws[f"B{r}"]
        c.value = text
        c.font = Font(bold=True, size=13, color=TEAL)
        r += 1

    def line(left, right=None, bold=False):
        nonlocal r
        ws[f"B{r}"] = left
        ws[f"B{r}"].font = Font(bold=bold)
        ws[f"B{r}"].alignment = Alignment(vertical="top", wrap_text=True)
        if right is not None:
            ws[f"C{r}"] = right
            ws[f"C{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    section("What it does")
    line("", "Shows what you really keep from each Etsy or Shopify sale after every fee, product cost and shipping. "
             "Compares Etsy and Shopify side by side, tells you the price to charge for a target margin, shows when a "
             "Shopify plan pays for itself, and tracks orders, refunds and overhead month by month.")

    section("Open it")
    line("Google Sheets", "Upload the .xlsx to Google Drive, open it, then File > Save as Google Sheets. "
                          "Work in that Google Sheets copy.")
    line("Excel", "Open the .xlsx in Excel 2016 or newer (Windows or Mac). If you see 'Enable Editing', click it.")

    section("5 steps to your first result")
    steps = [
        "Settings: enter your hourly rate, answer the Etsy $10k question, and pick your Shopify plan, billing and payment provider.",
        "Products: type each product with materials, packaging and labour minutes. Unit cost is worked out for you.",
        "Price Calculator: pick a product and enter price, shipping charged and shipping cost. Read Etsy vs Shopify profit, "
        "the price for your target margin, and when Shopify pays off.",
        "Orders Log: add one row per order (yellow columns A-L). Fees and profit fill in automatically.",
        "Expenses, then Dashboard: add monthly overhead, set the first month on the Dashboard, and read your monthly profit.",
    ]
    for i, s in enumerate(steps, 1):
        line(f"Step {i}", s)

    section("Colour key")
    ws[f"B{r}"].fill = INPUT_FILL
    ws[f"B{r}"].value = "Yellow cell"
    ws[f"B{r}"].border = BOX
    ws[f"C{r}"] = "You type here (inputs)."
    r += 1
    ws[f"B{r}"].fill = FORMULA_FILL
    ws[f"B{r}"].value = "Blue cell"
    ws[f"B{r}"].border = BOX
    ws[f"C{r}"] = "Calculated for you and locked, so it can't be overwritten by accident."
    r += 1
    line("Unlocking", "Sheets are protected with NO password. Excel: Review > Unprotect Sheet. Google Sheets: "
                      "Data > Protect sheets and ranges, then remove the protection. Edit formulas at your own risk.")

    section("Delete the sample data")
    line("", "The file ships with 5 sample products, 30 sample orders (Aug-Oct 2026) and 3 months of expenses so you "
             "can see it working. To start fresh, select these yellow ranges and press Delete (clear the contents; "
             "don't delete rows):")
    line("Orders Log", f"A{O_FIRST}:L{O_LAST}")
    line("Expenses", f"A{E_FIRST}:F{E_LAST}")
    line("Products", f"A{P_FIRST}:D{P_LAST} (after the orders, so nothing points at a missing product)")
    line("Dashboard", "Set 'First month to show' (C4) to your first month.")
    line("Price Calculator", "Pick one of your own products in C4.")

    section("Good to know")
    line("Sales tax", "Etsy collects US sales tax for you. It isn't your revenue, so leave it out of every amount.")
    line("Refunds", "Mark Refunded = Y. Leave 'Partial refund $' blank for a full refund. 'Etsy fees credited back?' = Y "
                    "credits Etsy's % fees in proportion to the refund (see the RULE rows on Settings).")
    line("Offsite Ads", "Set Y only on orders Etsy says came from an Offsite Ad. The 15%/12% rate and the $100 cap come from Settings.")
    line("Fee changes", "Every rate lives in the Settings fee table with its date, source and status. Update a value there "
                        "and the whole workbook follows. Rows marked UNVERIFIED or ASSUMPTION are best estimates.")

    section("Disclaimer")
    line("", f"All figures are estimates. Fees as of {FEES_AS_OF_TEXT}; Etsy fees are not yet verified against Etsy's own "
             "fee page. Check your own Etsy and Shopify statements. This is not tax or accounting advice.")
    section("Made with AI")
    line("", "MarginKite is a small shop. AI helped design and build this template; every formula was checked "
             "against hand-worked examples before release.")
    protect(ws)


# ================================================================ Lite calculator
def build_lite(ws):
    ws.title = "Etsy Fee Calculator"
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 44
    ws.column_dimensions["F"].width = 13
    ws["B1"] = "MarginKite Etsy Fee Calculator (Lite)"
    ws["B1"].font = Font(size=16, bold=True, color=NAVY)
    ws["B2"] = f"Free. US Etsy sellers. Fees as of {FEES_AS_OF_TEXT} (unverified). Estimates only; check your own statements."
    ws["B2"].font = Font(italic=True, color="595959")

    label(ws, "B4", "YOUR ITEM", bold=True, color=TEAL)
    ins = [(5, "Item price", 32.00, USD), (6, "Shipping charged to the buyer", 6.95, USD),
           (7, "Your shipping cost", 7.20, USD), (8, "Your cost to make it", 14.20, USD),
           (9, "Sold via Etsy Offsite Ad? (Y/N)", "N", None),
           (10, "Etsy sales in last 12 months $10,000+? (Y/N)", "N", None)]
    for r, lab, val, fmt in ins:
        label(ws, f"B{r}", lab)
        inp(ws[f"C{r}"], val, fmt)
    add_list_validation(ws, '"Y,N"', ["C9", "C10"])

    # fee settings block (all rates live here)
    label(ws, "B26", "FEE SETTINGS (edit if Etsy changes its fees)", bold=True, color=TEAL)
    header(ws, 27, 2, ["Fee", "Value", "As of", "Source", "Status"])
    fees = [
        (28, "Listing fee ($ per unit)", 0.20, USD, ETSY_SRC_A),
        (29, "Transaction fee (item + shipping)", 0.065, PCT2, ETSY_SRC_B),
        (30, "Payment processing %", 0.03, PCT2, ETSY_SRC_B),
        (31, "Payment processing fixed ($ per order)", 0.25, USD, ETSY_SRC_B),
        (32, "Offsite Ads rate, under $10k", 0.15, PCT2, ETSY_SRC_OFF),
        (33, "Offsite Ads rate, $10k+", 0.12, PCT2, ETSY_SRC_OFF),
        (34, "Offsite Ads cap ($ per order)", 100, USD, ETSY_SRC_OFF),
    ]
    for r, lab, val, fmt, src in fees:
        ws[f"B{r}"] = lab
        inp(ws[f"C{r}"], val, fmt)
        inp(ws[f"D{r}"], FEES_AS_OF, DATE_FMT)
        inp(ws[f"E{r}"], src)
        st = inp(ws[f"F{r}"], "UNVERIFIED")
        st.font = Font(bold=True, color="C00000")
    label(ws, "B35", "Official page to check: https://www.etsy.com/legal/fees/  (sales tax and discounts not modeled)",
          italic=True, color="595959")

    label(ws, "B12", "RESULT", bold=True, color=TEAL)
    off = 'IF($C$9="Y",IF($C$10="Y",$C$33,$C$32),0)'
    res = [
        (13, "Order total (buyer pays)", "=$C$5+$C$6", USD),
        (14, "Listing fee", "=$C$28", USD),
        (15, "Transaction fee", "=$C$29*C13", USD),
        (16, "Payment processing", "=$C$30*C13+$C$31", USD),
        (17, "Offsite Ads fee", f"=MIN({off}*C13,$C$34)", USD),
        (18, "Total Etsy fees", "=SUM(C14:C17)", USD),
        (19, "Your costs (making + shipping)", "=$C$7+$C$8", USD),
        (20, "NET PROFIT", "=C13-C18-C19", USD),
        (21, "Margin (profit / order total)", '=IF(C13>0,C20/C13,"")', PCT),
    ]
    for r, lab, f, fmt in res:
        bold = r in (18, 20, 21)
        label(ws, f"B{r}", lab, bold=bold)
        fx(ws[f"C{r}"], f, fmt, bold=bold)

    ws["B23"] = ("Want every order tracked, Etsy vs Shopify side by side and the price for your target margin? "
                 "That's the full MarginKite Seller Profit Tracker:")
    ws["B23"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("B23:E23")
    ws.row_dimensions[23].height = 30
    ws["B24"] = "[PLACEHOLDER: link to the full tracker goes here before launch]"
    ws["B24"].font = Font(bold=True, color="C00000")
    ws["B24"].fill = WARN_FILL
    ws.merge_cells("B24:E24")

    ws["E5"] = "Yellow = you type"
    ws["E5"].fill = INPUT_FILL
    ws["E6"] = "Blue = calculated (locked, no password)"
    ws["E6"].fill = FORMULA_FILL
    ws["E8"] = "Made by MarginKite with help from AI."
    ws["E8"].font = Font(italic=True, color="595959")
    protect(ws)


# ================================================================ save helpers
def save_deterministic(wb, path: Path):
    wb.properties.creator = "MarginKite"
    wb.properties.lastModifiedBy = "MarginKite"
    wb.properties.created = FIXED_STAMP
    wb.calculation.fullCalcOnLoad = True
    tmp = path.with_suffix(".tmp")
    wb.save(tmp)
    stamp = FIXED_STAMP.strftime("%Y-%m-%dT%H:%M:%SZ")
    with zipfile.ZipFile(tmp) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in sorted(zin.infolist(), key=lambda i: i.filename):
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)",
                              rb"\g<1>" + stamp.encode() + rb"\g<2>", data)
            zi = zipfile.ZipInfo(item.filename, date_time=(2026, 10, 21, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            zout.writestr(zi, data)
    tmp.unlink()


def build_tracker(path: Path):
    wb = Workbook()
    ws_start = wb.active
    ws_start.title = "Start Here"
    ws_set = wb.create_sheet("Settings")
    ws_prod = wb.create_sheet("Products")
    ws_calc = wb.create_sheet("Price Calculator")
    ws_ord = wb.create_sheet("Orders Log")
    ws_exp = wb.create_sheet("Expenses")
    ws_dash = wb.create_sheet("Dashboard")
    build_start(ws_start)
    build_settings(ws_set)
    build_products(ws_prod, SAMPLE_PRODUCTS)
    build_calculator(ws_calc)
    build_orders(ws_ord, sample_orders(SAMPLE_PRODUCTS))
    build_expenses(ws_exp, SAMPLE_EXPENSES)
    build_dashboard(ws_dash)
    for ws, color in ((ws_start, TEAL), (ws_set, "7F7F7F"), (ws_dash, NAVY)):
        ws.sheet_properties.tabColor = color
    save_deterministic(wb, path)


def build_lite_file(path: Path):
    wb = Workbook()
    build_lite(wb.active)
    save_deterministic(wb, path)


def main():
    if DIST.exists():
        shutil.rmtree(DIST)
    DIST.mkdir(parents=True)
    build_tracker(DIST / TRACKER_NAME)
    build_lite_file(DIST / LITE_NAME)
    print(f"Built {DIST / TRACKER_NAME}")
    print(f"Built {DIST / LITE_NAME}")


if __name__ == "__main__":
    main()
