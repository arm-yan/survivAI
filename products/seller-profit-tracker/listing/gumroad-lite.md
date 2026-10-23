# Gumroad listing: MarginKite Etsy Fee Calculator (Lite), free

**Status:** DRAFT for owner approval. Nothing is published.
**Source of truth:** `build_lite()` in `products/seller-profit-tracker/build.py` (fees as of **2026-10-21**).
**Price:** **$0+** (pay what you want, minimum $0). Gumroad asks for the buyer's email at checkout, and the description says so.

Limits: product name up to **255** characters on Gumroad (per Gumroad open-source docs, checked 2026-10-23). Kept under 70 so it isn't cut off on cards.

---

## Product name

**Recommended (A), 58 / 255 characters:**

```
Free Etsy Fee Calculator (Lite): Google Sheets + Excel, US
```

**Variant (B), 53 / 255 characters:**

```
MarginKite Etsy Fee Calculator Lite: Free Spreadsheet
```

Why A: "free etsy fee calculator" is the phrase people search and share. It's honest about what it is: one tab, Etsy only. B leads with the brand; use it if A looks too much like the big web calculators.

---

## Summary (Gumroad "Summary" field)

> Every Etsy fee on one item, plus your profit and margin. Free, one tab, US sellers, fees dated 2026-10-21. (106 characters)

**Additional details:**

| Key | Value |
|---|---|
| Format | 2 .xlsx (example + BLANK) + 1 Google Sheets copy link |
| Works in | Google Sheets, Excel 2016+ |
| Tabs | 1 |
| Fees as of | 2026-10-21 (US Etsy) |
| Price | Free (pay what you want) |

---

## Description (paste as is)

```
Type in one item and see every Etsy fee, what you keep and your margin.

WHO IT'S FOR
US Etsy sellers who want a quick, honest fee check before setting a price.

WHAT IT DOES
✔ Listing fee, transaction fee (on item + shipping) and payment processing
✔ Offsite Ads at 15% or 12% (answer one Y/N about your last 12 months), with the $100 cap
✔ Your costs: making it + shipping it
✔ Net profit and margin on the order
✔ A small fee table with the date and source of each rate, so you can update it when Etsy changes fees

HOW IT WORKS
1. Make your Google Sheets copy (link after checkout) or open an .xlsx in Excel.
2. Fill the yellow cells: item price, shipping charged, your shipping cost, your cost to make it, and two Y/N questions.
3. Read total Etsy fees, net profit and margin in the blue cells.

FORMAT
• Free digital download. Nothing is shipped.
• Two one-tab .xlsx files (one with an example item, one BLANK) + one Google Sheets "make a copy" link.
• Google Sheets (free account) or Excel 2016 or newer.
• No macros. Protected with no password.

GOOD TO KNOW
• US Etsy sellers. Fees as of 2026-10-21. [[PLACEHOLDER: keep "Etsy rates come from published 2026 fee guides and are marked 'unverified' until checked against Etsy's own fee page." until the owner verifies them; then replace with "checked against Etsy's fee page on [[PLACEHOLDER: date]]".]]
• Estimates only. Check your own Etsy statements.
• Not modeled in the lite version: sales tax, discounts, refunds, Shopify.
• Gumroad asks for your email to send the file. We may email you about MarginKite tools now and then; unsubscribe any time.

WANT MORE THAN ONE ITEM?
The full MarginKite Seller Profit Tracker ($12) adds:
+ Etsy and Shopify side by side
+ The price to charge for your target margin
+ When a Shopify plan pays for itself
+ An Orders Log for 500 orders, with refunds and fee credits
+ Products, Expenses and a monthly Dashboard
Get it here: [[PLACEHOLDER: Gumroad URL of the full tracker]]

MADE WITH AI
MarginKite is a small shop. AI helped build this calculator, and its formulas were checked against hand-worked examples.

Etsy is named only to describe compatibility. MarginKite isn't affiliated with or endorsed by Etsy.
```

---

## Tags (Gumroad Share tab > Discover)

| Tag | Why |
|---|---|
| etsy | Main platform |
| etsy fees | The exact problem it answers |
| fee calculator | What it is |
| free template | Free-product browsers |
| spreadsheet template | Format category |

Note: Gumroad shows products on Discover only after a sale, which may never happen for a $0 product. Traffic will come from the links we share (Pinterest, Reddit answers), so the name and summary matter more than tags here.

**Upsell line:** the "WANT MORE THAN ONE ITEM?" block, plus the link already inside the sheet (cell B24, also a placeholder in build.py).

---

## FAQ

**1. Is it really free?**
Yes. Enter $0 at checkout. You can pay something if you want to, but it unlocks nothing extra.

**2. Why do you need my email?**
Gumroad asks for it to deliver the file. We may send occasional MarginKite emails; every one has an unsubscribe link.

**3. Do I need Excel?**
No. Use the Google Sheets copy link with a free Google account (just type over the example numbers), or open either .xlsx in Excel 2016 or newer.

**4. Does it include Offsite Ads?**
Yes: 15% or 12% depending on your last 12 months of Etsy sales, capped at $100 per order.

**5. Does it work for sellers outside the US?**
No. It uses US Etsy fees only.

**6. What's the difference from the full tracker?**
Lite checks one Etsy item at a time. The full tracker ($12) adds Etsy vs Shopify, target-margin pricing, Shopify break-even, a 500-order log with refunds, products, expenses and a monthly dashboard.

**7. Was AI used?**
Yes. AI helped build it, and the formulas were checked against hand-worked examples.

---

## What's included / format / delivery

| | |
|---|---|
| Files | `MarginKite-Etsy-Fee-Calculator-Lite.xlsx` (example item) + `MarginKite-Etsy-Fee-Calculator-Lite-BLANK.xlsx` (empty) + one Google Sheets copy link (on the Gumroad download page) |
| Delivery | Instant on the Gumroad receipt page and by email. Nothing is shipped. |
| Works in | Google Sheets (free account), Excel 2016+ |
| Macros | None. Protected with no password. |
| Region | US Etsy sellers; fees as of 2026-10-21 |
| Placeholders | [[PLACEHOLDER: Google Sheets copy link (lite)]] · [[PLACEHOLDER: Gumroad URL of the full tracker]] (also in the sheet, cell B24) |

---

## Claim-to-feature checklist

| Claim | Backed by (`build_lite()`) | OK |
|---|---|---|
| Listing, transaction (item + shipping), processing fees | Rows 14–16; C13 = price + shipping | ✓ |
| Offsite Ads 15%/12% via one Y/N, $100 cap | C9, C10, rows 32–34, row 17 `MIN(...)` | ✓ |
| Your costs: making + shipping | C7, C8, row 19 | ✓ |
| Net profit and margin | Rows 20–21 | ✓ |
| Fee table with date and source per rate, editable | Rows 27–34, cols C–F | ✓ |
| One tab | `build_lite_file()` creates one sheet | ✓ |
| Two .xlsx: example + BLANK | `LITE_NAME`, `LITE_BLANK_NAME`, `build_lite(blank=True)`; both in dist/ | ✓ |
| One Google Sheets copy link | Tech Lead decision (Ticket #10); lite tab "HOW TO OPEN" rows 37–41 | Placeholder |
| Fees as of 2026-10-21, unverified | B2 text; Status "UNVERIFIED" | ✓ |
| Sales tax, discounts not modeled | B35 note | ✓ |
| Refunds, Shopify not in lite | Absent from `build_lite()` | ✓ (stated as absent) |
| No macros; no password | Docstring; `protect()` | ✓ |
| Made with help from AI | E8 "Made by MarginKite with help from AI." | ✓ |
| Full tracker features listed in upsell | `build_calculator()`, `build_orders()`, `build_products()`, `build_expenses()`, `build_dashboard()` | ✓ |
| Full tracker price $12 | notes/log.md owner decision | ✓ |
| NOT claimed: target-margin price in lite | Research suggested it, but `build_lite()` has no such cell | Left out |
