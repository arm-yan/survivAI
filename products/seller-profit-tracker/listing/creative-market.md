# Creative Market: shop profile and listing for the MarginKite Seller Profit Tracker

**Status:** DRAFT for owner approval. The owner submits; nothing is published.
**Product:** the $12 tracker only. The free lite calculator does **not** go on Creative Market: its cell B24 links to Gumroad.
**Source of truth:** `products/seller-profit-tracker/build.py` (fees as of **2026-10-21**).

## Creative Market rules that shape this draft (researched 2026-10-23)

Creative Market's help center returned HTTP 403 to direct fetches today, so the items below come from search-engine extracts of its own help pages (**S** = search extract of an official page) or from our earlier research, which opened the pages directly (**V**, `research/2026-10-23-marketplaces-armenia.md`). **U** = unverified.

| Rule | What it means for us | Source | Quality |
|---|---|---|---|
| No "links directing customers to purchase the same product elsewhere" in descriptions, titles or screenshots | No Gumroad link or mention anywhere in the listing or previews | [Shop Owner Guidelines](https://support.creativemarket.com/hc/en-us/articles/360002027893-Shop-Owner-Guidelines) | V (2026-10-09) + S |
| Products under 4 GB must be uploaded in full; delivering via a link in a text or PDF file risks removal | The **zip holds both complete .xlsx files**. The Google Sheets links in the PDF are only an optional shortcut, and the guide says so | [Shop Owner Upload Guidelines](https://support.creativemarket.com/hc/en-us/articles/1260803906370-Shop-Owner-Upload-Guidelines) | S |
| Deliver as a **.zip**; .rar and .7z may be removed | One zip: 2 .xlsx + PDF | same | S |
| Truthful pricing (FTC); no fake "introductory" or "limited-time" prices | One plain price per license, no strikethrough | Shop Owner Guidelines | V |
| Templates minimum prices: Personal **$9**, Commercial **$14**, Extended Commercial **$36** | See the price section | [Minimum prices](https://support.creativemarket.com/hc/en-us/articles/4405642917915) | V (2026-06-09) |
| Shop earns **50%** of the list price by default | ~$6.00 per $12 sale | [Earnings](https://support.creativemarket.com/hc/en-us/articles/201193714) | V (2026-08-10) |
| Upload asks Yes/No: was the product "or one of its key features … primarily created using generative AI tools"? An AI label shows on Yes | Recommended answer: **Yes** (see below) | [Navigating Our New AI Label](https://support.creativemarket.com/hc/en-us/articles/26926388691099-Navigating-Our-New-AI-Label) | S (article about 2 years old) |
| Preview images: recommended **1820 × 1214 px** (3:2) or larger; min 910 × 607; max 3640 × 10920; JPG, PNG or GIF; ideally under 5 MB; the first image is the cover shown site-wide | Our previews are 2000 × 2000, so they need re-rendering (see the designer spec) | [Listing a New Product](https://support.creativemarket.com/hc/en-us/articles/360000905593-Listing-a-New-Product), [Product Screenshots FAQ](https://support.creativemarket.com/hc/en-us/articles/201193634-Product-Screenshots-FAQ) | S |
| Description of at least **10 words**; SEO meta description about **155** characters; image alt text 20–100 characters (max 150) | Lengths below are counted | [Creating High-Quality PDPs](https://support.creativemarket.com/hc/en-us/articles/43335464121499-Creating-High-Quality-Product-Description-Pages-PDPs) | S |
| Title: "short, clear, specific". **No character limit found** | Title kept under 60 characters | [Pro Tips for Naming Products](https://creativemarket.com/blog/pro-tips-for-naming-your-creative-market-products) | U for any limit |
| Tags: **no maximum found**; advice is accurate tags over popular ones | 12 tags, in priority order | Listing a New Product (tag section cut off in the extract) | U |
| Category: Templates (spreadsheet templates already sell there); **the exact subcategory list wasn't visible** | Owner picks the closest Templates subcategory in the form | research file + [example listing](https://creativemarket.com/Resellacademy/291292393-1000-Google-Sheets-Templates-%C2%A0MRR) | U for subcategory |
| License tiers: Personal (non-commercial), Commercial (up to 5,000 end products for sale), Extended Commercial; reselling the template itself is prohibited under all of them | See the price section | [General Overview for Licensing](https://support.creativemarket.com/hc/en-us/articles/360021205213-General-Overview-for-Licensing), [License FAQ](https://creativemarket.com/licenses) | S |

---

## Shop profile

**Shop name:** MarginKite

**Tagline (short):**
> Spreadsheets that show online sellers what they really keep. (60 characters)

**About / bio:**
> MarginKite is a small shop making spreadsheet tools for Etsy and Shopify sellers. Our first product, the Seller Profit Tracker, shows what you keep from each sale after every fee, product cost and shipping, side by side for Etsy and Shopify. Every rate sits in a dated table you can update yourself. AI helps us design and build the templates; we check every formula against hand-worked examples before release. Questions: marginkite.shop@gmail.com

No links in the profile apart from the support email. Owner to confirm whether Creative Market's profile has a website field; if it does, leave it empty, because our website is the Gumroad shop that sells the same product.

**Avatar:** the MarginKite logo (kite mark), same as Gumroad.

---

## Listing

### Title

**Recommended (A), 57 characters:**
```
Etsy & Shopify Seller Profit Tracker, Google Sheets Excel
```

**Variant (B), 49 characters:**
```
Etsy Seller Profit Tracker & Fee Calculator Sheet
```

Why A: Creative Market buyers browse Templates, so the title names who it's for (Etsy and Shopify sellers), what it is (profit tracker) and the two apps. B drops Shopify in favour of "fee calculator", the phrase sellers search most elsewhere.

### SEO meta description (about 155 characters max)

> See your real profit after every Etsy and Shopify fee. 7-tab Google Sheets and Excel tracker for US sellers, fees dated 2026-10-21. (131 characters)

### Description (paste as is; contains no links)

```
See what you really keep from each Etsy or Shopify sale after fees, product cost and shipping.

WHO IT'S FOR
US Etsy sellers who want their real profit per order, and anyone asking "would a Shopify store leave me more?"

THE PROBLEM
Etsy charges a listing fee, a transaction fee on item plus shipping, payment processing and sometimes an Offsite Ads fee. Refunds give part of it back. Shopify swaps most of that for a monthly plan and a card fee. Comparing the two on your own numbers takes an afternoon with a calculator.

WHAT IT DOES
- Every Etsy fee for each order: listing, transaction, payment processing and Offsite Ads (15% or 12%, with the $100 cap)
- Etsy vs Shopify side by side for one item: every fee line, net profit and margin
- Reverse mode: type a target margin and get the item price to charge on Etsy and on Shopify
- "When does Shopify pay off?": the orders a month at which a Shopify plan pays for itself, at your price
- Orders Log for 500 orders, with discounts, refunds, partial refunds and Etsy fee credits
- Unit cost from materials, packaging and labour minutes
- Dashboard: 12 months of revenue, fees, costs, overhead and net profit; top 3 and bottom 3 products; 2 charts

WHAT'S INCLUDED (one .zip)
- MarginKite-Seller-Profit-Tracker.xlsx: the full workbook with sample data (5 products, 30 orders, 3 months of expenses), so you can see it working
- MarginKite-Seller-Profit-Tracker-BLANK.xlsx: the same workbook, empty, ready for your shop
- MarginKite-Buyer-Guide.pdf: a 1-page quick start, with optional Google Sheets "make a copy" shortcuts
7 tabs: Start Here, Settings (dated fee table), Products, Price Calculator, Orders Log, Expenses, Dashboard.

HOW IT WORKS
1. Open an .xlsx in Excel, or bring it into Google Sheets.
2. Settings: hourly rate, Etsy sales level, Shopify plan.
3. Products: add what you sell.
4. Orders Log: one row per order. Fees and profit fill in.
5. Dashboard: read your monthly profit.

COMPATIBILITY
- Microsoft Excel 2016 or newer (Windows or Mac) and Google Sheets (free Google account).
- No macros, no add-ons, no external data connections. Sheets are protected with no password, so you can unlock and change anything.
- Yellow cells are for typing; blue cells are calculated and locked against accidents.

GOOD TO KNOW
- US sellers only in this version. Fees are US Etsy and Shopify fees as of 2026-10-21, and each rate shows its date, source and status on the Settings tab. Edit one cell when a fee changes and the whole workbook follows.
- Some Etsy rates are still labelled "unverified" in the template until we confirm all of them, including US payment processing, against Etsy's own pages. Shopify plan prices come from Shopify's pricing page; check your card rate in your Shopify admin.
- All figures are estimates. Check them against your own Etsy and Shopify statements. Not tax or accounting advice.
- Not in this version: Etsy CSV import (you type orders in), non-US fees, tax set-aside.
- License: use the tracker to run your own shop. Reselling or sharing the template files isn't allowed.

MADE WITH AI
MarginKite is a small shop. AI helped design and build this template, and every formula was checked against hand-worked examples before release.

Questions? Email marginkite.shop@gmail.com.

Etsy and Shopify are named only to describe compatibility. MarginKite isn't affiliated with or endorsed by either.
```

### Tags (no limit found; enter in this order and stop where the form stops)

| # | Tag | Why |
|---|---|---|
| 1 | etsy seller spreadsheet | Who it's for, plus the format |
| 2 | profit tracker | What it is |
| 3 | etsy fee calculator | The phrase sellers search most |
| 4 | shopify fees | Our main difference: the comparison |
| 5 | etsy vs shopify | Same, as a comparison search |
| 6 | google sheets template | Format browsers on Creative Market |
| 7 | excel template | Format browsers |
| 8 | small business spreadsheet | Broad business-template browsers |
| 9 | pricing calculator | Backs reverse mode (target-margin price) |
| 10 | order tracker | Backs the 500-row Orders Log |
| 11 | profit and loss | Backs the monthly Dashboard |
| 12 | handmade business | Products tab costs materials, packaging and labour |

Left out: "bookkeeping" and "accounting" (it isn't accounting software), "2026" (goes stale), "free" (not free here), and "AI" (the AI label and description already disclose it).

### Category, compatibility and AI question

- **Category:** Templates, then the closest subcategory the form offers (business, office or planner). [[PLACEHOLDER: owner to pick the exact subcategory; the list wasn't visible to us]]
- **Compatible with / file types:** Microsoft Excel, Google Sheets, PDF (pick whichever the form offers).
- **"Primarily created using generative AI?"** Recommended answer: **Yes.** AI (Claude) designed and built the workbook, and a person reviewed it. Answering No would contradict our own "Made with AI" line. Owner decision: if the label hurts visibility, that is a cost of honesty we accept.

### Price and license (recommendation)

| License | Price | Why |
|---|---|---|
| **Personal** | **$12** | Same as on Gumroad. Above the $9 Templates minimum. No sale price or strikethrough. |
| **Commercial** | **$14** | The Templates minimum. Most buyers are sellers running a business and may prefer this license; $14 keeps the gap small and honest. |
| Extended Commercial | Leave off if the form allows it; otherwise $36 (the minimum) | Buyers don't need it for a tool they use to run their own shop. |

At 50% we earn about $6.00 per Personal sale and $7.00 per Commercial sale. Creative Market may add a buyer fee at checkout. **Open question:** does Creative Market require all three licenses to be priced? We couldn't confirm.

---

## What's included / format / delivery

| | |
|---|---|
| Upload | One file: `MarginKite-Seller-Profit-Tracker.zip` |
| Inside the zip | `MarginKite-Seller-Profit-Tracker.xlsx` (sample data) · `MarginKite-Seller-Profit-Tracker-BLANK.xlsx` (empty) · `MarginKite-Buyer-Guide.pdf` (the Creative Market version, from `listing/buyer-guide-creative-market.md`) |
| Delivery | Instant download from Creative Market after purchase. Nothing is shipped. |
| Google Sheets | Upload an .xlsx and Save as Google Sheets, or use the optional copy links in the PDF |
| Works in | Excel 2016+ (Windows or Mac), Google Sheets |
| Macros | None. Protected with no password. |
| Region | US sellers; fees as of 2026-10-21 |
| Off-platform links | None in the listing, previews or profile. The workbooks have no Gumroad link (checked in build.py; only the lite file links to Gumroad, and it isn't sold here). The PDF links only to the Google Sheets copies and the support email. |

---

## Claim-to-feature checklist

| Claim | Backed by (build.py / dist) | OK |
|---|---|---|
| Every Etsy fee per order incl. Offsite Ads 15%/12%, $100 cap | Orders Log cols N–Q; Settings `ETSY_OFFSITE_*`, C6 | ✓ |
| Etsy vs Shopify side by side, fee lines, profit, margin | Price Calculator rows 13–27 | ✓ |
| Reverse mode: target-margin price for both platforms | Price Calculator C11, rows 29–31 | ✓ |
| When Shopify pays off | Price Calculator rows 33–40 | ✓ |
| 500 orders; discounts, refunds, partial refunds, fee credits | `O_FIRST, O_LAST = 6, 505`; cols H, J, K, L, R, T | ✓ |
| Unit cost from materials, packaging, labour minutes | Products cols B–F; Settings C5 | ✓ |
| Dashboard: 12 months, top/bottom 3, 2 charts | `build_dashboard()` | ✓ |
| 7 tabs with the names listed | `build_tracker()` | ✓ |
| Sample-data version: 5 products, 30 orders, 3 months of expenses | `SAMPLE_PRODUCTS`, `sample_orders()`, `SAMPLE_EXPENSES` | ✓ |
| BLANK version, same default fees | `TRACKER_BLANK_NAME`, `build_tracker(blank=True)`; dist/ | ✓ |
| Zip with 2 .xlsx + PDF | dist/ has both .xlsx; the PDF needs a CM render; the zip is not built yet | Placeholder |
| Dated fee table with source and status; edit one cell | Settings fee table cols C, E, F, G; `S(key)` references | ✓ |
| Fees as of 2026-10-21; US only | `FEES_AS_OF_TEXT`; Start Here banner | ✓ |
| Some Etsy rates labelled unverified | Settings Status column ("UNVERIFIED" on Etsy rows) | ✓ |
| Excel 2016+ and Google Sheets; no macros or add-ons | Docstring compatibility rules; Start Here "Open it"; Tech Lead Sheets export check (notes/log.md) | ✓ |
| No external data connections | Test S3 (no external links) | ✓ |
| No password; formula cells locked; yellow/blue key | `protect()`, `fx()`, `inp()`; Start Here colour key | ✓ |
| Not included: CSV import, non-US fees, tax set-aside | Absent from build.py | ✓ (stated as absent) |
| License: no resale | Creative Market license terms (all tiers) | ✓ (platform rule) |
| AI helped build it; formulas checked by hand | Start Here "Made with AI"; tests/cases.md | ✓ |
| Support email | notes/log.md: marginkite.shop@gmail.com | ✓ |
