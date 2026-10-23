# Lite follow-up email sequence (DRAFT)

**Status:** DRAFT for owner approval. Nothing has been set up or sent.
**Audience:** people who got the free MarginKite Etsy Fee Calculator (Lite) on Gumroad and have not bought the tracker.
**Goal:** be useful first; introduce the $12 Seller Profit Tracker in email 4.
**Source of truth for every number and feature:** `products/seller-profit-tracker/build.py` (fee rows `ETSY_*`, fees as of 2026-10-21) and `build_lite()`.
**Unsubscribe:** per the ticket, Gumroad adds the unsubscribe link to workflow emails automatically, so none is written into the copy. Gumroad's Workflows article doesn't mention it (see below), so check for it in the preview email before publishing.

Rates used (build.py, the current published US Etsy rates as of 2026-10-21): listing fee $0.20; transaction fee 6.5% of item price + shipping; payment processing 3% + $0.25 on the order total; Offsite Ads 15% (shops under $10k in the last 12 months), capped at $100 per order.

---

## Email 1: Welcome (send immediately after the trigger)

**Subject:** Your free Etsy fee calculator (how to use it) *(45 characters)*
**Preview text:** Three steps to see every Etsy fee, your profit and your margin on one item. *(75 characters)*

**Body:**

> Hi,
>
> Thanks for getting the MarginKite Etsy Fee Calculator. Here's how to use it in 3 steps:
>
> 1. Make your own Google Sheets copy (or open the .xlsx in Excel 2016 or newer):
> https://docs.google.com/spreadsheets/d/1f_ZIgOi_mZsbkiiI2cuk_XiTdiTpllCazWIeVIYob00/copy
> 2. Fill the yellow cells: item price, shipping you charge, your shipping cost, your cost to make it, and two Y/N questions about Offsite Ads.
> 3. Read the blue cells: total Etsy fees, net profit and margin.
>
> The fees use the current published US Etsy rates as of 2026-10-21. If Etsy changes a rate, edit it in the fee table at the bottom of the sheet.
>
> Over the next ten days I'll send three short emails with Etsy fee tips. Reply any time with questions.
>
> MarginKite
> (A small shop. AI helps us build our templates; we check every formula by hand.)

*(134 words)*

**Notes (not in the email):**
- The copy link is the lite master from notes/log.md (2026-10-23). It is the only link.
- The yellow inputs are `build_lite()` C5–C10: item price, shipping charged, your shipping cost, your cost to make it, "Sold via Etsy Offsite Ad? (Y/N)", "Etsy sales in last 12 months $10,000+? (Y/N)".
- The blue outputs are rows 18, 20 and 21: Total Etsy fees, NET PROFIT and Margin. The fee table is rows 27–34.
- This is the AI disclosure for the whole sequence; it is included once, as the ticket asks.
- No numbers, so no arithmetic.

---

## Email 2: The tip (day 3)

**Subject:** Why Etsy fees hit small orders harder *(37 characters)*
**Preview text:** An $8 sale loses about 15% to fees. A $40 sale loses about 11%. Here's why. *(75 characters)*

**Body:**

> Hi,
>
> One Etsy fee tip that surprises a lot of sellers: two of the fees are fixed, so they cost cheap items more.
>
> Using the current published US Etsy rates as of 2026-10-21:
>
> - $8 item, no shipping charged: fees are $1.21, about 15% of the sale.
> - $40 item, no shipping charged: fees are $4.25, about 11%.
>
> The $0.20 listing fee and the $0.25 processing fee are the same on both orders. On the $8 order they are over a third of the fees.
>
> Also: Etsy's transaction and processing fees apply to the shipping you charge too. Charge $5 for shipping and about $0.48 of it goes to fees, so set shipping a little above your label cost.
>
> Try your own prices in your calculator copy.
>
> MarginKite
>
> P.S. Next week: the full tracker, which does this for every order you log.

*(142 words)*

**Notes (not in the email):**
- $8 item, $0 shipping: listing 0.20 + transaction 0.065 × 8 = 0.52 + processing 0.03 × 8 + 0.25 = 0.49 → **$1.21**. 1.21 / 8 = 15.1% → "about 15%".
- $40 item, $0 shipping: 0.20 + 0.065 × 40 = 2.60 + (0.03 × 40 + 0.25 = 1.45) → **$4.25**. 4.25 / 40 = 10.6% → "about 11%".
- Fixed part: 0.20 + 0.25 = 0.45. 0.45 / 1.21 = 37% → "over a third of the fees" on the $8 order (on the $40 order: 0.45 / 4.25 = 11%).
- Shipping: (0.065 + 0.03) × 5 = 0.475 → **about $0.48**. The basis is build.py `ETSY_TXN` ("of item price + shipping") and `ETSY_PROC_PCT` ("of order total"); lite row 13 = price + shipping.
- No link in the body (the P.S. has none), so the email stays a tip.

---

## Email 3: The costs sellers forget (day 6)

**Subject:** The Etsy costs a per-item calculator can't see *(46 characters)*
**Preview text:** Offsite Ads, refunds, Etsy Ads and monthly overhead, and why a monthly view helps. *(82 characters)*

**Body:**

> Hi,
>
> A per-item calculator shows one sale. Some costs only show up across a month:
>
> - **Offsite Ads.** On a $32 item with $6.95 shipping, an Offsite Ad sale adds $5.84 in fees. With the calculator's example costs, profit drops from $13.40 to $7.56.
> - **Refunds.** Etsy credits back part of its fees, not all of them. The listing fee doesn't come back.
> - **Etsy Ads and overhead.** Ad spend, tools and supplies are monthly bills. $63 a month over 10 orders is about $6.34 an order.
>
> Rates: current published US Etsy rates as of 2026-10-21.
>
> That's why we built the MarginKite Seller Profit Tracker: an order log, refunds and monthly overhead in one dashboard, plus Etsy vs Shopify. I'll show you what's inside in a few days.
>
> MarginKite

*(129 words)*

**Notes (not in the email):**
- Example = the lite sample inputs (`build_lite()` C5–C8): item 32.00, shipping charged 6.95, shipping cost 7.20, cost to make 14.20. Order total 38.95.
- Without Offsite Ads: listing 0.20 + transaction 0.065 × 38.95 = 2.5318 + processing 0.03 × 38.95 + 0.25 = 1.4185 → fees 4.1503. Profit 38.95 − 4.1503 − (7.20 + 14.20) = **13.40**. This matches lite preview lite-01.
- With Offsite Ads (15%, under $10k): 0.15 × 38.95 = **5.8425 → $5.84** (under the $100 cap). Fees 9.9928. Profit 38.95 − 9.9928 − 21.40 = **7.5573 → $7.56**.
- Refunds: build.py `RULE_REFUND_LISTING` = "Never" (listing fee not credited); the transaction and % processing fees are credited in proportion when the seller marks "Etsy fees credited back? = Y". The email states only the listing-fee fact and "part, not all", which both the research and the rule rows support.
- Overhead: the August sample row in `SAMPLE_EXPENSES` = 12.99 + 27.00 + 8.40 + 15.00 + 0 = **63.39** ("$63 a month"). 63.39 / 10 orders = **6.339 → about $6.34**. It is an illustration, not a typical figure.
- Tracker features named: Orders Log, refunds (cols J–L, R, T), Expenses tab, Dashboard, Price Calculator Etsy vs Shopify. All are in build.py.
- No link: the soft intro only.

---

## Email 4: The tracker (day 10)

**Subject:** The full MarginKite tracker, $12 one-time *(41 characters)*
**Preview text:** What it adds to your free calculator: every order, Etsy vs Shopify, a monthly dashboard. *(88 characters)*

**Body:**

> Hi,
>
> If the free calculator has been useful, here's what the full MarginKite Seller Profit Tracker adds:
>
> - **Orders Log** for 500 orders. Fees, refunds and profit fill in per order.
> - **Etsy vs Shopify** side by side for any product, plus how many orders a month a Shopify plan needs to pay for itself.
> - **Price for your target margin**, for Etsy and Shopify.
> - **Products and Expenses**: unit cost from materials, packaging and labour, plus monthly overhead.
> - **Dashboard**: 12 months of profit, your top and bottom products, 2 charts.
>
> $12, one-time. Google Sheets and Excel 2016+, with a sample-data version and a BLANK one. US sellers; fees dated 2026-10-21; estimates, so check your own statements.
>
> https://marginkite.gumroad.com/l/profit-tracker
>
> No rush and no sale: the price is the same whenever you look. This is the last email in this series.
>
> MarginKite

*(140 words)*

**Notes (not in the email):**
- Orders Log 500 rows: `O_FIRST, O_LAST = 6, 505`. Per-order fees, refunds and profit: cols N–V.
- Etsy vs Shopify and break-even: Price Calculator rows 13–27 and 33–40. Target-margin price: rows 29–31.
- Products: cols B–F, hourly rate on Settings C5. Expenses tab. Dashboard: rows 7–18, top/bottom 3 rows 22–26, two bar charts.
- Sample-data + BLANK versions: `TRACKER_NAME`, `TRACKER_BLANK_NAME`. Google Sheets copy links are in the tracker's PDF guide.
- Price $12, one-time: the live Gumroad product (notes/log.md). There is no discount; the owner hasn't approved one.
- "This is the last email" is accurate: the workflow has 4 emails.

---

## Gumroad Workflow setup (for the owner; don't set up until approved)

Facts below are from Gumroad's help article **"Send automated emails with Workflows"**, read directly on 2026-10-23: https://gumroad.com/help/article/131-using-workflows-to-send-automated-updates (**V** unless marked).

| Setting | Value | Gumroad rule |
|---|---|---|
| **Blocker: can we send yet?** | **Not yet.** "You cannot send emails until you've earned $100 after fees and received a payout." At about $9.65 net per tracker sale that is about 11 tracker sales, plus one payout. | V |
| Trigger | **Purchase** | Triggers: Purchase, New subscriber, Member cancels, New affiliate, Abandoned cart (V). That a $0 "purchase" fires the Purchase trigger is **U**; it's very likely, since free downloads appear as $0 sales, but confirm with a test. |
| Filter: Has bought | **MarginKite Etsy Fee Calculator (Lite)** | V |
| Filter: Has not yet bought | **MarginKite Seller Profit Tracker** | V: only **one** product can go in this filter, and "if a customer purchases this product mid-way through the workflow, they will not receive any subsequent emails in the series". This is exactly the exclusion we want. |
| Past downloaders | Owner decision: Gumroad offers "send emails to previously eligible audience members". Recommended: yes, for lite downloads since launch (2026-10-23), since the series is useful on its own. | V (option exists) |
| Email 1 delay | 0 hours after the trigger (or 1 hour, if 0 isn't allowed) | Delays are "hours, days, weeks, or months after the trigger" (V) |
| Email 2 delay | 3 days | Measured from the trigger, not from the previous email (V: "after the trigger") |
| Email 3 delay | 6 days | same |
| Email 4 delay | 10 days | same |
| Sender | Profile name "MarginKite", from an @gumroad.com address; replies go to the support email (marginkite.shop@gmail.com), or the product's own support email if set | V |
| Locked after publishing | "Besides the workflow's name, you cannot update any of its settings once a workflow has been published for the first time." Emails stay editable. **So set the trigger and filters right the first time.** | V |
| Testing | Test purchases don't fire workflows. Get the free lite product with a second email address instead (Gumroad suggests a 100% discount code for paid products). | V |
| Unsubscribe link | Assumed to be added by Gumroad (per the ticket). **Not mentioned in this article (U)**: confirm it appears in the preview email. | U |

**Until the $100 + payout gate is met:** keep this file as the draft. The lite receipt and download page already cross-sell the tracker ("More from this creator", noted in the log) and lite cell B24 links to it, so nothing is lost in the meantime.
