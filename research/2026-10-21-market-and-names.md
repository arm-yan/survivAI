# Market research, pricing and brand names — Ticket #3

**Author:** Brand & Market Researcher · **Date checked:** 2026-10-21 (all checks below were made on this date unless noted)
**Product:** "Seller Profit Tracker" (working name), Google Sheets + Excel, for Etsy and Shopify sellers.

## Verdict (summary)

The market is crowded and cheap, but thin where we plan to compete. Etsy holds dozens of Etsy-only profit and bookkeeping sheets, and proven sellers sit at an effective price of about **$8–13** once their nearly permanent sales are applied. **We found no paid template that puts Shopify plan and payment fees beside Etsy fees, and none of the popular free calculators has a "price for target margin" mode.** Those two features, plus a dated fee table that is easy to update, are the angle. Gumroad shows almost no organic demand: every one of the 9 seller-tool products we checked there has **0 ratings**. That supports keeping Etsy as the main channel.

**Evidence limits, read first:**
- **Etsy blocks automated access.** `etsy.com` returned HTTP 403 for the official fees page, the help center, listing pages and shop URLs, including a known existing shop. So Etsy listing prices and reviews come from **search-engine snippets** and are marked *unverified*. All Etsy shop-name checks are also *unverified*.
- **Reddit blocked our fetch tool**, and searches didn't surface r/EtsySellers threads. The buyer-complaint evidence is therefore thinner than the persona standard asks for. It comes from listing review snippets, listing FAQs and the Etsy Community forum.
- We did not get around any block, per the ticket.

Legend: **V** = verified directly from the source page today. **U** = unverified (search snippet, third-party source, or check blocked).

---

## Part A — Competitors

### Competitor table

| # | Product | Platform | Price | Format | Reviews | Rating | What it does / gap we can exploit | Link | V/U |
|---|---|---|---|---|---|---|---|---|---|
| 1 | Etsy Seller Bookkeeping Spreadsheet POD (PrioriDigitalStudio) | Etsy | $13.19 on sale (from $26.39) | Sheets + Excel | 218 (structured data). Page also shows "11.2k", probably shop-wide | 4.9 | Category leader. Bookkeeping plus Printify/POD costs. **Etsy-only. No Shopify comparison and no reverse pricing.** | https://www.etsy.com/listing/1657487524 | U |
| 2 | "Etsy Seller Spreadsheet – Updated 2025 Fees" (shop not captured) | Etsy | $8.45 on sale (from $16.91) | Sheets/Excel | 3.1k ratings (likely shop-level) | n/a | High-volume seller still labelled **"2025 fees"** in Oct 2026, so fees go stale | https://www.etsy.com/market/etsy_seller_bookkeeping | U |
| 3 | Etsy Seller Bookkeeping Spreadsheet Fee & Profit Calculator, US Version | Etsy | $2.75 (one snippet said $12.50, conflicting) | Excel + Sheets | 3 visible, all 5★. Snippet cache about 4.5 years old | n/a | Old, cheap, US-only. Shows the bottom of the price band | https://www.etsy.com/listing/993076115 | U |
| 4 | Etsy Seller Profit Tracker Spreadsheet, Fee Calculator | Etsy | $9.00 on sale (from $12.00) | Excel + Sheets, blank + example | not captured | n/a | Close to our spec: 2026 US fee defaults, monthly P&L, product ranking, pricing calculator. **No Shopify, no reverse mode. Seller states it doesn't track sales tax/VAT.** | https://www.etsy.com/listing/4597094291 | U |
| 5 | Etsy Profit Calculator Spreadsheet | Etsy | not captured | Sheets/Excel | not captured | n/a | Offsite Ads on/off toggle (15% or 12%) with a cap field. That is table stakes for us | https://www.etsy.com/listing/4434905011 | U |
| 6 | Etsy Ads profit / ROAS tracker | Etsy | not captured | Sheets/Excel | not captured | n/a | 18 KPIs, break-even ROAS per listing. Ads-focused, not a full P&L | https://www.etsy.com/listing/4582032900 | U |
| 7 | Etsy Bookkeeping Spreadsheet / Profit Tracker (Growtoria) | Own site | $16 ($19 for a 3-sheet kit) | Sheets + Excel | 0 | — | Dashboard, orders, editable fee tab, per-product profit, tax set-aside. **Etsy-only. No Shopify, no reverse mode** | https://growtoria.com/product/etsy-bookkeeping-spreadsheet/ | V |
| 8 | Etsy Seller Automatic Bookkeeping (SpreadsheetsHub; also on MonsterOne and TemplateMonster) | Own site | $22.99 on sale (from $55.99) | Sheets + Excel | 0 | — | Etsy CSV import, charts, forecasts. **Needs Microsoft 365, since Excel 2019 lacks its formulas. Some cells and tabs locked.** | https://spreadsheetshub.com/products/etsy-seller-automatic-bookkeeping | V |
| 9 | Etsy Shop Profit Tracker (cavaso) | Gumroad | $15.00 | .xlsx + Sheets link | 0 | — | Listing, transaction and processing fees only. 30 rows | https://cavaso.gumroad.com/l/Etsy-Shop-Profit-Tracker | V (price, ratings) |
| 10 | Etsy Seller Toolkit – Excel (Ardent Workshop) | Gumroad | $22.95 | Excel | 0 | — | Excel-only | https://ardentworkshop.gumroad.com/l/kxwlhh | V |
| 11 | ETSY Sellers Hub (Serenatas Store) | Gumroad | $22.99 | n/a | 0 | — | — | https://serenatasstore.gumroad.com/l/etsysellershub | V |
| 12 | Profit Calculator for Etsy Sellers (AMEtech) | Gumroad | $7.99 | n/a | 0 (seller has 4.5★ from 14 ratings across 5 products) | — | — | https://ametech.gumroad.com/l/profit-calculator | V |
| 13 | Digital Product Sales Tracker (supunsandeep) | Gumroad | $19.00 | Excel + Sheets | 0 | — | Multi-platform (Gumroad, Etsy) for digital sellers | https://supunsandeep.gumroad.com/l/zxjrf | V |
| 14 | E-commerce Profit & Pricing Calculator for Shopify, Etsy & eBay (Planoryx) | Gumroad | **Free ($0)** | Sheets + Excel | 0 | — | **The closest feature overlap we found:** multi-platform, break-even price. Free, so it competes with our lite version | https://planoryxstudio.gumroad.com/l/profit-pricing-tool | V |
| 15 | Google Sheets Sales P&L Tracker (VIP Graphics) | Gumroad | $39.00 | Sheets | 0 | — | Generic P&L, quarterly and annual dashboards. Top of the price band | https://vipgraphics.gumroad.com/l/sales-tracker | V |
| 16 | Etsy Sales & Profit Dashboard (SB Data Services) | Gumroad | Pay what you want, from $0 | Excel | 0 | — | Built on Etsy CSV export | https://sbdataservices.gumroad.com/l/Etsy_Seller_Dashboard | V |
| 17 | Etsy Fees Cheat Sheet 2026 | Payhip | $6.99 on sale | PDF + mini sheet | n/a | n/a | Fee explainer + mini calculator | https://payhip.com/b/8QWMh | U |
| 18 | ProfitBird (LeakShield Technologies) | Shopify App Store | Free / $9 / $19 per month | Shopify app | 0 | 0.0 | **Phase 2 competitor:** net profit per Shopify order. Also a **name collision risk** (see Part D) | https://apps.shopify.com/profitbird | V |

**Free web calculators (the lead-magnet competition):**

| # | Tool | Price | Signup | Covers | Lacks | Link | V/U |
|---|---|---|---|---|---|---|---|
| 19 | Craftybase Etsy Fee Calculator | Free (spreadsheet version needs an email) | No (web) | Listing, transaction, country processing (57 countries), Regulatory Operating Fee. Updated March 2026 | **Excludes Offsite Ads. No reverse or target-margin mode. One order at a time.** Shopify is on a separate page. Sample figures don't reconcile ($0.98 shown vs $1.00 expected) | https://craftybase.com/etsy-fee-calculator | V |
| 20 | Listadum Etsy Fee Calculator | Free | No | Listing, transaction, processing, Etsy Ads (as TACOS %), Offsite Ads | No Regulatory Operating Fee, **no Shopify, no reverse mode** | https://www.listadum.com/etsy-fee-calculator | V |
| 21 | Zik Analytics Etsy Fee Calculator | Free | n/a | Basic fees | Excludes Offsite Ads and Regulatory Operating Fee | https://www.zikanalytics.com/free-tools/etsy-fee-calculator | U |
| 22 | Vendoo Etsy Fee Calculator | Free | n/a | Includes Offsite Ads | n/a | https://www.vendoo.co/etsy-fee-calculator | U |
| 23 | Printful Etsy Calculator | Free | n/a | Aimed at print-on-demand | POD-specific | https://www.printful.com/etsy-calculator | U |
| 24 | Aeolidia Etsy vs Shopify Fee Calculator | Free | n/a | Etsy vs Shopify plans by order value and volume. US only | Comparison only, no tracking | https://aeolidia.com/resource/etsy-vs-shopify-fee-calculator/ | U |

### What buyers miss (evidence)

All items below are **U**: they are snippets or forum posts, not a representative sample.
- **Refunds and fee reimbursement are confusing.** Etsy Community threads report transaction fees not coming back on refunded shipping and on Offsite Ads orders. Posters also say the processing fee is credited at once but the transaction fee later. Sources: https://community.etsy.com/t5/Technical-Issues/Etsy-not-refunding-fees-on-refunded-shipping-or-canceled-orders/m-p/146788931/highlight/true and https://community.etsy.com/t5/Technical-Issues/Potential-bug-with-refund-processing-system/m-p/147724289. A tracker that logs refunds with a "fees reimbursed? Y/partial/N" column addresses a real pain.
- **Period structure is too rigid.** A review of one bookkeeping sheet complains its tabs hold only one month and one week at a time (https://www.etsy.com/listing/263864723).
- **Access and quality problems.** A listing FAQ addresses the sheet opening "view only" (https://www.etsy.com/listing/1588204479). A review on the same listing says "I spent more on another shop's sheet and [it] doesn't even work this well."
- **Compatibility and lock-in.** SpreadsheetsHub needs Microsoft 365 and locks cells (row 8, **V**). Buyers on older Excel can't use such files.
- **Stale fees.** A high-volume seller is still labelled "Updated 2025 Fees" (row 2). Blogs also disagree on Shopify Basic's price ($29 vs $39), while Shopify's own page today says $27 monthly / $19 yearly (Part B). Fee data from blogs goes out of date fast.
- **Offsite Ads gets missed.** Several free tools leave it out (Craftybase, Zik). One guide notes that a single missed Offsite Ad can cut a sale's profit by more than 20% (https://www.webgility.com/blog/etsy-fees).

### Table stakes vs gap

- **Table stakes:** Sheets + Excel files, a blank and an example copy, an editable fee tab, per-order profit, a monthly dashboard, per-product profit, an Offsite Ads toggle, and an instructions tab. Many leaders also offer Etsy CSV import.
- **Gap (nobody we found does all of these):** (1) Etsy **and** Shopify fees side by side on the seller's own numbers. (2) Reverse "price for target margin" mode in a paid tracker. (3) A **dated** fee table with "last checked" sources. (4) Refund handling that models fee reimbursement. (5) No Microsoft 365-only formulas and no locked cells.

---

## Part B — Current fees (for the Settings tab)

### Etsy (US sellers)

**Status: all U.** Etsy's official fee policy (https://www.etsy.com/legal/fees/) and the help center returned HTTP 403 to our fetch on 2026-10-21. The figures below agree across several 2026 third-party guides. **The owner should confirm them on Etsy's page in a browser before the Settings table is final.**

| Fee | Value | Source (third-party) | V/U |
|---|---|---|---|
| Listing fee | $0.20 per listing; renews every 4 months or when the item sells (charged again for each extra unit of a multi-quantity sale) | https://checkoutpage.com/blog/etsy-fees, https://www.listadum.com/etsy-fee-calculator | U |
| Transaction fee | 6.5% of item price **plus shipping and gift wrap** | https://checkoutpage.com/blog/etsy-fees, https://craftybase.com/etsy-fee-calculator | U |
| Payment processing (Etsy Payments, US) | 3% + $0.25 per order, on the order total | https://craftybase.com/etsy-fee-calculator, https://checkoutpage.com/blog/etsy-fees | U |
| Offsite Ads | 15% if the shop made under $10k in the past 12 months (opt-out allowed). 12% at $10k or more (mandatory). Capped at **$100 per order**. Attributed sales only | https://checkoutpage.com/blog/etsy-fees, https://blog.marmalead.com/etsy-vs-shopify/ | U |
| Regulatory Operating Fee | **None for US sellers.** Applies in some other countries (e.g. UK 0.32%) | https://checkoutpage.com/blog/etsy-fees, https://craftybase.com/etsy-fee-calculator | U |
| Etsy Ads (onsite) | Seller-set daily budget, pay per click. Variable, so it goes in as a user input | https://www.listadum.com/etsy-fee-calculator | U |
| Share & Save | Credit to the seller on sales from shared links. Rate not confirmed | snippet referenced in a competitor listing | U |
| Currency conversion | Not verified this session. From prior knowledge: 2.5% when the listing currency differs from the payment account currency | — | U (prior knowledge) |
| Etsy Plus subscription | Not verified this session. From prior knowledge: $10/month, optional | — | U (prior knowledge) |
| Refund behaviour | Transaction and processing fees are usually credited back on refunds within 180 days through Etsy Payments. Listing fee is not refunded. Sellers report delays and exceptions | https://ecomcalctools.com/blog/fees-etsy/etsy-refund-returns/, Etsy Community threads above | U |

### Shopify (US)

| Item | Value | Source | V/U |
|---|---|---|---|
| Basic plan | $27/mo billed monthly; $19/mo billed yearly | https://www.shopify.com/pricing (fetched 2026-10-21) | **V** |
| Grow plan | $72/mo monthly; $54/mo yearly | same | **V** |
| Advanced plan | $399/mo monthly; $299/mo yearly | same | **V** |
| Plus | From $2,300/mo | same | **V** |
| Third-party payment provider fee (not using Shopify Payments) | Basic 2%, Grow 1%, Advanced 0.6%, Plus 0.2% | same | **V** |
| Starter plan | The page's code references a "Starter" plan card. Its price was not shown in the fetched content | same | U |
| Shopify Payments online card rate | Basic 2.9% + 30¢; Grow 2.7% + 30¢; Advanced 2.5% + 30¢; Plus about 2.25% + 30¢ | Search-engine index of https://www.shopify.com/pricing, plus https://qualimero.com/en/blog/shopify-transaction-fees. **The official page loads rates client-side by region (`{{cardRate}}` placeholders), so we couldn't read them directly.** Some sources list Grow 2.6% and Advanced 2.4% | U |
| Premium/Amex online cards | 3.5% / 3.3% / 3.1% / 2.95% + 30¢ (Basic/Grow/Advanced/Plus) | search index of shopify.com/pricing | U |
| International cards | +1% surcharge online | search index of shopify.com/pricing | U |
| In-person (POS) | Basic about 2.6% + 10¢ | https://ecommerce-platforms.com/articles/how-much-does-shopify-take-per-sale | U |
| Exact rate for a given store | Shown in Settings > Payments > Shopify Payments > Manage | https://help.shopify.com/en/manual/payments/shopify-payments/payouts/pay-periods-and-fees | U (per search summary) |

**Recommendation for the template:** every fee is an editable cell with "Value / Source URL / Date checked" columns. Card rates default to the U figures above, labelled "check your Shopify admin".

---

## Part C — Positioning, price and spec (RECOMMENDATIONS)

Everything in this section is opinion based on the evidence above.

### Positioning (recommendation)

> **For Etsy sellers who don't know what they really keep after fees — or whether Shopify would leave them more — Seller Profit Tracker is the only template that runs Etsy and Shopify fees side by side on your own orders and tells you the price to charge to hit your target margin.**

### Price (recommendation): **$12 at launch, then $15 after about 15 reviews**

- **Evidence band:** Proven Etsy sellers' effective (sale) prices are $8.45–$13.19 (rows 1, 2, 4). The cheap end goes down to $2.75 (row 3). Gumroad and own-site sellers list at $15–$39 but show 0 ratings, so those prices aren't proven to convert.
- **Why $12 at launch:** With zero reviews we compete with the $9–13 listings on trust. $12 is inside the draft range and close to the leader's sale price.
- **Why $15 later:** The Shopify comparison and reverse pricing are features none of the cheap listings has. $15 sits under the $16–23 own-site sellers and is still inside the $12–19 range.
- **Avoid the "inflated original price, permanent 50% off" pattern** many listings use. It's common, but it's a credibility and policy risk. Use Etsy's sale tool only for real, time-boxed promotions. **Owner decision; see open questions.**
- **Bundle idea (v1.1):** "Seller Profit Kit" = tracker + pricing calculator for handmade materials and labour + a tax set-aside sheet, at $19–24 (Growtoria sells a similar kit for $19).

### Spec changes (recommendations)

**Add to v1.0:**
1. **"When does Shopify pay off?" box on the Dashboard.** At your actual monthly orders and average order value, show the volume at which Shopify Basic or Grow beats Etsy fees. Of the sources we found, only the free Aeolidia web tool does this, and not on the seller's own data. This makes the side-by-side feature tangible.
2. **Offsite Ads handling:** a per-order flag, a 15%/12% switch driven by "last-12-months sales ≥ $10k", and the $100 cap in the formula.
3. **Etsy Ads (onsite) spend** as a monthly expense line, with optional allocation per listing.
4. **Refunds column with fee-reimbursement status** (full / partial / none), so profit reflects what Etsy actually gave back.
5. **Shopify monthly vs yearly billing switch** and a "third-party gateway" switch (2% / 1% / 0.6% extra fee).
6. **Dated Settings table** with Value, Source URL and Date checked on every row, plus a visible "Fees last checked: 2026-10" badge on Start Here. This is also a listing-copy point.
7. **Compatibility rule:** no Microsoft 365-only functions (e.g. avoid LAMBDA and dynamic-array-only formulas in the .xlsx), no password locks. Protect formula cells with warnings only.
8. **Ship a blank copy and an example copy** (table stakes).
9. **Sales tax note:** Etsy collects US marketplace sales tax, so it is excluded from revenue by default, with an explanatory note.

**Move to v1.1:**
- **Etsy CSV import tab.** Several competitors have one (rows 8, 16, and MonsterOne). It's a real value-add but the most build and support risk. v1.0 ships with manual entry plus a "paste rows here" area.
- **Non-US fee presets** (UK, CA, AU). Competitors claim 13–57 countries, but the US-first target market doesn't need them at launch.
- Tax set-aside estimate and the bundle.

**Cut / don't build:** multi-marketplace support beyond Etsy and Shopify (eBay, Amazon). It dilutes the positioning and adds fee tables to maintain.

### Lite calculator: spreadsheet or web page? (recommendation: **one-tab Google Sheet now, web page in Phase 2**)

- **Evidence:** The free web calculators that rank today are run by funded companies (Craftybase, Printful, Zik, Vendoo, Listadum), and they need no signup. Competing with them on SEO from day 1 is unrealistic. Craftybase itself gates its **spreadsheet** behind an email, which is the model we want.
- **Why a sheet:** zero hosting, no build time, and it shows off the paid product's look. It also delivers through Gumroad as a $0 product, which collects the email automatically.
- **Why not on Etsy:** from prior knowledge, Etsy has a minimum price of $0.20 and doesn't allow free listings (**U**). The lite version should therefore live on Gumroad and be linked from Pinterest and Reddit answers.
- **Make the lite sheet different from free tools:** include Offsite Ads with the cap (Craftybase and Zik lack it) and a one-cell "price for target margin" (Craftybase and Listadum lack it). Lock Shopify comparison and tracking behind the paid version.
- **Phase 2:** rebuild the lite version as a small web page. It doubles as the landing page for the browser extension.

---

## Part D — Brand names

**Method:**
- **`.com`:** Verisign RDAP (`https://rdap.verisign.com/com/v1/domain/<name>.com`). We started with rdap.org, but it rate-limited us with HTTP 429 after a few queries, so we switched to the registry directly. Controls: google.com → 200, a nonsense domain → 404.
- **Fallback TLDs:** `.app` (Google Registry RDAP) and `.shop` (GMO Registry RDAP), with the same controls. **`.co` has no RDAP server in IANA's bootstrap file, so `.co` results from rdap.org are meaningless and were discarded.**
- **Gumroad:** `https://<name>.gumroad.com`. Controls: existing seller → 200, nonsense → 404.
- **Etsy:** `https://www.etsy.com/shop/<Name>` returned 403 even for a known existing shop, so all Etsy results are **U**.
- **Instagram/Pinterest:** both return HTTP 200 for nonexistent handles, so we compared logged-out page titles instead. A taken handle shows a profile title, a free one shows a generic title. This is a **heuristic, so marked U.**
- **Trademark:** web search only. **No USPTO search was done**, because tmsearch.uspto.gov needs an interactive browser.

**Screening:** 56 unique candidate `.com`s were checked, and 38 were already registered. Taken examples: feelens, keepmargin, marginmint, netpenny, tallyhouse, pricewell, shopsum, sellsum, margincraft, tallywise, netkite, profitkite, tidymargin, sellerledger and brightmargin. These five survived.

### Name shortlist (ranked)

| Rank | Name | Why it works | Etsy shop name | .com | Fallback TLDs | Gumroad | Handles (IG / Pinterest) | Trademark / brand risk | Verified? |
|---|---|---|---|---|---|---|---|---|---|
| **1** | **MarginKite** | "Margin" says profit and fits any seller tool (sheet, extension, app). "Kite" suggests lift and lightness. Two common words, easy to spell, 10 characters (within Etsy's limit) | Unverified: Etsy 403 (https://www.etsy.com/shop/MarginKite) | **Unregistered** (404, https://rdap.verisign.com/com/v1/domain/marginkite.com) | .app 404, .shop 404 (unregistered) | **Free** (404, https://marginkite.gumroad.com) | Both appear free (generic title) | **Low–medium.** No "MarginKite" brand found. "Kite" is Zerodha's trading platform (India), a different market but finance-adjacent (search: "MarginKite" OR "Margin Kite") | .com/.app/.shop/Gumroad **V**; Etsy, handles, TM **U** |
| **2** | **KeepLedge** | "What you keep" + a nod to ledger. Unique, no collisions. Weaker meaning: may be misread as "Keep Ledger" | U (403) — https://www.etsy.com/shop/KeepLedge | **Unregistered** (404) | .app 404, .shop 404 | **Free** (404) | Both appear free | **Low.** Search found only geography/dictionary results for "ledge" | as above |
| 3 | **ProfitFinch** | Friendly, memorable, suits a bird mark (a mascot, not a fake person). "Profit" says exactly what it does | U (403) | **Unregistered** (404) | .app 404, .shop 404 | **Free** (404) | Both appear free | **Medium.** **ProfitBird** is a Shopify profit-tracking app in our exact Phase 2 category (https://apps.shopify.com/profitbird). Finch (finch.com) is an e-commerce performance-marketing firm. "Profit + bird" is too close | as above |
| 4 | **FeeFinch** | Catchy, alliterative-ish, says "fees". Narrower: ties the brand to fees, not profit or tools broadly | U (403) | **Unregistered** (404) | .app 404, .shop 404 | **Free** (404) | Both appear free | **Low–medium.** Many unrelated "Finch" fintechs (Finch Money, Finch expense app). No "FeeFinch" found | as above |
| 5 | **ProfitPine** | Calm, natural, easy to say | U (403) | **Unregistered** (404) | .app 404, .shop 404 | **Free** (404) | **Instagram @profitpine is TAKEN** (profile title "ProfitPine (@profitpine)"). Pinterest appears free | **Low** for trademark (no brand found; "Pine" also means TradingView's Pine Script), but the handle conflict hurts | as above |

Checked as a spare: **MarginFinch**. `.com`, `.app`, `.shop` and Gumroad are all unregistered, and IG and Pinterest appear free. It carries the same Finch/ProfitBird closeness as #3.

---

## Recommendation

*Everything in this section is a recommendation.*

- **Positioning:** "For Etsy sellers who don't know what they really keep after fees — or whether Shopify would leave them more — Seller Profit Tracker is the only template that runs Etsy and Shopify fees side by side on your own orders and tells you the price to charge to hit your target margin."
- **Price:** $12 launch, then $15 after about 15 reviews, on both Etsy and Gumroad. A $19–24 bundle comes in v1.1.
- **Top name pick: MarginKite.**
  - Says "profit tools", not "spreadsheets", so it carries into the Phase 2 extension ("MarginKite for Chrome").
  - `.com`, `.app` and `.shop` are unregistered, and the Gumroad username is free.
  - Search showed no brand collision in e-commerce.
- **Runner-up: KeepLedge.** It has the cleanest collision profile and the same availability, but a weaker meaning.
- **Not recommended despite appeal:** ProfitFinch, because of how close it is to ProfitBird in the same Shopify-profit category.
- **Product name** can stay descriptive under the brand, e.g. "MarginKite Seller Profit Tracker".
- **Lead magnet:** a one-tab Google Sheet delivered as a $0 Gumroad product, with Offsite Ads cap and target-margin cell. Build a web page in Phase 2.
- **Spec:** add the Shopify-break-even box, Offsite Ads logic, Etsy Ads spend, refund-reimbursement column, Shopify billing and gateway switches, dated Settings table, and no M365-only formulas. Move CSV import, non-US presets and tax set-aside to v1.1.

## Open questions for the owner

1. **Name approval:** do you approve **MarginKite** (or prefer KeepLedge)? Once you choose, these checks need you logged in, since they were blocked for us:
   - the Etsy shop name, which Etsy confirms during shop setup;
   - the Instagram and Pinterest handles.
   We have registered or reserved nothing.
2. **Domain purchase:** do you want to buy `marginkite.com` now, before announcing the name? We did not check a registrar price. Expect a typical `.com` first-year price, and confirm the amount before buying, per the spending rule.
3. **Trademark:** should we accept web-search-level clearance, or will you (or a lawyer) run a USPTO search (https://tmsearch.uspto.gov) for "MarginKite" in classes 9, 35 and 42? We couldn't run it automatically.
4. **Etsy fee confirmation:** Etsy's fee page blocked our checks. Could you open https://www.etsy.com/legal/fees/ in your browser and confirm the US figures in Part B, or paste the page text? We'll mark them verified after that.
5. **Shopify card rates:** if you have any Shopify login, the exact rates are shown in the admin. Otherwise we'll ship the template's defaults labelled "check your admin".
6. **Pricing style:** is an honest $12 launch price OK, or do you want the common "high original price + sale" pattern? We advise against it.
7. **Scope:** OK to ship v1.0 **US-only**, with CSV import and other countries in v1.1?
8. **Lead magnet on Gumroad, not Etsy:** this assumes Etsy doesn't allow free listings (unverified). Is that fine?
9. **Buyer complaints:** Reddit blocked our tools. If you'd like stronger evidence, browsing r/EtsySellers for "profit spreadsheet" yourself for 10 minutes would help. Otherwise we'll proceed on the evidence above.
