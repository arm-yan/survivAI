# Creative Market shop application: ready-to-paste answers

**Status:** DRAFT. The **owner** fills and submits the form while logged in. We have not signed up, logged in or applied.

## What the application asks for (researched 2026-10-23)

Creative Market's help center returned HTTP 403 to direct fetches today. Points marked **V** were read directly in our earlier research (`research/2026-10-23-marketplaces-armenia.md`); **S** = taken from a search-engine extract of the official page; **U** = unverified or third-party.

| Requirement | Detail | Quality |
|---|---|---|
| Online portfolio link | Must be a public, linkable web page. Attachments, zip files, Dropbox or Google Drive links aren't reviewed. No link shorteners. | S ([Open a Shop](https://support.creativemarket.com/hc/en-us/articles/201251700-Open-a-Shop-on-Creative-Market)); V for "portfolio link" |
| Your account email visible on the linked page | Used to prove ownership. The email on the Creative Market account must appear on the portfolio page. | V + S |
| Amount of work | Creative Market "likes to see at least 10–20 examples of your best work, ideally pertaining to what you plan to sell". Third-party guides say 3–5 products can be enough. | S (official); U (third-party) |
| Human review, possible waitlist | Applications are reviewed by staff and may be waitlisted. Third-party guides say 1–4 weeks. | V (review and waitlist); U (timing) |
| After approval | Tipalti payout set-up (wire to Armenia, fee waived), W-8BEN tax form and ID | V |

Exact form field names weren't visible to us (U), so the answers below cover the questions such forms usually ask. Use whichever ones the form shows.

---

## Before submitting (owner checklist)

1. **Create the Creative Market account with marginkite.shop@gmail.com**, so the email matches what the portfolio shows.
2. **Show that email on the portfolio page.** Recommended: add a line to the Gumroad profile bio at https://marginkite.gumroad.com, such as `Questions: marginkite.shop@gmail.com`. Check that it's visible when logged out.
3. **Decide whether two products are enough.** Our Gumroad shop has 2 products (the tracker and the free lite calculator) and 9 preview images. Creative Market asks for 10–20 examples. Options:
   - (a) apply with the Gumroad shop as is and accept a higher rejection or waitlist risk; or
   - (b) also publish a free portfolio page (Behance or Dribbble) showing the 9 preview images as separate pieces, with the email visible, and give that as the main link. **Recommended: (b), with Gumroad as the second link.** It needs the owner's sign-up, which is an owner action.
4. Creative Market's rules forbid listings that send buyers elsewhere to buy, but linking to where you already sell is normal in an **application**. Gumroad is fine here; it must never appear in the listing.

---

## Answers (paste as is)

**Shop name**
```
MarginKite
```

**Portfolio / website URL**
```
https://marginkite.gumroad.com
```
If option (b) is used: [[PLACEHOLDER: Behance or Dribbble portfolio URL]] first, Gumroad second.

**Where else do you sell?**
```
Gumroad: https://marginkite.gumroad.com (live since October 2026).
```

**What do you plan to sell? / Category**
```
Templates: spreadsheet templates for small online sellers, built for Google Sheets and Microsoft Excel.
```

**Tell us about yourself and your work (bio)**
```
MarginKite is a small shop making spreadsheet tools for Etsy and Shopify sellers. Our first product, the Seller Profit Tracker, shows what a seller really keeps from each sale after every fee, product cost and shipping, with Etsy and Shopify side by side, a price-for-target-margin calculator, a 500-order log with refunds and a monthly dashboard. It works in Google Sheets and Excel 2016+, with no macros.

We're open about how we work: AI helps us design and build the templates, and we check every formula against hand-worked test cases before release. Every fee rate sits in a dated table with its source, so buyers can see exactly what the numbers are based on and update them when fees change.

On Creative Market we'd start with the Seller Profit Tracker and add related seller templates over time.
```
(about 140 words)

**Sample work / what makes your products stand out (if asked)**
```
- Etsy and Shopify fees compared side by side on the seller's own numbers
- Reverse pricing: the item price needed for a target margin, for both platforms
- "When does Shopify pay off?": the monthly order count at which a Shopify plan pays for itself
- Dated fee table with sources; no macros; no password locks; a sample-data copy and a BLANK copy
```

**Do you use AI in your products? (if asked)**
```
Yes. AI helps design and build our templates, and we review and test every formula by hand before release. We'll answer the AI question honestly on each product upload.
```

**Email**
```
marginkite.shop@gmail.com
```

---

## Notes

- Everything above is true today: the 2 Gumroad products are live, the features exist in `build.py`, and the 25 test cases are in `tests/cases.md`. No sales counts, reviews or follower numbers are claimed.
- "Related seller templates over time" states an intention, not a product. The owner can delete it if they'd rather not signal plans.
