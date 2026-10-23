# Hand-calculated test cases

`run_tests.py` checks every case below against the generated workbooks. Money must match within $0.01 and margins within 0.01 percentage points.

**Fee values** (Settings, as of 2026-10-21, research Part B):
- Etsy: listing $0.20 per unit, transaction 6.5%, processing 3% + $0.25, Offsite Ads 15% (12% at $10k+), cap $100.
- Shopify Basic: $27/mo monthly or $19/mo yearly, Shopify Payments 2.9% + $0.30, third-party extra fee 2%.
- Third-party provider placeholder: 2.9% + $0.30.

**Rules** (defaults):
- Etsy fees use the price after the discount (Y), per Etsy's fees help article https://help.etsy.com/hc/en-us/articles/115014483627 (unverified). The alternative N (pre-discount base) is tested in C19.
- Average sales tax rate: 0%. It feeds only the processing-fee base (tested at 8% in C18).
- On a refund with "fees credited back" = Y, Etsy credits transaction + % processing in proportion to refund ÷ order total. The $0.25, the listing fee and Offsite Ads are not credited.
- Shopify fees are never credited.

**Test setup:**
- Hourly rate $20.
- "Test Mug": materials 4.00 + packaging 1.00 + 30 min × $20/60 = 10.00, so **unit cost $15.00** (case C00).
- "Test Print" (no orders): unit cost $2.50.

Profit = order total − refund − total fees − product cost − shipping cost. Margin = profit ÷ (order total − refund).

## Pass A — Basic, Monthly, Shopify Payments, shop under $10k

| Case | Inputs | Arithmetic | Expected |
|---|---|---|---|
| **C01 Basic Etsy sale** | qty 1, price 30, ship charged 5, ship cost 4 | total 35. Listing 0.20. Txn 0.065×35 = 2.275. Proc 0.03×35 + 0.25 = 1.30. Fees 3.775. Profit 35 − 3.775 − 15 − 4 = 12.225. Margin 12.225/35 | fees **3.775**, profit **12.225**, margin **34.93%** |
| **C02 Free shipping** | price 30, ship 0, ship cost 4 | total 30. Txn 1.95. Proc 0.90 + 0.25 = 1.15. Fees 0.20 + 1.95 + 1.15 = 3.30. Profit 30 − 3.30 − 19 | fees **3.30**, profit **7.70** |
| **C03 Offsite Ads 15%** | C01 + Offsite = Y | Offsite 0.15×35 = 5.25. Fees 3.775 + 5.25 = 9.025. Profit 35 − 9.025 − 19 | offsite **5.25**, profit **6.975** |
| **C04 $100 cap** | price 800, ship 0, ship cost 10, Offsite = Y | 0.15×800 = 120, capped to **100**. Txn 52. Proc 24 + 0.25 = 24.25. Fees 0.20 + 52 + 24.25 + 100 = 176.45. Profit 800 − 176.45 − 15 − 10 | fees **176.45**, profit **598.55** |
| **C05 Full refund** | C01, refunded Y, partial blank, credited Y | refund 35. Credit (2.275 + 0.03×35)×35/35 = 3.325. Fees 3.775 − 3.325 = 0.45. Profit 35 − 35 − 0.45 − 15 − 4 | credit **3.325**, profit **−19.45** |
| **C06 Partial refund, credited** | C01, partial 10, credited Y | credit 3.325×10/35 = 0.95. Fees 2.825. Profit 35 − 10 − 2.825 − 19 | credit **0.95**, profit **3.175** |
| **C07 Partial refund, not credited** | C01, partial 10, credited N | fees 3.775. Profit 35 − 10 − 3.775 − 19 | profit **2.225** |
| **C08 Discount** | qty 2, price 30, ship 5, ship cost 6, discount 6 | total 60 + 5 − 6 = 59. Listing 2×0.20 = 0.40. Txn 0.065×59 = 3.835 (after-discount default). Proc 0.03×59 + 0.25 = 2.02. Fees 6.255. Cost 2×15 = 30. Profit 59 − 6.255 − 30 − 6. Margin 16.745/59 | fees **6.255**, profit **16.745**, margin **28.38%** |
| **C09 Shopify order (log)** | Shopify, C01 numbers | proc 0.029×35 + 0.30 = 1.315. Profit 35 − 1.315 − 19 | profit **14.685** |
| **C10 Shopify sale, monthly billing (calculator)** | price 30, ship 5, ship cost 4, 50 orders/mo | plan/order 27/50 = 0.54. Fees 1.315 + 0.54 = 1.855. Profit 35 − 1.855 − 19. Margin 14.145/35. Etsy column = C01 | Shopify profit **14.145**, margin **40.41%**; Etsy **12.225** |
| **C12 Reverse margin 30% (monthly)** | as C10, target 30% | **Etsy:** total = (0.20 + 0.25 + 15 + 4)/(1 − 0.065 − 0.03 − 0.30) = 19.45/0.605 = 32.1488. Price = 32.1488 − 5 = 27.1488 → rounded up **27.15**. Check: at 27.15 the total is 32.15, fees 0.20 + 2.0898 + 1.2145 = 3.5043, profit 9.6458, margin 30.00%. **Shopify:** (0.30 + 0.54 + 19)/(1 − 0.029 − 0.30) = 19.84/0.671 = 29.5678, minus 5 = 24.5678 → **24.57** | Etsy **27.15**, Shopify **24.57** |
| **C13 Shopify break-even (monthly)** | as C10 | Etsy fees 3.775 − Shopify variable fees 1.315 = 2.46 saved per order. 27/2.46 = 10.98 → rounded up | **11 orders/month** |
| **C17 Dashboard roll-up** | all Pass A orders in Sep 2026 + one expense row (10 + 27 + 5 + 8 + 0 = 50) | revenue 6×35 + 30 + 800 + 59 = 1099. Σ profit 12.225 + 7.70 + 6.975 + 598.55 − 19.45 + 3.175 + 2.225 + 16.745 + 14.685 = 642.83. Net 642.83 − 50 | 9 orders, revenue **1099**, net **592.83**. Fee chart total = Fees column. Top product Test Mug (642.83), bottom Test Print |

## Pass B — Basic, **Yearly**, sales tax rate **8%** (Pass A orders)

| Case | Arithmetic | Expected |
|---|---|---|
| **C18 Sales tax 8%** (C01 order + calculator) | Txn unchanged: 0.065×35 = 2.275. Proc 0.03×35×1.08 + 0.25 = 1.134 + 0.25 = 1.384. Fees 0.20 + 2.275 + 1.384 = 3.859. Profit 35 − 3.859 − 19. Shopify processing is unaffected (1.315) | proc **1.384**, profit **12.141** |
| **C11 Shopify sale, yearly billing** | plan $19/mo, so 19/50 = 0.38 per order. Fees 1.315 + 0.38 = 1.695. Profit 35 − 1.695 − 19 | **14.305** |
| **C12 Reverse (yearly)** | (0.30 + 0.38 + 19)/0.671 = 29.3294, minus 5 = 24.3294 → rounded up | **24.33** |
| **C13 Break-even (yearly)** | 19/2.46 = 7.72 → rounded up | **8 orders/month** |

## Pass C — shop at $10k+, **third-party provider**, discount rule **N**

| Case | Arithmetic | Expected |
|---|---|---|
| **C15 Offsite Ads 12%** | rate 0.12. 0.12×35 = 4.20. Fees 3.775 + 4.20 = 7.975. Profit 35 − 7.975 − 19 | **8.025** |
| **C16 Shopify, third-party (log)** | (0.029 + 0.02)×35 + 0.30 = 2.015. Profit 35 − 2.015 − 19 | **13.985** |
| **C16 calculator** | 2.015 + 0.54 = 2.555. Profit 35 − 2.555 − 19 | **13.445** |
| **C19 Discount, alternative rule N** | C08 inputs. Txn 0.065×(59 + 6) = 4.225 (pre-discount). Proc 2.02 (always on what the buyer paid). Fees 0.40 + 4.225 + 2.02 = 6.645. Profit 59 − 6.645 − 30 − 6 | fees **6.645**, profit **16.355** |

## Lite calculator

| Case | Inputs | Arithmetic | Expected |
|---|---|---|---|
| **C14 Lite** | price 30, ship 5, ship cost 4, cost 15, Offsite Y, under $10k | listing 0.20, txn 2.275, proc 1.30, offsite 5.25. Fees 9.025. Profit 35 − 9.025 − 19. Margin 6.975/35 | profit **6.975**, margin **19.93%** |

## Static checks
- **S1:** formulas use only allowed functions.
- **S2:** no numeric literals other than 0, 1, 2, 3, 60 and the tie-break 1e9 (so no hardcoded rates).
- **S3:** no external links.
- **S4:** shipped files with sample data evaluate with no `#` errors.
- **S5:** the BLANK workbooks evaluate with no `#` errors on any formula cell. S1–S3 also cover the BLANK files.
- **S0:** every sheet is protected without a password, and every formula cell is locked. S0 is reported only if it fails.

Pending for next workbook release: mark Etsy listing $0.20, transaction 6.5%, Offsite Ads 15% VERIFIED (Etsy sell page, viewed 2026-10-23).
