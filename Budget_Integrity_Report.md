# Task 1 - Budget integrity review: Sales_Budget vs actuals

Original dataset not modified. All figures below are computed from the source file by `mis_engine.py` and are reproduced live in the workbook sheet `Budget_Integrity`.

**Headline:** Sales_Budget revenue for the 21 actual months is 7.05x actual net revenue. On the Region x Category cells that have any sales it is 4.10x.

## Tests performed
| Test | Method | Finding | Verdict |
|---|---|---|---|
| Reporting periods | Set of Sales_Budget months vs Calendar months | Aligned: same 24 months | NO PERIOD MISMATCH |
| Aggregation level | Budget rows per month; grain of each table | Budget = Month x Region x Category (12 cells/month, no product or customer dimension). Transactions = invoice x product, aggregated here to the same grain. Product-level budget cannot be derived. | COMPARABLE ONLY AT REGION x CATEGORY |
| Units / currency / scale factor | Budget/actual ratio for each in-scope Month x Region x Category cell | 147 cell-months: ratio ranges 1.5x to 15.3x (median 4.6x, CV 51%). A uniform unit or currency conversion would give one constant ratio. | RULED OUT as sole cause |
| Periodicity (annual / quarterly figures) | Budget divided by 12 or by 3 vs actual | Budget/12 = 0.59x actual; Budget/3 = 2.35x. Neither is close to 1.0x. | RULED OUT |
| Business scope (coverage of Region x Category cells) | Cells with budget but no actual sales in any month | 5 of 12 cells have zero sales in the source (East/Personal Care, North/Foods, North/Home Care, South/Personal Care, West/Foods); they hold 41.9% of budget for the actual months. Cause in the data: each product is sold in one region only, so each region covers only 7/12 of the grid. Restricting budget to cells that have sales reduces the gap from 7.05x to 4.10x. | CONFIRMED (explains part of the gap) |
| Transaction coverage | Invoices per month, customers invoiced, avg invoice value | 36 invoices/month, avg INR 56,340; 48/48 customers invoiced. Budget of INR 14,304,171/month implies about 254 invoices/month at that average. Cannot tell whether the transaction table is a sample or the budget is overstated. | UNRESOLVED |
| Data-generation signature | Fit of budget to a closed-form grid | Budget is constant within each year (no seasonality), gross margin is exactly 38% in every row, 2nd-year uplift is exactly 1.08x for every cell, and revenue is additive in Region + Category effects (max residual INR 0.00). It is formula-generated, not built from transactions. | CONFIRMED |
| Margin and trend consistency | Budget GP % vs actual GP %; budget growth vs actual growth | Budget GP % 38.0% vs actual 35.7%. Budget steps up 8% in year 2; actual YTD revenue moved -1.8%. | INFORMATION |

## Root cause
PARTIALLY DEMONSTRATED. Demonstrated: (1) the budget is a closed-form formula grid with no link to the transaction table; (2) 42% of it sits in 5 Region x Category cells where the source has no sales at all, because each product is sold in one region only. Removing those cells narrows the gap from 7.1x to 4.1x. NOT demonstrated: the remaining 4.1x. It is not a constant factor (1.5x-15.3x), so it is not a unit/currency issue. Most likely the two tables were generated independently without calibration, but that cannot be proven from the data. Status: UNRESOLVED - business confirmation needed.

## Hypotheses considered and what the data says
- **Different currency or unit (e.g. thousands/lakhs):** a uniform conversion would give one constant ratio. The cell-month ratio runs from 1.5x to 15.3x (CV 51%). Ruled out as the sole cause.
- **Budget stated at a different period (annual or quarterly rows):** budget/12 = 0.59x and budget/3 = 2.35x of actual. Neither is near 1.0x. Ruled out.
- **Different business scope:** confirmed in part. 5 of 12 Region x Category cells have no sales in the source and carry 41.9% of the budget.
- **Incomplete transaction coverage (sample):** possible but not provable. 36 invoices a month at an average of INR 56,340 would need about 254 invoices a month to reach the budget.
- **Independent synthetic generation:** consistent with the evidence. The budget is an exact formula grid (constant within each year, flat 38%% margin, exactly +8%% in 2026, additive Region and Category effects). Not provable as intent.

## Why the 5 cells are empty (structural finding that also affects the MIS)
Each product is sold in exactly one region: East: Breakfast Cereal, Dishwash Gel; North: Body Lotion, Soap 4-pack; South: Biscuits 500g, Detergent 1kg; West: Floor Cleaner, Shampoo 400ml. So every region covers only two products, and region and product effects cannot be separated in this dataset.

## Controlled correction (separate version, not an approved budget)
`Budget_v2_Synthetic` multiplies every original budget cell by one factor k = 0.2541 (FY2025 actual revenue / original FY2025 budget in cells that have sales). The original sheet is carried unchanged; v2 is a separate sheet.
**Assumptions:** (1) the intended budget scale is the same order as FY2025 actuals - an assumption, not a fact; (2) the budget margin stays at 38%, not calibrated; (3) cells without sales are kept but excluded from variance; (4) FY2025 variance is zero by construction, so only 2026 is a meaningful, out-of-sample view; (5) one factor cannot fix the uneven cell-level ratios, so cell variances are indicative only.
**Illustrative result (YTD Sep-2026, in-scope cells):** actual INR 1.81 crore vs v2 budget INR 1.98 crore, -8.8%. This is driven by the original budget's +8% 2026 step against actual -1.8%; it is only as reliable as assumption (1).

## Decision needed from the business (cannot be resolved from the data)
1. Confirm the intended budget scale (or supply a corrected Sales_Budget).
2. Confirm whether the five empty Region x Category cells should have sales, i.e. whether the transaction table is incomplete.
3. Until then, report no budget variance as approved. The v2 sheets are labelled ILLUSTRATIVE ONLY.
