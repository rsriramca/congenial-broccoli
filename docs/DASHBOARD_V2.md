# v2 Executive FP&A Dashboard (enhancement - separate from Provisional v1)

A static, self-contained HTML dashboard built from the **original synthetic dataset**. It does not change, relabel or replace any Provisional v1 file. v1 remains the baseline.

## What it shows
- KPI tiles: YTD net revenue, gross profit and gross margin vs prior-year YTD; latest actual month.
- Monthly net revenue as stacked columns (COGS + gross profit) with the revenue total above each column and the segment values inside, and monthly gross margin as a labelled line (axis does not start at zero - labelled).
- Current YTD (Jan-Sep 2026) vs prior-year YTD (Jan-Sep 2025): chart and table.
- Region view and product view with a caveat that **region and product effects overlap** (each product is sold in exactly one region; each region carries only two products), so the views are confounded.
- A prominent **data-quality warning at the top**: the original budget is about 7.1x actuals, 42% of it sits in cells with no sales, cause unresolved. **No budget variance is calculated or shown.** Any variance in the v1 workbook's `Budget_v2_Synthetic` / `Budget_vs_Actual_v2` sheets is illustrative only and not an approved budget.
- Every bar carries a visible value label (above or inside; checked by a test). Colours: blue = COGS / current-period revenue, green = gross profit and margin; both pairs pass the dataviz palette validator in light and dark.
- Every chart has hover tooltips (native SVG titles) and a table alternative; light and dark themes follow the OS setting; charts scroll sideways on narrow screens.

## Files
| File | Role |
|---|---|
| `src/build_dashboard.py` | Builder: reads the source workbook read-only, recomputes figures in pandas, writes the two outputs |
| `outputs/v2/Nova_FPA_Dashboard.html` | The dashboard (inline SVG + CSS, no scripts, no network requests) |
| `outputs/v2/dashboard_data.json` | Every figure shown, for reconciliation |
| `tests/test_dashboard.py` | Reconciliation and content checks |

## Run / refresh (manual)
```
pip install pandas numpy openpyxl
python src/build_dashboard.py --input data/source/Nova_CFO_MIS_Complete_Synthetic_Dataset.xlsx --out outputs/v2
python tests/test_dashboard.py
```
Open `outputs/v2/Nova_FPA_Dashboard.html` in any browser. The output is deterministic (no timestamp): the same input gives byte-identical files, and a test fails if the committed output is stale.
To use a newer synthetic extract: add it under `data/`, build and review the v1 MIS candidate first (`python run_mis.py run ...`; approval stays a separate human step), then rerun the builder on that file and review the diff before committing.

## Assumptions and limits
- YTD uses calendar months 1-9 of 2026 vs the same months of 2025. The source calendar has no fiscal year, so calendar-year YTD is an assumption; the figures match the v1 workbook's YTD summary.
- Figures are actuals only; forecast months and budget are excluded. Net revenue excludes GST, returns, rebates and intercompany.
- It is a **point-in-time snapshot**. There is no live refresh, scheduling, autonomous publishing or data connection.
- Reconciled to the v1 workbook's LibreOffice-calculated values. **Not tested in Microsoft Excel Desktop**; browser rendering was checked in Chromium only. Not production-ready.
- Differences shown (for example YTD revenue -1.8%) are small and come from one synthetic series; the data cannot explain causes.
