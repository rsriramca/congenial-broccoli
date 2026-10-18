# Validation report - run 20261017-175501_46de2957
Status: **READY_FOR_APPROVAL**  |  input SHA-256 `46de2957611b9534d77f43628ee22a45b11b094f7f2ae46fb690f1c3d690ea3c`  |  code SHA-256 `5afbee66355edddb...`

## Publication gates (critical)
| Gate | Result | Detail |
|---|---|---|
| input_schema | PASS | all required sheets/columns present; Status values valid |
| selftest_record | PASS | valid for current code hash 5afbee66355e...: 68/68 tests passed at 2026-10-17T17:54:33Z |
| commentary_valid | PASS | 128 words; numbers traceable to fact pack; rule-based commentary |
| recalc_completed | PASS | LibreOffice headless recalculation succeeded (pass 1) |
| workbook_controls | PASS | ALL 51 CHECKS PASS; 0 failing:  |
| independent_recompute | PASS | 325 figures compared, 0 differences |
| error_scan | PASS | 0 error cells |
| pass1_equals_final | PASS | control results identical between build pass 1 and final build |
| source_unchanged | PASS | source SHA-256 identical before and after the run (original and working copy) |
| audit_chain | PASS | hash chain valid (edit-detecting only; not externally anchored) |

## Workbook controls
ALL 51 CHECKS PASS. Groups: Record counts: 5/5, Keys: 7/7, Status: 1/1, Row arithmetic: 4/4, Reconciliation: 22/22, Scope: 4/4, Budget v2: 6/6, Integrity: 2/2

## Warnings (need acknowledgement at approval)
- O3 Region vs product: Maximum number of regions in which any single product sold (actual) -> value 1.00
- Budget integrity: original budget is 7.1x actuals, 42% in cells with no sales; root cause UNRESOLVED; only SYNTHETIC v2 (unapproved) variance is shown.

## Budget integrity
Original budget / actual = 7.05x (4.10x on cells that have sales). Root cause: PARTIALLY DEMONSTRATED. Demonstrated: (1) the budget is a closed-form formula grid with no link to the transaction table; (2) 42% of it sits in 5 Region x Category cells where the source has no sales at all, because each product is sold in one region only. Removing those cells narrows the gap from 7.1x to 4.1x. NOT demonstrated: the remaining 4.1x. It is not a constant factor (1.5x-15.3x), so it is not a unit/currency issue. Most likely the two tables were generated independently without calibration, but that cannot be proven from the data. Status: UNRESOLVED - business confirmation needed.

## Tested in this run
- Input schema, source immutability (hash before/after), recalculation in LibreOffice, all workbook controls, independent pandas recompute, formula-error scan, commentary length and number traceability, audit-chain integrity.

## NOT tested - requires Microsoft Excel Desktop
- [ ] Open the candidate in Microsoft Excel Desktop (Windows or Mac) - expected: File opens without a repair prompt or "update links" warning.
- [ ] Force a full recalculation (Ctrl+Alt+F9) and read Control_Checks!C3 - expected: Shows "ALL <n> CHECKS PASS"; no #NAME?/#VALUE!/#N/A anywhere.
- [ ] Compare Monthly_MIS totals and Executive_Summary KPIs with the run's validation_report.md - expected: Values equal the LibreOffice-recalculated values to the rupee.
- [ ] Check month-filter formulas (SUMIFS on numeric Year/Month_No; LOOKUP(2,1/...) for latest month) - expected: Monthly_MIS!B4 shows the latest actual month; monthly rows non-zero.
- [ ] Check Indian lakh/crore number format, conditional formatting (PASS green / FAIL red) and chart rendering - expected: Formats and 5-6 charts display as intended.
- [ ] Check calculation time on a normal laptop - expected: Full recalculation completes in a few seconds (SUMPRODUCT/COUNTIF arrays over ~900 rows).

## Audit trail and publication boundary
- The audit log is a hash chain: it DETECTS EDITS to earlier lines. It is NOT tamper-proof and is NOT externally anchored. It becomes tamper-evident only if the final hash (see `run_mis.py anchor`) is stored outside this workspace; that has NOT been done for this run.
- `approve` verifies the self-test record, audit chain, candidate manifest, archived input hash and code hash, then copies to a new folder `published/<month>/<run_id>/`. It never overwrites a release.
- The script cannot prevent files being copied by hand into `published/` or elsewhere. That needs folder permissions or a signed-release process outside this script. `verify-release` can detect later edits to a release made through this workflow, not a wholly forged one.
- The approver name is typed text, not an authenticated identity.

## Not implemented
- Scheduling, unattended execution, live refresh, distribution, AI-model call (commentary is rule-based unless a commentary file is supplied), authenticated approver identity, cryptographic signing, external anchoring.
