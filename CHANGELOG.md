# Changelog

## Corrective: self-test evidence refresh (post PR #4)
- Regenerated `selftest_result.json` by running the full self-test (69/69 pass, was 68/68) on the corrected `run_mis.py`. Record was not edited by hand.
- Fix: run folders could collide when the same input was run twice within one second (found when the full self-test crashed); run IDs now get a `-2`, `-3` suffix. New self-test S7 covers it. No financial logic changed.
- Added `.gitattributes` (LF pinned) because the record is bound to the SHA-256 of the code files and CRLF conversion makes it read as stale; clearer stale-record message in `tests/test_reconciliation.py`; dashboard builder writes LF on every OS. Dashboard figures, formulas and disclosures unchanged.
- Historical v1 evidence (`Validation_Report_v2.md`, root `manifest.json`/`status.json`/run-folder zip) still describes the earlier 68/68 run at code hash `5afbee66...` and is intentionally unchanged. That candidate cannot be approved with the current code; a new candidate run is needed (none was approved).

## v2 dashboard enhancement (separate; Provisional v1 baseline unchanged)
- Added `src/build_dashboard.py`, `outputs/v2/` (static HTML dashboard + figures JSON), `tests/test_dashboard.py`, `docs/DASHBOARD_V2.md`. No v1 file modified. No budget variance shown; manual snapshot only.

## Provisional v1 (current)
- Initial prototype delivered to `main`: workbook, automation scripts, validation evidence, commentary, budget-integrity report, runbook (root-level files, unchanged).
- Documentation structure and data policy added (`docs/`, `data/`, `outputs/`, `src/` guides).
- Added: original synthetic source dataset with SHA-256 (`data/source/`), fast reconciliation tests (`tests/`), `.gitignore`, this changelog.
- Status: candidate run `20261017-175501_46de2957` is READY_FOR_APPROVAL but **not approved and not published**. No release exists. Budget comparability unresolved.
