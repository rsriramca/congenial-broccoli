# Nova CFO MIS Automation — Provisional v1

A portfolio project demonstrating an AI-assisted monthly CFO Sales MIS workflow using verified synthetic data.

> **Status:** Provisional v1. The workflow is not a fully autonomous refresh solution and requires human finance review before any release or management use.

## What is included

- **CFO monthly MIS workbook:** `CFO_Monthly_Sales_MIS.xlsx`
- **Automation:** `run_mis.py` and `mis_engine.py`
- **Validation evidence:** `Validation_Report_v2.md`, `selftest_result.json`, `run_validation_report.md`
- **Budget integrity review:** `Budget_Integrity_Report.md`
- **Operating boundaries:** `Runbook_and_Automation_Boundaries.md`
- **Management commentary:** `CFO_commentary.md`

## Test and reference additions
- **Original synthetic source dataset:** `data/source/Nova_CFO_MIS_Complete_Synthetic_Dataset.xlsx` with `data/source/SHA256SUMS` (the file the v1 evidence was produced from; never edit it).
- **Fast reconciliation tests:** `tests/test_reconciliation.py` (source hash, self-test record currency, row arithmetic and keys, dataset control baselines).
- **Change history:** `CHANGELOG.md`.

## Project structure

```text
data/       Synthetic sample-data policy and future approved inputs
docs/       Project documentation
outputs/    Future approved synthetic outputs and release artefacts
src/        Future refactored application modules
```

The original v1 files remain at the repository root intentionally, preserving their existing evidence paths while establishing a clear structure for future work.

## Financial integrity and scope

- The repository contains synthetic, non-sensitive example material only.
- The original budget remains unchanged.
- Budget-versus-actual comparability is unresolved; any `Budget_v2_Synthetic` output is illustrative only, not an approved budget.
- The documentation records automated checks, but it does not claim Microsoft Excel Desktop testing, scheduled refresh, live-data integration, unattended execution, or autonomous publication.

## Known limitations (Provisional v1)

- Budget-versus-actual comparability is **unresolved**: the original budget is about 7.1x actuals and 42% of it sits in cells with no sales; the cause is not established, so no original-budget variance is reported as approved.
- The approver name is typed text, not an authenticated identity; the audit log is edit-detecting only (not tamper-proof, not externally anchored); the script cannot prevent files being copied by hand outside its controlled folders. See the [runbook](Runbook_and_Automation_Boundaries.md).
- Validated only in LibreOffice. No Microsoft Excel Desktop test. Not production-ready.

## Run the checks

```bash
python tests/test_reconciliation.py     # fast reconciliation checks
python run_mis.py selftest --input data/source/Nova_CFO_MIS_Complete_Synthetic_Dataset.xlsx   # full fault-injection suite (about 10 minutes)
```
Changing `run_mis.py` or `mis_engine.py` makes `selftest_result.json` stale; runs are then blocked until the self-test is rerun.

## Run a candidate workflow

Requirements: Python 3, pandas, numpy, openpyxl, and LibreOffice for headless recalculation.

```bash
python run_mis.py run --input <dataset.xlsx> --root mis_workspace
```

This produces a candidate for review. Publication requires a separate named human approver as documented in the [runbook](Runbook_and_Automation_Boundaries.md).

## v2 enhancement: executive FP&A dashboard (separate from Provisional v1)

A static HTML dashboard built from the same synthetic source data is documented in [docs/DASHBOARD_V2.md](docs/DASHBOARD_V2.md) (`outputs/v2/Nova_FPA_Dashboard.html`). It carries the budget-comparability warning, shows no budget variance, is a manual snapshot (no live refresh), and does not change the Provisional v1 baseline above.

## Before using this project

Read the [budget integrity report](Budget_Integrity_Report.md) and [validation report](Validation_Report_v2.md). Do not upload real financial, customer, or credential-bearing files.
