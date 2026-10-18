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

## Run a candidate workflow

Requirements: Python 3, pandas, numpy, openpyxl, and LibreOffice for headless recalculation.

```bash
python run_mis.py run --input <dataset.xlsx> --root mis_workspace
```

This produces a candidate for review. Publication requires a separate named human approver as documented in the [runbook](Runbook_and_Automation_Boundaries.md).

## Before using this project

Read the [budget integrity report](Budget_Integrity_Report.md) and [validation report](Validation_Report_v2.md). Do not upload real financial, customer, or credential-bearing files.
