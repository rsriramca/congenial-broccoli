# congenial-broccoli - Nova CFO Monthly Sales MIS (AI-assisted automation, synthetic data)

A portfolio project showing a controlled, auditable monthly sales and gross-profit MIS workflow for a **fictional** company (Nova Consumer Products). All data is synthetic. It is not financial advice and not a production system.

## Status
- Current pipeline revision: **v3 (release controls)** - self-test 68/68 passed for the committed code.
- Latest run on the original dataset: **READY_FOR_APPROVAL - not approved, not published.** No validated release exists in this repository yet (`releases/` is created only on approval).
- Not tested: Microsoft Excel Desktop (LibreOffice only). Not implemented: scheduling, unattended runs, live refresh, authenticated approver identity, cryptographic signing, external audit anchoring. Details: [docs/RUNBOOK.md](docs/RUNBOOK.md).

## Layout
| Path | Contents |
|---|---|
| `pipeline/run_mis.py` | CLI: `run`, `approve`, `selftest`, `anchor`, `verify-anchor`, `verify-release` |
| `pipeline/mis_engine.py` | Workbook builder (formula-driven, 51 controls) |
| `pipeline/selftest_result.json` | Self-test record, bound to the SHA-256 of the two code files. Stale record = runs blocked, approval refused |
| `data/source/` | Original synthetic dataset (never modified) + `SHA256SUMS` |
| `docs/RUNBOOK.md` | How to run, approve, anchor, verify; what the script does not do |
| `docs/BUDGET_INTEGRITY.md` | Why the original budget is not used for variance (about 7.1x actuals; root cause unresolved) |
| `docs/validation/` | Validation reports per candidate/release |
| `tests/` | Fast reconciliation tests (source hash, record currency, row arithmetic, control baselines) |
| `CHANGELOG.md` | Versions and approval status |

## Quick start
```
pip install pandas numpy openpyxl          # LibreOffice is also required for recalculation
python tests/test_reconciliation.py        # fast checks
python pipeline/run_mis.py run --input data/source/Nova_CFO_MIS_Complete_Synthetic_Dataset.xlsx --root mis_workspace
python pipeline/run_mis.py approve --run <run_id> --by "<name>" [--ack-warnings]   # human step only
```
If you change either file in `pipeline/`, the self-test record becomes stale; rerun `python pipeline/run_mis.py selftest --input data/source/...xlsx` (about 10 minutes) and commit the refreshed record.

## Working rules
- Original source data is never edited; `tests/` fails if its hash changes.
- Validated outputs are never overwritten: releases are immutable `releases/<month>/<run_id>/` folders; new versions go in through a pull request after approval.
- No credentials, personal or confidential data in the repo (synthetic data only; `.gitignore` blocks common secret files). Note: this repository is public.
- Generated run folders (`mis_workspace/`) are not committed; only approved releases and their validation reports are.
