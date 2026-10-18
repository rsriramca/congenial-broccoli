# Validation report - Nova CFO MIS (revision 3: release controls)

Candidate run: `20261017-175501_46de2957` - status **READY_FOR_APPROVAL**. **Not approved. Not published.** No `published/` folder exists in the delivered workspace.

## What changed in this revision
1. **Self-test record is a critical gate.** Bound to the combined SHA-256 of `mis_engine.py` + `run_mis.py`. Missing, unreadable, stale, failing, too-short, inconsistent or fixture records BLOCK the run and make `approve` refuse; `--ack-warnings` cannot override (tests I1-I6).
2. **Immutable releases.** `approve` publishes to `published/<latest-month>/<run-id>/` via `os.mkdir` (never overwrites), keeps prior releases, appends a hash-chained `release_register.jsonl`, and updates `LATEST.json` atomically only after everything else succeeded (P1-P7).
3. **Audit trail wording and optional anchor.** The log is described as edit-detecting, not tamper-proof. New optional `anchor` / `verify-anchor` commands prepare and check an external hash record; A5 shows a fully re-chained log passes the internal check but fails the external hash. Not anchored in this delivery.
4. **Explicit publication boundary.** `approve` verifies self-test record, audit chain, manifest-vs-audit hash, candidate file hashes, archived input hash, code hash and run status before copying (M1-M4). Manual copying outside the controlled folder is NOT prevented; `verify-release` flags unregistered copies (M5, V1-V3). Folder permissions or a signed-release process are required.
5. Fixed during testing: `verify_audit` accepted an EMPTY log as valid (found by my own test A4); now rejected.
6. Preserved: original budget unchanged; `Budget_v2_Synthetic` labelled illustrative/unapproved; region/product overlap flagged; no original-budget variance reported.

## Candidate run gates (10/10 pass)
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

51 workbook controls PASS; 325 report figures independently recomputed, 0 differences; 0 error cells; warnings (2): budget integrity (original budget 7.1x actuals, root cause unresolved) and region/product overlap. Recalculation: LibreOffice only.

## Pipeline self-test: 68/68 passed (code hash `5afbee66355edddb...`, 2026-10-17T17:54:33Z)
| # | Test | Result | Evidence |
|---|---|---|---|
| 1 | S1 baseline run on the original dataset | PASS | status=READY_FOR_APPROVAL; warnings=2; gates 10/10 |
| 2 | S2 approval refused while warnings unacknowledged | PASS | refused: run has warnings; review them and re-run approve with --ack-warnings:
  - O3 Reg |
| 3 | S3 approval refused without approver name | PASS | --by <approver name> is required |
| 4 | S4 audit-log edit detected, approval refused | PASS | audit log integrity check failed: content altered at line 2 |
| 5 | S5 candidate file edit detected, approval refused | PASS | candidate file CFO_Monthly_Sales_MIS.xlsx changed since the run (hash mismatch) |
| 6 | S6 approval with acknowledgement publishes and writes APPROVAL.json | PASS | cratchpad/stroot2/published/2026-09/20261017-174734_46de2957 |
| 7 | F: duplicate Invoice_ID -> BLOCKED with specific control | PASS | status=BLOCKED; detects 'Duplicate Invoice_ID': True; no candidate folder: True |
| 8 | F: duplicate Invoice_ID -> approval refused | PASS | run status is BLOCKED; only READY_FOR_APPROVAL runs can be approved |
| 9 | F: unmatched Product_ID -> BLOCKED with specific control | PASS | status=BLOCKED; detects 'Unmatched Product_ID': True; no candidate folder: True |
| 10 | F: unmatched Product_ID -> approval refused | PASS | run status is BLOCKED; only READY_FOR_APPROVAL runs can be approved |
| 11 | F: unmatched Customer_ID -> BLOCKED with specific control | PASS | status=BLOCKED; detects 'Unmatched Customer_ID': True; no candidate folder: True |
| 12 | F: unmatched Customer_ID -> approval refused | PASS | run status is BLOCKED; only READY_FOR_APPROVAL runs can be approved |
| 13 | F: revenue off by INR 500 -> BLOCKED with specific control | PASS | status=BLOCKED; detects 'Net_Revenue <> Units': True; no candidate folder: True |
| 14 | F: revenue off by INR 500 -> approval refused | PASS | run status is BLOCKED; only READY_FOR_APPROVAL runs can be approved |
| 15 | F: COGS off by INR 250 -> BLOCKED with specific control | PASS | status=BLOCKED; detects 'COGS <> Units': True; no candidate folder: True |
| 16 | F: COGS off by INR 250 -> approval refused | PASS | run status is BLOCKED; only READY_FOR_APPROVAL runs can be approved |
| 17 | F: gross profit off by INR 100 -> BLOCKED with specific control | PASS | status=BLOCKED; detects 'Gross_Profit <>': True; no candidate folder: True |
| 18 | F: gross profit off by INR 100 -> approval refused | PASS | run status is BLOCKED; only READY_FOR_APPROVAL runs can be approved |
| 19 | F: Forecast row relabelled Actual -> BLOCKED with specific control | PASS | status=BLOCKED; detects 'Transaction Status vs Calendar': True; no candidate folder: True |
| 20 | F: Forecast row relabelled Actual -> approval refused | PASS | run status is BLOCKED; only READY_FOR_APPROVAL runs can be approved |
| 21 | F: month missing from Calendar -> BLOCKED with specific control | PASS | status=BLOCKED; detects 'Unmatched Month': True; no candidate folder: True |
| 22 | F: month missing from Calendar -> approval refused | PASS | run status is BLOCKED; only READY_FOR_APPROVAL runs can be approved |
| 23 | F: negative units (self-consistent row) -> BLOCKED with specific control | PASS | status=BLOCKED; detects 'Non-positive': True; no candidate folder: True |
| 24 | F: negative units (self-consistent row) -> approval refused | PASS | run status is BLOCKED; only READY_FOR_APPROVAL runs can be approved |
| 25 | F: blank Net_Price cell -> BLOCKED with specific control | PASS | status=BLOCKED; detects 'Blank cells': True; no candidate folder: True |
| 26 | F: blank Net_Price cell -> approval refused | PASS | run status is BLOCKED; only READY_FOR_APPROVAL runs can be approved |
| 27 | F: duplicate budget key -> BLOCKED with specific control | PASS | status=BLOCKED; detects 'Duplicate Month/Region/Category': True; no candidate folder: True |
| 28 | F: duplicate budget key -> approval refused | PASS | run status is BLOCKED; only READY_FOR_APPROVAL runs can be approved |
| 29 | S: missing column -> REJECTED_INPUT, no workbook built | PASS | Sheet Customer_Master is missing columns ['Region'] |
| 30 | S: unexpected Status value -> REJECTED_INPUT | PASS | Unexpected Status values in Sales_Transactions: {'Draft'} |
| 31 | R: internally consistent restatement of a published month -> READY but RESTATEMENT warning | PASS | status=READY_FOR_APPROVAL; restatement warning: True |
| 32 | R: restated run cannot be approved without acknowledgement | PASS | run has warnings; review them and re-run approve with --ack-warnings:
 |
| 33 | N: next-month simulation with STALE source baselines -> BLOCKED | PASS | status=BLOCKED |
| 34 | N: next-month simulation with refreshed baselines -> READY, 22 months, no restatement | PASS | status=READY_FOR_APPROVAL; warnings=['O3 Region vs product: Maximum number of ', 'Budget integrity: original b |
| 35 | N: simulation workbook lists 22 actual months | PASS | Monthly_MIS months counted |
| 36 | O: dataset without Sales_Budget -> READY, no budget sheets | PASS | status=READY_FOR_APPROVAL |
| 37 | C: commentary with an invented number -> BLOCKED | PASS | Commentary override rejected: numbers not found in verified fact pack: ['41.3%'] |
| 38 | C: commentary over 150 words -> BLOCKED | PASS | Commentary override rejected: 188 words (limit 150) |
| 39 | C: commentary without UNVERIFIED section -> BLOCKED | PASS | Commentary override rejected: must contain separate OBSERVED and UNVERIFIED CAUSES sections |
| 40 | C: valid externally supplied commentary -> READY with disclosure warning | PASS | status=READY_FOR_APPROVAL |
| 41 | I1 stale self-test record (code hash differs) -> run BLOCKED by selftest_record gate | PASS | self-test record is STALE: run_mis.py/mis_engine.py changed since it was produced - run "run_mis.py  |
| 42 | I2 failing self-test record for current code -> run BLOCKED | PASS | self-test record shows failures (71/73 passed) - run "run_mis.py selftest" on the current code |
| 43 | I3 missing self-test record -> run BLOCKED | PASS | no self-test record found - run "run_mis.py selftest" on the current code |
| 44 | I4 hand-edited failing record (results hash no longer matches) -> run BLOCKED | PASS | self-test record is internally inconsistent (edited?) - run "run_mis.py selftest" on the current cod |
| 45 | I5 (control) valid fixture record -> READY (so the I1-I4 blocks are caused by the record) | PASS | READY_FOR_APPROVAL |
| 46 | I5 approval refused with --ack-warnings when record is stale / failing / missing | PASS | CRITICAL - self-test record invalid for the current code: self-test re |
| 47 | I6 default record path: missing or fixture record -> approve refused and run BLOCKED; fixture rejected as genuine | PASS | CRITICAL - self-test record invalid for  / CRITICAL - self-test record invalid for the curren |
| 48 | P1 first release is published/<latest-month>/<run-id>/ with workbook, manifest and APPROVAL.json | PASS | published/2026-09/20261017-174734_46de2957 |
| 49 | P2 re-approving an already published run is refused | PASS | run status is PUBLISHED; only READY_FOR_APPROVAL runs can be approved |
| 50 | P3 second approval of the same month goes to a NEW folder; first release files unchanged; both kept; LATEST -> second | PASS | published/2026-09/20261017-175314_46de2957 |
| 51 | P4 pre-existing release folder is never overwritten; approval refused; sentinel untouched; run stays READY | PASS | release folder already exists and will never be overwritten: /tmp/claude-0/-home |
| 52 | P5 LATEST.json unchanged by every refused approval | PASS | bytes identical before/after refusals |
| 53 | P6 released files are marked read-only (best effort; NOT a security control) | PASS | mode bits |
| 54 | P7 release register is a valid hash chain with one entry per release | PASS | 2 entries |
| 55 | M1 workbook AND manifest edited consistently -> refused (manifest hash recorded in audit log) | PASS | candidate manifest.json does not match the manifest hash recorded in the audit log (manife |
| 56 | M2 archived input modified -> refused | PASS | archived input does not match the input hash recorded at run time |
| 57 | M3 code differs from the code that built the candidate -> refused | PASS | code hash mismatch: the code that built this candidate is not the code now installed (re-r |
| 58 | M4 BLOCKED run relabelled as candidate with edited status.json -> refused | PASS | audit log / manifest record this run as BLOCKED/BLOCKED, not READY_FOR_APPROVAL |
| 59 | M5 boundary: a MANUAL copy into published/ is not prevented by the script, but verify-release flags it as unregistered | PASS | no APPROVAL.json: not a release produced by "approve" |
| 60 | A1 anchor record written, states NOT ANCHORED until stored externally, and audit log itself unchanged | PASS | anchor_record_16_entries.txt |
| 61 | A2 verify-anchor passes with the correct externally stored hash | PASS | externally stored hash found at entry 15; log is unchanged since the anchor. Thi |
| 62 | A3 verify-anchor fails with a wrong hash | PASS | the externally stored hash does not appear anywhere in this log (log rewritten o |
| 63 | A4 edited audit log, and an emptied audit log -> verify-anchor fails | PASS | audit chain is internally broken: content altered  / audit chain is internally broken: audit  |
| 64 | A5 WHOLE chain recomputed after an edit passes the internal check (why it is not tamper-proof) but fails against the externally stored hash | PASS | internal chain valid=True; anchor check code=5 |
| 65 | V1 verify-release passes on an untouched release | PASS | release 20261017-174734_46de2957 matches APPROVAL.json and the register (approve |
| 66 | V2 verify-release fails when a released file is modified | PASS | release file CFO_Monthly_Sales_MIS.xlsx missing or modified since approval |
| 67 | V3 verify-release fails when APPROVAL.json is edited | PASS | APPROVAL.json differs from the version recorded in the register |
| 68 | X: original dataset unchanged by the whole self-test | PASS | SHA-256 46de2957611b9534... before == after |

## NOT tested / NOT done
- **Microsoft Excel Desktop:** no test of any kind (opening, recalculation, formula behaviour, formats, charts). LibreOffice only.
- **Approval and publication:** not performed on the delivered workspace. Self-test approvals ran only in a throwaway workspace.
- **Scheduling, unattended execution, live refresh, distribution, any AI-model call in the script:** not implemented.
- **Authenticated approver identity:** not implemented; approver is typed text.
- **Externally anchored audit log / cryptographic release signing:** not implemented; `anchor` only prepares a record you must store yourself.
- **Folder permissions / signed-release process:** outside the script; not set up.
- Self-test record is unsigned JSON: a deliberately forged complete record is not detected.

## Limits of the evidence
- Fault injection covers 11 data-defect types and the listed release-control attacks, not every possible one. Restatement and next-month tests are simulations from existing rows.
- Self-test results are valid only for the recorded code hash; any code edit makes them stale and blocks runs until the self-test is rerun.
