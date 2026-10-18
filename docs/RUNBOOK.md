# Nova CFO MIS - runbook, release controls and automation boundaries

## One command (build a candidate; nothing is published)
```
python run_mis.py run --input <dataset.xlsx> --root mis_workspace
```
Needs Python 3 with pandas, numpy, openpyxl, and LibreOffice (headless recalculation). Takes about 20 seconds. Exit codes: 0 = ready for approval, 3 = blocked, 2 = input rejected.
Steps (each written to the audit log): archive input read-only + SHA-256; reject bad schema; **check the self-test record**; build; recalculate in LibreOffice; read back all 51 controls; independent pandas recompute; error scan; restatement check against the last published release; write candidate (`candidate/`) or, if any critical gate failed, `blocked/` with a `BLOCKED_` workbook.

## Critical gate: the self-test record
`selftest_result.json` is bound to the **combined code hash**: SHA-256 of the two text lines `mis_engine.py:<sha256 of file>` and `run_mis.py:<sha256 of file>` (joined by a newline).
A run is BLOCKED, and `approve` refuses, if the record is missing, unreadable, not for the current code (stale), shows any failure, has too few tests / lacks required test groups, is internally inconsistent (results hash mismatch), or is a test fixture.
`--ack-warnings` cannot override any of this. Fix: `python run_mis.py selftest --input <known-good dataset>` (about 10 minutes), then re-run.
Limit: the record is a plain JSON file with an internal hash. It catches stale code and casual hand-edits; it is not signed, so someone deliberately forging a complete record is not detected (test A5-style limits apply).

## Approval (human, separate, never automatic)
```
python run_mis.py approve --run <run_id> --by "<name>" [--ack-warnings] [--external-ref "<ticket/email id>"]
```
Hard refusals (none overridable by `--ack-warnings`): approver name blank; run not READY_FOR_APPROVAL; self-test record invalid for current code; audit chain broken or empty; `manifest.json` differs from the manifest hash logged in the audit; candidate file hash or file list differs from the manifest; unlisted files in the candidate folder; archived input differs from the hash logged at run time; code hash differs from the code that built the candidate; manifest/audit do not record READY_FOR_APPROVAL; manifest run id differs from the folder; release folder already exists.
Only run **warnings** (budget integrity, region/product overlap, restatement) can be acknowledged with `--ack-warnings`.

On success: creates `published/<latest-month>/<run_id>/` with `os.mkdir` (fails if it exists), copies the candidate files and `manifest.json`, re-verifies the copies against the manifest hashes, writes `APPROVAL.json` (approver, time, optional external reference, file hashes, input/code hashes, audit hash before approval), best-effort marks files read-only, logs the approval, appends to the hash-chained `published/release_register.jsonl`, and **only then** atomically replaces `published/LATEST.json`. On any error the partial release is removed and LATEST is untouched.
Releases are never overwritten or deleted by the script. A later release for the same month gets a different `<run_id>` folder; earlier ones stay.

## Publication boundary - what the script does NOT do
- It does **not** stop anyone copying files by hand into `published/` or anywhere else, or editing a released file. The read-only flag is a convenience, not security. To enforce the boundary use folder permissions (write access for the approver only / read-only share / write-once storage) or a signed-release verification process outside this script.
- `python run_mis.py verify-release --release published/<month>/<run_id>` recomputes the release hashes against `APPROVAL.json` and the register. It detects later edits and manual copies that are not in the register. It cannot detect a wholly forged release **and** register unless the register hash is held outside the workspace.
- No cryptographic signing is implemented. The approver name is typed text, not an authenticated identity. `--external-ref` only records a reference to an approval that you hold elsewhere.

## Audit trail: edit-detecting, not tamper-proof
`runs/<run_id>/audit/audit_log.jsonl` is a hash chain: changing or deleting an earlier line (or emptying the file) breaks verification. Someone with write access to the workspace can recompute the whole chain (self-test A5 demonstrates this), so it is **not tamper-proof and not tamper-evident on its own**. It becomes tamper-evident only if the final hash is stored outside the workspace.
Optional external step (not performed in this delivery):
1. `python run_mis.py anchor --run <run_id>` writes `audit/anchor_record_<n>_entries.txt` (final hash, input/code/manifest hashes). Run it after `run`, and again after `approve`.
2. You store the `final_audit_hash` somewhere the workspace owner cannot silently edit (email to an auditor, ticket, write-once store, signed message). The script cannot do or verify this.
3. Later: `python run_mis.py verify-anchor --run <run_id> --hash <stored value>`.
Until step 2 is done, treat the log as edit-detecting only.

## Commentary
Rule-based text from the verified results (no AI model is called by the script), max 150 words, OBSERVED and UNVERIFIED sections. Optional `--commentary-file` is checked for length and number traceability only; the check does not validate sign, direction or causal wording.

## Automated vs human
| Step | Status |
|---|---|
| Schema validation, build, recalculation (LibreOffice), 51 controls, independent recompute, error scan, restatement check | Automated, tested |
| Blocking on any critical failure, incl. stale/failing self-test | Automated, tested |
| Immutable release folders, approval verification, release register | Automated, tested (self-test groups P, M, V) |
| Interpretation of warnings; budget decision; approval to publish | Human |
| Folder permissions / signed-release process; external anchoring | Human / not implemented |
| Opening and checking the file in Microsoft Excel Desktop | Human - NOT tested |
| Scheduling, unattended runs, live refresh, distribution, authenticated approver identity | NOT implemented |

## Folder layout
```
mis_workspace/
  runs/<run_id>/input/       read-only copy of the source
  runs/<run_id>/candidate/   workbook, CFO_commentary.md, validation_report.md, fact_pack.json, manifest.json   (or blocked/)
  runs/<run_id>/audit/       audit_log.jsonl, optional anchor_record_*.txt
  runs/<run_id>/status.json
  published/<month>/<run_id>/   immutable release + APPROVAL.json   (created only by approve)
  published/release_register.jsonl, published/LATEST.json
```
