# Changelog

## v3 - release controls (candidate run 20261017-175501_46de2957: READY_FOR_APPROVAL, NOT approved)
- Self-test record is a critical gate bound to the combined hash of `run_mis.py` + `mis_engine.py`; `--ack-warnings` cannot override.
- Immutable releases (`<month>/<run_id>/`), release register, atomic `LATEST.json`.
- `approve` verifies manifest vs audit hash, candidate hashes, archived input, code hash; new `anchor`, `verify-anchor`, `verify-release`.
- Audit log described as edit-detecting, not tamper-proof; empty log now rejected (defect found by own test).
- Self-test 41 -> 68 tests, all pass. Candidate: 10/10 gates, 51 controls, 325 figures recomputed, 0 differences.

## v2 / v1 (earlier rounds, superseded; not committed)
- Budget integrity analysis, region/product confounding observation, one-command pipeline, commentary grounding. Round-1 scripts (`build_cfo_mis.py`, `validate_cfo_mis.py`) are superseded and intentionally omitted.
