"""Fast financial reconciliation checks (no LibreOffice needed). Run: python -m pytest tests  (or: python tests/test_reconciliation.py)

1. The original synthetic source is byte-identical to the recorded SHA-256 (it must never be modified).
2. The self-test record matches the CURRENT pipeline code (stale record = pipeline refuses to approve).
3. Source reconciliation: row arithmetic, key integrity and the dataset's own Control_Checks baselines.
The full 51-control workbook check, independent recompute and 68-test fault-injection suite run via run_mis.py (see Runbook_and_Automation_Boundaries.md).
"""
import hashlib, os, sys
import pandas as pd
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'data', 'source', 'Nova_CFO_MIS_Complete_Synthetic_Dataset.xlsx')
sys.path.insert(0, ROOT)   # v1 scripts stay at the repository root

def _sheets(): return pd.read_excel(SRC, sheet_name=None)

def test_source_unchanged():
    want = open(os.path.join(ROOT, 'data', 'source', 'SHA256SUMS')).read().split()[0]
    assert hashlib.sha256(open(SRC, 'rb').read()).hexdigest() == want

def test_selftest_record_current():
    import run_mis
    rec, problem = run_mis.load_selftest_record()
    if problem:
        files = {f: open(os.path.join(ROOT, f), 'rb').read() for f in run_mis.CODE_FILES}
        hint = {f: (run_mis.sha256(os.path.join(ROOT, f))[:12], (rec or {}).get('code_files', {}).get(f, '?')[:12], b.count(b'\r\n')) for f, b in files.items()}
        raise AssertionError(f'{problem}. (file: current sha12, recorded sha12, CRLF line count) = {hint}. A non-zero CRLF count means the checkout converted '
                             'line endings (git core.autocrlf on Windows): re-checkout with the repository .gitattributes, or run "git add --renormalize ." and re-checkout the .py files.')

def test_row_arithmetic_and_keys():
    x = _sheets(); s = x['Sales_Transactions']
    assert s.Invoice_ID.is_unique
    assert s.Product_ID.isin(x['Product_Master'].Product_ID).all() and s.Customer_ID.isin(x['Customer_Master'].Customer_ID).all()
    assert s.Month.isin(x['Calendar'].Month).all()
    assert ((s.Net_Revenue_INR - s.Units * s.Net_Price_INR).abs() < 0.05).all()
    assert ((s.COGS_INR - s.Units * s.Unit_Cost_INR).abs() < 0.05).all()
    assert ((s.Gross_Profit_INR - (s.Net_Revenue_INR - s.COGS_INR)).abs() < 0.05).all()

def test_dataset_control_baselines():
    x = _sheets(); s = x['Sales_Transactions']; a = s[s.Status == 'Actual']; cc = x['Control_Checks'].set_index('Control').Expected_Value
    got = {'Sales transactions': len(s), 'Unique invoice IDs': s.Invoice_ID.nunique(), 'Actual sales rows': len(a), 'Forecast sales rows': (s.Status == 'Forecast').sum(),
           'Actual net revenue INR': round(a.Net_Revenue_INR.sum(), 2), 'Actual COGS INR': round(a.COGS_INR.sum(), 2), 'Actual gross profit INR': round(a.Gross_Profit_INR.sum(), 2)}
    for k, v in got.items(): assert abs(float(cc[k]) - float(v)) < 0.01, (k, cc[k], v)

if __name__ == '__main__':
    for n, f in list(globals().items()):
        if n.startswith('test_'): f(); print('PASS', n)
