#!/usr/bin/env python3
"""Nova CFO MIS - controlled one-command workflow.

  python run_mis.py run            --input <dataset.xlsx> [--root DIR] [--commentary-file FILE]
  python run_mis.py approve        --run <run_id> --by "<name>" [--ack-warnings] [--external-ref TEXT] [--root DIR]
  python run_mis.py selftest       --input <known-good dataset.xlsx>
  python run_mis.py anchor         --run <run_id> [--root DIR]          (optional: prepare a record to store OUTSIDE the workspace)
  python run_mis.py verify-anchor  --run <run_id> --hash <externally stored hash> [--root DIR]
  python run_mis.py verify-release --release <published/<month>/<run_id>> [--root DIR]

run      : validate input -> build workbook -> recalculate (LibreOffice) -> independent recompute -> gates -> commentary -> audit log.
           Produces a CANDIDATE only. Nothing is published. Exit 0 = READY_FOR_APPROVAL, 3 = BLOCKED, 2 = input rejected.
           A missing, stale or failing self-test record (for the CURRENT code) is a critical gate failure.
approve  : human step. Re-verifies self-test record, audit chain, manifest, archived input and code hash, then copies the candidate into a NEW
           immutable folder published/<month>/<run_id>/ (never overwrites). --ack-warnings cannot override any critical condition.
Limits (read these): the audit log is a hash chain - it DETECTS EDITS to earlier lines but is NOT tamper-proof; it is only externally anchored if
YOU store the final hash outside the workspace. The approver name is typed text, not an authenticated identity. The script cannot stop someone
copying files by hand into published/ or elsewhere; that needs folder permissions or a signed-release process outside this script.
Not implemented: scheduling, unattended triggers, live data connections, any AI-model call, cryptographic signing.
"""
import argparse, datetime as dt, hashlib, json, os, platform, re, shutil, subprocess, sys, tempfile
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import numpy as np, openpyxl, pandas as pd
import mis_engine as E

OUT_NAME = 'CFO_Monthly_Sales_MIS.xlsx'
EXCEL_DESKTOP_TODO = [
    ('Open the candidate in Microsoft Excel Desktop (Windows or Mac)', 'File opens without a repair prompt or "update links" warning.'),
    ('Force a full recalculation (Ctrl+Alt+F9) and read Control_Checks!C3', 'Shows "ALL <n> CHECKS PASS"; no #NAME?/#VALUE!/#N/A anywhere.'),
    ('Compare Monthly_MIS totals and Executive_Summary KPIs with the run\'s validation_report.md', 'Values equal the LibreOffice-recalculated values to the rupee.'),
    ('Check month-filter formulas (SUMIFS on numeric Year/Month_No; LOOKUP(2,1/...) for latest month)', 'Monthly_MIS!B4 shows the latest actual month; monthly rows non-zero.'),
    ('Check Indian lakh/crore number format, conditional formatting (PASS green / FAIL red) and chart rendering', 'Formats and 5-6 charts display as intended.'),
    ('Check calculation time on a normal laptop', 'Full recalculation completes in a few seconds (SUMPRODUCT/COUNTIF arrays over ~900 rows).'),
]

# ----------------------------------------------------------------------------- helpers
def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for ch in iter(lambda: f.read(1 << 20), b''): h.update(ch)
    return h.hexdigest()

def utcnow(): return dt.datetime.now(dt.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')

CODE_FILES = ('mis_engine.py', 'run_mis.py')
SELFTEST_PATH = os.path.join(HERE, 'selftest_result.json')
MIN_SELFTEST_TESTS = 68
REQUIRED_TAGS = ('S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'F:', 'S:', 'R:', 'N:', 'O:', 'C:', 'X:', 'I1', 'I2', 'I3', 'I4', 'I5', 'I6', 'P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'P7',
                 'M1', 'M2', 'M3', 'M4', 'M5', 'A1', 'A2', 'A3', 'A4', 'A5', 'V1', 'V2', 'V3')

def code_hash():
    """Combined code hash = SHA-256 of the two text lines 'mis_engine.py:<sha256>' and 'run_mis.py:<sha256>' joined by a newline."""
    return hashlib.sha256('\n'.join(f'{f}:{sha256(os.path.join(HERE, f))}' for f in CODE_FILES).encode()).hexdigest()

def results_hash(results): return hashlib.sha256(json.dumps(results, sort_keys=True).encode()).hexdigest()

def make_record(results, fixture=False, code=None):
    n = sum(1 for r in results if r['passed'])
    return dict(ts=utcnow(), code_sha256=code or code_hash(), code_files={f: sha256(os.path.join(HERE, f)) for f in CODE_FILES}, fixture=fixture, total=len(results), passed=n,
                all_passed=(n == len(results)), results=results, results_sha256=results_hash(results))

def load_selftest_record(path=None, allow_fixture=False):
    """Return (record, problem). problem is None only if the record is present, genuine-looking, passes in full and matches the CURRENT combined code hash.
    The results hash and consistency checks detect casual hand edits of the JSON only; they are not a signature."""
    path = path or SELFTEST_PATH
    if not os.path.exists(path): return None, 'no self-test record found'
    try: r = json.load(open(path))
    except Exception as e: return None, f'self-test record unreadable: {e}'
    if r.get('fixture') and not allow_fixture: return r, 'self-test record is a test fixture, not a genuine record'
    if r.get('code_sha256') != code_hash(): return r, 'self-test record is STALE: run_mis.py/mis_engine.py changed since it was produced'
    if r.get('fixture'): return r, None   # throwaway fixture, only accepted when a path is passed explicitly (self-test workspace)
    res = r.get('results')
    if not isinstance(res, list) or len(res) < MIN_SELFTEST_TESTS: return r, f'self-test record has too few tests (need at least {MIN_SELFTEST_TESTS})'
    names = [str(t.get('test', '')) for t in res]
    miss = [t for t in REQUIRED_TAGS if not any(n.startswith(t) for n in names)]
    if miss: return r, f'self-test record lacks required test groups: {miss}'
    if r.get('results_sha256') != results_hash(res) or r.get('total') != len(res) or r.get('passed') != sum(1 for t in res if t.get('passed')): return r, 'self-test record is internally inconsistent (edited?)'
    if not r.get('all_passed') or not all(t.get('passed') for t in res): return r, f"self-test record shows failures ({r.get('passed')}/{r.get('total')} passed)"
    return r, None

class Audit:
    """Append-only hash-chained JSON-lines log. EDIT-DETECTING only: a changed or deleted earlier line breaks the chain. It is NOT tamper-proof - anyone who can
    write the workspace can rebuild the whole chain. It becomes tamper-evident only if the final hash is stored outside the workspace (see 'anchor')."""
    def __init__(self, path, run_id):
        self.path, self.run_id = path, run_id; self.prev = '0'*64; self.seq = 0
        if os.path.exists(path):
            ok, msg = verify_audit(path); assert ok, msg
            lines = [json.loads(l) for l in open(path) if l.strip()]; self.prev = lines[-1]['hash']; self.seq = lines[-1]['seq']+1
    def log(self, step, status, detail=None):
        rec = dict(seq=self.seq, ts=utcnow(), run_id=self.run_id, step=step, status=status, detail=detail or {}, prev=self.prev)
        rec['hash'] = hashlib.sha256(json.dumps(rec, sort_keys=True, default=str).encode()).hexdigest()
        with open(self.path, 'a') as f: f.write(json.dumps(rec, default=str)+'\n')
        self.prev = rec['hash']; self.seq += 1; return rec['hash']

def verify_audit(path):
    prev = '0'*64; n = 0
    try:
        for i, l in enumerate(open(path)):
            if not l.strip(): continue
            rec = json.loads(l); h = rec.pop('hash')
            if rec['prev'] != prev or rec['seq'] != i: return False, f'chain broken at line {i+1}'
            if hashlib.sha256(json.dumps(rec, sort_keys=True, default=str).encode()).hexdigest() != h: return False, f'content altered at line {i+1}'
            prev = h; n += 1
    except Exception as e: return False, f'unreadable audit log: {e}'
    if n == 0: return False, 'audit log is empty'
    return True, prev

def recalc(src, outdir):
    os.makedirs(outdir, exist_ok=True); prof = tempfile.mkdtemp(prefix='nova_lo_')
    try:
        subprocess.run(['soffice', f'-env:UserInstallation=file://{prof}', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', outdir, src],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=280, check=True)
    finally: shutil.rmtree(prof, ignore_errors=True)
    p = os.path.join(outdir, os.path.basename(src))
    if not os.path.exists(p): raise RuntimeError('recalculation produced no file')
    return p

def lo_version():
    try: return subprocess.run(['soffice', '--version'], capture_output=True, text=True, timeout=60).stdout.strip()
    except Exception: return 'unknown'

# ----------------------------------------------------------------------------- reading + independent verification
def read_controls(path):
    wb = openpyxl.load_workbook(path, data_only=True); cw = wb['Control_Checks']
    ctr = [dict(n=r[0], group=r[1], name=r[2], exp=r[4], act=r[5], res=r[8]) for r in cw.iter_rows(min_row=6, values_only=True) if isinstance(r[0], int)]
    obs = {r[0]: dict(area=r[1], text=r[2], value=r[4], status=r[8]) for r in cw.iter_rows(min_row=6, values_only=True) if isinstance(r[0], str) and re.fullmatch(r'O\d+', r[0])}
    return wb, cw['C3'].value, ctr, obs

def independent_recompute(src, wb, info):
    """Re-derive report figures from the source file with a separate pandas code path and compare with cached workbook values."""
    x = pd.read_excel(src, sheet_name=None)
    s = x['Sales_Transactions'].merge(x['Customer_Master'], on='Customer_ID', how='left').merge(x['Product_Master'], on='Product_ID', how='left')
    a = s[s.Status == 'Actual'].copy(); a['y'] = a.Month.str[:4].astype(int); a['m'] = a.Month.str[5:].astype(int)
    ly = int(a.y.max()); lm = int(a[a.y == ly].m.max()); cur = a[(a.y == ly) & (a.m <= lm)]; pri = a[(a.y == ly-1) & (a.m <= lm)]
    n = 0; bad = []
    def cmp(tag, got, exp, tol=0.01):
        nonlocal n; n += 1
        if got is None or isinstance(got, str) or abs(got-exp) > tol: bad.append((tag, got, exp))
    mm = a.groupby('Month')[['Units', 'Net_Revenue_INR', 'COGS_INR', 'Gross_Profit_INR']].sum(); cnt = a.groupby('Month').size(); ws = wb['Monthly_MIS']
    for i, mo in enumerate(info['months']):
        r = info['m0']+i; e = mm.loc[mo]
        assert ws.cell(r, 1).value == mo, 'month order mismatch'
        for c, ev, t in ((3, cnt[mo], .01), (4, e.Units, .01), (5, e.Net_Revenue_INR, .01), (6, e.COGS_INR, .01), (7, e.Gross_Profit_INR, .01), (8, e.Gross_Profit_INR/e.Net_Revenue_INR, 1e-9)): cmp(f'Monthly {mo} c{c}', ws.cell(r, c).value, ev, t)
        y, mn = int(mo[:4]), int(mo[5:]); py = f'{y-1}-{mn:02d}'
        if py in mm.index: cmp(f'Monthly {mo} YoY', ws.cell(r, 10).value, e.Net_Revenue_INR/mm.loc[py].Net_Revenue_INR-1, 1e-9)
    for sheet, key, c0 in (('Region_Performance', 'Region', 2), ('Product_Performance', 'Product_Name', 3)):
        w = wb[sheet]; g = a.groupby(key)[['Net_Revenue_INR', 'COGS_INR', 'Gross_Profit_INR']].sum(); gc = cur.groupby(key)[['Net_Revenue_INR', 'Gross_Profit_INR']].sum(); gp = pri.groupby(key)[['Net_Revenue_INR', 'Gross_Profit_INR']].sum()
        rows = info['RG'] if key == 'Region' else info['PR']
        for row in range(rows['r0'], rows['r1']+1):
            k = w.cell(row, 1).value; e = g.loc[k]
            exp = [(e.Net_Revenue_INR, .01), (e.COGS_INR, .01), (e.Gross_Profit_INR, .01), (e.Gross_Profit_INR/e.Net_Revenue_INR, 1e-9), (e.Net_Revenue_INR/g.Net_Revenue_INR.sum(), 1e-9),
                   (gc.loc[k].Net_Revenue_INR, .01), (gp.loc[k].Net_Revenue_INR, .01), (gc.loc[k].Net_Revenue_INR/gp.loc[k].Net_Revenue_INR-1, 1e-9),
                   (gc.loc[k].Gross_Profit_INR/gc.loc[k].Net_Revenue_INR, 1e-9), (gp.loc[k].Gross_Profit_INR/gp.loc[k].Net_Revenue_INR, 1e-9),
                   ((gc.loc[k].Gross_Profit_INR/gc.loc[k].Net_Revenue_INR-gp.loc[k].Gross_Profit_INR/gp.loc[k].Net_Revenue_INR)*100, 1e-7)]
            for j, (ev, t) in enumerate(exp): cmp(f'{sheet} {k} c{c0+j}', w.cell(row, c0+j).value, ev, t)
            if key == 'Product_Name': cmp(f'{sheet} {k} regions', w.cell(row, 14).value, a[a.Product_Name == k].Region.nunique(), 0)
    ps = info['PS']; w = wb['Monthly_MIS']
    for off, d_ in ((1, cur), (2, pri)):
        cmp(f'Period {off} rev', w.cell(ps+off, 5).value, d_.Net_Revenue_INR.sum()); cmp(f'Period {off} cogs', w.cell(ps+off, 6).value, d_.COGS_INR.sum())
    cmp('Full prior year rev', w.cell(ps+3, 5).value, a[a.y == ly-1].Net_Revenue_INR.sum())
    bd = info['bd']
    if bd and bd['v2']:
        b = x['Sales_Budget']; cov = set(zip(a.Region, a.Category)); b = b.assign(inc=[(r, c) in cov for r, c in zip(b.Region, b.Category)], y=b.Month.str[:4].astype(int), m=b.Month.str[5:].astype(int))
        cy = bd['calib_year']; k = a[a.y == cy].Net_Revenue_INR.sum()/b[(b.y == cy) & b.inc].Revenue_Budget_INR.sum()
        cmp('Budget k', wb['Budget_v2_Synthetic']['B7'].value, k, 1e-9)
        ba = wb['Budget_vs_Actual_v2']
        for i, mo in enumerate(info['months']):
            cmp(f'BvA {mo} budget', ba.cell(5+i, 2).value, k*b[(b.Month == mo) & b.inc].Revenue_Budget_INR.sum()); cmp(f'BvA {mo} actual', ba.cell(5+i, 3).value, mm.loc[mo].Net_Revenue_INR)
        yb = b[(b.y == ly) & (b.m <= lm) & b.inc & b.Month.isin(mm.index)].Revenue_Budget_INR.sum()*k
        cmp('BvA YTD budget', ba.cell(info['bva']['YTDB'], 2).value, yb); cmp('BvA YTD actual', ba.cell(info['bva']['YTDB'], 3).value, cur.Net_Revenue_INR.sum())
    return n, bad, dict(monthly={mo: [float(mm.loc[mo].Net_Revenue_INR), float(mm.loc[mo].COGS_INR)] for mo in mm.index}, latest=f'{ly}-{lm:02d}')

def error_scan(wb):
    return [(w.title, c.coordinate, c.value) for w in wb for row in w.iter_rows() for c in row if isinstance(c.value, str) and re.fullmatch(r'#(N/A|REF!|DIV/0!|VALUE!|NAME\?|NUM!|NULL!)', c.value)]

# ----------------------------------------------------------------------------- run
def run(inp, root, commentary_file=None, quiet=False, record_path=None):
    inp = os.path.abspath(inp); root = os.path.abspath(root)
    if not os.path.isfile(inp): print('ERROR: input not found', file=sys.stderr); return 2, None
    h_in = sha256(inp); run_id = dt.datetime.now(dt.timezone.utc).strftime('%Y%m%d-%H%M%S') + '_' + h_in[:8]
    rd = os.path.join(root, 'runs', run_id); os.makedirs(os.path.join(rd, 'input')); os.makedirs(os.path.join(rd, 'work')); os.makedirs(os.path.join(rd, 'audit'))
    audit = Audit(os.path.join(rd, 'audit', 'audit_log.jsonl'), run_id)
    ip = os.path.join(rd, 'input', os.path.basename(inp)); shutil.copy2(inp, ip); os.chmod(ip, 0o444)
    audit.log('run_started', 'OK', dict(input=os.path.basename(inp), input_sha256=h_in, input_bytes=os.path.getsize(inp), code_sha256=code_hash(), python=platform.python_version(),
                                        pandas=pd.__version__, openpyxl=openpyxl.__version__, libreoffice=lo_version()))
    gates = []; warnings = []
    def gate(name, ok, detail): gates.append(dict(name=name, ok=bool(ok), detail=detail)); audit.log('gate:'+name, 'PASS' if ok else 'FAIL', dict(detail=detail))
    def finish(status, extra=None):
        st = dict(run_id=run_id, status=status, gates=gates, warnings=warnings, input_sha256=h_in, code_sha256=code_hash(),
                  audit_integrity='hash chain (edit-detecting only; NOT tamper-proof)', audit_externally_anchored=False, approver_identity_authenticated=False, **(extra or {}))
        json.dump(st, open(os.path.join(rd, 'status.json'), 'w'), indent=1, default=str); return st
    # 1 preflight
    try: x = E.load(ip)
    except E.EngineError as e:
        gate('input_schema', False, str(e)); audit.log('run_finished', 'REJECTED_INPUT'); st = finish('REJECTED_INPUT')
        if not quiet: print(f'REJECTED_INPUT: {e}\nRun folder: {rd}')
        return 2, st
    except Exception as e:
        gate('input_schema', False, f'unreadable workbook: {e}'); audit.log('run_finished', 'REJECTED_INPUT'); st = finish('REJECTED_INPUT')
        if not quiet: print(f'REJECTED_INPUT: unreadable workbook: {e}')
        return 2, st
    gate('input_schema', True, 'all required sheets/columns present; Status values valid')
    rec_, prob_ = load_selftest_record(record_path, allow_fixture=record_path is not None)
    gate('selftest_record', prob_ is None, (f"valid for current code hash {code_hash()[:12]}...: {rec_['passed']}/{rec_['total']} tests passed at {rec_['ts']}" if prob_ is None else prob_ + ' - run "run_mis.py selftest" on the current code'))
    override = open(commentary_file).read() if commentary_file else None
    work = os.path.join(rd, 'work')
    # 2 pass 1 build + recalc + checks
    try:
        info = E.build(ip, os.path.join(work, 'p1.xlsx'), None, override)
    except E.EngineError as e:
        gate('commentary_valid', False, str(e)); audit.log('run_finished', 'BLOCKED'); st = finish('BLOCKED')
        if not quiet: print(f'BLOCKED: {e}')
        return 3, st
    except Exception as e:   # never crash on dirty data: block and record
        import traceback
        gate('engine_build', False, f'{type(e).__name__}: {e}'); audit.log('engine_exception', 'FAIL', dict(trace=traceback.format_exc()[-1500:])); audit.log('run_finished', 'BLOCKED'); st = finish('BLOCKED')
        if not quiet: print(f'BLOCKED: engine could not process the source ({type(e).__name__}: {e})')
        return 3, st
    gate('commentary_valid', True, f"{info['words']} words; numbers traceable to fact pack" + ('; AI/externally supplied commentary file used' if override else '; rule-based commentary'))
    if override: warnings.append('Commentary was supplied from an external file (checked for length and number traceability only, not for wording or causal claims).')
    try:
        rc1 = recalc(os.path.join(work, 'p1.xlsx'), os.path.join(work, 'rc1')); gate('recalc_completed', True, 'LibreOffice headless recalculation succeeded (pass 1)')
    except Exception as e:
        gate('recalc_completed', False, str(e)); audit.log('run_finished', 'BLOCKED'); st = finish('BLOCKED'); return 3, st
    wb1, status1, ctr, obs = read_controls(rc1)
    n_rc, bad, facts = independent_recompute(ip, wb1, info); errs = error_scan(wb1)
    npass = sum(1 for c in ctr if c['res'] == 'PASS'); fails = [c for c in ctr if c['res'] != 'PASS']
    # 3 validation records for the sheet
    val = [dict(test='Source inspection and schema check', method='All required sheets/columns validated; Status values validated; source file hashed before and after.', result='PASS', evidence=f'SHA-256 {h_in[:16]}...'),
           dict(test='Workbook control checks', method='Recalculated headlessly in LibreOffice; every Control_Checks row read back.', result='PASS' if not fails else 'FAIL', evidence=f'{status1}; PASS={npass}, FAIL={len(fails)} of {len(ctr)}.'),
           dict(test='Independent recompute of report figures', method='Re-derived Monthly_MIS, Region, Product, period summary' + (' and budget v2 variance' if info['bd'] and info['bd']['v2'] else '') + ' from the source file in a separate pandas code path; compared to workbook values.', result='PASS' if not bad else 'FAIL', evidence=f'{n_rc} figures compared; {len(bad)} differences.'),
           dict(test='Formula error scan', method='Scanned every cell of the recalculated workbook for Excel error values.', result='PASS' if not errs else 'FAIL', evidence=f'{len(errs)} error cells.'),
           dict(test='Pipeline self-test record (fault injection, gating, release controls)', method='Record bound to the combined SHA-256 of run_mis.py + mis_engine.py; must be current and fully passing. A stale or failing record blocks the run and blocks approval.',
                result='PASS' if prob_ is None else 'FAIL', evidence=(f"{rec_['passed']}/{rec_['total']} self-tests passed at {rec_['ts']}; code hash {rec_['code_sha256'][:16]}..." if prob_ is None else prob_))]
    for t, e in EXCEL_DESKTOP_TODO: val.append(dict(test='EXCEL DESKTOP REQUIRED: ' + t, method='Not performed by this pipeline (no Microsoft Excel available).', result='NOT DONE', evidence='Expected: ' + e))
    val.append(dict(test='Audit trail strength', method='Hash chain verified at run time.', result='LIMITED', evidence='Edit-detecting only. NOT externally anchored: no anchor record has been stored outside the workspace.'))
    val.append(dict(test='Not implemented', method='Scheduling, unattended execution, live data refresh, email distribution, AI-model call, authenticated approver identity, cryptographic release signing.', result='NOT DONE', evidence='Run is started manually; approval is a separate manual step; approver name is typed text.'))
    # 4 pass 2 (final) build embedding the validation sheet
    info2 = E.build(ip, os.path.join(work, OUT_NAME), val, override)
    rc2 = recalc(os.path.join(work, OUT_NAME), os.path.join(work, 'rc2')); wb2, status2, ctr2, obs2 = read_controls(rc2)
    n2, bad2, facts2 = independent_recompute(ip, wb2, info2); errs2 = error_scan(wb2)
    same = [(c['n'], c['res'], c['act']) for c in ctr] == [(c['n'], c['res'], c['act']) for c in ctr2]
    fails2 = [c for c in ctr2 if c['res'] != 'PASS']
    audit.log('recalc_and_checks', 'DONE', dict(controls=len(ctr2), control_fails=len(fails2), recompute_figures=n2, recompute_diffs=len(bad2), error_cells=len(errs2), pass1_equals_pass2=same))
    # 5 gates
    gate('workbook_controls', (not fails2) and len(ctr2) == info2['nchecks'] and status2 == f'ALL {len(ctr2)} CHECKS PASS', f'{status2}; {len(fails2)} failing: ' + '; '.join(f"#{c['n']} {c['name']}" for c in fails2[:8]))
    gate('independent_recompute', not bad2, f'{n2} figures compared, {len(bad2)} differences' + (f': {bad2[:3]}' if bad2 else ''))
    gate('error_scan', not errs2, f'{len(errs2)} error cells' + (f': {errs2[:3]}' if errs2 else ''))
    gate('pass1_equals_final', same, 'control results identical between build pass 1 and final build')
    gate('source_unchanged', sha256(inp) == h_in and sha256(ip) == h_in, 'source SHA-256 identical before and after the run (original and working copy)')
    ok_a, msg_a = verify_audit(audit.path); gate('audit_chain', ok_a, 'hash chain valid (edit-detecting only; not externally anchored)' if ok_a else msg_a)
    # 6 warnings
    for k_, o in obs2.items():
        if o['status'] == 'FLAG' and not (k_ == 'O1' and info2['bd']): warnings.append(f"{k_} {o['area']}: {o['text']} -> value {o['value']:.2f}" if isinstance(o['value'], (int, float)) else f"{k_} {o['area']}: {o['text']}")
    if info2['bd'] and info2['bd']['flag']: warnings.append(f"Budget integrity: original budget is {info2['bd']['ratio_all']:.1f}x actuals, {info2['bd']['unc_share']:.0%} in cells with no sales; root cause UNRESOLVED; only SYNTHETIC v2 (unapproved) variance is shown.")
    pub = os.path.join(root, 'published', 'LATEST.json')
    if os.path.exists(pub):
        prev = json.load(open(pub)); diffs = []
        for mo, (rv, cg) in prev.get('monthly', {}).items():
            if mo not in facts2['monthly']: diffs.append(f'{mo}: previously published actual month missing now')
            else:
                cr, cc = facts2['monthly'][mo]
                if abs(cr-rv) > 0.01 or abs(cc-cg) > 0.01: diffs.append(f'{mo}: revenue {rv:,.2f} -> {cr:,.2f}; COGS {cg:,.2f} -> {cc:,.2f}')
        if diffs: warnings.append('RESTATEMENT of previously published months: ' + ' | '.join(diffs[:6]) + (f' (+{len(diffs)-6} more)' if len(diffs) > 6 else ''))
        audit.log('restatement_check', 'WARN' if diffs else 'OK', dict(previous_run=prev.get('run_id'), differences=len(diffs)))
    else: audit.log('restatement_check', 'SKIPPED', dict(reason='no previously published run'))
    # 7 outputs
    critical_ok = all(g['ok'] for g in gates); status = 'READY_FOR_APPROVAL' if critical_ok else 'BLOCKED'
    outdir = os.path.join(rd, 'candidate' if critical_ok else 'blocked'); os.makedirs(outdir)
    fn = OUT_NAME if critical_ok else 'BLOCKED_' + OUT_NAME
    shutil.copy2(rc2, os.path.join(outdir, fn))
    obs_t, unv_t = info2['obs'], info2['unv']
    open(os.path.join(outdir, 'CFO_commentary.md'), 'w').write(f"# CFO commentary - actuals to {facts2['latest']}\n\n{obs_t}\n\n{unv_t}\n\n_{info2['words']} words. Generated from verified results; numbers checked against fact_pack.json. Status: {status}._\n")
    json.dump(info2['pack'], open(os.path.join(outdir, 'fact_pack.json'), 'w'), indent=1)
    bt = ''
    if info2['bd']:
        bd = info2['bd']; bt = f"\n## Budget integrity\nOriginal budget / actual = {bd['ratio_all']:.2f}x ({bd['ratio_in']:.2f}x on cells that have sales). Root cause: {bd['root']}\n"
    rep = [f'# Validation report - run {run_id}', f'Status: **{status}**  |  input SHA-256 `{h_in}`  |  code SHA-256 `{code_hash()[:16]}...`', '', '## Publication gates (critical)', '| Gate | Result | Detail |', '|---|---|---|']
    rep += [f"| {g['name']} | {'PASS' if g['ok'] else 'FAIL'} | {g['detail']} |" for g in gates]
    rep += ['', '## Workbook controls', f'{status2}. Groups: ' + ', '.join(f"{g}: {sum(1 for c in ctr2 if c['group']==g and c['res']=='PASS')}/{sum(1 for c in ctr2 if c['group']==g)}" for g in dict.fromkeys(c['group'] for c in ctr2))]
    rep += ['', '## Warnings (need acknowledgement at approval)'] + ([f'- {w}' for w in warnings] or ['- none']) + [bt]
    rep += ['## Tested in this run', '- Input schema, source immutability (hash before/after), recalculation in LibreOffice, all workbook controls, independent pandas recompute, formula-error scan, commentary length and number traceability, audit-chain integrity.',
            '', '## NOT tested - requires Microsoft Excel Desktop'] + [f'- [ ] {t} - expected: {e}' for t, e in EXCEL_DESKTOP_TODO]
    rep += ['', '## Audit trail and publication boundary',
            '- The audit log is a hash chain: it DETECTS EDITS to earlier lines. It is NOT tamper-proof and is NOT externally anchored. It becomes tamper-evident only if the final hash (see `run_mis.py anchor`) is stored outside this workspace; that has NOT been done for this run.',
            '- `approve` verifies the self-test record, audit chain, candidate manifest, archived input hash and code hash, then copies to a new folder `published/<month>/<run_id>/`. It never overwrites a release.',
            '- The script cannot prevent files being copied by hand into `published/` or elsewhere. That needs folder permissions or a signed-release process outside this script. `verify-release` can detect later edits to a release made through this workflow, not a wholly forged one.',
            '- The approver name is typed text, not an authenticated identity.',
            '', '## Not implemented', '- Scheduling, unattended execution, live refresh, distribution, AI-model call (commentary is rule-based unless a commentary file is supplied), authenticated approver identity, cryptographic signing, external anchoring.']
    open(os.path.join(outdir, 'validation_report.md'), 'w').write('\n'.join(rep)+'\n')
    manifest = dict(run_id=run_id, created=utcnow(), status=status, input=dict(name=os.path.basename(inp), sha256=h_in), code_sha256=code_hash(),
                    selftest_record=dict(code_sha256=rec_.get('code_sha256'), passed=rec_.get('passed'), total=rec_.get('total'), ts=rec_.get('ts'), results_sha256=rec_.get('results_sha256'), valid=prob_ is None) if rec_ else None,
                    files={f: sha256(os.path.join(outdir, f)) for f in sorted(os.listdir(outdir))}, monthly_facts=facts2)
    json.dump(manifest, open(os.path.join(outdir, 'manifest.json'), 'w'), indent=1)
    audit.log('outputs_written', 'OK', dict(folder=os.path.basename(outdir), files=manifest['files'], manifest_sha256=sha256(os.path.join(outdir, 'manifest.json'))))
    audit.log('run_finished', status, dict(warnings=len(warnings)))
    st = finish(status, dict(output_dir=outdir, workbook=os.path.join(outdir, fn), warnings_count=len(warnings), failed_controls=[c['name'] for c in fails2], recompute_differences=len(bad2), error_cells=len(errs2)))
    shutil.rmtree(work, ignore_errors=True)
    if not quiet:
        print(f'{status}  run={run_id}'); print(f'  critical gates: {sum(g["ok"] for g in gates)}/{len(gates)} pass; warnings: {len(warnings)}')
        for g in gates:
            if not g['ok']: print('  FAILED GATE:', g['name'], '-', g['detail'][:200])
        for w in warnings: print('  WARNING:', w[:220])
        print(f'  folder: {outdir}')
        print(('  Next (human step, NOT done): review validation_report.md, then: python run_mis.py approve --run ' + run_id + ' --by "<your name>"' + (' --ack-warnings' if warnings else '')) if critical_ok else '  Publication is blocked. Fix the source or investigate the failed gates, then rerun.')
    return (0 if critical_ok else 3), st

# ----------------------------------------------------------------------------- approve / release controls
def _entries(path): return [json.loads(l) for l in open(path) if l.strip()]

def _rmtree_force(d):
    for r_, ds, fs in os.walk(d):
        for n in ds + fs:
            try: os.chmod(os.path.join(r_, n), 0o755 if n in ds else 0o644)
            except OSError: pass
    shutil.rmtree(d, ignore_errors=True)

def approve(run_id, by, ack, root, quiet=False, record_path=None, code_sha_override=None, external_ref=None):
    """Human approval step. Every condition below is a hard refusal that --ack-warnings can NOT override; only the run's WARNINGS can be acknowledged.
    record_path / code_sha_override are Python-only test hooks and are not exposed on the command line."""
    root = os.path.abspath(root)
    def refuse(msg):
        if not quiet: print('REFUSED:', msg, file=sys.stderr)
        return 4, msg
    if not by or not by.strip(): return refuse('--by <approver name> is required')
    if not run_id or run_id != os.path.basename(run_id) or run_id in ('.', '..'): return refuse('invalid run id')
    rd = os.path.join(root, 'runs', run_id)
    if not os.path.isdir(rd): return refuse('run not found')
    try: st = json.load(open(os.path.join(rd, 'status.json')))
    except Exception as e: return refuse(f'status.json unreadable: {e}')
    if st.get('status') != 'READY_FOR_APPROVAL': return refuse(f"run status is {st.get('status')}; only READY_FOR_APPROVAL runs can be approved")
    rec_, prob_ = load_selftest_record(record_path, allow_fixture=record_path is not None)
    if prob_: return refuse(f'CRITICAL - self-test record invalid for the current code: {prob_}. --ack-warnings cannot override this. Run "run_mis.py selftest", then re-run the MIS.')
    ap_ = os.path.join(rd, 'audit', 'audit_log.jsonl')
    ok, msg = verify_audit(ap_)
    if not ok: return refuse('audit log integrity check failed: ' + msg)
    cd = os.path.join(rd, 'candidate'); mp = os.path.join(cd, 'manifest.json')
    if not os.path.isfile(mp): return refuse('no candidate manifest (this run has no candidate)')
    try:
        man = json.load(open(mp)); ents = _entries(ap_)
        started = next(e for e in ents if e['step'] == 'run_started'); written = [e for e in ents if e['step'] == 'outputs_written'][-1]; fin = [e for e in ents if e['step'] == 'run_finished'][-1]
    except Exception as e: return refuse(f'manifest or audit log incomplete/unreadable: {type(e).__name__}')
    if sha256(mp) != written['detail'].get('manifest_sha256'): return refuse('candidate manifest.json does not match the manifest hash recorded in the audit log (manifest edited?)')
    if written['detail'].get('files') != man.get('files'): return refuse('manifest file list differs from the audit log record')
    for f, h in man['files'].items():
        fp = os.path.join(cd, f)
        if not os.path.isfile(fp) or sha256(fp) != h: return refuse(f'candidate file {f} changed since the run (hash mismatch)')
    extra = sorted(set(os.listdir(cd)) - set(man['files']) - {'manifest.json'})
    if extra: return refuse(f'candidate folder contains files not in the manifest: {extra}')
    arch = os.path.join(rd, 'input', man['input']['name'])
    if not os.path.isfile(arch) or sha256(arch) != man['input']['sha256'] or man['input']['sha256'] != started['detail'].get('input_sha256'): return refuse('archived input does not match the input hash recorded at run time')
    cur = code_sha_override or code_hash()
    if man.get('code_sha256') != started['detail'].get('code_sha256') or man.get('code_sha256') != cur: return refuse('code hash mismatch: the code that built this candidate is not the code now installed (re-run the MIS)')
    if man.get('status') != 'READY_FOR_APPROVAL' or fin.get('status') != 'READY_FOR_APPROVAL': return refuse(f"audit log / manifest record this run as {fin.get('status')}/{man.get('status')}, not READY_FOR_APPROVAL")
    if not (man.get('selftest_record') or {}).get('valid'): return refuse('run was not built under a valid self-test record')
    if man.get('run_id') != run_id: return refuse('manifest run id does not match the run folder')
    latest = man['monthly_facts']['latest']; pdir = os.path.join(root, 'published', latest); rel = os.path.join(pdir, run_id)
    if os.path.exists(rel): return refuse(f'release folder already exists and will never be overwritten: {rel}')
    if st.get('warnings') and not ack: return refuse('run has warnings; review them and re-run approve with --ack-warnings:\n  - ' + '\n  - '.join(st['warnings']))
    audit = Audit(ap_, run_id); hash_before = audit.prev
    os.makedirs(pdir, exist_ok=True)
    try:
        os.mkdir(rel)                                               # fails if it exists -> never overwrites
    except FileExistsError: return refuse('release folder already exists')
    LATEST = os.path.join(root, 'published', 'LATEST.json'); latest_written = False
    try:
        copied = {}
        for f in list(man['files']) + ['manifest.json']:
            shutil.copy2(os.path.join(cd, f), os.path.join(rel, f)); copied[f] = sha256(os.path.join(rel, f))
        for f, h in man['files'].items():
            if copied[f] != h: raise RuntimeError(f'copy of {f} does not match manifest hash')
        if copied['manifest.json'] != sha256(mp): raise RuntimeError('copy of manifest.json differs')
        appr = dict(run_id=run_id, approved_by=by.strip(), approved_at=utcnow(), external_reference=external_ref, warnings_acknowledged=st['warnings'] if ack else [],
                    files=copied, input_sha256=man['input']['sha256'], code_sha256=man['code_sha256'], selftest_record=man.get('selftest_record'), audit_hash_before_approval=hash_before, month=latest,
                    note='approver name is self-declared text, not an authenticated identity; this record is not cryptographically signed')
        ap_json = os.path.join(rel, 'APPROVAL.json'); json.dump(appr, open(ap_json, 'w'), indent=1); appr_sha = sha256(ap_json)
        for n in os.listdir(rel):
            try: os.chmod(os.path.join(rel, n), 0o444)             # best effort; NOT a security control
            except OSError: pass
        audit.log('approved_and_published', 'OK', dict(by=by.strip(), acknowledged_warnings=bool(ack), release=os.path.relpath(rel, root), approval_sha256=appr_sha, external_reference=external_ref))
        Audit(os.path.join(root, 'published', 'release_register.jsonl'), 'register').log('release', 'OK', dict(run_id=run_id, month=latest, release=os.path.relpath(rel, root), approval_sha256=appr_sha, manifest_sha256=copied['manifest.json'], approved_by=by.strip()))
        tmp_ = LATEST + '.tmp'
        json.dump(dict(run_id=run_id, latest=latest, monthly=man['monthly_facts']['monthly'], approved_by=by.strip(), approved_at=appr['approved_at'], release=os.path.relpath(rel, root), approval_sha256=appr_sha), open(tmp_, 'w'), indent=1)
        os.replace(tmp_, LATEST); latest_written = True             # LATEST moves only after everything else succeeded
    except Exception as e:
        if not latest_written: _rmtree_force(rel)
        return refuse(f'publication failed and was rolled back: {type(e).__name__}: {e}')
    st['status'] = 'PUBLISHED'; json.dump(st, open(os.path.join(rd, 'status.json'), 'w'), indent=1)
    if not quiet: print(f'PUBLISHED (new immutable release) {rel}\n  approved by "{by.strip()}" (typed text, not authenticated). Prior releases are kept. Audit log is edit-detecting only and NOT externally anchored: see "run_mis.py anchor".')
    return 0, rel

def anchor(run_id, root, quiet=False):
    """Optional: write a plain-text record of the run's current final audit hash for the user to store OUTSIDE the workspace. The script cannot do the external storing."""
    root = os.path.abspath(root); rd = os.path.join(root, 'runs', run_id); ap_ = os.path.join(rd, 'audit', 'audit_log.jsonl')
    if not os.path.isfile(ap_): print('REFUSED: run/audit log not found', file=sys.stderr); return 4, None
    ok, h = verify_audit(ap_)
    if not ok: print('REFUSED: audit chain invalid: ' + h, file=sys.stderr); return 4, h
    ents = _entries(ap_); started = ents[0]['detail']; written = [e for e in ents if e['step'] == 'outputs_written']
    fp = os.path.join(rd, 'audit', f'anchor_record_{len(ents)}_entries.txt')
    txt = '\n'.join(['NOVA CFO MIS - AUDIT ANCHOR RECORD', f'run_id: {run_id}', f'created_utc: {utcnow()}', f'audit_entries: {len(ents)}', f'final_audit_hash: {h}',
                     f"input_sha256: {started.get('input_sha256')}", f"code_sha256: {started.get('code_sha256')}", f"manifest_sha256: {written[-1]['detail'].get('manifest_sha256') if written else 'n/a'}", f"last_step: {ents[-1]['step']} ({ents[-1]['status']})", '',
                     'STATUS: NOT ANCHORED. This file is inside the workspace and proves nothing by itself.',
                     'It becomes an anchor only when YOU store final_audit_hash somewhere the workspace owner cannot silently edit (e.g. email to the approver/auditor, a ticket, a write-once store,',
                     'a signed message). The script cannot verify that you did. Check later with: run_mis.py verify-anchor --run ' + run_id + ' --hash <stored value>', ''])
    open(fp, 'w').write(txt)
    if not quiet: print(txt + f'Written to {fp}')
    return 0, h

def verify_anchor(run_id, ext_hash, root, quiet=False):
    root = os.path.abspath(root); ap_ = os.path.join(root, 'runs', run_id, 'audit', 'audit_log.jsonl')
    def out(c, m):
        if not quiet: print(('OK: ' if c == 0 else 'FAILED: ') + m)
        return c, m
    if not os.path.isfile(ap_): return out(4, 'audit log not found')
    ok, h = verify_audit(ap_)
    if not ok: return out(5, 'audit chain is internally broken: ' + h)
    ents = _entries(ap_); idx = [e['seq'] for e in ents if e['hash'] == ext_hash.strip().lower()]
    if not idx: return out(5, 'the externally stored hash does not appear anywhere in this log (log rewritten or wrong run/hash)')
    extra = len(ents) - 1 - idx[0]
    return out(0, f'externally stored hash found at entry {idx[0]}; {"log is unchanged since the anchor" if extra == 0 else f"{extra} later entries were appended after the anchor (expected if approval followed)"}. This proves the earlier part was not rewritten only if that hash really was stored outside the workspace.')

def verify_release(rel, root, quiet=False):
    """Recompute a release folder's file hashes against its APPROVAL.json and the release register. Detects later edits to a release made through this workflow;
    it cannot detect a wholly forged release + register unless the register or an anchor is held outside the workspace."""
    rel = os.path.abspath(rel); root = os.path.abspath(root)
    def out(c, m):
        if not quiet: print(('OK: ' if c == 0 else 'FAILED: ') + m)
        return c, m
    apf = os.path.join(rel, 'APPROVAL.json')
    if not os.path.isfile(apf): return out(5, 'no APPROVAL.json: not a release produced by "approve"')
    try: appr = json.load(open(apf))
    except Exception: return out(5, 'APPROVAL.json unreadable')
    for f, h in appr.get('files', {}).items():
        fp = os.path.join(rel, f)
        if not os.path.isfile(fp) or sha256(fp) != h: return out(5, f'release file {f} missing or modified since approval')
    extra = sorted(set(os.listdir(rel)) - set(appr.get('files', {})) - {'APPROVAL.json'})
    if extra: return out(5, f'release contains unlisted files: {extra}')
    reg = os.path.join(root, 'published', 'release_register.jsonl')
    if not os.path.isfile(reg): return out(5, 'no release register found')
    ok, h = verify_audit(reg)
    if not ok: return out(5, 'release register chain broken: ' + h)
    m = [e for e in _entries(reg) if e['detail'].get('run_id') == appr.get('run_id')]
    if not m: return out(5, 'release is not recorded in the release register (manually copied?)')
    if m[-1]['detail'].get('approval_sha256') != sha256(apf): return out(5, 'APPROVAL.json differs from the version recorded in the register')
    if os.path.relpath(rel, root) != m[-1]['detail'].get('release'): return out(5, 'release folder location differs from the register')
    return out(0, f"release {appr['run_id']} matches APPROVAL.json and the register (approved by '{appr.get('approved_by')}', typed text). Not externally anchored unless you stored the register/log hash elsewhere.")

# ----------------------------------------------------------------------------- selftest
def _write_variant(src, dst, mutate):
    x = pd.read_excel(src, sheet_name=None); r_ = mutate(x); x = r_ if isinstance(r_, dict) else x
    with pd.ExcelWriter(dst) as w:
        for k, v in x.items(): v.to_excel(w, sheet_name=k, index=False)

def _refresh_baselines(x):
    s = x['Sales_Transactions']; a = s[s.Status == 'Actual']; cc = x['Control_Checks'].set_index('Control')
    vals = {'Sales transactions': len(s), 'Unique invoice IDs': s.Invoice_ID.nunique(), 'Actual sales rows': len(a), 'Forecast sales rows': (s.Status == 'Forecast').sum(),
            'Actual net revenue INR': round(a.Net_Revenue_INR.sum(), 2), 'Actual COGS INR': round(a.COGS_INR.sum(), 2), 'Actual gross profit INR': round(a.Gross_Profit_INR.sum(), 2)}
    for k, v in vals.items(): cc.loc[k, 'Expected_Value'] = v
    x['Control_Checks'] = cc.reset_index()

def selftest(src, root):
    src = os.path.abspath(src); root = os.path.abspath(root); h0 = sha256(src); results = []; tmp = tempfile.mkdtemp(prefix='nova_selftest_')
    FX = os.path.join(tmp, 'fixture_record.json')
    json.dump(make_record([dict(test='fixture', passed=True, evidence='throwaway record used only inside the self-test workspace')], fixture=True), open(FX, 'w'))
    def _run(i, r, cf=None, quiet=False, record_path=FX): return run(i, r, cf, quiet=quiet, record_path=record_path)
    def _approve(rid, by, ack, r, quiet=False, **k):
        k.setdefault('record_path', FX); return approve(rid, by, ack, r, quiet=quiet, **k)
    def rec(name, cond, evidence): results.append(dict(test=name, passed=bool(cond), evidence=evidence)); print(('PASS ' if cond else 'FAIL ') + name + ' | ' + evidence, flush=True)
    def variant(name, mut):
        p = os.path.join(tmp, name + '.xlsx'); _write_variant(src, p, mut); return p
    def failed_names(st): return ' | '.join([g['detail'] for g in st['gates'] if not g['ok']] + st.get('failed_controls', []))
    # S1 baseline
    code, st = _run(src, root, quiet=True); runs_base = st['run_id']
    rec('S1 baseline run on the original dataset', code == 0 and st['status'] == 'READY_FOR_APPROVAL', f"status={st['status']}; warnings={len(st['warnings'])}; gates {sum(g['ok'] for g in st['gates'])}/{len(st['gates'])}")
    c, msg = _approve(runs_base, 'selftest', False, root, quiet=True)
    rec('S2 approval refused while warnings unacknowledged', c == 4, 'refused: ' + str(msg)[:80])
    c, msg = _approve(runs_base, '', True, root, quiet=True); rec('S3 approval refused without approver name', c == 4, str(msg))
    # tamper tests on a copy of the run folder
    shutil.copytree(os.path.join(root, 'runs', runs_base), os.path.join(root, 'runs', 'tamper_audit')); shutil.copytree(os.path.join(root, 'runs', runs_base), os.path.join(root, 'runs', 'tamper_file'))
    ap = os.path.join(root, 'runs', 'tamper_audit', 'audit', 'audit_log.jsonl'); L_ = open(ap).read().replace('"PASS"', '"FAIL"', 1); open(ap, 'w').write(L_)
    c, msg = _approve('tamper_audit', 'selftest', True, root, quiet=True); rec('S4 audit-log edit detected, approval refused', c == 4 and 'audit' in msg, str(msg)[:90])
    wp = os.path.join(root, 'runs', 'tamper_file', 'candidate', OUT_NAME); open(wp, 'ab').write(b'x')
    c, msg = _approve('tamper_file', 'selftest', True, root, quiet=True); rec('S5 candidate file edit detected, approval refused', c == 4 and 'hash' in msg, str(msg)[:90])
    c, pubdir = _approve(runs_base, 'selftest', True, root, quiet=True); rec('S6 approval with acknowledgement publishes and writes APPROVAL.json', c == 0 and os.path.exists(os.path.join(str(pubdir), 'APPROVAL.json')), str(pubdir)[-60:])
    # fault injection, one defect per run
    def setv(x, idx, col, val, sheet='Sales_Transactions'): x[sheet].loc[idx, col] = val
    def addv(x, idx, col, d, sheet='Sales_Transactions'): x[sheet].loc[idx, col] = x[sheet].loc[idx, col] + d
    faults = {
      'duplicate Invoice_ID': (lambda x: setv(x, 5, 'Invoice_ID', x['Sales_Transactions'].loc[4, 'Invoice_ID']), 'Duplicate Invoice_ID'),
      'unmatched Product_ID': (lambda x: setv(x, 10, 'Product_ID', 'P999'), 'Unmatched Product_ID'),
      'unmatched Customer_ID': (lambda x: setv(x, 20, 'Customer_ID', 'C999'), 'Unmatched Customer_ID'),
      'revenue off by INR 500': (lambda x: addv(x, 30, 'Net_Revenue_INR', 500.0), 'Net_Revenue <> Units'),
      'COGS off by INR 250': (lambda x: addv(x, 31, 'COGS_INR', 250.0), 'COGS <> Units'),
      'gross profit off by INR 100': (lambda x: addv(x, 32, 'Gross_Profit_INR', 100.0), 'Gross_Profit <>'),
      'Forecast row relabelled Actual': (lambda x: setv(x, 800, 'Status', 'Actual'), 'Transaction Status vs Calendar'),
      'month missing from Calendar': (lambda x: x.__setitem__('Calendar', x['Calendar'][x['Calendar'].Month != '2025-03'].copy()), 'Unmatched Month'),
      'negative units (self-consistent row)': (None, 'Non-positive'),
      'blank Net_Price cell': (lambda x: setv(x, 40, 'Net_Price_INR', np.nan), 'Blank cells'),
      'duplicate budget key': (lambda x: x.__setitem__('Sales_Budget', pd.concat([x['Sales_Budget'], x['Sales_Budget'].iloc[[3]]], ignore_index=True)), 'Duplicate Month|Region|Category'),
    }
    def neg(x):
        s = x['Sales_Transactions']; i = 50; s.loc[i, 'Units'] = -5; s.loc[i, 'Net_Revenue_INR'] = round(-5*s.loc[i, 'Net_Price_INR'], 2); s.loc[i, 'COGS_INR'] = round(-5*s.loc[i, 'Unit_Cost_INR'], 2); s.loc[i, 'Gross_Profit_INR'] = round(s.loc[i, 'Net_Revenue_INR']-s.loc[i, 'COGS_INR'], 2)
    faults['negative units (self-consistent row)'] = (neg, 'Non-positive')
    for nm, (fn, key) in faults.items():
        p = variant('f_' + nm.split()[0] + str(abs(hash(nm)) % 1000), (lambda x, fn=fn: (fn(x), x)[1]))
        code, st = _run(p, root, quiet=True); det = failed_names(st) if st else ''
        rec(f'F: {nm} -> BLOCKED with specific control', code == 3 and st['status'] == 'BLOCKED' and key in det and not os.path.exists(os.path.join(root, 'runs', st['run_id'], 'candidate')), f"status={st['status']}; detects '{key}': {key in det}; no candidate folder: {not os.path.exists(os.path.join(root, 'runs', st['run_id'], 'candidate'))}")
        c, msg = _approve(st['run_id'], 'selftest', True, root, quiet=True); rec(f'F: {nm} -> approval refused', c == 4, str(msg)[:70])
    # schema failure
    p = variant('schema', lambda x: x.__setitem__('Customer_Master', x['Customer_Master'].drop(columns=['Region'])))
    code, st = _run(p, root, quiet=True); rec('S: missing column -> REJECTED_INPUT, no workbook built', code == 2 and st['status'] == 'REJECTED_INPUT' and not any(f.endswith('.xlsx') for _, _, fs in os.walk(os.path.join(root, 'runs', st['run_id'])) for f in fs if f != os.path.basename(p)), st['gates'][0]['detail'][:90])
    p = variant('badstatus', lambda x: setv(x, 3, 'Status', 'Draft'))
    code, st = _run(p, root, quiet=True); rec('S: unexpected Status value -> REJECTED_INPUT', code == 2, st['gates'][0]['detail'][:90])
    # restatement
    def restate(x):
        s = x['Sales_Transactions']; m = s.Month == '2025-03'
        s.loc[m, 'Units'] = (s.loc[m, 'Units']*1.1).round(0); s.loc[m, 'Net_Revenue_INR'] = (s.loc[m, 'Units']*s.loc[m, 'Net_Price_INR']).round(2); s.loc[m, 'COGS_INR'] = (s.loc[m, 'Units']*s.loc[m, 'Unit_Cost_INR']).round(2)
        s.loc[m, 'Gross_Profit_INR'] = (s.loc[m, 'Net_Revenue_INR']-s.loc[m, 'COGS_INR']).round(2); _refresh_baselines(x)
    p = variant('restate', restate); code, st = _run(p, root, quiet=True)
    rec('R: internally consistent restatement of a published month -> READY but RESTATEMENT warning', code == 0 and any('RESTATEMENT' in w for w in st['warnings']), f"status={st['status']}; restatement warning: {any('RESTATEMENT' in w for w in st['warnings'])}")
    c, msg = _approve(st['run_id'], 'selftest', False, root, quiet=True); rec('R: restated run cannot be approved without acknowledgement', c == 4, str(msg)[:70])
    # next month simulation
    def nextm(x):
        c_ = x['Calendar']; nxt = c_[c_.Reporting_Status == 'Forecast'].Month.iloc[0]; c_.loc[c_.Month == nxt, 'Reporting_Status'] = 'Actual'; s = x['Sales_Transactions']; s.loc[s.Month == nxt, 'Status'] = 'Actual'
    def nextm_fresh(x): nextm(x); _refresh_baselines(x)
    p = variant('next_stale', nextm); code, st = _run(p, root, quiet=True)
    rec('N: next-month simulation with STALE source baselines -> BLOCKED', code == 3 and 'baseline' in failed_names(st).lower(), f"status={st['status']}")
    p = variant('next_ok', nextm_fresh); code, st = _run(p, root, quiet=True)
    rec('N: next-month simulation with refreshed baselines -> READY, 22 months, no restatement', code == 0 and not any('RESTATEMENT' in w for w in st['warnings']), f"status={st['status']}; warnings={[w[:40] for w in st['warnings']]}")
    if st['status'] == 'READY_FOR_APPROVAL':
        wbn = openpyxl.load_workbook(os.path.join(st['output_dir'], OUT_NAME), data_only=True); rec('N: simulation workbook lists 22 actual months', sum(1 for r in range(9, 40) if str(wbn['Monthly_MIS'].cell(r, 1).value or '')[:2] == '20') == 22, 'Monthly_MIS months counted')
    # optional budget sheet absent
    p = variant('nobudget', lambda x: (x.pop('Sales_Budget'), None)[1]); code, st = _run(p, root, quiet=True)
    rec('O: dataset without Sales_Budget -> READY, no budget sheets', code == 0 and st['status'] == 'READY_FOR_APPROVAL', f"status={st['status']}")
    # commentary override
    base = open(os.path.join(root, 'runs', runs_base, 'candidate', 'CFO_commentary.md')).read().split('\n\n'); obs_t, unv_t = base[1], base[2]
    p_bad = os.path.join(tmp, 'c_bad.txt'); open(p_bad, 'w').write(obs_t.replace('35.7%', '41.3%', 1) + '\n' + unv_t); code, st = _run(src, root, p_bad, quiet=True)
    rec('C: commentary with an invented number -> BLOCKED', code == 3 and 'fact pack' in failed_names(st), failed_names(st)[:110])
    p_long = os.path.join(tmp, 'c_long.txt'); open(p_long, 'w').write(obs_t + ' ' + ' '.join(['Filler']*60) + '\n' + unv_t); code, st = _run(src, root, p_long, quiet=True)
    rec('C: commentary over 150 words -> BLOCKED', code == 3 and 'words' in failed_names(st), failed_names(st)[:110])
    p_nu = os.path.join(tmp, 'c_nounv.txt'); open(p_nu, 'w').write(obs_t); code, st = _run(src, root, p_nu, quiet=True); rec('C: commentary without UNVERIFIED section -> BLOCKED', code == 3, failed_names(st)[:110])
    p_ok = os.path.join(tmp, 'c_ok.txt'); open(p_ok, 'w').write(obs_t.replace('OBSERVED', 'OBSERVED (supplied)', 1) + '\n' + unv_t); code, st = _run(src, root, p_ok, quiet=True)
    rec('C: valid externally supplied commentary -> READY with disclosure warning', code == 0 and any('external' in w.lower() for w in st['warnings']), f"status={st['status']}")
    # ---------------- release-control tests (all inside this throwaway workspace; nothing here touches your real workspace)
    import stat
    def synth(nfail=0, code=None):
        names = list(REQUIRED_TAGS) + [f'filler {i}' for i in range(MIN_SELFTEST_TESTS + 5 - len(REQUIRED_TAGS))]
        rs = [dict(test=n + ' synthetic', passed=True, evidence='synthetic') for n in names]
        for i in range(nfail): rs[i]['passed'] = False
        return make_record(rs, code=code)
    def wrec(name, r): pth = os.path.join(tmp, name); json.dump(r, open(pth, 'w'), indent=1); return pth
    pub = os.path.join(root, 'published'); LAT = os.path.join(pub, 'LATEST.json')
    # I: self-test record gate
    rp = wrec('stale.json', synth(code='0'*64)); code, st = _run(src, root, quiet=True, record_path=rp)
    rec('I1 stale self-test record (code hash differs) -> run BLOCKED by selftest_record gate', code == 3 and 'STALE' in failed_names(st) and st['status'] == 'BLOCKED', failed_names(st)[:100])
    rp_fail = wrec('failing.json', synth(nfail=2)); code, st = _run(src, root, quiet=True, record_path=rp_fail)
    rec('I2 failing self-test record for current code -> run BLOCKED', code == 3 and 'failures' in failed_names(st), failed_names(st)[:100])
    code, st = _run(src, root, quiet=True, record_path=os.path.join(tmp, 'nope.json'))
    rec('I3 missing self-test record -> run BLOCKED', code == 3 and 'no self-test record' in failed_names(st), failed_names(st)[:100])
    cheat = json.load(open(rp_fail)); cheat['all_passed'] = True; cheat['passed'] = cheat['total']
    for r_ in cheat['results']: r_['passed'] = True
    rp_cheat = wrec('cheat.json', cheat); code, st = _run(src, root, quiet=True, record_path=rp_cheat)
    rec('I4 hand-edited failing record (results hash no longer matches) -> run BLOCKED', code == 3 and 'inconsistent' in failed_names(st), failed_names(st)[:100])
    code, st = _run(src, root, quiet=True)
    ready_run = st['run_id']; rec('I5 (control) valid fixture record -> READY (so the I1-I4 blocks are caused by the record)', code == 0, st['status'])
    lat_before = open(LAT, 'rb').read()
    c1, m1 = _approve(ready_run, 'selftest', True, root, quiet=True, record_path=rp); c2, m2 = _approve(ready_run, 'selftest', True, root, quiet=True, record_path=rp_fail)
    c3, m3 = _approve(ready_run, 'selftest', True, root, quiet=True, record_path=os.path.join(tmp, 'nope.json'))
    rec('I5 approval refused with --ack-warnings when record is stale / failing / missing', c1 == c2 == c3 == 4 and all('self-test' in m for m in (m1, m2, m3)) and 'cannot override' in m1, m1[:70])
    global SELFTEST_PATH
    saved = SELFTEST_PATH
    try:
        SELFTEST_PATH = os.path.join(tmp, 'default_missing.json')
        c4, m4 = approve(ready_run, 'selftest', True, root, quiet=True); code5, st5 = run(src, root, quiet=True)
        SELFTEST_PATH = FX; c6, m6 = approve(ready_run, 'selftest', True, root, quiet=True); code7, st7 = run(src, root, quiet=True)
    finally: SELFTEST_PATH = saved
    r_fx, p_fx = load_selftest_record(FX, allow_fixture=False)
    rec('I6 default record path: missing or fixture record -> approve refused and run BLOCKED; fixture rejected as genuine', c4 == 4 and c6 == 4 and code5 == 3 and code7 == 3 and p_fx and 'fixture' in p_fx, f'{m4[:40]} | {m6[:50]}')
    # P: immutable releases
    man0 = json.load(open(os.path.join(str(pubdir), 'manifest.json'))); h_first = {f: sha256(os.path.join(str(pubdir), f)) for f in os.listdir(str(pubdir))}
    latest_m = man0['monthly_facts']['latest']
    rec('P1 first release is published/<latest-month>/<run-id>/ with workbook, manifest and APPROVAL.json', str(pubdir) == os.path.join(pub, latest_m, runs_base) and all(os.path.exists(os.path.join(str(pubdir), f)) for f in (OUT_NAME, 'manifest.json', 'APPROVAL.json')), os.path.relpath(str(pubdir), root))
    c, m = _approve(runs_base, 'selftest', True, root, quiet=True); rec('P2 re-approving an already published run is refused', c == 4 and 'PUBLISHED' in m, m[:80])
    c, rel2 = _approve(ready_run, 'selftest2', True, root, quiet=True)
    lat = json.load(open(LAT))
    rec('P3 second approval of the same month goes to a NEW folder; first release files unchanged; both kept; LATEST -> second', c == 0 and rel2 != str(pubdir) and os.path.isdir(str(pubdir)) and h_first == {f: sha256(os.path.join(str(pubdir), f)) for f in os.listdir(str(pubdir))} and lat['run_id'] == ready_run and lat['release'] == os.path.relpath(rel2, root), os.path.relpath(str(rel2), root))
    code, st3 = _run(src, root, quiet=True); run3 = st3['run_id']; sent = os.path.join(pub, latest_m, run3); os.makedirs(sent); open(os.path.join(sent, 'sentinel.txt'), 'w').write('keep')
    lat_b = open(LAT, 'rb').read(); c, m = _approve(run3, 'selftest3', True, root, quiet=True)
    rec('P4 pre-existing release folder is never overwritten; approval refused; sentinel untouched; run stays READY', c == 4 and 'already exists' in m and open(os.path.join(sent, 'sentinel.txt')).read() == 'keep' and os.listdir(sent) == ['sentinel.txt'] and json.load(open(os.path.join(root, 'runs', run3, 'status.json')))['status'] == 'READY_FOR_APPROVAL', m[:80])
    rec('P5 LATEST.json unchanged by every refused approval', open(LAT, 'rb').read() == lat_b, 'bytes identical before/after refusals')
    ro = all((os.stat(os.path.join(str(pubdir), f)).st_mode & 0o222) == 0 for f in os.listdir(str(pubdir)))
    rec('P6 released files are marked read-only (best effort; NOT a security control)', ro, 'mode bits')
    ok_r, _ = verify_audit(os.path.join(pub, 'release_register.jsonl')); nreg = len(_entries(os.path.join(pub, 'release_register.jsonl')))
    rec('P7 release register is a valid hash chain with one entry per release', ok_r and nreg == 2, f'{nreg} entries')
    # M: approve verification and publication boundary (each tampered copy is a pristine copy of run3, which is READY and unpublished)
    def clone(name): d = os.path.join(root, 'runs', name); shutil.copytree(os.path.join(root, 'runs', run3), d); return d
    d = clone('m_manifest'); wbp = os.path.join(d, 'candidate', OUT_NAME); open(wbp, 'ab').write(b'x'); mj = os.path.join(d, 'candidate', 'manifest.json'); mm = json.load(open(mj)); mm['files'][OUT_NAME] = sha256(wbp); json.dump(mm, open(mj, 'w'), indent=1)
    c, m = _approve('m_manifest', 'selftest', True, root, quiet=True); rec('M1 workbook AND manifest edited consistently -> refused (manifest hash recorded in audit log)', c == 4 and 'manifest' in m, m[:90])
    d = clone('m_input'); ipf = os.path.join(d, 'input', os.listdir(os.path.join(d, 'input'))[0]); os.chmod(ipf, 0o644); open(ipf, 'ab').write(b'x')
    c, m = _approve('m_input', 'selftest', True, root, quiet=True); rec('M2 archived input modified -> refused', c == 4 and 'archived input' in m, m[:90])
    clone('m_code'); c, m = _approve('m_code', 'selftest', True, root, quiet=True, code_sha_override='f'*64); rec('M3 code differs from the code that built the candidate -> refused', c == 4 and 'code hash' in m, m[:90])
    pb, stb = _run(write_blk := variant('blk_m4', lambda x: setv(x, 5, 'Invoice_ID', x['Sales_Transactions'].loc[4, 'Invoice_ID'])), root, quiet=True)
    d = os.path.join(root, 'runs', 'm_blocked'); shutil.copytree(os.path.join(root, 'runs', stb['run_id']), d); os.rename(os.path.join(d, 'blocked'), os.path.join(d, 'candidate')); sj = os.path.join(d, 'status.json'); q = json.load(open(sj)); q['status'] = 'READY_FOR_APPROVAL'; json.dump(q, open(sj, 'w'))
    c, m = _approve('m_blocked', 'selftest', True, root, quiet=True); rec('M4 BLOCKED run relabelled as candidate with edited status.json -> refused', pb == 3 and c == 4, m[:100])
    man_dir = os.path.join(pub, latest_m, 'manual_copy'); shutil.copytree(os.path.join(root, 'runs', run3, 'candidate'), man_dir)
    vc, vm = verify_release(man_dir, root, quiet=True)
    rec('M5 boundary: a MANUAL copy into published/ is not prevented by the script, but verify-release flags it as unregistered', os.path.isdir(man_dir) and vc == 5, vm[:100])
    # A: optional external anchor
    rd0 = os.path.join(root, 'runs', runs_base); ap0 = os.path.join(rd0, 'audit', 'audit_log.jsonl'); b0 = sha256(ap0); c, h_anchor = anchor(runs_base, root, quiet=True)
    af = [f for f in os.listdir(os.path.join(rd0, 'audit')) if f.startswith('anchor_record')]
    rec('A1 anchor record written, states NOT ANCHORED until stored externally, and audit log itself unchanged', c == 0 and af and 'NOT ANCHORED' in open(os.path.join(rd0, 'audit', af[0])).read() and sha256(ap0) == b0, af[0] if af else '')
    c, m = verify_anchor(runs_base, h_anchor, root, quiet=True); rec('A2 verify-anchor passes with the correct externally stored hash', c == 0, m[:80])
    c, m = verify_anchor(runs_base, 'a'*64, root, quiet=True); rec('A3 verify-anchor fails with a wrong hash', c == 5, m[:80])
    shutil.copytree(rd0, os.path.join(root, 'runs', 'a_edit')); pa = os.path.join(root, 'runs', 'a_edit', 'audit', 'audit_log.jsonl'); txt_ = open(pa).read(); open(pa, 'w').write(txt_.replace('"PASS"', '"FAIL"', 1))
    c, m = verify_anchor('a_edit', h_anchor, root, quiet=True); open(pa, 'w').write('')
    c_e, m_e = verify_anchor('a_edit', h_anchor, root, quiet=True)
    rec('A4 edited audit log, and an emptied audit log -> verify-anchor fails', c == 5 and 'broken' in m and c_e == 5 and 'empty' in m_e, m[:50] + ' | ' + m_e[:40])
    shutil.copytree(rd0, os.path.join(root, 'runs', 'a_rechain')); pr = os.path.join(root, 'runs', 'a_rechain', 'audit', 'audit_log.jsonl'); es = _entries(pr); prev = '0'*64; outl = []
    es[3]['detail'] = dict(es[3]['detail'], forged=True)
    for e in es: e.pop('hash', None); e['prev'] = prev; e['hash'] = hashlib.sha256(json.dumps(e, sort_keys=True, default=str).encode()).hexdigest(); prev = e['hash']; outl.append(json.dumps(e, default=str))
    open(pr, 'w').write('\n'.join(outl) + '\n'); chain_ok, _ = verify_audit(pr); c, m = verify_anchor('a_rechain', h_anchor, root, quiet=True)
    rec('A5 WHOLE chain recomputed after an edit passes the internal check (why it is not tamper-proof) but fails against the externally stored hash', chain_ok and c == 5, f'internal chain valid={chain_ok}; anchor check code={c}')
    # V: verify-release
    c, m = verify_release(str(pubdir), root, quiet=True); rec('V1 verify-release passes on an untouched release', c == 0, m[:80])
    cp = os.path.join(pub, latest_m, 'v_copy_mod'); shutil.copytree(str(pubdir), cp); wbm = os.path.join(cp, OUT_NAME); os.chmod(wbm, 0o644); open(wbm, 'ab').write(b'x')
    c, m = verify_release(cp, root, quiet=True); rec('V2 verify-release fails when a released file is modified', c == 5 and 'modified' in m, m[:80])
    cp2 = os.path.join(pub, latest_m, 'v_copy_appr'); shutil.copytree(str(pubdir), cp2); aj = os.path.join(cp2, 'APPROVAL.json'); os.chmod(aj, 0o644); q = json.load(open(aj)); q['approved_by'] = 'someone else'; json.dump(q, open(aj, 'w'))
    c, m = verify_release(cp2, root, quiet=True); rec('V3 verify-release fails when APPROVAL.json is edited', c == 5, m[:80])
    rec('X: original dataset unchanged by the whole self-test', sha256(src) == h0, f'SHA-256 {h0[:16]}... before == after')
    npass = sum(r['passed'] for r in results)
    out = make_record(results, fixture=False); out['source_sha256'] = h0
    json.dump(out, open(SELFTEST_PATH, 'w'), indent=1); shutil.rmtree(tmp, ignore_errors=True)
    print(f'\nSELFTEST {npass}/{len(results)} passed'); return 0 if out['all_passed'] else 1

# ----------------------------------------------------------------------------- CLI
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter); sp = ap.add_subparsers(dest='cmd', required=True)
    r = sp.add_parser('run'); r.add_argument('--input', required=True); r.add_argument('--root', default='mis_workspace'); r.add_argument('--commentary-file')
    a_ = sp.add_parser('approve'); a_.add_argument('--run', required=True); a_.add_argument('--by', required=True); a_.add_argument('--ack-warnings', action='store_true'); a_.add_argument('--root', default='mis_workspace')
    a_.add_argument('--external-ref', default=None, help='optional: reference to an approval record held outside this workspace (ticket no., email id, signed-message id)')
    t = sp.add_parser('selftest'); t.add_argument('--input', required=True); t.add_argument('--root', default=None)
    n_ = sp.add_parser('anchor'); n_.add_argument('--run', required=True); n_.add_argument('--root', default='mis_workspace')
    va = sp.add_parser('verify-anchor'); va.add_argument('--run', required=True); va.add_argument('--hash', required=True); va.add_argument('--root', default='mis_workspace')
    vr = sp.add_parser('verify-release'); vr.add_argument('--release', required=True); vr.add_argument('--root', default='mis_workspace')
    a = ap.parse_args()
    if a.cmd == 'run': sys.exit(run(a.input, a.root, a.commentary_file)[0])
    if a.cmd == 'approve': sys.exit(approve(a.run, a.by, a.ack_warnings, a.root, external_ref=a.external_ref)[0])
    if a.cmd == 'anchor': sys.exit(anchor(a.run, a.root)[0])
    if a.cmd == 'verify-anchor': sys.exit(verify_anchor(a.run, a.hash, a.root)[0])
    if a.cmd == 'verify-release': sys.exit(verify_release(a.release, a.root)[0])
    root = a.root or tempfile.mkdtemp(prefix='nova_selftest_root_'); sys.exit(selftest(a.input, root))

if __name__ == '__main__': main()
