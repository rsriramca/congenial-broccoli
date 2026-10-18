"""Nova CFO MIS - v2 executive FP&A dashboard builder (static, self-contained HTML; synthetic data only).

    python src/build_dashboard.py --input data/source/Nova_CFO_MIS_Complete_Synthetic_Dataset.xlsx --out outputs/v2

Reads the ORIGINAL synthetic dataset (read-only), recomputes every figure with pandas and writes:
    Nova_FPA_Dashboard.html   one file, inline SVG + CSS, no scripts, no network requests
    dashboard_data.json       every figure shown, for reconciliation against the v1 MIS workbook
It does NOT modify any v1 file, does not call any AI model, and is a point-in-time snapshot: re-running it is a manual step (no live refresh).
Output is deterministic (no timestamp), so identical input gives byte-identical output.
"""
import argparse, hashlib, html, json, os
import pandas as pd

LAKH = 1e5
def sha256(p): return hashlib.sha256(open(p, 'rb').read()).hexdigest()

# ----------------------------------------------------------------------------- figures
def compute(src):
    x = pd.read_excel(src, sheet_name=None)
    cal, pm, cm, tx = x['Calendar'], x['Product_Master'], x['Customer_Master'], x['Sales_Transactions']
    last = cal[cal.Reporting_Status == 'Actual'].Month.max(); Y, M = int(last[:4]), int(last[5:7])
    a = tx[tx.Status == 'Actual'].merge(pm[['Product_ID', 'Product_Name', 'Category']], on='Product_ID', how='left').merge(cm[['Customer_ID', 'Region']], on='Customer_ID', how='left')
    a['Y'] = a.Month.str[:4].astype(int); a['M'] = a.Month.str[5:7].astype(int)
    def agg(df):
        r, c, g = df.Net_Revenue_INR.sum(), df.COGS_INR.sum(), df.Gross_Profit_INR.sum()
        return dict(revenue=round(float(r), 2), cogs=round(float(c), 2), gp=round(float(g), 2), gm=float(g / r) if r else None)
    ytd = lambda df, y: df[(df.Y == y) & (df.M <= M)]
    out = dict(as_of=last, current_year=Y, ytd_through_month=M, source_sha256=sha256(src), source_file=os.path.basename(src))
    out['monthly'] = [dict(month=m, **agg(g)) for m, g in a.groupby('Month')]
    out['total'] = agg(a); out['ytd_current'] = agg(ytd(a, Y)); out['ytd_prior'] = agg(ytd(a, Y - 1))
    yc, yp = out['ytd_current'], out['ytd_prior']
    out['ytd_change'] = dict(revenue=yc['revenue'] / yp['revenue'] - 1, gp=yc['gp'] / yp['gp'] - 1, cogs=yc['cogs'] / yp['cogs'] - 1, gm_pp=(yc['gm'] - yp['gm']) * 100)
    def dim(col, extra=None):
        rows = []
        for k, g in a.groupby(col):
            d = dict(name=k, **agg(g)); d['share'] = d['revenue'] / out['total']['revenue']
            c, p = agg(ytd(g, Y)), agg(ytd(g, Y - 1))
            d.update(ytd_revenue=c['revenue'], ytd_revenue_prior=p['revenue'], ytd_yoy=(c['revenue'] / p['revenue'] - 1) if p['revenue'] else None, ytd_gm=c['gm'], ytd_gm_prior=p['gm'])
            if extra: d.update(extra(g))
            rows.append(d)
        return rows
    out['regions'] = dim('Region', lambda g: dict(products=sorted(g.Product_Name.unique())))
    out['products'] = dim('Product_Name', lambda g: dict(regions=sorted(g.Region.unique()), category=g.Category.iloc[0]))
    out['overlap'] = dict(max_regions_per_product=int(a.groupby('Product_Name').Region.nunique().max()), max_products_per_region=int(a.groupby('Region').Product_Name.nunique().max()),
                          n_regions=int(a.Region.nunique()), n_products=int(a.Product_Name.nunique()))
    # original budget comparability (diagnostic only - no variance is computed or shown)
    b = x['Sales_Budget']; am = set(a.Month.unique()); cov = set(zip(a.Region, a.Category))
    bm = b[b.Month.isin(am)]; inc = bm[[(r, c) in cov for r, c in zip(bm.Region, bm.Category)]]
    out['budget'] = dict(ratio_vs_actual=float(bm.Revenue_Budget_INR.sum() / a.Net_Revenue_INR.sum()), ratio_in_scope=float(inc.Revenue_Budget_INR.sum() / a.Net_Revenue_INR.sum()),
                         share_in_cells_without_sales=float(1 - inc.Revenue_Budget_INR.sum() / bm.Revenue_Budget_INR.sum()),
                         cells_total=len(set(zip(b.Region, b.Category))), cells_without_sales=len(set(zip(b.Region, b.Category)) - cov))
    return out

# ----------------------------------------------------------------------------- formatting + svg helpers
E = html.escape
def lk(v): return f'{v / LAKH:,.1f}'
def pct(v, d=1): return f'{v * 100:.{d}f}%'
def sg(v, d=1): return f'{v * 100:+.{d}f}%'

def nice_ticks(lo, hi, n=5):
    import math
    span = hi - lo; raw = span / (n - 1); mag = 10 ** math.floor(math.log10(raw)); step = next(m * mag for m in (1, 2, 2.5, 5, 10) if m * mag >= raw)
    t0 = math.floor(lo / step + 1e-9) * step; t = []; v = t0
    while v < hi - 1e-9: t.append(v); v += step
    t.append(v); return t

def line_svg(labels, series, ymin, ymax, fmt, w=760, h=280, yname='', lfmt=None):
    lfmt = lfmt or fmt
    L, R, T, B = 52, 120, 14, 34; pw, ph = w - L - R, h - T - B; n = len(labels)
    ticks = nice_ticks(ymin, ymax); ymin, ymax = ticks[0], ticks[-1]
    X = lambda i: L + pw * i / (n - 1); Y = lambda v: T + ph * (1 - (v - ymin) / (ymax - ymin))
    s = [f'<svg viewBox="0 0 {w} {h}" role="img" aria-label="{E(yname)}" class="chart">']
    for v in ticks:
        y = Y(v)
        s.append(f'<line x1="{L}" x2="{L + pw}" y1="{y:.1f}" y2="{y:.1f}" class="grid"/><text x="{L - 8}" y="{y + 4:.1f}" class="tick" text-anchor="end">{fmt(v)}</text>')
    for i, lab in enumerate(labels):
        if i % 3 == 0 or i == n - 1: s.append(f'<text x="{X(i):.1f}" y="{h - 12}" class="tick" text-anchor="middle">{E(lab)}</text>')
    for name, cls, vals in series:
        pts = ' '.join(f'{X(i):.1f},{Y(v):.1f}' for i, v in enumerate(vals)); s.append(f'<polyline points="{pts}" class="ln {cls}"/>')
        for i, v in enumerate(vals): s.append(f'<circle cx="{X(i):.1f}" cy="{Y(v):.1f}" r="{4.5 if i == n - 1 else 3}" class="dot {cls}"><title>{E(name)} {E(labels[i])}: {lfmt(v)}</title></circle>')
        s.append(f'<text x="{X(n - 1) + 10:.1f}" y="{Y(vals[-1]) + 4:.1f}" class="dl">{E(name)} {lfmt(vals[-1])}</text>')
    return '<div class="cw">' + '\n'.join(s) + '</svg></div>'

def bars_svg(rows, w=760, rh=30, name_w=150):
    """Horizontal bars, zero baseline, one hue; value labels at bar end. rows = [(label, value, text, tooltip)]"""
    mx = max(v for _, v, _, _ in rows) * 1.0; pw = w - name_w - 200; h = rh * len(rows) + 8
    s = [f'<svg viewBox="0 0 {w} {h}" class="chart" role="img" aria-label="bar chart">']
    for i, (lab, v, txt, tip) in enumerate(rows):
        y = 4 + i * rh; bw = max(2, pw * v / mx)
        s.append(f'<text x="{name_w - 8}" y="{y + 17}" class="tick lbl" text-anchor="end">{E(lab)}</text><rect x="{name_w}" y="{y + 4}" width="{bw:.1f}" height="16" rx="3" class="bar"><title>{E(tip)}</title></rect><text x="{name_w + bw + 6:.1f}" y="{y + 17}" class="dl">{E(txt)}</text>')
    return '<div class="cw">' + '\n'.join(s) + '</svg></div>'

def ytd_bars_svg(items, w=760, h=250):
    """Grouped vertical bars: prior-year YTD (hollow, dashed) vs current YTD (solid). items=[(label, cur, prior, fmtfn)]"""
    L, T, B = 52, 20, 44; ph = h - T - B; gw = (w - L - 20) / len(items); ticks = nice_ticks(0, max(max(c, p) for _, c, p, _ in items) * 1.08, 5); mx = ticks[-1]
    s = [f'<svg viewBox="0 0 {w} {h}" class="chart" role="img" aria-label="YTD current versus prior year">']
    for v in ticks:
        y = T + ph * (1 - v / mx); s.append(f'<line x1="{L}" x2="{w - 20}" y1="{y:.1f}" y2="{y:.1f}" class="grid"/><text x="{L - 8}" y="{y + 4:.1f}" class="tick" text-anchor="end">{v / LAKH:,.0f}</text>')
    for i, (lab, c, p, f) in enumerate(items):
        x0 = L + i * gw + gw * 0.22; bw = gw * 0.26
        for j, (v, cls, nm) in enumerate(((p, 'bar-prior', 'Prior-year YTD'), (c, 'bar', 'Current YTD'))):
            bh = ph * v / mx; x = x0 + j * (bw + 4); y = T + ph - bh
            s.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bw:.1f}" height="{bh:.1f}" rx="3" class="{cls}"><title>{E(lab)} {nm}: INR {f(v)} lakh</title></rect><text x="{x + bw / 2:.1f}" y="{y - 5:.1f}" class="dl" text-anchor="middle">{f(v)}</text>')
        s.append(f'<text x="{L + i * gw + gw / 2:.1f}" y="{h - 22}" class="tick lbl" text-anchor="middle">{E(lab)}</text>')
    return '<div class="cw">' + '\n'.join(s) + '</svg></div>'

def table(head, rows, cls=''):
    return f'<table class="{cls}"><thead><tr>' + ''.join(f'<th>{E(h)}</th>' for h in head) + '</tr></thead><tbody>' + ''.join('<tr>' + ''.join(f'<td>{c}</td>' for c in r) + '</tr>' for r in rows) + '</tbody></table>'

CSS = """
:root{color-scheme:light;--bg:#f6f5f2;--surface:#fcfcfb;--ink:#0b0b0b;--ink2:#52514e;--grid:#e1e0d9;--edge:rgba(11,11,11,.10);--s1:#2a78d6;--s2:#eb6834;--warnbg:#fff3e6;--warnbd:#d03b3b;--warnfg:#7a1d1d;--note:#eef3fb}
@media (prefers-color-scheme:dark){:root:where(:not([data-theme=light])){color-scheme:dark;--bg:#121211;--surface:#1a1a19;--ink:#f0efec;--ink2:#c3c2b7;--grid:#2c2c2a;--edge:rgba(255,255,255,.10);--s1:#3987e5;--s2:#d95926;--warnbg:#3a1a18;--warnbd:#ec835a;--warnfg:#ffd9cf;--note:#17202c}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
main{max-width:980px;margin:0 auto;padding:20px 16px 48px}h1{font-size:24px;margin:4px 0 2px}h2{font-size:18px;margin:0 0 4px}p{margin:6px 0}
.sub{color:var(--ink2);font-size:14px}.tag{display:inline-block;border:1px solid var(--edge);border-radius:999px;padding:2px 10px;font-size:12px;color:var(--ink2);margin:2px 4px 2px 0;background:var(--surface)}
.warn{background:var(--warnbg);border:2px solid var(--warnbd);border-radius:10px;padding:14px 16px;margin:16px 0;color:var(--warnfg)}.warn b.t{font-size:16px}
.note{background:var(--note);border:1px solid var(--edge);border-radius:10px;padding:10px 14px;margin:12px 0;font-size:14px}
section.card{background:var(--surface);border:1px solid var(--edge);border-radius:12px;padding:16px;margin:16px 0}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px}.kpi{background:var(--surface);border:1px solid var(--edge);border-radius:12px;padding:12px 14px}
.kpi .v{font-size:26px;font-weight:650}.kpi .l{color:var(--ink2);font-size:13px}.kpi .d{font-size:13px;color:var(--ink2)}
.chart{width:100%;min-width:600px;height:auto;display:block}.cw{overflow-x:auto}.grid{stroke:var(--grid);stroke-width:1}.tick{fill:var(--ink2);font-size:11px}.lbl{font-size:12px}.dl{fill:var(--ink);font-size:11.5px}
.ln{fill:none;stroke-width:2}.s1{stroke:var(--s1)}.s2{stroke:var(--s2)}.dot{stroke:var(--surface);stroke-width:2}.dot.s1{fill:var(--s1)}.dot.s2{fill:var(--s2)}
.bar{fill:var(--s1)}.bar-prior{fill:none;stroke:var(--s1);stroke-width:2;stroke-dasharray:5 3}
.legend{display:flex;gap:16px;flex-wrap:wrap;font-size:13px;color:var(--ink2);margin:4px 0 6px}.legend i{display:inline-block;width:14px;height:0;border-top:3px solid;vertical-align:middle;margin-right:6px}
.legend i.s1{border-color:var(--s1)}.legend i.s2{border-color:var(--s2)}.legend i.solid{height:10px;border:0;background:var(--s1);border-radius:2px}.legend i.hollow{height:10px;border:2px dashed var(--s1);border-radius:2px}
table{border-collapse:collapse;width:100%;font-size:13px;margin-top:8px}th,td{padding:6px 8px;border-bottom:1px solid var(--grid);text-align:right;white-space:nowrap}th:first-child,td:first-child,th:nth-child(2).tx,td:nth-child(2).tx{text-align:left}th{color:var(--ink2);font-weight:600}
.scroll{overflow-x:auto}details{margin-top:8px}summary{cursor:pointer;color:var(--ink2);font-size:13px}code,pre{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:12.5px}pre{background:var(--bg);padding:10px;border-radius:8px;overflow-x:auto}
footer{color:var(--ink2);font-size:12.5px;margin-top:24px}
"""

def render(d):
    mo = d['monthly']; labels = [m['month'] for m in mo]; yc, yp, ch = d['ytd_current'], d['ytd_prior'], d['ytd_change']; Y, M = d['current_year'], d['ytd_through_month']
    bud = d['budget']; ov = d['overlap']; last = mo[-1]; per = f'Jan-{["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][M - 1]}'
    kp = lambda l, v, dl: f'<div class="kpi"><div class="l">{E(l)}</div><div class="v">{v}</div><div class="d">{dl}</div></div>'
    kpis = ''.join([kp(f'YTD net revenue {Y} ({per})', f'INR {lk(yc["revenue"])} lakh', f'{sg(ch["revenue"])} vs prior-year YTD INR {lk(yp["revenue"])} lakh'),
                    kp(f'YTD gross profit {Y}', f'INR {lk(yc["gp"])} lakh', f'{sg(ch["gp"])} vs INR {lk(yp["gp"])} lakh'),
                    kp(f'YTD gross margin {Y}', pct(yc['gm']), f'{ch["gm_pp"]:+.1f} pp vs {pct(yp["gm"])} prior-year YTD'),
                    kp(f'Latest actual month {d["as_of"]}', f'INR {lk(last["revenue"])} lakh', f'margin {pct(last["gm"])}; gross profit INR {lk(last["gp"])} lakh')])
    warn = (f'<div class="warn" role="alert"><b class="t">&#9888; DATA-QUALITY WARNING - budget is NOT comparable</b>'
            f'<p>The original <code>Sales_Budget</code> is about <b>{bud["ratio_vs_actual"]:.1f}x</b> actual net revenue for the same months, and <b>{pct(bud["share_in_cells_without_sales"], 0)}</b> of it sits in '
            f'Region x Category cells that have no sales at all ({bud["cells_without_sales"]} of {bud["cells_total"]} cells). The cause is <b>unresolved</b>. '
            f'<b>This dashboard therefore shows no budget-vs-actual variance.</b> Any variance in the v1 workbook\'s <code>Budget_v2_Synthetic</code> / <code>Budget_vs_Actual_v2</code> sheets is <b>illustrative only</b> '
            f'(calibrated to prior-year actuals, not an approved budget) and must not be used for performance assessment.</p></div>')
    scope = (f'<div class="note"><b>Scope.</b> Synthetic data only (fictional company). Actuals only through {E(d["as_of"])}; forecast months excluded. Net revenue excludes GST, returns, rebates and intercompany. '
             f'YTD = calendar months 1-{M} of {Y} versus the same months of {Y - 1} (the source calendar has no separate fiscal year; calendar-year YTD is an assumption). '
             f'Static snapshot rebuilt manually: <b>no live refresh, no autonomous publishing</b>.</div>')
    trend1 = line_svg(labels, [('Net revenue', 's1', [m['revenue'] for m in mo]), ('Gross profit', 's2', [m['gp'] for m in mo])], 0, max(m['revenue'] for m in mo) * 1.05, lambda v: f'{v / LAKH:,.0f}', yname='Monthly net revenue and gross profit in INR lakh', lfmt=lambda v: f'{v / LAKH:,.1f}')
    gmv = [m['gm'] for m in mo]; lo, hi = min(gmv), max(gmv); ymin, ymax = nice_ticks(lo - 0.003, hi + 0.003, 5)[0], nice_ticks(lo - 0.003, hi + 0.003, 5)[-1]
    trend2 = line_svg(labels, [('Gross margin', 's1', gmv)], ymin, ymax, lambda v: pct(v, 1), yname='Monthly gross margin percent')
    mtab = table(['Month', 'Net revenue (INR lakh)', 'COGS', 'Gross profit', 'Gross margin'], [[m['month'], lk(m['revenue']), lk(m['cogs']), lk(m['gp']), pct(m['gm'])] for m in mo])
    ybars = ytd_bars_svg([('Net revenue', yc['revenue'], yp['revenue'], lk), ('COGS', yc['cogs'], yp['cogs'], lk), ('Gross profit', yc['gp'], yp['gp'], lk)])
    ytab = table(['Measure (INR lakh)', f'{Y} YTD', f'{Y - 1} YTD', 'Change'], [['Net revenue', lk(yc['revenue']), lk(yp['revenue']), sg(ch['revenue'])], ['COGS', lk(yc['cogs']), lk(yp['cogs']), sg(ch['cogs'])],
                 ['Gross profit', lk(yc['gp']), lk(yp['gp']), sg(ch['gp'])], ['Gross margin', pct(yc['gm']), pct(yp['gm']), f'{ch["gm_pp"]:+.1f} pp']])
    def dimsec(rows, key_label, extra_head, extra_cell):
        rows = sorted(rows, key=lambda r: -r['revenue'])
        bars = bars_svg([(r['name'], r['revenue'], f'{lk(r["revenue"])}  ({pct(r["gm"])} margin)', f'{r["name"]}: net revenue INR {lk(r["revenue"])} lakh, gross margin {pct(r["gm"])}') for r in rows])
        tb = table([key_label, extra_head, 'Net revenue (INR lakh)', 'Share', 'Gross margin', f'{Y} YTD revenue', f'{Y - 1} YTD revenue', 'YTD change', f'YTD margin {Y} / {Y - 1}'],
                   [[E(r['name']), extra_cell(r), lk(r['revenue']), pct(r['share']), pct(r['gm']), lk(r['ytd_revenue']), lk(r['ytd_revenue_prior']), sg(r['ytd_yoy']), f'{pct(r["ytd_gm"])} / {pct(r["ytd_gm_prior"])}'] for r in rows])
        return bars, f'<div class="scroll">{tb}</div>'
    rb, rt = dimsec(d['regions'], 'Region', 'Products sold', lambda r: E(', '.join(r['products'])))
    pb, pt = dimsec(d['products'], 'Product', 'Region sold in', lambda r: E(', '.join(r['regions'])))
    overlap = (f'<div class="note"><b>Caveat - region and product effects overlap.</b> In this dataset each product is sold in exactly {ov["max_regions_per_product"]} region (max regions per product = {ov["max_regions_per_product"]}), '
               f'and each region carries only {ov["max_products_per_region"]} of the {ov["n_products"]} products. A region\'s result is therefore its two products\' result, and a product\'s result is its region\'s. '
               f'The two views are <b>confounded</b>: do not read a regional difference as a regional effect, or a product difference as a product effect. The data cannot separate them.</div>')
    legend_rg = '<div class="legend"><span><i class="solid"></i>Net revenue, actuals to ' + E(d['as_of']) + '</span></div>'
    refresh = f"""<pre>pip install pandas numpy openpyxl
python src/build_dashboard.py --input data/source/Nova_CFO_MIS_Complete_Synthetic_Dataset.xlsx --out outputs/v2
python tests/test_dashboard.py          # reconciliation checks</pre>
<p>To use a newer synthetic extract: place it under <code>data/</code>, first build and review the v1 MIS candidate with <code>python run_mis.py run</code> (human approval remains separate), then rerun the commands above on that file. See <code>docs/DASHBOARD_V2.md</code>.</p>"""
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Nova FP&A Dashboard v2</title><style>{CSS}</style></head><body><main>
<h1>Nova Consumer Products - Executive FP&amp;A Dashboard</h1>
<p class="sub">Actuals to {E(d["as_of"])} &middot; INR &middot; built from the v1 synthetic source data</p>
<p><span class="tag">v2 dashboard enhancement (separate from Provisional v1)</span><span class="tag">Synthetic data</span><span class="tag">Static snapshot - not live</span><span class="tag">Not validated in Excel Desktop</span><span class="tag">Not production-ready</span></p>
{warn}{scope}
<div class="kpis">{kpis}</div>
<section class="card"><h2>Monthly net revenue and gross profit</h2><p class="sub">INR lakh, actuals, zero-based axis.</p><div class="legend"><span><i class="s1"></i>Net revenue</span><span><i class="s2"></i>Gross profit</span></div>{trend1}</section>
<section class="card"><h2>Monthly gross margin</h2><p class="sub">Percent of net revenue. <b>Axis does not start at zero</b> ({pct(ymin, 0)}-{pct(ymax, 0)}) so the month-to-month range is visible; the total range is only {pct(lo)} to {pct(hi)}.</p>{trend2}
<details><summary>Show monthly table</summary><div class="scroll">{mtab}</div></details></section>
<section class="card"><h2>YTD {Y} ({per}) versus prior-year YTD</h2><div class="legend"><span><i class="solid"></i>{Y} YTD</span><span><i class="hollow"></i>{Y - 1} YTD (same months)</span></div><p class="sub">INR lakh. Revenue {sg(ch["revenue"])}, gross profit {sg(ch["gp"])}, margin {ch["gm_pp"]:+.1f} pp. Differences are small and this is one synthetic series; no cause can be inferred from the data.</p>{ybars}<div class="scroll">{ytab}</div></section>
<section class="card"><h2>Region view</h2>{overlap}{legend_rg}{rb}{rt}</section>
<section class="card"><h2>Product view</h2>{overlap}{legend_rg}{pb}{pt}</section>
<section class="card"><h2>Budget comparability (no variance shown)</h2><p>Original budget / actual net revenue (same months): <b>{bud["ratio_vs_actual"]:.1f}x</b>; {bud["ratio_in_scope"]:.1f}x on the cells that do have sales. {pct(bud["share_in_cells_without_sales"], 0)} of budget is in cells with no sales. Cause unresolved; the original budget is unchanged in the source and the v1 documents.
Budget variance is intentionally <b>not</b> calculated here. See <code>Budget_Integrity_Report.md</code>.</p></section>
<section class="card"><h2>How to refresh (manual)</h2>{refresh}</section>
<footer>Source file <code>{E(d["source_file"])}</code> SHA-256 <code>{d["source_sha256"][:16]}...</code>. Recomputed in pandas from the source; reconciled to the v1 MIS workbook by <code>tests/test_dashboard.py</code> (LibreOffice-calculated values; not tested in Microsoft Excel Desktop). Synthetic, non-sensitive data. Not live, not autonomous, not production-ready.</footer>
</main></body></html>"""

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--input', required=True); ap.add_argument('--out', default='outputs/v2'); a = ap.parse_args()
    d = compute(a.input); os.makedirs(a.out, exist_ok=True)
    json.dump(d, open(os.path.join(a.out, 'dashboard_data.json'), 'w'), indent=1, sort_keys=True)
    open(os.path.join(a.out, 'Nova_FPA_Dashboard.html'), 'w', encoding='utf-8').write(render(d))
    print(f'wrote {a.out}/Nova_FPA_Dashboard.html and dashboard_data.json (as of {d["as_of"]})')

if __name__ == '__main__': main()
