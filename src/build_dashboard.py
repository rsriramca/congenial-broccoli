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

def line_svg(labels, series, ymin, ymax, fmt, w=900, h=300, yname='', lfmt=None, plabels=False):
    lfmt = lfmt or fmt
    L, R, T, B = 68, 118, 20, 34; pw, ph = w - L - R, h - T - B; n = len(labels)
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
        if plabels:
            for i, v in enumerate(vals):   # label on the side of the point away from the line, so text does not sit on it
                nb = [vals[j] for j in (i - 1, i + 1) if 0 <= j < n]; below = v < sum(nb) / len(nb)
                s.append(f'<text x="{X(i) + (4 if i == 0 else 0):.1f}" y="{Y(v) + (17 if below else -9):.1f}" class="vl sm" text-anchor="{"start" if i == 0 else "middle"}">{lfmt(v)}</text>')
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
            s.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bw:.1f}" height="{bh:.1f}" rx="3" class="{cls}"><title>{E(lab)} {nm}: INR {f(v)} lakh</title></rect><text x="{x + bw / 2:.1f}" y="{y - 6:.1f}" class="vl" text-anchor="middle">{f(v)}</text>')
        s.append(f'<text x="{L + i * gw + gw / 2:.1f}" y="{h - 22}" class="tick lbl" text-anchor="middle">{E(lab)}</text>')
    return '<div class="cw">' + '\n'.join(s) + '</svg></div>'

def stacked_svg(mo, w=1000, h=340):
    """Monthly stacked columns: COGS (bottom) + gross profit (top). Labels: net revenue above each column, COGS and gross profit inside their segments."""
    L, R, T, B = 44, 10, 28, 34; pw, ph = w - L - R, h - T - B; n = len(mo); slot = pw / n; bw = slot * 0.66
    ticks = nice_ticks(0, max(m['revenue'] for m in mo) * 1.06, 5); mx = ticks[-1]; Y = lambda v: T + ph * (1 - v / mx)
    s = [f'<svg viewBox="0 0 {w} {h}" class="chart wide" role="img" aria-label="Monthly net revenue split into COGS and gross profit, INR lakh">']
    for v in ticks: s.append(f'<line x1="{L}" x2="{w - R}" y1="{Y(v):.1f}" y2="{Y(v):.1f}" class="grid"/><text x="{L - 6}" y="{Y(v) + 4:.1f}" class="tick" text-anchor="end">{v / LAKH:,.0f}</text>')
    for i, m in enumerate(mo):
        x = L + i * slot + (slot - bw) / 2; yc = Y(m['cogs']); yr = Y(m['revenue']); cx = x + bw / 2
        s.append(f'<rect x="{x:.1f}" y="{yc:.1f}" width="{bw:.1f}" height="{Y(0) - yc:.1f}" class="seg-cogs"><title>{E(m["month"])} COGS: INR {lk(m["cogs"])} lakh</title></rect>')
        s.append(f'<rect x="{x:.1f}" y="{yr:.1f}" width="{bw:.1f}" height="{yc - yr - 1.5:.1f}" rx="3" class="seg-gp"><title>{E(m["month"])} gross profit: INR {lk(m["gp"])} lakh ({pct(m["gm"])} margin); net revenue INR {lk(m["revenue"])} lakh</title></rect>')
        s.append(f'<text x="{cx:.1f}" y="{yr - 6:.1f}" class="vl" text-anchor="middle">{lk(m["revenue"])}</text>')
        s.append(f'<text x="{cx:.1f}" y="{(yr + yc) / 2 + 3.5:.1f}" class="vli dk" text-anchor="middle">{lk(m["gp"])}</text><text x="{cx:.1f}" y="{(yc + Y(0)) / 2 + 3.5:.1f}" class="vli" text-anchor="middle">{lk(m["cogs"])}</text>')
        s.append(f'<text x="{cx:.1f}" y="{h - 12}" class="tick" text-anchor="middle">{m["month"][5:]}/{m["month"][2:4]}</text>')
    return '<div class="cw">' + '\n'.join(s) + '</svg></div>'

def cols_svg(rows, w=760, h=310):
    """Vertical columns, zero baseline. rows=[(name, value, top_label, inside_label, tooltip)]. Value above each bar, secondary measure inside it."""
    L, T, B = 20, 30, 56; ph = h - T - B; slot = (w - L - 10) / len(rows); bw = min(78, slot * 0.6); mx = max(r[1] for r in rows) * 1.08
    s = [f'<svg viewBox="0 0 {w} {h}" class="chart" role="img" aria-label="column chart">', f'<line x1="{L}" x2="{w - 10}" y1="{T + ph}" y2="{T + ph}" class="grid"/>']
    for i, (nm, v, top, ins, tip) in enumerate(rows):
        bh = ph * v / mx; x = L + i * slot + (slot - bw) / 2; y = T + ph - bh; cx = x + bw / 2
        s.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bw:.1f}" height="{bh:.1f}" rx="4" class="bar"><title>{E(tip)}</title></rect><text x="{cx:.1f}" y="{y - 7:.1f}" class="vl" text-anchor="middle">{E(top)}</text><text x="{cx:.1f}" y="{y + 18:.1f}" class="vli" text-anchor="middle">{E(ins)}</text>')
        parts = nm.split(' ') if len(nm) > 11 and ' ' in nm else [nm]
        if len(parts) > 1: k = max(1, len(parts) // 2); parts = [' '.join(parts[:k]), ' '.join(parts[k:])]
        for j, t_ in enumerate(parts): s.append(f'<text x="{cx:.1f}" y="{T + ph + 18 + 14 * j}" class="tick lbl" text-anchor="middle">{E(t_)}</text>')
    return '<div class="cw">' + '\n'.join(s) + '</svg></div>'

def table(head, rows, cls=''):
    return f'<table class="{cls}"><thead><tr>' + ''.join(f'<th>{E(h)}</th>' for h in head) + '</tr></thead><tbody>' + ''.join('<tr>' + ''.join(f'<td>{c}</td>' for c in r) + '</tr>' for r in rows) + '</tbody></table>'

CSS = """
:root{color-scheme:light;--bg:#eef1f5;--surface:#ffffff;--ink:#14202e;--ink2:#4d5a69;--grid:#e3e8ee;--edge:rgba(20,32,46,.10);--s1:#2a78d6;--s2:#eb6834;--s3:#1baf7a;--onblue:#ffffff;--onaqua:#0b2a1f;
--warnbg:#fff1e8;--warnbd:#d03b3b;--warnfg:#6b1a1a;--note:#eaf1fb;--hero1:#10294a;--hero2:#1d4f91;--herotxt:#ffffff;--shadow:0 1px 2px rgba(20,32,46,.06),0 4px 14px rgba(20,32,46,.06)}
@media (prefers-color-scheme:dark){:root:where(:not([data-theme=light])){color-scheme:dark;--bg:#0e1319;--surface:#171e27;--ink:#eef2f7;--ink2:#aab6c4;--grid:#26303c;--edge:rgba(255,255,255,.09);--s1:#3987e5;--s2:#d95926;--s3:#199e70;--onblue:#ffffff;--onaqua:#06231a;
--warnbg:#3a1a18;--warnbd:#ec835a;--warnfg:#ffd9cf;--note:#14202e;--hero1:#0b1c33;--hero2:#16427a;--shadow:none}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
main{max-width:1040px;margin:0 auto;padding:0 16px 48px}
.hero{background:linear-gradient(135deg,var(--hero1),var(--hero2));color:var(--herotxt);padding:26px 0 22px;margin:0 0 8px}.hero .in{max-width:1040px;margin:0 auto;padding:0 16px}
h1{font-size:25px;margin:0 0 2px;letter-spacing:.2px}h2{font-size:18px;margin:0 0 4px}p{margin:6px 0}.sub{color:var(--ink2);font-size:14px}.hero .sub{color:#d3e0f2}
.tag{display:inline-block;border:1px solid rgba(255,255,255,.35);border-radius:999px;padding:2px 11px;font-size:12px;color:#f1f6ff;margin:3px 5px 0 0;background:rgba(255,255,255,.10)}
.warn{background:var(--warnbg);border:2px solid var(--warnbd);border-left-width:8px;border-radius:10px;padding:14px 16px;margin:16px 0;color:var(--warnfg);box-shadow:var(--shadow)}.warn b.t{font-size:16px}
.note{background:var(--note);border:1px solid var(--edge);border-left:4px solid var(--s1);border-radius:8px;padding:10px 14px;margin:12px 0;font-size:14px}
section.card{background:var(--surface);border:1px solid var(--edge);border-radius:14px;padding:18px;margin:16px 0;box-shadow:var(--shadow)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(215px,1fr));gap:12px}.kpi{background:var(--surface);border:1px solid var(--edge);border-top:4px solid var(--s1);border-radius:12px;padding:12px 14px;box-shadow:var(--shadow)}
.kpi.g{border-top-color:var(--s3)}.kpi .v{font-size:27px;font-weight:700;letter-spacing:-.3px}.kpi .l{color:var(--ink2);font-size:13px}.kpi .d{font-size:13px;color:var(--ink2)}
.chart{width:100%;min-width:600px;height:auto;display:block}.chart.wide{min-width:900px}.cw{overflow-x:auto}.grid{stroke:var(--grid);stroke-width:1}.tick{fill:var(--ink2);font-size:11px}.lbl{font-size:12px}.dl{fill:var(--ink);font-size:11.5px}
.vl{fill:var(--ink);font-size:11.5px;font-weight:650}.vl.sm{font-size:9.5px;font-weight:600;fill:var(--ink2)}.vli{fill:var(--onblue);font-size:10px;font-weight:600}.vli.dk{fill:var(--onaqua)}
.ln{fill:none;stroke-width:2.5}.s1{stroke:var(--s1)}.s2{stroke:var(--s2)}.s3{stroke:var(--s3)}.dot{stroke:var(--surface);stroke-width:2}.dot.s1{fill:var(--s1)}.dot.s2{fill:var(--s2)}.dot.s3{fill:var(--s3)}
.bar{fill:var(--s1)}.bar-prior{fill:var(--s1);fill-opacity:.14;stroke:var(--s1);stroke-width:2;stroke-dasharray:5 3}.seg-cogs{fill:var(--s1)}.seg-gp{fill:var(--s3)}
.legend{display:flex;gap:16px;flex-wrap:wrap;font-size:13px;color:var(--ink2);margin:4px 0 6px}.legend i{display:inline-block;width:14px;height:0;border-top:3px solid;vertical-align:middle;margin-right:6px}
.legend i.s1{border-color:var(--s1)}.legend i.s3{border-color:var(--s3)}.legend i.solid{height:10px;border:0;background:var(--s1);border-radius:2px}.legend i.gpk{background:var(--s3)}.legend i.hollow{height:10px;border:2px dashed var(--s1);border-radius:2px;background:rgba(42,120,214,.14)}
table{border-collapse:collapse;width:100%;font-size:13px;margin-top:8px}th,td{padding:7px 8px;border-bottom:1px solid var(--grid);text-align:right;white-space:nowrap}th:first-child,td:first-child,.tx th:nth-child(2),.tx td:nth-child(2){text-align:left}th{color:var(--ink2);font-weight:600;background:var(--bg)}tbody tr:hover{background:var(--bg)}
.scroll{overflow-x:auto}details{margin-top:8px}summary{cursor:pointer;color:var(--ink2);font-size:13px}code,pre{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:12.5px}pre{background:var(--bg);padding:10px;border-radius:8px;overflow-x:auto}
footer{color:var(--ink2);font-size:12.5px;margin-top:24px}
"""

def render(d):
    mo = d['monthly']; labels = [m['month'] for m in mo]; yc, yp, ch = d['ytd_current'], d['ytd_prior'], d['ytd_change']; Y, M = d['current_year'], d['ytd_through_month']
    bud = d['budget']; ov = d['overlap']; last = mo[-1]; per = f'Jan-{["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][M - 1]}'
    kp = lambda l, v, dl, c='': f'<div class="kpi {c}"><div class="l">{E(l)}</div><div class="v">{v}</div><div class="d">{dl}</div></div>'
    kpis = ''.join([kp(f'YTD net revenue {Y} ({per})', f'INR {lk(yc["revenue"])} lakh', f'{sg(ch["revenue"])} vs prior-year YTD INR {lk(yp["revenue"])} lakh'),
                    kp(f'YTD gross profit {Y}', f'INR {lk(yc["gp"])} lakh', f'{sg(ch["gp"])} vs INR {lk(yp["gp"])} lakh', 'g'),
                    kp(f'YTD gross margin {Y}', pct(yc['gm']), f'{ch["gm_pp"]:+.1f} pp vs {pct(yp["gm"])} prior-year YTD', 'g'),
                    kp(f'Latest actual month {d["as_of"]}', f'INR {lk(last["revenue"])} lakh', f'margin {pct(last["gm"])}; gross profit INR {lk(last["gp"])} lakh')])
    warn = (f'<div class="warn" role="alert"><b class="t">&#9888; DATA-QUALITY WARNING - budget is NOT comparable</b>'
            f'<p>The original <code>Sales_Budget</code> is about <b>{bud["ratio_vs_actual"]:.1f}x</b> actual net revenue for the same months, and <b>{pct(bud["share_in_cells_without_sales"], 0)}</b> of it sits in '
            f'Region x Category cells that have no sales at all ({bud["cells_without_sales"]} of {bud["cells_total"]} cells). The cause is <b>unresolved</b>. '
            f'<b>This dashboard therefore shows no budget-vs-actual variance.</b> Any variance in the v1 workbook\'s <code>Budget_v2_Synthetic</code> / <code>Budget_vs_Actual_v2</code> sheets is <b>illustrative only</b> '
            f'(calibrated to prior-year actuals, not an approved budget) and must not be used for performance assessment.</p></div>')
    scope = (f'<div class="note"><b>Scope.</b> Synthetic data only (fictional company). Actuals only through {E(d["as_of"])}; forecast months excluded. Net revenue excludes GST, returns, rebates and intercompany. '
             f'YTD = calendar months 1-{M} of {Y} versus the same months of {Y - 1} (the source calendar has no separate fiscal year; calendar-year YTD is an assumption). '
             f'Static snapshot rebuilt manually: <b>no live refresh, no autonomous publishing</b>.</div>')
    trend1 = stacked_svg(mo)
    gmv = [m['gm'] for m in mo]; lo, hi = min(gmv), max(gmv); ymin, ymax = nice_ticks(lo - 0.003, hi + 0.003, 5)[0], nice_ticks(lo - 0.003, hi + 0.003, 5)[-1]
    trend2 = line_svg(labels, [('Gross margin', 's3', gmv)], ymin, ymax, lambda v: pct(v, 1), yname='Monthly gross margin percent', plabels=True)
    mtab = table(['Month', 'Net revenue (INR lakh)', 'COGS', 'Gross profit', 'Gross margin'], [[m['month'], lk(m['revenue']), lk(m['cogs']), lk(m['gp']), pct(m['gm'])] for m in mo])
    ybars = ytd_bars_svg([('Net revenue', yc['revenue'], yp['revenue'], lk), ('COGS', yc['cogs'], yp['cogs'], lk), ('Gross profit', yc['gp'], yp['gp'], lk)])
    ytab = table(['Measure (INR lakh)', f'{Y} YTD', f'{Y - 1} YTD', 'Change'], [['Net revenue', lk(yc['revenue']), lk(yp['revenue']), sg(ch['revenue'])], ['COGS', lk(yc['cogs']), lk(yp['cogs']), sg(ch['cogs'])],
                 ['Gross profit', lk(yc['gp']), lk(yp['gp']), sg(ch['gp'])], ['Gross margin', pct(yc['gm']), pct(yp['gm']), f'{ch["gm_pp"]:+.1f} pp']])
    def dimsec(rows, key_label, extra_head, extra_cell):
        rows = sorted(rows, key=lambda r: -r['revenue'])
        bars = cols_svg([(r['name'], r['revenue'], lk(r['revenue']), pct(r['gm']), f'{r["name"]}: net revenue INR {lk(r["revenue"])} lakh, gross margin {pct(r["gm"])}') for r in rows])
        tb = table(cls='tx', head=[key_label, extra_head, 'Net revenue (INR lakh)', 'Share', 'Gross margin', f'{Y} YTD revenue', f'{Y - 1} YTD revenue', 'YTD change', f'YTD margin {Y} / {Y - 1}'],
                   rows=[[E(r['name']), extra_cell(r), lk(r['revenue']), pct(r['share']), pct(r['gm']), lk(r['ytd_revenue']), lk(r['ytd_revenue_prior']), sg(r['ytd_yoy']), f'{pct(r["ytd_gm"])} / {pct(r["ytd_gm_prior"])}'] for r in rows])
        return bars, f'<div class="scroll">{tb}</div>'
    rb, rt = dimsec(d['regions'], 'Region', 'Products sold', lambda r: E(', '.join(r['products'])))
    pb, pt = dimsec(d['products'], 'Product', 'Region sold in', lambda r: E(', '.join(r['regions'])))
    overlap = (f'<div class="note"><b>Caveat - region and product effects overlap.</b> In this dataset each product is sold in exactly {ov["max_regions_per_product"]} region (max regions per product = {ov["max_regions_per_product"]}), '
               f'and each region carries only {ov["max_products_per_region"]} of the {ov["n_products"]} products. A region\'s result is therefore its two products\' result, and a product\'s result is its region\'s. '
               f'The two views are <b>confounded</b>: do not read a regional difference as a regional effect, or a product difference as a product effect. The data cannot separate them.</div>')
    legend_rg = '<div class="legend"><span><i class="solid"></i>Column = net revenue, INR lakh (label above); white figure inside = gross margin %. Actuals to ' + E(d['as_of']) + '</span></div>'
    refresh = f"""<pre>pip install pandas numpy openpyxl
python src/build_dashboard.py --input data/source/Nova_CFO_MIS_Complete_Synthetic_Dataset.xlsx --out outputs/v2
python tests/test_dashboard.py          # reconciliation checks</pre>
<p>To use a newer synthetic extract: place it under <code>data/</code>, first build and review the v1 MIS candidate with <code>python run_mis.py run</code> (human approval remains separate), then rerun the commands above on that file. See <code>docs/DASHBOARD_V2.md</code>.</p>"""
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Nova FP&A Dashboard v2</title><style>{CSS}</style></head><body><header class="hero"><div class="in"><h1>Nova Consumer Products - Executive FP&amp;A Dashboard</h1>
<p class="sub">Actuals to {E(d["as_of"])} &middot; INR &middot; built from the v1 synthetic source data</p>
<p><span class="tag">v2 dashboard enhancement (separate from Provisional v1)</span><span class="tag">Synthetic data</span><span class="tag">Static snapshot - not live</span><span class="tag">Not validated in Excel Desktop</span><span class="tag">Not production-ready</span></p></div></header>
<main>
{warn}{scope}
<div class="kpis">{kpis}</div>
<section class="card"><h2>Monthly net revenue, split into COGS and gross profit</h2><p class="sub">INR lakh, actuals, zero-based axis. Number above each column = net revenue; inside = gross profit (top, green) and COGS (bottom, blue). Months shown as MM/YY.</p><div class="legend"><span><i class="solid"></i>COGS</span><span><i class="solid gpk"></i>Gross profit</span></div>{trend1}</section>
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
    json.dump(d, open(os.path.join(a.out, 'dashboard_data.json'), 'w', newline='\n'), indent=1, sort_keys=True)
    open(os.path.join(a.out, 'Nova_FPA_Dashboard.html'), 'w', encoding='utf-8', newline='\n').write(render(d))   # LF on every OS (keeps output byte-identical)
    print(f'wrote {a.out}/Nova_FPA_Dashboard.html and dashboard_data.json (as of {d["as_of"]})')

if __name__ == '__main__': main()
