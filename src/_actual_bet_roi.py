# coding: utf-8
"""実際の馬券購入ROIレポート -> G:/マイドライブ/競馬AI/actual_bet_roi.html
data/tohyo/ 以下の投票CSVを集計して日別・式別ROIを表示する。
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import pandas as pd
import numpy as np
import os, glob

base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ── CSVロード ──────────────────────────────────────────────────
def read_tohyo(f):
    for enc in ['utf-8-sig', 'cp932', 'utf-8']:
        try: return pd.read_csv(f, encoding=enc)
        except: pass
    return None

files = sorted(glob.glob(os.path.join(base_dir, 'data', 'tohyo', '*.csv')))
dfs = [read_tohyo(f) for f in files]
df = pd.concat([d for d in dfs if d is not None], ignore_index=True)

# 合計行（式別がNaN）を除外
df = df[df['式別'].notna()].copy()

def parse_amount(s):
    """300／900 形式は合計（後ろ）を使う、通常数値はそのまま"""
    s = str(s).strip()
    if '／' in s:
        return float(s.split('／')[-1])
    try: return float(s)
    except: return 0.0

df['購入金額'] = df['購入金額'].apply(parse_amount)
df['払戻金額'] = pd.to_numeric(df['払戻金額'], errors='coerce').fillna(0)
df['返還金額'] = pd.to_numeric(df['返還金額'], errors='coerce').fillna(0)
df['日付_num'] = pd.to_numeric(df['日付'], errors='coerce')
# 返還分は投資から除く（未的中でも返還 → 実質投資 = 購入 - 返還）
df['実質投資'] = df['購入金額'] - df['返還金額']
df['損益']    = df['払戻金額'] - df['実質投資']

# 式別グループ（表示用に整理）
def shiki_group(s):
    s = str(s)
    if '単勝' in s: return '単勝'
    if '複勝' in s: return '複勝'
    if 'ワイド' in s: return 'ワイド'
    if '馬連' in s: return '馬連'
    if '馬単' in s: return '馬単'
    if '枠連' in s: return '枠連'
    if '3連複' in s or '３連複' in s: return '3連複'
    if '3連単' in s or '３連単' in s: return '3連単'
    return 'その他'

df['式別G'] = df['式別'].apply(shiki_group)

d = str(int(df['日付_num'].max()))
last_date = f"{d[:4]}/{d[4:6]}/{d[6:8]}"

print(f"読込: {len(files)}ファイル / {len(df)}行")
print(f"日付範囲: {df['日付_num'].min():.0f} ~ {df['日付_num'].max():.0f}")

# ── 日別集計 ───────────────────────────────────────────────────
daily_rows = []
cum_pf = 0
for dnum, grp in df.groupby('日付_num'):
    invest = int(grp['実質投資'].sum())
    ret    = int(grp['払戻金額'].sum())
    pf     = ret - invest
    roi    = ret / invest - 1.0 if invest > 0 else 0
    cum_pf += pf

    # 式別内訳
    shiki_detail = {}
    for sg, sg_grp in grp.groupby('式別G'):
        si = int(sg_grp['実質投資'].sum())
        sr = int(sg_grp['払戻金額'].sum())
        if si > 0:
            shiki_detail[sg] = {'投資': si, '回収': sr, 'pf': sr-si, 'roi': sr/si-1}

    d_str = str(int(dnum))
    date_str = f"{d_str[:4]}/{d_str[4:6]}/{d_str[6:8]}"
    daily_rows.append({
        '日付': date_str, '日付_num': dnum,
        '投資': invest, '回収': ret, '損益': pf, 'ROI': roi,
        '累計損益': cum_pf,
        '式別': shiki_detail,
    })
    sign = '+' if pf >= 0 else ''
    print(f"{date_str}  投資{invest:,}円  回収{ret:,}円  {sign}{pf:,}円 ({roi:+.1%})")

# ── 式別累計集計 ───────────────────────────────────────────────
shiki_total = {}
for sg, sg_grp in df.groupby('式別G'):
    si = int(sg_grp['実質投資'].sum())
    sr = int(sg_grp['払戻金額'].sum())
    hits = int((sg_grp['払戻金額'] > 0).sum())
    bets = int((sg_grp['実質投資'] > 0).sum())
    if si > 0:
        shiki_total[sg] = {'投資': si, '回収': sr, 'pf': sr-si, 'roi': sr/si-1,
                           'hits': hits, 'bets': bets}

total_invest = sum(r['投資'] for r in daily_rows)
total_ret    = sum(r['回収'] for r in daily_rows)
total_pf     = total_ret - total_invest
total_roi    = total_ret / total_invest - 1.0 if total_invest > 0 else 0
plus_days    = sum(1 for r in daily_rows if r['損益'] >= 0)
total_days   = len(daily_rows)

print(f"\n累計: {('+' if total_pf>=0 else '')}{total_pf:,}円  ROI{total_roi:+.1%}  {plus_days}/{total_days}日プラス")

# ── HTML生成（keiba-dragon.github.io のライト×ブルー基調に合わせる）──────
SHIKI_ORDER = ['単勝','複勝','ワイド','馬連','枠連','馬単','3連複','3連単','その他']
GREEN = '#16a34a'
RED   = '#dc2626'
BLUE  = '#2563eb'
col_total = GREEN if total_pf >= 0 else RED

def pf_td(pf, roi):
    sign = '+' if pf >= 0 else ''
    col  = GREEN if pf >= 0 else RED
    return f'<td style="color:{col};font-weight:700;text-align:right">{sign}{pf:,}円<br><small>({roi:+.1%})</small></td>'

def cum_td(pf):
    col = GREEN if pf >= 0 else RED
    sign = '+' if pf >= 0 else ''
    return f'<td style="color:{col};font-weight:700;text-align:right">{sign}{pf:,}円</td>'

# 式別ミニバッジ
def shiki_badges(detail):
    badges = []
    for sg in SHIKI_ORDER:
        if sg not in detail: continue
        d = detail[sg]
        col = GREEN if d['pf'] >= 0 else RED
        sign = '+' if d['pf'] >= 0 else ''
        badges.append(
            f'<span style="display:inline-block;margin:1px 2px;padding:1px 6px;'
            f'border-radius:999px;font-size:9px;font-weight:600;background:#eff6ff;color:{col};border:1px solid #dbeafe">'
            f'{sg} {sign}{d["pf"]:,}</span>'
        )
    return ''.join(badges)

# 日別行
rows_html = ''
for r in daily_rows:
    rows_html += f'''<tr>
<td style="text-align:center;white-space:nowrap">{r["日付"]}</td>
<td style="text-align:right">{r["投資"]:,}円</td>
<td style="text-align:right">{r["回収"]:,}円</td>
{pf_td(r["損益"], r["ROI"])}
<td style="text-align:left;font-size:9px">{shiki_badges(r["式別"])}</td>
{cum_td(r["累計損益"])}
</tr>'''

# 式別サマリー行
shiki_rows = ''
for sg in SHIKI_ORDER:
    if sg not in shiki_total: continue
    s = shiki_total[sg]
    hit_rate = f'{s["hits"]}/{s["bets"]}' if s["bets"] > 0 else '-'
    col = GREEN if s['pf'] >= 0 else RED
    sign = '+' if s['pf'] >= 0 else ''
    shiki_rows += f'''<tr>
<td style="font-weight:700">{sg}</td>
<td style="text-align:right">{s["投資"]:,}円</td>
<td style="text-align:right">{s["回収"]:,}円</td>
<td style="color:{col};font-weight:700;text-align:right">{sign}{s["pf"]:,}円</td>
<td style="color:{col};font-weight:700;text-align:right">{s["roi"]:+.1%}</td>
<td style="text-align:center">{hit_rate}</td>
</tr>'''

# ── グラフ用データ ──────────────────────────────────────────────
import json
chart_labels = [r['日付'][5:] for r in daily_rows]
chart_daily  = [r['損益'] for r in daily_rows]
chart_cumul  = [r['累計損益'] for r in daily_rows]
chart_colors = [GREEN if v >= 0 else RED for v in chart_daily]
# 式別ごとの日別 投資/回収（グラフの絞り込み用）
daily_by_shiki = [
    {sg: [d['投資'], d['回収']] for sg, d in r['式別'].items()}
    for r in daily_rows
]
shiki_checkboxes = ''.join(
    f'<label class="chk"><input type="checkbox" value="{sg}" checked>{sg}</label>'
    for sg in SHIKI_ORDER
)

html = f'''<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>実際の馬券ROI 2026</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4/dist/chart.umd.min.js"></script>
<style>
:root{{--blue:{BLUE};--ink:#1e293b;--muted:#64748b;--border:#e2e8f0;--bg2:#f8fafc}}
*{{box-sizing:border-box}}
body{{font-family:"Hiragino Kaku Gothic Pro","Noto Sans JP",Meiryo,sans-serif;background:#ffffff;color:var(--ink);padding:20px;max-width:960px;margin:0 auto}}
h2{{color:var(--ink);font-weight:800;margin-top:28px;font-size:1.3em}}
h3{{color:var(--blue);font-weight:800;margin-top:24px;font-size:1.05em}}
.updated{{color:var(--muted);font-size:0.85em;font-weight:500}}
.summary{{display:flex;gap:12px;justify-content:center;margin:14px 0 24px;flex-wrap:wrap}}
.card{{background:var(--bg2);border:1px solid var(--border);border-radius:16px;padding:14px 20px;text-align:center;min-width:110px;flex:1 1 110px;max-width:170px}}
.card .lbl{{color:var(--muted);font-size:0.8em;font-weight:600}}
.card .val{{font-size:1.4em;font-weight:800;margin-top:2px}}
.chart-wrap{{background:var(--bg2);border:1px solid var(--border);border-radius:16px;padding:14px;margin-bottom:20px}}
.chart-wrap canvas{{display:block;width:100%!important;height:220px!important}}
@media(min-width:600px){{.chart-wrap canvas{{height:260px!important}}}}
.filter-bar{{background:var(--bg2);border:1px solid var(--border);border-radius:16px;padding:12px 14px;margin-bottom:14px}}
.filter-bar .filter-title{{color:var(--muted);font-size:0.8em;font-weight:600;margin-bottom:8px}}
.filter-chks{{display:flex;flex-wrap:wrap;gap:6px 10px}}
.chk{{display:inline-flex;align-items:center;gap:4px;font-size:0.85em;font-weight:600;color:var(--ink);cursor:pointer;user-select:none}}
.chk input{{accent-color:var(--blue);width:15px;height:15px;cursor:pointer}}
.filter-btns{{display:flex;gap:8px;margin-top:10px}}
.filter-btns button{{font-family:inherit;font-size:0.8em;font-weight:700;color:var(--blue);background:#fff;border:1px solid var(--blue);border-radius:999px;padding:4px 12px;cursor:pointer}}
.filter-btns button:hover{{background:#eff6ff}}
.tbl-wrap{{overflow-x:auto;-webkit-overflow-scrolling:touch;margin-bottom:24px;border:1px solid var(--border);border-radius:16px}}
table{{width:100%;border-collapse:collapse;font-size:0.85em;white-space:nowrap}}
th{{background:var(--bg2);color:var(--blue);font-weight:800;padding:8px 10px;text-align:center;position:sticky;top:0}}
td{{padding:7px 10px;border-bottom:1px solid var(--border)}}
tr:hover{{background:#f1f5f9}}
@media(max-width:480px){{
  body{{padding:12px}}
  .card .val{{font-size:1.15em}}
  table{{font-size:0.78em}}
}}
</style></head><body>
<h2>実際の馬券ROI <span class="updated">（最終更新: {last_date}）</span></h2>
<div class="summary">
  <div class="card"><div class="lbl">累計損益</div><div class="val" style="color:{col_total}">{("+" if total_pf>=0 else "")}{total_pf:,}円</div></div>
  <div class="card"><div class="lbl">累計ROI</div><div class="val" style="color:{col_total}">{total_roi:+.1%}</div></div>
  <div class="card"><div class="lbl">総投資</div><div class="val">{total_invest:,}円</div></div>
  <div class="card"><div class="lbl">総回収</div><div class="val">{total_ret:,}円</div></div>
  <div class="card"><div class="lbl">プラス日数</div><div class="val">{plus_days}/{total_days}日</div></div>
</div>

<h3>グラフ（式別を選んで絞り込み）</h3>
<div class="filter-bar">
  <div class="filter-title">表示する式別</div>
  <div class="filter-chks" id="shikiChks">{shiki_checkboxes}</div>
  <div class="filter-btns">
    <button type="button" id="btnAll">全選択</button>
    <button type="button" id="btnNone">全解除</button>
  </div>
</div>
<div class="summary" id="filterSummary">
  <div class="card"><div class="lbl">絞込損益</div><div class="val" id="fSumPf">-</div></div>
  <div class="card"><div class="lbl">絞込ROI</div><div class="val" id="fSumRoi">-</div></div>
  <div class="card"><div class="lbl">絞込投資</div><div class="val" id="fSumInvest">-</div></div>
  <div class="card"><div class="lbl">絞込回収</div><div class="val" id="fSumRet">-</div></div>
</div>
<h4 style="color:var(--ink);font-size:0.95em;margin:0 0 4px">日別損益グラフ</h4>
<div class="chart-wrap"><canvas id="dailyChart"></canvas></div>
<h4 style="color:var(--ink);font-size:0.95em;margin:16px 0 4px">累計損益グラフ</h4>
<div class="chart-wrap"><canvas id="cumulChart"></canvas></div>

<h3>式別ROI</h3>
<div class="tbl-wrap"><table>
<thead><tr>
<th>式別</th><th>投資</th><th>回収</th><th>損益</th><th>ROI</th><th>的中/買い目数</th>
</tr></thead><tbody>
{shiki_rows}
</tbody></table></div>

<h3>日別損益</h3>
<div class="tbl-wrap"><table>
<thead><tr>
<th>日付</th><th>投資</th><th>回収</th><th>損益</th><th>式別内訳</th><th>累計損益</th>
</tr></thead><tbody>
{rows_html}
</tbody></table></div>

<script>
const labels = {json.dumps(chart_labels, ensure_ascii=False)};
// 日別・式別ごとの [投資, 回収]。式別がその日に無ければキー無し。
const dailyByShiki = {json.dumps(daily_by_shiki, ensure_ascii=False)};
const shikiOrder = {json.dumps(SHIKI_ORDER, ensure_ascii=False)};

const isMobile = window.innerWidth < 600;
const maxTicks = isMobile ? 8 : 20;
const ptRadius = isMobile ? 2 : 3;
const tickSz = isMobile ? 10 : 12;
const gridColor = '#e2e8f0';
const tickColor = '#64748b';
const GREEN = '{GREEN}', RED = '{RED}', BLUE = '{BLUE}';

function makeScales() {{
  return {{
    x: {{ ticks: {{ color: tickColor, maxRotation: 45, maxTicksLimit: maxTicks, font: {{ size: tickSz }} }}, grid: {{ color: gridColor }} }},
    y: {{ ticks: {{ color: tickColor, font: {{ size: tickSz }}, callback: v => (v>=0?'+':'')+v.toLocaleString()+'円' }}, grid: {{ color: gridColor }} }}
  }};
}}

const dailyChart = new Chart(document.getElementById('dailyChart'), {{
  type: 'bar',
  data: {{ labels: labels, datasets: [{{ label: '日別損益（円）', data: [], backgroundColor: [], borderRadius: 4 }}] }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ labels: {{ color: '#1e293b', font: {{ size: tickSz }} }} }},
      tooltip: {{ callbacks: {{ label: ctx => (ctx.raw >= 0 ? '+' : '') + ctx.raw.toLocaleString() + '円' }} }}
    }},
    scales: makeScales()
  }}
}});

const cumulChart = new Chart(document.getElementById('cumulChart'), {{
  type: 'line',
  data: {{ labels: labels, datasets: [{{
    label: '累計損益（円）', data: [],
    borderColor: BLUE, backgroundColor: 'rgba(37,99,235,0.08)',
    borderWidth: 2, fill: true, tension: 0.15,
    pointRadius: ptRadius, pointBackgroundColor: BLUE
  }}] }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ labels: {{ color: '#1e293b', font: {{ size: tickSz }} }} }},
      tooltip: {{ callbacks: {{ label: ctx => (ctx.raw >= 0 ? '+' : '') + ctx.raw.toLocaleString() + '円' }} }}
    }},
    scales: makeScales()
  }}
}});

function selectedShiki() {{
  return [...document.querySelectorAll('#shikiChks input:checked')].map(el => el.value);
}}

function recompute() {{
  const chosen = selectedShiki();
  let cum = 0, totalInvest = 0, totalRet = 0, plusDays = 0, activeDays = 0;
  const dailyPf = [], cumulPf = [], colors = [];
  dailyByShiki.forEach(dayMap => {{
    let invest = 0, ret = 0, hasAny = false;
    chosen.forEach(sg => {{
      if (dayMap[sg]) {{ invest += dayMap[sg][0]; ret += dayMap[sg][1]; hasAny = true; }}
    }});
    const pf = ret - invest;
    cum += pf;
    dailyPf.push(hasAny ? pf : 0);
    cumulPf.push(cum);
    colors.push(pf >= 0 ? GREEN : RED);
    if (hasAny) {{
      activeDays++;
      totalInvest += invest; totalRet += ret;
      if (pf >= 0) plusDays++;
    }}
  }});

  dailyChart.data.datasets[0].data = dailyPf;
  dailyChart.data.datasets[0].backgroundColor = colors;
  dailyChart.update();
  cumulChart.data.datasets[0].data = cumulPf;
  cumulChart.update();

  const pf = totalRet - totalInvest;
  const roi = totalInvest > 0 ? (totalRet / totalInvest - 1) : 0;
  const col = pf >= 0 ? GREEN : RED;
  const sign = pf >= 0 ? '+' : '';
  document.getElementById('fSumPf').textContent = sign + pf.toLocaleString() + '円';
  document.getElementById('fSumPf').style.color = col;
  document.getElementById('fSumRoi').textContent = (roi >= 0 ? '+' : '') + (roi*100).toFixed(1) + '%';
  document.getElementById('fSumRoi').style.color = col;
  document.getElementById('fSumInvest').textContent = totalInvest.toLocaleString() + '円';
  document.getElementById('fSumRet').textContent = totalRet.toLocaleString() + '円';
}}

document.getElementById('shikiChks').addEventListener('change', recompute);
document.getElementById('btnAll').addEventListener('click', () => {{
  document.querySelectorAll('#shikiChks input').forEach(el => el.checked = true);
  recompute();
}});
document.getElementById('btnNone').addEventListener('click', () => {{
  document.querySelectorAll('#shikiChks input').forEach(el => el.checked = false);
  recompute();
}});

recompute();
</script>
</body></html>'''

out = r'G:\マイドライブ\競馬AI\actual_bet_roi.html'
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"\n出力: {out}")
