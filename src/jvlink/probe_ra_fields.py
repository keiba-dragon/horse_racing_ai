# -*- coding: utf-8 -*-
"""
RA レコードの距離・芝ダ・クラスフィールド位置探索
4/19/2026 データを netkeibanote CSVと照合して確定する。
"""
import sys, io, time, collections
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import pythoncom
pythoncom.CoInitialize()
import win32com.client as wc
import pandas as pd, os

BASE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# netkeibanote 4/19 CSVから正解データを読む
CSV_PATH = os.path.join(BASE, 'data/raw/results/出馬表形式4月19日結果確認用.csv')
nk = pd.read_csv(CSV_PATH, encoding='cp932')

# netkeibanote の場R → JV venue code
NK_VENUE = {'福': '03', '新': '04', '東': '05', '中': '06', '中京': '07', '京': '08', '阪': '09', '小': '10', '札': '01', '函': '02'}

def nk_to_venue(ba_r):
    """'福1' → venue='03', race_no='01'"""
    ba_r = str(ba_r).strip()
    for name, code in sorted(NK_VENUE.items(), key=lambda x: -len(x[0])):
        if ba_r.startswith(name):
            rn = ba_r[len(name):]
            return code, f"{int(rn):02d}"
    return None, None

# netkeibanote の正解テーブル: (venue, race_no) → (距離, 芝ダ, トラックコードJV)
truth = {}
df_unique = nk[['場 R','芝ダ','距離','トラックコード(JV)']].drop_duplicates()
for _, row in df_unique.iterrows():
    v, rn = nk_to_venue(row['場 R'])
    if v and rn:
        truth[(v, rn)] = {
            'dist': int(row['距離']) if pd.notna(row['距離']) else None,
            'surface': row['芝ダ'] if pd.notna(row['芝ダ']) else None,
            'track_jv': int(row['トラックコード(JV)']) if pd.notna(row['トラックコード(JV)']) else None,
        }

print("netkeibanote 正解テーブル (4/19):")
for k, v in sorted(truth.items())[:20]:
    print(f"  venue={k[0]} R={k[1]}: {v}")

# JV-Link RA レコードを取得
print("\nJVLink RA 取得中...")
jv = wc.gencache.EnsureDispatch("JVDTLab.JVLink")
jv.JVInit("UNKNOWN")
rc, readcnt, dldcnt, lts = jv.JVOpen("RACE", "20260419000000", 1, 0, 0, "")
print(f"JVOpen: rc={rc} readcnt={readcnt}")

buf = " " * 110000
ra_recs = {}  # (venue, race_no) → RA record

while True:
    ret, data, sz, fname = jv.JVRead(buf, 110000, "")
    if ret == 0: break
    if ret == -1: continue
    if ret == -3: time.sleep(0.05); continue
    if ret < 0: break
    if data[:2] != "RA": continue
    kd = data[11:19].strip()
    if kd != "20260419": continue
    vc = data[19:21].strip()
    rn = data[25:27].strip()
    ra_recs[(vc, rn)] = data[:ret]

jv.JVClose()
print(f"RA レコード取得数: {len(ra_recs)}")

# 照合して距離・芝ダ位置を確認
print()
print("=== RA フィールド照合 ===")
print(f"{'venue':>5} {'R':>3} {'正解距離':>6} {'正解芝ダ':>5} {'TrkJV':>5} | RA[558:562] RA[33:37] RA[559:563] RA[560:564]")
print("-" * 80)

match_dist = total = 0
for (vc, rn) in sorted(ra_recs.keys()):
    rec = ra_recs[(vc, rn)]
    t = truth.get((vc, rn), {})
    if not t: continue
    total += 1
    d_ra558 = rec[558:562] if len(rec) > 562 else '????'
    d_ra33  = rec[33:37]   if len(rec) > 37  else '????'
    d_ra559 = rec[559:563] if len(rec) > 563 else '????'
    d_ra560 = rec[560:564] if len(rec) > 564 else '????'

    actual_dist = str(t.get('dist',''))
    if d_ra558 == actual_dist.zfill(4):
        match_dist += 1
        flag = " ✓"
    else:
        flag = ""

    JYO = {"03":"福","04":"新","05":"東","06":"中","07":"中京","08":"京","09":"阪","10":"小"}
    jyo = JYO.get(vc, vc)
    print(f"  {jyo}{rn}  {t.get('dist','?'):>6}  {t.get('surface','?'):>5}  {t.get('track_jv','?'):>5}  | {d_ra558}  {d_ra33}  {d_ra559}  {d_ra560}{flag}")

print(f"\n距離マッチ率: {match_dist}/{total} = {match_dist/total*100:.1f}% (RA[558:562])")

# トラックコード探索: 17,18,24 を全RA内でスキャン
print()
print("=== RA内トラックコード(17,18,24,52)探索 ===")
for (vc, rn), rec in sorted(ra_recs.items()):
    t = truth.get((vc, rn), {})
    if not t: continue
    expected_track = str(t.get('track_jv', '')).zfill(2)
    if not expected_track.strip('0'): continue

    # 2桁数字として探す
    hits = [i for i in range(len(rec)-1) if rec[i:i+2] == expected_track and (i==0 or not rec[i-1].isdigit()) and (i+2>=len(rec) or not rec[i+2].isdigit())]
    JYO = {"03":"福","04":"新","05":"東","06":"中","07":"中京","08":"京","09":"阪","10":"小"}
    jyo = JYO.get(vc, vc)
    if hits:
        print(f"  {jyo}{rn} 芝ダ={t.get('surface')} track={expected_track}: pos={hits[:10]}")

if __name__ == '__main__':
    pass  # already ran
