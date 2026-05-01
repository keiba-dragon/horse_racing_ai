# -*- coding: utf-8 -*-
"""
RA レコードの芝ダ（トラックコード）フィールド位置を特定する
4/19/2026 のRA全フィールドダンプ + track_jv対応位置を探す
"""
import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import pythoncom
pythoncom.CoInitialize()
import win32com.client as wc
import pandas as pd, os

BASE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
CSV_PATH = os.path.join(BASE, 'data/raw/results/出馬表形式4月19日結果確認用.csv')
nk = pd.read_csv(CSV_PATH, encoding='cp932')

NK_VENUE = {'福': '03', '新': '04', '東': '05', '中': '06', '中京': '07',
            '京': '08', '阪': '09', '小': '10', '札': '01', '函': '02'}

def nk_to_venue(ba_r):
    ba_r = str(ba_r).strip()
    for name, code in sorted(NK_VENUE.items(), key=lambda x: -len(x[0])):
        if ba_r.startswith(name):
            rn = ba_r[len(name):]
            return code, f"{int(rn):02d}"
    return None, None

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

jv = wc.gencache.EnsureDispatch("JVDTLab.JVLink")
jv.JVInit("UNKNOWN")
rc, readcnt, dldcnt, lts = jv.JVOpen("RACE", "20260419000000", 1, 0, 0, "")

buf = " " * 110000
ra_recs = {}

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

# 代表レース: 福01(ダ1150, track=24) と 福03(芝2000, track=17) を詳細ダンプ
SAMPLES = [("03","01"), ("03","03"), ("06","05"), ("09","04")]  # ダ、芝、芝、障

for (vc, rn) in SAMPLES:
    rec = ra_recs.get((vc, rn), None)
    if rec is None: continue
    t = truth.get((vc, rn), {})
    JYO = {"03":"福","04":"新","05":"東","06":"中","07":"中京","08":"京","09":"阪","10":"小"}
    jyo = JYO.get(vc, vc)
    print(f"\n{'='*60}")
    print(f"{jyo}{rn}R  dist={t.get('dist')}  surface={t.get('surface')}  track_jv={t.get('track_jv')}  len={len(rec)}")
    print(f"{'='*60}")
    # 540-580 付近を詳細表示
    print("pos[540:590]:")
    for i in range(540, min(590, len(rec)), 4):
        chunk = rec[i:i+4]
        print(f"  [{i}:{i+4}] = {repr(chunk)}")
    # 全非スペース位置（アスキーのみ）
    non_sp = [(i, repr(rec[i])) for i in range(len(rec))
              if rec[i] not in (' ','　','\r','\n') and ord(rec[i]) < 256]
    print(f"\n全ASCII非スペース位置 (先頭60件):")
    for pos, ch in non_sp[:60]:
        print(f"  [{pos}] = {ch}")

# 全レースで「track_jv」値が一意に特定できるポジションを探す
print()
print("="*60)
print("track_jv 2桁コードを単独で持つポジション探索")
print("="*60)

# 各レースで2桁コード(track_jv)が含まれる全ポジション
from collections import defaultdict
pos_to_vals = defaultdict(list)  # pos → list of (venue, rno, surface, track_jv, val_at_pos)

for (vc, rn), rec in sorted(ra_recs.items()):
    t = truth.get((vc, rn), {})
    if not t.get('track_jv'): continue
    track_str = f"{t['track_jv']:02d}"
    surface = t['surface']
    for i in range(len(rec) - 1):
        if rec[i:i+2] == track_str:
            pos_to_vals[i].append((vc, rn, surface, t['track_jv']))

# track_jvが芝ダで異なる値を持つ位置を探す
print("全レースでtrack_jvと一致し、芝とダで値が異なるポジション:")
for pos in sorted(pos_to_vals.keys()):
    entries = pos_to_vals[pos]
    # 全レース数に占める割合
    if len(entries) < 10: continue  # マイナー位置除外
    surfaces = {e[2] for e in entries}
    track_jvs = {e[3] for e in entries}
    if len(track_jvs) > 1 and len(surfaces) > 1:
        print(f"  pos[{pos}]: count={len(entries)} tracks={track_jvs} surfaces={surfaces}")
