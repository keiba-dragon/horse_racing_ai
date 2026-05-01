# -*- coding: utf-8 -*-
"""
SE7 距離・芝ダ フィールド交差検証
複数の既知レースで pos[364:368]=距離, pos[60]=芝ダ を確認する。

使い方:
  python src/jvlink/probe_se7_crosscheck.py
"""
import sys, io, time, collections
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import pythoncom
pythoncom.CoInitialize()
import win32com.client as wc

# 検証対象：（日付, 場コード, R番号, 期待距離, 期待芝ダ）
# netkeibanote確認済みレース
TARGETS = [
    ("20260321", "06", "01", "1200", "ダ"),  # 中山1R ダ1200
    ("20260321", "06", "07", None,   None),  # 中山7R → 何か
    ("20260321", "05", "01", None,   None),  # 東京1R → 何か
    ("20260104", "06", "01", None,   None),  # 中山1R 1/4
    ("20260104", "05", "01", None,   None),  # 東京1R 1/4（芝があれば）
]

def main():
    jv = wc.gencache.EnsureDispatch("JVDTLab.JVLink")
    rc = jv.JVInit("UNKNOWN")
    if rc != 0:
        print("JVInit失敗"); sys.exit(1)

    # 3/21前後を一括取得
    from_dt = "20260101000000"
    rc, readcnt, dldcnt, lts = jv.JVOpen("RACE", from_dt, 1, 0, 0, "")
    print(f"JVOpen: rc={rc} readcnt={readcnt}")

    buf = " " * 110000
    # (date, venue, race_no) → first SE record
    race_recs = collections.defaultdict(list)

    while True:
        ret, data, sz, fname = jv.JVRead(buf, 110000, "")
        if ret == 0: break
        if ret == -1: continue
        if ret == -3: time.sleep(0.05); continue
        if ret < 0: break

        if data[:2] != "SE": continue
        kd = data[11:19].strip()
        vc = data[19:21].strip()
        rn = data[25:27].strip()
        key = (kd, vc, rn)
        if len(race_recs[key]) < 3:
            race_recs[key].append(data[:ret])

    jv.JVClose()

    JYO = {"05":"東京","06":"中山","07":"中京","08":"京都","09":"阪神","10":"小倉","01":"札幌","04":"新潟"}

    print()
    print(f"{'レース':<12} {'[60:62]':8} {'[360:370]':12} {'[364:368]':8} {'[362:366]':8} {'[58:62]':8}")
    print("-" * 70)

    for kd, vc, rn in sorted(race_recs.keys()):
        recs = race_recs[(kd, vc, rn)]
        if not recs: continue
        rec = recs[0]
        jyo = JYO.get(vc, vc)
        label = f"{kd[2:]}/{jyo}{rn}R"
        if len(rec) < 370: continue
        print(f"{label:<12} {repr(rec[60:62]):8} {repr(rec[360:370]):14} {repr(rec[364:368]):8} {repr(rec[362:366]):8} {repr(rec[58:62]):8}")

    # 詳細表示
    print()
    print("=== 各レース先頭馬 詳細 ===")
    for kd, vc, rn in sorted(race_recs.keys()):
        recs = race_recs[(kd, vc, rn)]
        if not recs: continue
        rec = recs[0]
        if len(rec) < 370: continue
        jyo = JYO.get(vc, vc)
        uma = rec[40:58].strip()
        print(f"\n[{kd}/{jyo}{rn}R] 馬:{uma}")
        print(f"  [58:68]  = {repr(rec[58:68])}")
        print(f"  [60:62]  = {repr(rec[60:62])}  ← 芝ダ候補")
        print(f"  [355:375]= {repr(rec[355:375])}")
        print(f"  [360:370]= {repr(rec[360:370])}")
        print(f"  [364:368]= {repr(rec[364:368])}  ← 距離候補")
        # 4桁数字スキャン（300-380付近）
        hits = []
        for i in range(300, min(len(rec)-3, 380)):
            s = rec[i:i+4]
            if s.isdigit() and 800 <= int(s) <= 3600:
                hits.append(f"[{i}:{i+4}]={s}")
        if hits:
            print(f"  距離候補(800-3600): {', '.join(hits)}")

if __name__ == '__main__':
    main()
