# -*- coding: utf-8 -*-
"""
SE7 同レース全頭比較 - 距離・芝ダ候補フィールドが頭ごとに変わるか調べる
中山1R ダ1200 (3/21/2026) と中山2R (芝?) を並べる
"""
import sys, io, time, collections
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import pythoncom
pythoncom.CoInitialize()
import win32com.client as wc

WATCH = [
    ("20260321", "06", "01"),  # 中山1R ダ1200
    ("20260321", "06", "02"),  # 中山2R 何か
    ("20260321", "06", "11"),  # 中山11R 何か
]

def main():
    jv = wc.gencache.EnsureDispatch("JVDTLab.JVLink")
    jv.JVInit("UNKNOWN")

    rc, readcnt, dldcnt, lts = jv.JVOpen("RACE", "20260321000000", 1, 0, 0, "")
    print(f"JVOpen: rc={rc} readcnt={readcnt}")

    buf = " " * 110000
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
        if kd == "20260321" and (vc, rn) in [("06","01"), ("06","02"), ("06","11")]:
            race_recs[(vc, rn)].append(data[:ret])

    jv.JVClose()

    JYO = {"06": "中山"}

    for (vc, rn) in [("06","01"), ("06","02"), ("06","11")]:
        recs = race_recs[(vc, rn)]
        if not recs:
            print(f"\n{JYO.get(vc,vc)}{rn}R: データなし")
            continue
        print(f"\n{'='*70}")
        print(f"中山{rn}R  計{len(recs)}頭  [60:62]={repr(recs[0][60:62])}")
        print(f"{'馬名':12} [360:370]       [360:364] [364:368] [368:372]")
        for rec in recs:
            uma = rec[40:58].strip()[:8]
            print(f"  {uma:10} {repr(rec[360:370]):16} {repr(rec[360:364]):10} {repr(rec[364:368]):10} {repr(rec[368:372]):10}")

        # 距離候補として800-3600の4桁数字をスキャン（全頭に共通の値を探す）
        print()
        print("  全頭に共通の4桁数字(800-3600)候補:")
        common = None
        for rec in recs:
            cands = set()
            for i in range(50, min(len(rec)-3, 380)):
                s = rec[i:i+4]
                if s.isdigit() and 800 <= int(s) <= 3600:
                    cands.add((i, s))
            if common is None:
                common = cands
            else:
                common &= cands
        if common:
            for pos, val in sorted(common):
                print(f"    pos[{pos}:{pos+4}] = '{val}'")
        else:
            print("    なし（全頭共通の距離候補なし）")

if __name__ == '__main__':
    main()
