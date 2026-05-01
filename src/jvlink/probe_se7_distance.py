# -*- coding: utf-8 -*-
"""
SE7 距離フィールドの体系的探索
同レース内で全頭が同値（レース固定）で、かつ別レースと異なる位置を探す。
中山01R (ダ1200) vs 中山02R (芝?) vs 中山05R (距離不明) を比較。
"""
import sys, io, time, collections
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import pythoncom
pythoncom.CoInitialize()
import win32com.client as wc

# 既知ペア: 3/21中山の複数レースを距離で識別できるかチェック
WATCH_RACES = [
    ("20260321", "06", "01"),  # 中山1R  ダ1200
    ("20260321", "06", "02"),  # 中山2R  未知
    ("20260321", "06", "05"),  # 中山5R  未知
    ("20260321", "06", "11"),  # 中山11R 未知
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
        key = (kd, vc, rn)
        if any(kd == w[0] and vc == w[1] and rn == w[2] for w in WATCH_RACES):
            race_recs[key].append(data[:ret])

    jv.JVClose()

    # レースごとに「全頭共通の値」を収集（ASCII数字のみ、4文字）
    race_constants = {}  # key → {pos: value}
    for key in WATCH_RACES:
        key = tuple(key)
        recs = race_recs.get(key, [])
        if len(recs) < 3:
            print(f"{key}: データ不足 ({len(recs)}頭)")
            continue
        # 全頭で共通のASCII値を持つ位置を探す
        L = min(len(r) for r in recs)
        common = {}
        for pos in range(L - 3):
            vals = set()
            for rec in recs:
                vals.add(rec[pos:pos+4])
            if len(vals) == 1:
                v = next(iter(vals))
                # ASCII数字のみ、かつ全部スペースでない
                if v.isdigit() or (v.strip() and all(c.isdigit() or c == ' ' for c in v)):
                    common[pos] = v
        race_constants[tuple(key)] = common

    # レース間で「値が異なる」位置 = 距離候補
    print()
    print("=== レース間で値が異なるASCII4桁固定フィールド ===")
    all_keys = [k for k in WATCH_RACES if tuple(k) in race_constants]
    if len(all_keys) < 2:
        print("レースデータ不足")
        return

    # 全レースに存在し、かつレース間で値が異なる位置
    first = race_constants[tuple(all_keys[0])]
    diff_positions = {}
    for pos, v0 in first.items():
        vals = {tuple(all_keys[0]): v0}
        for key in all_keys[1:]:
            rc2 = race_constants.get(tuple(key), {})
            if pos in rc2:
                vals[tuple(key)] = rc2[pos]
        if len(vals) == len(all_keys) and len(set(vals.values())) > 1:
            diff_positions[pos] = vals

    print(f"{'pos':>5}  " + "  ".join(f"{k[2]}R" for k in all_keys))
    for pos in sorted(diff_positions.keys()):
        row = diff_positions[pos]
        vals_str = "  ".join(f"[{row.get(tuple(k), '????')}]" for k in all_keys)
        print(f"{pos:>5}  {vals_str}")

    # さらに、4桁数字(800-3600)に限定
    print()
    print("=== 距離候補(800-3600のみ) ===")
    for pos in sorted(diff_positions.keys()):
        row = diff_positions[pos]
        all_numeric = all(v.isdigit() and 800 <= int(v) <= 3600 for v in row.values())
        if all_numeric:
            vals_str = "  ".join(f"[{row.get(tuple(k), '????')}]" for k in all_keys)
            print(f"{pos:>5}  {vals_str}")

    # RA レコードも確認
    print()
    print("=== RA レコード 再確認（3/21 中山） ===")
    jv2 = wc.gencache.EnsureDispatch("JVDTLab.JVLink")
    jv2.JVInit("UNKNOWN")
    rc2, readcnt2, dldcnt2, lts2 = jv2.JVOpen("RACE", "20260321000000", 1, 0, 0, "")
    buf2 = " " * 110000
    ra_recs = []
    while len(ra_recs) < 5:
        ret2, data2, sz2, fname2 = jv2.JVRead(buf2, 110000, "")
        if ret2 == 0: break
        if ret2 < 0 and ret2 not in (-1, -3): break
        if ret2 <= 0: continue
        if data2[:2] == "RA":
            kd = data2[11:19].strip()
            vc = data2[19:21].strip()
            rn = data2[25:27].strip()
            if kd == "20260321" and vc == "06" and rn in ["01","02","05","11"]:
                ra_recs.append((rn, data2[:ret2]))
    jv2.JVClose()

    for rn, rec in ra_recs:
        print(f"\n中山{rn}R RA  len={len(rec)}")
        # 非スペース箇所を探す
        non_space = [(i, rec[i]) for i in range(len(rec)) if rec[i] not in (' ', '　', '\r', '\n')]
        print(f"  非スペース位置: {non_space[:60]}")
        # 4桁数字(800-3600)スキャン
        dist_hits = []
        for i in range(len(rec)-3):
            s = rec[i:i+4]
            if s.isdigit() and 800 <= int(s) <= 3600:
                dist_hits.append((i, s))
        if dist_hits:
            print(f"  距離候補(RA): {dist_hits[:20]}")
        else:
            print("  距離候補なし")

if __name__ == '__main__':
    main()
