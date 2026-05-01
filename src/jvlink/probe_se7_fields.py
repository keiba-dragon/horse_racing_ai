# -*- coding: utf-8 -*-
"""
SE7 フィールド位置探索スクリプト
既知レース（中山1R 2026/3/21 ダ1200 未勝利）を基準に
距離・馬場・クラスのフィールド位置を特定する。

使い方:
  python src/jvlink/probe_se7_fields.py
"""
import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import pythoncom
pythoncom.CoInitialize()
import win32com.client as wc

TARGET_DATE   = "20260321"  # 中山1R：ダ1200 未勝利
TARGET_VENUE  = "06"        # 中山
TARGET_RACE   = "01"        # 1R

def main():
    jv = wc.gencache.EnsureDispatch("JVDTLab.JVLink")
    rc = jv.JVInit("UNKNOWN")
    if rc != 0:
        print("JVInit失敗")
        sys.exit(1)

    from_dt = TARGET_DATE + "000000"
    rc, readcnt, dldcnt, lts = jv.JVOpen("RACE", from_dt, 1, 0, 0, "")
    print(f"JVOpen: rc={rc} readcnt={readcnt}")

    buf = " " * 110000
    found = 0
    target_recs = []

    while found < 10:
        ret, data, sz, fname = jv.JVRead(buf, sz if False else 110000, "")
        if ret == 0:
            break
        if ret == -1:
            continue
        if ret == -3:
            time.sleep(0.05)
            continue
        if ret < 0:
            break

        rt = data[:2]
        if rt != "SE":
            continue

        # 当該レースのSEだけ
        kd = data[11:19].strip()
        vc = data[19:21].strip()
        rn = data[25:27].strip()

        if kd != TARGET_DATE or vc != TARGET_VENUE or rn != TARGET_RACE:
            continue

        target_recs.append(data[:ret])
        found += 1

    jv.JVClose()

    if not target_recs:
        print("対象レコードが見つかりません")
        return

    rec = target_recs[0]
    print(f"\n中山1R 2026/3/21 SE7  len={len(rec)}")
    print(f"馬名: {rec[40:58].strip()}")
    print()

    # 全フィールドを10文字ずつ表示
    print("=== 全フィールドダンプ ===")
    for i in range(0, min(len(rec), 560), 10):
        chunk = rec[i:i+10]
        print(f"  [{i:03d}:{i+10:03d}] {repr(chunk)}")

    print()
    print("=== '1200' 探索 ===")
    for i in range(len(rec) - 3):
        if rec[i:i+4] == '1200':
            print(f"  pos [{i}:{i+4}] → '1200' 前後: {repr(rec[max(0,i-5):i+9])}")

    print()
    print("=== '0120' '1201' '120 ' 探索（距離候補）===")
    for pat in ['0120', '1201', '120 ']:
        for i in range(len(rec) - 3):
            if rec[i:i+4] == pat:
                print(f"  pat={pat!r} pos [{i}:{i+4}] → {repr(rec[max(0,i-5):i+9])}")

    print()
    print("=== 既知位置周辺詳細 ===")
    print(f"  [058:075] → {repr(rec[58:75])}  ← 馬名直後")
    print(f"  [060:070] → {repr(rec[60:70])}")
    print(f"  [058:068] → {repr(rec[58:68])}")
    print(f"  [058:062] → {repr(rec[58:62])}  (馬名直後4文字)")
    print(f"  [058:060] → {repr(rec[58:60])}  (馬名直後2文字)")

    # 全レコード比較（複数頭）
    print()
    print("=== 同レース全頭 比較（距離はレース固定なので同値のはず）===")
    print(f"{'頭':>4} {'[58:62]':8} {'[62:66]':8} {'[66:70]':8} {'[70:74]':8} {'[74:78]':8} {'[78:82]':8}")
    for rec in target_recs:
        uma = rec[40:58].strip()[:6]
        print(f"{uma:>6} {rec[58:62]:8} {rec[62:66]:8} {rec[66:70]:8} {rec[70:74]:8} {rec[74:78]:8} {rec[78:82]:8}")

    # 芝ダ探索：中山1Rはダ=2想定
    print()
    print("=== 芝ダコード探索（'2'=ダ, '1'=芝 候補位置）===")
    for pos in [58, 59, 60, 61, 62, 63, 64, 65, 27, 30, 31, 32]:
        c = rec[pos] if pos < len(rec) else '?'
        print(f"  pos[{pos}] = {repr(c)}")

if __name__ == '__main__':
    main()
