# -*- coding: utf-8 -*-
"""
JG の race_no フィールド確認
3/21 中山RA race_no='01','11' に対応するJGを探す
"""
import sys, io, time, collections
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import pythoncom
pythoncom.CoInitialize()
import win32com.client as wc

def main():
    jv = wc.gencache.EnsureDispatch("JVDTLab.JVLink")
    jv.JVInit("UNKNOWN")
    rc, readcnt, dldcnt, lts = jv.JVOpen("RACE", "20260321000000", 1, 0, 0, "")
    print(f"JVOpen rc={rc} readcnt={readcnt}")

    buf = " " * 110000
    # 中山(06)のJGをすべて集める
    jg_06 = collections.defaultdict(list)  # rec[25:27] → [(horse_name, blood_id)]
    ra_06 = {}  # race_no → RA record

    while True:
        ret, data, sz, fname = jv.JVRead(buf, 110000, "")
        if ret == 0: break
        if ret == -1: continue
        if ret == -3: time.sleep(0.05); continue
        if ret < 0: break

        rt = data[:2]
        kd = data[11:19].strip()
        if kd != "20260321": continue
        vc = data[19:21].strip()
        if vc != "06": continue  # 中山のみ

        if rt == "JG":
            rn = data[25:27].strip()
            horse = data[37:55].strip() if len(data) > 55 else '?'
            blood = data[29:37].strip() if len(data) > 37 else '?'
            jg_06[rn].append((horse, blood, data[:ret]))
        elif rt == "RA":
            rn = data[25:27].strip()
            ra_06[rn] = data[:ret]
        elif rt == "SE":
            rn = data[25:27].strip()
            # SEの馬名も参考に
            horse = data[40:58].strip() if len(data) > 58 else '?'

    jv.JVClose()

    print("\n=== RA race_no ===")
    for rn in sorted(ra_06.keys()):
        rec = ra_06[rn]
        dist = rec[558:562] if len(rec) > 562 else '???'
        track = rec[566:568] if len(rec) > 568 else '??'
        print(f"  RA race_no={rn}  dist={dist}  track={track}")

    print("\n=== JG race_no ===")
    for rn in sorted(jg_06.keys()):
        entries = jg_06[rn]
        print(f"  JG race_no={rn}  頭数={len(entries)}")
        for horse, blood, rec in entries[:3]:
            print(f"    {horse} (blood={blood})")
        if len(entries) > 3:
            print(f"    ... +{len(entries)-3}")

    print("\n=== 対応確認 ===")
    ra_race_nos = sorted(ra_06.keys())
    jg_race_nos = sorted(jg_06.keys())
    print(f"RA: {ra_race_nos}")
    print(f"JG: {jg_race_nos}")

    # JG race_no vs RA race_no の対応
    for ra_rn in ra_race_nos:
        # 直接対応
        if ra_rn in jg_06:
            n = len(jg_06[ra_rn])
            print(f"  RA[{ra_rn}] → JG[{ra_rn}] ({n}頭) ✓")
        else:
            # +1 オフセット試行
            jg_rn_minus1 = f"{int(ra_rn)-1:02d}"
            if jg_rn_minus1 in jg_06:
                n = len(jg_06[jg_rn_minus1])
                print(f"  RA[{ra_rn}] → JG[{jg_rn_minus1}] ({n}頭) (offset -1)")
            else:
                print(f"  RA[{ra_rn}] → JG: not found")

if __name__ == '__main__':
    main()
