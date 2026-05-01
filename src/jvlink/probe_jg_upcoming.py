# -*- coding: utf-8 -*-
"""
JG レコード（出馬表）の詳細フィールド探索
upcoming/historical 両方の JG を全フィールドダンプする。
"""
import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import pythoncom
pythoncom.CoInitialize()
import win32com.client as wc

def probe_jg(from_dt, label):
    print(f"\n{'='*60}")
    print(f"JG probe: {label}  from={from_dt}")
    print(f"{'='*60}")
    jv = wc.gencache.EnsureDispatch("JVDTLab.JVLink")
    jv.JVInit("UNKNOWN")
    rc, readcnt, dldcnt, lts = jv.JVOpen("RACE", from_dt, 1, 0, 0, "")
    print(f"JVOpen: rc={rc} readcnt={readcnt}")

    buf = " " * 110000
    jg_samples = []
    se_samples = []

    while len(jg_samples) < 5 or len(se_samples) < 3:
        ret, data, sz, fname = jv.JVRead(buf, 110000, "")
        if ret == 0: break
        if ret == -1: continue
        if ret == -3: time.sleep(0.05); continue
        if ret < 0: break
        rt = data[:2]
        if rt == "JG" and len(jg_samples) < 5:
            jg_samples.append((ret, data[:ret]))
        elif rt == "SE" and len(se_samples) < 3:
            se_samples.append((ret, data[:ret]))

    jv.JVClose()

    print(f"\nJG レコード {len(jg_samples)} 件:")
    for i, (ret, rec) in enumerate(jg_samples):
        print(f"\n  JG[{i+1}]  len={ret}")
        for j in range(0, min(ret, 200), 10):
            chunk = rec[j:j+10]
            print(f"    [{j:03d}:{j+10:03d}] {repr(chunk)}")
        # 全非スペースASCII
        ns = [(p, repr(rec[p])) for p in range(min(ret, len(rec))) if rec[p] not in (' ','　','\r','\n') and ord(rec[p]) < 256]
        print(f"  非スペースASCII: {ns[:40]}")
        # 数字スキャン(400-750 斤量候補)
        for p in range(len(rec)-2):
            s = rec[p:p+3]
            if s.isdigit() and 400 <= int(s) <= 750:
                print(f"  斤量候補(400-750) pos[{p}:{p+3}] = '{s}'")

    print(f"\nSE レコード比較 (斤量位置特定用):")
    for i, (ret, rec) in enumerate(se_samples):
        uma = rec[40:58].strip()
        print(f"\n  SE[{i+1}]  len={ret}  馬名={uma}")
        # 斤量候補（SE7の既知エリア）
        for p in range(60, min(ret, 120)):
            s = rec[p:p+3]
            if s.isdigit() and 400 <= int(s) <= 750:
                print(f"    pos[{p}:{p+3}] = '{s}'  context=[{p-3}:{p+6}]={repr(rec[max(0,p-3):p+6])}")

def main():
    # 過去 (4/19) のJGレコード
    probe_jg("20260419000000", "過去 4/19")

    # 直近 (4/26 週末) のJGレコード
    probe_jg("20260426000000", "4/26 週末")

if __name__ == '__main__':
    main()
