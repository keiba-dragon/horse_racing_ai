# -*- coding: utf-8 -*-
"""
JVLink 過去レース結果一括取得
SE（競走成績）を取得して result CSV に保存する。

使い方:
  python src/jvlink/jvlink_fetch_history.py                      # 2025年全件
  python src/jvlink/jvlink_fetch_history.py --from 20250701      # 開始日指定
  python src/jvlink/jvlink_fetch_history.py --probe              # 生データ確認

出力:
  data/raw/results/jvlink/YYYYMMDD.csv  (日ごとに分割)

SE7 フィールド位置は 2026/3/21 中山7R で検証済み。
"""
import sys, io, os, csv, argparse, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import pythoncom
pythoncom.CoInitialize()
import win32com.client as wc
from collections import defaultdict

base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# netkeibanote の '場 R' 列に合わせた短縮名（PKL の race_key と一致させる）
JYO_NAME = {
    "01": "札",   # 札幌
    "02": "函",   # 函館
    "03": "福",   # 福島
    "04": "新",   # 新潟
    "05": "東",   # 東京
    "06": "中",   # 中山
    "07": "中京", # 中京（2文字）
    "08": "京",   # 京都
    "09": "阪",   # 阪神
    "10": "小",   # 小倉
}

# 検証済み SE7 フィールド位置（Python文字列スライス、0ベース）
# 2026-03-21 中山7R で着順・オッズ一致を確認済み
SE_POS = {
    'kaisai_year': (11, 15),  # YYYY
    'kaisai_md':   (15, 19),  # MMDD
    'venue_cd':    (19, 21),  # 場コード 01〜10
    'race_no':     (25, 27),  # レース番号 01〜12
    'umaban':      (28, 30),  # 馬番
    'horse_name':  (40, 58),  # 馬名 全角18文字固定
    'chakujun':    (212, 214),  # 確定着順（数字 or 取/中/除）
    'tan_odds':    (237, 241),  # 単勝オッズ 4桁整数（÷10 = 倍率）
}


def parse_se(rec):
    try:
        if len(rec) < 242:
            return None
        kaisai_date = rec[11:19].strip()
        venue_cd    = rec[19:21].strip()
        race_no_s   = rec[25:27].strip()
        umaban      = rec[28:30].strip()
        horse_name  = rec[40:58].strip()
        chakujun    = rec[212:214].strip()
        tan_odds_s  = rec[237:241].strip()

        if not kaisai_date or not venue_cd:
            return None

        race_no = int(race_no_s) if race_no_s.isdigit() else race_no_s

        if tan_odds_s.isdigit() and int(tan_odds_s) > 0:
            tan_odds = int(tan_odds_s) / 10.0
        else:
            tan_odds = ''

        return {
            'kaisai_date': kaisai_date,
            'venue_cd':    venue_cd,
            'race_no':     race_no,
            'umaban':      umaban,
            'horse_name':  horse_name,
            'chakujun':    chakujun,
            'tan_odds':    tan_odds,
        }
    except Exception:
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--from', dest='from_date', default='20250101',
                    help='取得開始日 YYYYMMDD (デフォルト: 20250101)')
    ap.add_argument('--to', dest='to_date', default='20251231',
                    help='取得終了日 YYYYMMDD (デフォルト: 20251231)')
    ap.add_argument('--probe', action='store_true',
                    help='SE/HRレコード生データを表示して終了')
    args = ap.parse_args()

    from_dt = args.from_date + "000000"

    jv = wc.gencache.EnsureDispatch("JVDTLab.JVLink")
    rc_init = jv.JVInit("UNKNOWN")
    print(f"JVInit: {rc_init}")
    if rc_init != 0:
        print("JVInit失敗。ターゲットFrontierが起動しているか確認してください。")
        sys.exit(1)

    rc, readcnt, dldcnt, lts = jv.JVOpen("RACE", from_dt, 1, 0, 0, "")
    print(f"JVOpen(RACE): rc={rc} readcnt={readcnt} dldcnt={dldcnt}")
    if rc != 0:
        print(f"JVOpenエラー rc={rc}")
        jv.JVClose()
        sys.exit(1)

    buf  = " " * 110000
    size = 110000

    if args.probe:
        se_n = hr_n = 0
        while se_n < 3 or hr_n < 2:
            ret, data, sz, fname = jv.JVRead(buf, size, "")
            if ret == 0:
                break
            if ret < 0 and ret not in (-1, -3):
                break
            if ret <= 0:
                continue
            rt = data[:2].strip()
            if rt == "SE" and se_n < 3:
                print(f"\n[SE {se_n+1}] len={ret}")
                for i in range(0, min(ret, 280), 10):
                    print(f"  [{i:03d}:{i+10:03d}] {repr(data[i:i+10])}")
                se_n += 1
            if rt == "HR" and hr_n < 2:
                print(f"\n[HR {hr_n+1}] len={ret}")
                for i in range(0, min(ret, 500), 10):
                    print(f"  [{i:03d}:{i+10:03d}] {repr(data[i:i+10])}")
                hr_n += 1
        jv.JVClose()
        return

    # 本番取得
    daily   = defaultdict(list)
    n_se = n_hr = n_skip = 0

    print(f"読み込み中: {args.from_date} 〜 {args.to_date} ...")
    while True:
        try:
            ret, data, sz, fname = jv.JVRead(buf, size, "")
        except Exception as e:
            print(f"JVRead例外: {e}")
            break
        if ret == 0:
            print("EOF")
            break
        if ret == -1:
            continue
        if ret == -3:
            time.sleep(0.05)
            continue
        if ret < 0:
            print(f"JVReadエラー: {ret}")
            break

        rt = data[:2].strip()

        if rt == "SE":
            d = parse_se(data)
            if d is None:
                continue
            kd = d['kaisai_date']
            if kd < args.from_date or kd > args.to_date:
                continue
            venue_name = JYO_NAME.get(d['venue_cd'], d['venue_cd'])
            daily[kd].append({
                '開催場R':   venue_name + str(d['race_no']),
                '馬名S':     d['horse_name'],
                '着':        d['chakujun'],
                '単勝オッズ': d['tan_odds'],
                '馬番':      d['umaban'],
            })
            n_se += 1
        elif rt == "HR":
            n_hr += 1
        else:
            n_skip += 1

    jv.JVClose()
    print(f"SE: {n_se}件 / HR: {n_hr}件 / その他: {n_skip}件")

    if n_se == 0:
        print("SEレコードが取れませんでした。--probe オプションで確認してください。")
        sys.exit(1)

    out_dir = os.path.join(base_dir, 'data', 'raw', 'results', 'jvlink')
    os.makedirs(out_dir, exist_ok=True)

    total_rows = 0
    for kd in sorted(daily.keys()):
        out_path = os.path.join(out_dir, f'{kd}.csv')
        rows = daily[kd]
        with open(out_path, 'w', encoding='cp932', newline='') as f:
            writer = csv.DictWriter(
                f, fieldnames=['開催場R', '馬名S', '着', '単勝オッズ', '馬番'])
            writer.writeheader()
            writer.writerows(rows)
        total_rows += len(rows)

    print(f"\n完了: {len(daily)}日 {total_rows}行 → {out_dir}/")
    print("先頭ファイルを確認してください:")
    first = sorted(daily.keys())[0] if daily else None
    if first:
        print(f"  {first}.csv 先頭5行:")
        for r in daily[first][:5]:
            print(f"    {r}")


if __name__ == '__main__':
    main()
