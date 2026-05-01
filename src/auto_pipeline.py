# -*- coding: utf-8 -*-
"""
競馬AI 完全自動パイプライン

モード:
  predict  木〜金に実行。カード取得→予測→新聞→docs/index更新→push
  watch    土日に実行。オッズポーリング（odds_watcher.py に委譲）
  auto     曜日で自動判定（金→predict / 土日→watch / それ以外→次週末のpredict）

タスクスケジューラからの実行を想定。
setup_scheduler.py で Task Scheduler に登録できる。

使い方:
  python src/auto_pipeline.py --mode auto
  python src/auto_pipeline.py --mode predict
  python src/auto_pipeline.py --mode watch
  python src/auto_pipeline.py --mode predict --date 20260503
"""
import sys, io, os, re, glob, shutil, subprocess, argparse
from datetime import datetime, timedelta

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

BASE_DIR     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR      = os.path.join(BASE_DIR, 'src')
DOCS_DIR     = os.path.join(BASE_DIR, 'docs')
OUTPUT_DIR   = os.path.join(BASE_DIR, 'd_core', 'predict', 'output')
CACHE_DIR    = os.path.join(BASE_DIR, 'data', 'raw', 'cache')
VENV_PYTHON  = os.path.join(BASE_DIR, '.venv_new', 'Scripts', 'python.exe')
PYTHON       = VENV_PYTHON if os.path.exists(VENV_PYTHON) else sys.executable

LOG_DIR      = os.path.join(BASE_DIR, 'logs')


# ── ユーティリティ ─────────────────────────────────────────────────────────

def run(cmd: list, label: str = '') -> int:
    label_s = f'[{label}] ' if label else ''
    print(f"\n{'='*60}")
    print(f"  {label_s}{' '.join(str(c) for c in cmd[-3:])}")
    print(f"{'='*60}")
    r = subprocess.run(cmd, cwd=BASE_DIR, text=True, encoding='utf-8', errors='replace')
    return r.returncode


def next_weekend_dates() -> list[str]:
    """今週または来週の土日の日付を YYYYMMDD で返す"""
    today = datetime.now()
    wd = today.weekday()  # 0=月 ... 5=土 6=日
    if wd == 5:   # 土曜
        sat = today
    elif wd == 6:  # 日曜
        sat = today - timedelta(days=1)
    else:
        days_to_sat = 5 - wd
        sat = today + timedelta(days=days_to_sat)
    sun = sat + timedelta(days=1)
    return [sat.strftime('%Y%m%d'), sun.strftime('%Y%m%d')]


def copy_newspaper_to_docs(date_str: str) -> str | None:
    """最新の newspaper HTML を docs/ にコピーし、コピー先パスを返す"""
    yymmdd = date_str[2:]
    pattern = os.path.join(OUTPUT_DIR, f'd_newspaper_{yymmdd}_*.html')
    files = sorted(glob.glob(pattern))
    if not files:
        print(f"  [WARN] newspaper HTML 見つからず: {pattern}")
        return None
    latest = files[-1]
    dest = os.path.join(DOCS_DIR, f'd_newspaper_{yymmdd}.html')
    shutil.copy2(latest, dest)
    print(f"  docs/ にコピー: d_newspaper_{yymmdd}.html  ({os.path.getsize(dest)//1024}KB)")
    return dest


def update_index_html(dates: list[str]):
    """docs/index.html に新エントリを先頭追加（既存は重複しない）・更新日も更新"""
    index_path = os.path.join(DOCS_DIR, 'index.html')
    with open(index_path, 'r', encoding='utf-8') as f:
        content = f.read()

    for date_str in sorted(dates, reverse=True):
        yymmdd = date_str[2:]
        href = f'd_newspaper_{yymmdd}.html'
        if href in content:
            continue

        y = date_str[:4]
        m = int(date_str[4:6])
        d = int(date_str[6:8])
        label = f'{y}/{m:02d}/{d:02d}'

        new_entry = (
            f'\n      <a class="item" href="{href}">\n'
            f'        <span class="item-icon">📋</span>\n'
            f'        <span class="item-label">{label} D指標競馬新聞</span>\n'
            f'        <span class="item-arrow">›</span>\n'
            f'      </a>'
        )
        content = re.sub(
            r'(<div class="section-title">週次予想レポート</div>)',
            r'\1' + new_entry,
            content, count=1
        )
        print(f"  index.html に追加: {label} D指標競馬新聞")

    # 更新日を今日に
    today_str = datetime.now().strftime('%Y-%m-%d')
    content = re.sub(r'更新: \d{4}-\d{2}-\d{2}', f'更新: {today_str}', content)

    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(content)


def git_commit_push(dates: list[str], suffix: str = ''):
    """変更ファイルを git commit して push"""
    doc_files = []
    for d in dates:
        f = os.path.join('docs', f'd_newspaper_{d[2:]}.html')
        if os.path.exists(os.path.join(BASE_DIR, f)):
            doc_files.append(f)

    targets = doc_files + ['docs/index.html']
    run(['git', 'add'] + targets, 'git add')

    dates_str = '/'.join(d[4:] for d in dates)
    msg = f'auto: D指標新聞 {dates_str}{suffix}'
    rc = subprocess.run(
        ['git', 'commit', '-m', msg],
        cwd=BASE_DIR, capture_output=True, text=True
    )
    if rc.returncode != 0:
        print("  変更なし（コミットスキップ）")
        return
    print(f"  コミット: {msg}")

    run(['git', 'push', 'origin', 'main'], 'git push')


# ── predict モード ──────────────────────────────────────────────────────────

def run_predict(target_dates: list[str]):
    """
    カード取得 → 予測 → 新聞生成 → docs コピー → index 更新 → push
    """
    dates_str = ','.join(target_dates)
    print(f"\n【predict モード】対象: {target_dates}")

    from_date = (datetime.now() - timedelta(days=7)).strftime('%Y%m%d')
    to_date   = max(target_dates)

    # Step 1: predict_weekend.py
    predict_script = os.path.join(SRC_DIR, 'predict_weekend.py')
    rc = run(
        [PYTHON, predict_script, '--from', from_date, '--to', to_date],
        'predict_weekend'
    )
    if rc != 0:
        print(f"[ERROR] predict_weekend.py 失敗 (rc={rc})")
        return

    # Step 2: docs/ にコピー
    copied = []
    for d in target_dates:
        if copy_newspaper_to_docs(d):
            copied.append(d)

    if not copied:
        print("[WARN] コピーする HTML がなかった。終了。")
        return

    # Step 3: index.html 更新
    update_index_html(copied)

    # Step 4: commit & push
    git_commit_push(copied, ' (印候補)')
    print(f"\n【predict 完了】{copied}")


# ── watch モード ────────────────────────────────────────────────────────────

def run_watch(target_dates: list[str], interval: int = 10, until: str = '17:30'):
    """odds_watcher.py を呼び出す（ブロッキング）"""
    print(f"\n【watch モード】対象: {target_dates}")
    watcher = os.path.join(SRC_DIR, 'odds_watcher.py')
    from_date = (datetime.now() - timedelta(days=7)).strftime('%Y%m%d')
    run(
        [PYTHON, watcher,
         '--date', ','.join(target_dates),
         '--from', from_date,
         '--interval', str(interval),
         '--until', until],
        'odds_watcher'
    )


# ── main ────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--mode', choices=['auto', 'predict', 'watch'], default='auto')
    ap.add_argument('--date', default=None,
                    help='対象日 YYYYMMDD（カンマ区切りで複数可）。省略時は今週末')
    ap.add_argument('--interval', type=int, default=10, help='watch ポーリング間隔（分）')
    ap.add_argument('--until', default='17:30', help='watch 終了時刻 HH:MM')
    args = ap.parse_args()

    if args.date:
        target_dates = [d.strip() for d in args.date.split(',')]
    else:
        target_dates = next_weekend_dates()

    mode = args.mode
    if mode == 'auto':
        wd = datetime.now().weekday()
        if wd in (5, 6):    # 土日
            mode = 'watch'
        else:               # 平日（金曜想定）
            mode = 'predict'
        print(f"auto モード: 曜日={wd} → {mode}")

    os.makedirs(LOG_DIR, exist_ok=True)

    if mode == 'predict':
        run_predict(target_dates)
    elif mode == 'watch':
        today = datetime.now().strftime('%Y%m%d')
        today_dates = [d for d in target_dates if d == today]
        if not today_dates:
            today_dates = [today]
        run_watch(today_dates, args.interval, args.until)


if __name__ == '__main__':
    main()
