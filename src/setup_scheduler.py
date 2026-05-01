# -*- coding: utf-8 -*-
"""
Windows タスクスケジューラに競馬AI パイプラインを登録する。

作成するタスク:
  KeibAI-Predict   金曜 21:00  predict_weekend + docs push（印候補）
  KeibAI-Watch-Sat 土曜 08:00  オッズポーリング → 確定版 push
  KeibAI-Watch-Sun 日曜 08:00  同上

実行:
  python src/setup_scheduler.py           # 登録
  python src/setup_scheduler.py --delete  # 削除
  python src/setup_scheduler.py --status  # 確認

注意:
  - JV-Link (ターゲットFrontier) が起動している状態でないと失敗します
  - 「ログオン時のみ実行」で登録するため、PC がスリープ中は実行されません
    → スリープ無効推奨（コントロールパネル → 電源オプション）
"""
import sys, io, os, subprocess, argparse

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

BASE_DIR    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VENV_PYTHON = os.path.join(BASE_DIR, '.venv_new', 'Scripts', 'python.exe')
PYTHON      = VENV_PYTHON if os.path.exists(VENV_PYTHON) else sys.executable
SCRIPT      = os.path.join(BASE_DIR, 'src', 'auto_pipeline.py')
LOG_DIR     = os.path.join(BASE_DIR, 'logs')

TASKS = [
    {
        'name':  'KeibAI-Predict',
        'desc':  '競馬AI 週末予測（金曜・印候補新聞生成）',
        'day':   'FRI',
        'time':  '21:00',
        'mode':  'predict',
    },
    {
        'name':  'KeibAI-Watch-Sat',
        'desc':  '競馬AI 土曜オッズ更新',
        'day':   'SAT',
        'time':  '08:00',
        'mode':  'watch',
    },
    {
        'name':  'KeibAI-Watch-Sun',
        'desc':  '競馬AI 日曜オッズ更新',
        'day':   'SUN',
        'time':  '08:00',
        'mode':  'watch',
    },
]


def schtasks(*args) -> tuple[int, str]:
    cmd = ['schtasks'] + list(args)
    print(f"  $ {' '.join(cmd)}")
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='cp932', errors='replace')
    out = r.stdout.strip()
    if out:
        print(f"    {out}")
    if r.returncode != 0 and r.stderr.strip():
        print(f"    [ERR] {r.stderr.strip()}", file=sys.stderr)
    return r.returncode, out


def build_tr(mode: str, task_name: str) -> str:
    """/tr パラメータ: python auto_pipeline.py --mode MODE > log.txt 2>&1"""
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.join(LOG_DIR, f'{task_name}.log')
    # cmd /c でリダイレクト付き実行
    inner = f'"{PYTHON}" "{SCRIPT}" --mode {mode}'
    # タスクスケジューラは /tr でリダイレクト不可→ cmd /c 経由
    return f'cmd /c "{inner} >> "{log_path}" 2>&1"'


def register():
    print("=== タスクスケジューラ 登録 ===\n")
    for t in TASKS:
        tr = build_tr(t['mode'], t['name'])
        print(f"登録: {t['name']}  ({t['day']} {t['time']})")
        rc, _ = schtasks(
            '/create', '/f',
            '/tn', t['name'],
            '/tr', tr,
            '/sc', 'WEEKLY',
            '/d', t['day'],
            '/st', t['time'],
            '/rl', 'HIGHEST',
        )
        if rc == 0:
            print(f"  ✓ 登録完了\n")
        else:
            print(f"  ✗ 登録失敗\n")


def delete():
    print("=== タスクスケジューラ 削除 ===\n")
    for t in TASKS:
        print(f"削除: {t['name']}")
        schtasks('/delete', '/f', '/tn', t['name'])
        print()


def status():
    print("=== タスクスケジューラ 確認 ===\n")
    for t in TASKS:
        print(f"── {t['name']} ──")
        rc, out = schtasks('/query', '/tn', t['name'], '/fo', 'LIST')
        if rc != 0:
            print("  (未登録)")
        print()


def main():
    ap = argparse.ArgumentParser()
    grp = ap.add_mutually_exclusive_group()
    grp.add_argument('--delete', action='store_true', help='タスクを削除')
    grp.add_argument('--status', action='store_true', help='登録状況を確認')
    args = ap.parse_args()

    if args.delete:
        delete()
    elif args.status:
        status()
    else:
        register()
        print("\n登録済みタスク:")
        status()
        print("\n補足:")
        print(f"  ログ出力先: {LOG_DIR}\\")
        print("  手動実行: schtasks /run /tn KeibAI-Predict")
        print("  注意: PC がスリープ中は実行されません（スリープ無効推奨）")


if __name__ == '__main__':
    main()
