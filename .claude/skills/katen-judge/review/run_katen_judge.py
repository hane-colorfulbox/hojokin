# -*- coding: utf-8 -*-
"""加点判定ローカル実行ドライバ（katen-judge スキル同梱・API不使用）。

記入済みの「加点判定用賃金台帳」を読み、加点措置①②を判定してレポートを
標準出力する。--fill を付けると公式様式①②を外科的パッチ（原本無改変）で
生成する。判定エンジンは引き継ぎパック同梱の hojokin モジュール
（ツール本番と同一コード）を使う。依存: Python3 + openpyxl のみ。

使い方:
  python run_katen_judge.py 台帳.xlsx
  python run_katen_judge.py 台帳.xlsx --app-month 2026/08
  python run_katen_judge.py 台帳.xlsx --fill 通常 --company 株式会社サンプル
  python run_katen_judge.py 台帳.xlsx --fill インボイス --outdir 出力先

終了コード: 0=正常 / 2=ファイル・環境エラー
"""
import argparse
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# 補助金フォルダ（ZIP展開先 or 開発リポジトリ）を親方向に探索して hojokin を import 可能にする
_HERE = Path(__file__).resolve()
ROOT = next((p for p in _HERE.parents if (p / 'hojokin' / 'wage_reader.py').exists()), None)
if ROOT is None:
    print('❌ hojokin/（判定エンジン）が見つかりません。引き継ぎパック'
          '（hojokin-handoff.zip）を補助金フォルダ直下に展開し、スキルは'
          'そのフォルダ配下に置いたまま実行してください。', file=sys.stderr)
    sys.exit(2)
sys.path.insert(0, str(ROOT))

from hojokin.wage_reader import (  # noqa: E402
    BONUS2_BASE_YM,
    BONUS_THRESHOLD_YEN,
    fill_bonus_sheet_1,
    fill_bonus_sheet_2,
    judge_bonus_points,
    read_bonus_wage_ledger,
    ym_label,
)

# 公式様式（申請枠で①のテンプレが変わる。②は全枠共通）
KATEN_DIR = ROOT / '補助金加点'
TEMPLATE_S1 = {
    '通常': KATEN_DIR / '補助率引き上げ・加点措置①用.xlsx',
    'インボイス': KATEN_DIR / '加点措置①用.xlsx',
    'セキュリティ': KATEN_DIR / '加点措置①用.xlsx',
}
TEMPLATE_S2 = KATEN_DIR / '加点措置②用.xlsx'


def _parse_app_month(text: str) -> tuple[int, int]:
    m = re.search(r'(\d{4})[/\-年.](\d{1,2})', text)
    if not m:
        raise argparse.ArgumentTypeError(f'交付申請月の形式が不正です: {text}（例: 2026/08）')
    year, month = int(m.group(1)), int(m.group(2))
    if not 1 <= month <= 12:
        raise argparse.ArgumentTypeError(f'月が不正です: {text}')
    return (year, month)


def _print_report(res) -> None:
    print('=' * 60)
    print('加点判定レポート（判定エンジン: hojokin.wage_reader＝ツールと同一）')
    print('=' * 60)
    print(f'都道府県        : {res.prefecture or "（未入力）"}'
          f'（地域別最賃 改定前 {res.min_wage_r6}円 / R7改定後 {res.min_wage_r7}円）')
    if res.application_ym:
        print(f'交付申請月      : {ym_label(res.application_ym)}'
              f' → 直近月: {ym_label(res.latest_ym)}')
    else:
        print('交付申請月      : （未入力。加点②は判定不能）')
    print()
    print('── 加点措置①（加点項目14・補助率1/2→2/3トリガー） ──')
    print('  月別の「時間換算給与 < R7改定後最賃」該当者（該当/母数）:')
    for d in res.bonus1_details:
        if not d['has_data']:
            mark, body = '－', 'データなし'
        else:
            mark = '○' if d['meets_30pct'] else '×'
            body = f"{d['under_r7']}/{d['total']}名 ({d['ratio'] * 100:.1f}%)"
        print(f"    {d['label']:<10}: {body} {mark}")
    n_met = len(res.bonus1_months_met)
    print(f'  30%以上の月: {n_met} か月（3か月以上で対象・連続不要）')
    print(f'  判定: {"◎ 対象（補助率2/3）" if res.bonus1_eligible else "対象外"}')
    print()
    print('── 加点措置②（加点項目15） ──')
    print(f'  基準月 {ym_label(BONUS2_BASE_YM)} の事業場内最低賃金: '
          f'{res.bonus2_min_wage_july:,.0f}円')
    if res.latest_ym:
        print(f'  直近月 {ym_label(res.latest_ym)} の事業場内最低賃金: '
              f'{res.bonus2_min_wage_latest:,.0f}円')
        print(f'  差: {res.bonus2_diff:+,.0f}円（{BONUS_THRESHOLD_YEN}円以上で対象）')
    print(f'  判定: {"◎ 対象" if res.bonus2_eligible else "対象外"}')
    if res.notes:
        print()
        print('── 注意（必ず全件確認する） ──')
        for note in res.notes:
            print(f'  ⚠ {note}')
    print()


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description='加点措置①②のローカル判定と公式様式生成')
    ap.add_argument('ledger', type=Path, help='記入済みの加点判定用賃金台帳 .xlsx')
    ap.add_argument('--app-month', type=_parse_app_month, default=None,
                    help='交付申請月（台帳C3が未入力の場合のフォールバック。例: 2026/08）')
    ap.add_argument('--fill', choices=sorted(TEMPLATE_S1), default=None,
                    help='公式様式①②を生成する（申請枠を指定。①のテンプレが枠で変わる）')
    ap.add_argument('--outdir', type=Path, default=None, help='様式の出力先（既定: 台帳と同じ場所）')
    ap.add_argument('--company', default=None, help='出力ファイル名の会社名（既定: 台帳ファイル名）')
    args = ap.parse_args(argv)

    if not args.ledger.exists():
        print(f'❌ 台帳が見つかりません: {args.ledger}', file=sys.stderr)
        return 2

    ledger = read_bonus_wage_ledger(args.ledger, application_ym_fallback=args.app_month)
    res = judge_bonus_points(ledger)
    print(f'台帳: {args.ledger.name}（従業員 {len(ledger.employees)}名・役員は台帳に含めない運用）')
    _print_report(res)

    if args.fill:
        outdir = args.outdir or args.ledger.parent
        outdir.mkdir(parents=True, exist_ok=True)
        company = args.company or args.ledger.stem
        t1 = TEMPLATE_S1[args.fill]
        missing = [p for p in (t1, TEMPLATE_S2) if not p.exists()]
        if missing:
            print(f'❌ 公式様式テンプレが見つかりません: {[str(p) for p in missing]}。'
                  '引き継ぎパックを最新に更新してください。', file=sys.stderr)
            return 2
        out1 = outdir / f'{company}_{t1.stem}.xlsx'
        fill_bonus_sheet_1(t1, out1, res)
        print(f'様式①生成: {out1}')
        if res.application_ym:
            out2 = outdir / f'{company}_{TEMPLATE_S2.stem}.xlsx'
            fill_bonus_sheet_2(TEMPLATE_S2, out2, res)
            print(f'様式②生成: {out2}')
        else:
            print('⚠ 交付申請月が無いため様式②は生成しません（--app-month で指定可）')
        print('（様式は原本無改変の外科的パッチ。判定式は Excel で開いた時に再計算されます）')

    return 0


if __name__ == '__main__':
    sys.exit(main())
