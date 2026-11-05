# -*- coding: utf-8 -*-
"""加点判定ローカル実行ドライバ（katen-judge スキル同梱・API不使用）。

記入済みの「加点判定用賃金台帳」を読み、加点措置①②を判定してレポートを
標準出力する。--fill を付けると公式様式①②を外科的パッチ（原本無改変）で
生成する。判定エンジンはスキルに同梱した hojokin モジュール
（ツール本番と同一コード）を使う。依存: Python3 + openpyxl のみ。

使い方（--call-round で実際の公募回を必ず指定する。判定だけの実行でも必須）:
  python run_katen_judge.py 台帳.xlsx --call-round 7
  python run_katen_judge.py 台帳.xlsx --call-round 7 --app-month 2026/11
  python run_katen_judge.py 台帳.xlsx --call-round 7 --fill 通常 --company 株式会社サンプル
  python run_katen_judge.py 台帳.xlsx --call-round 7 --fill インボイス --outdir 出力先
  第1〜6次（旧回）の訂正だけは --call-round 6 と旧版の台帳を使う。
  交付申請月が2026年11月以降なのに公募回が6以下のときは、旧ルールで判定せずに止める。

終了コード: 0=確定 / 2=入力・版・環境エラー / 3=未確定（該当様式は生成しない）
"""
import argparse
import re
import sys
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):   # Windows でもエラーの日本語が化けないように
    _stream.reconfigure(encoding='utf-8')

# スキルのフォルダ（配布版）か開発リポジトリを親方向に探索して hojokin を import 可能にする
_HERE = Path(__file__).resolve()
ROOT = next((p for p in _HERE.parents if (p / 'hojokin' / 'wage_reader.py').exists()), None)
if ROOT is None:
    print('❌ hojokin/（判定エンジン）が見つかりません。katen-judge スキルのフォルダが'
          '欠けています。補助金プラグインを入れ直すか、管理者に連絡してください。',
          file=sys.stderr)
    sys.exit(2)
sys.path.insert(0, str(ROOT))

from hojokin.subsidy_rules import rules_for_round  # noqa: E402

from hojokin.wage_reader import (  # noqa: E402
    fill_bonus_sheet_1,
    fill_bonus_sheet_2,
    judge_bonus_points,
    read_bonus_wage_ledger,
    ym_label,
)

# 第6次は2026年10月30日の締切で終了した（制度リファレンスの記録。第7次の締切は2026年12月1日）。
# 2026年11月以降に交付申請する案件は第7次以降のルール（加点②は令和8年6月比55円等）で判定する。
ROUND7_FIRST_CALL = 7
ROUND7_FIRST_APPLICATION_YM = (2026, 11)

# 公式様式（申請枠で①のテンプレが変わる。②は全枠共通）。値は公式のファイル名。
KATEN_DIR = ROOT / '補助金加点'
FORM_S1 = {
    '通常': '補助率引き上げ・加点措置①用.xlsx',
    'インボイス': '加点措置①用.xlsx',
    'セキュリティ': '加点措置①用.xlsx',
}
FORM_S2 = '加点措置②用.xlsx'
# 配布版（組織プラグイン）では claude.ai が記号入りのファイル名を受け付けないため、
# ①②・【】 を置き換えた名前で置いてある（scripts/distribution_manifest.py の置き換えと同じ）
DIST_NAME_REPLACEMENTS = (('【', ''), ('】', '_'), ('①', '1'), ('②', '2'), ('・', '_'))


def _dist_name(name: str) -> str:
    for old, new in DIST_NAME_REPLACEMENTS:
        name = name.replace(old, new)
    return name


def form_path(name: str, call_round: int) -> Path:
    """公式様式の実際の置き場所。開発リポジトリでは公式の名前、配布版では置き換えた名前。"""
    directory = ROOT / rules_for_round(call_round).forms_dir
    official = directory / name
    return official if official.exists() else directory / _dist_name(name)


def _parse_app_month(text: str) -> tuple[int, int]:
    m = re.search(r'(\d{4})[/\-年.](\d{1,2})', text)
    if not m:
        raise argparse.ArgumentTypeError(f'交付申請月の形式が不正です: {text}（例: 2026/08）')
    year, month = int(m.group(1)), int(m.group(2))
    if not 1 <= month <= 12:
        raise argparse.ArgumentTypeError(f'月が不正です: {text}')
    return (year, month)


def _print_report(res) -> None:
    rules = rules_for_round(res.call_round)
    print(f'公募回: {res.call_round} / 制度版: {rules.rule_id}')
    print('=' * 60)
    print('加点判定レポート（判定エンジン: hojokin.wage_reader＝ツールと同一）')
    print('=' * 60)
    print(f'都道府県        : {res.prefecture or "（未入力）"}'
          f'（地域別最賃 改定前 {res.min_wage_r6}円 / 改定後 {res.min_wage_r7}円）')
    if res.application_ym:
        print(f'交付申請月      : {ym_label(res.application_ym)}'
              f' → 直近月: {ym_label(res.latest_ym)}')
    else:
        print('交付申請月      : （未入力。加点②は判定不能）')
    print()
    print('── 加点措置①（通常枠のみ補助率引上げを兼ねる） ──')
    print('  月別の「時間換算給与 < 適用版の改定後最賃」該当者（該当/母数）:')
    for d in res.bonus1_details:
        if not d['has_data'] or not d.get('complete', False):
            mark, body = '－', 'データなし'
        else:
            mark = '○' if d['meets_30pct'] else '×'
            body = f"{d['under_r7']}/{d['total']}名 ({d['ratio'] * 100:.1f}%)"
        print(f"    {d['label']:<10}: {body} {mark}")
    n_met = len(res.bonus1_months_met)
    print(f'  30%以上の月: {n_met} か月（3か月以上で対象・連続不要）')
    print(f'  判定: {res.bonus1_status}')
    print()
    print('── 加点措置② ──')
    if res.bonus2_status == '未確定':
        print('  基準月・直近月の最低賃金と差額: 算定未確定（母集団・時間・賃金等の確認待ち）')
    else:
        print(f'  基準月 {ym_label(rules.bonus2_base_ym)} の事業場内最低賃金: '
              f'{res.bonus2_min_wage_july:,.0f}円')
        if res.latest_ym:
            print(f'  直近月 {ym_label(res.latest_ym)} の事業場内最低賃金: '
                  f'{res.bonus2_min_wage_latest:,.0f}円')
            print(f'  差: {res.bonus2_diff:+,.0f}円（{rules.bonus2_threshold_yen}円以上で対象）')
    print(f'  判定: {res.bonus2_status}')
    if res.notes:
        print()
        print('── 注意（必ず全件確認する） ──')
        for note in res.notes:
            print(f'  ⚠ {note}')
    print()


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description='加点措置①②のローカル判定と公式様式生成')
    ap.add_argument('ledger', type=Path, help='記入済みの加点判定用賃金台帳 .xlsx')
    ap.add_argument('--call-round', type=int, default=None,
                    help='実際の公募回（必須。7次は7。第1〜6次の訂正だけ6以下）')
    ap.add_argument('--app-month', type=_parse_app_month, default=None,
                    help='交付申請月（台帳C3が未入力の場合のフォールバック。例: 2026/08）')
    ap.add_argument('--fill', choices=sorted(FORM_S1), default=None,
                    help='公式様式①②を生成する（申請枠を指定。①のテンプレが枠で変わる）')
    ap.add_argument('--outdir', type=Path, default=None, help='様式の出力先（既定: 台帳と同じ場所）')
    ap.add_argument('--company', default=None, help='出力ファイル名の会社名（既定: 台帳ファイル名）')
    args = ap.parse_args(argv)

    if not args.ledger.exists():
        print(f'❌ 台帳が見つかりません: {args.ledger}', file=sys.stderr)
        return 2

    if args.call_round is None:
        print('❌ --call-round で実際の公募回を指定してください（7次は --call-round 7）。'
              '判定だけの実行でも必須です。公募回によって判定期間・基準月・金額が変わるため、'
              '指定がないまま旧ルールで判定しません。', file=sys.stderr)
        return 2
    call_round = args.call_round
    try:
        ledger = read_bonus_wage_ledger(args.ledger, application_ym_fallback=args.app_month, call_round=call_round)
    except (ValueError, TypeError) as exc:
        print(f'❌ 台帳を確認してください: {exc}', file=sys.stderr)
        return 2
    if (call_round < ROUND7_FIRST_CALL and ledger.application_ym
            and tuple(ledger.application_ym) >= ROUND7_FIRST_APPLICATION_YM):
        print(f'❌ 交付申請月が{ym_label(ledger.application_ym)}なのに、公募回が{call_round}です。'
              f'{ym_label(ROUND7_FIRST_APPLICATION_YM)}以降の申請は第7次以降のルールで判定します。'
              f'--call-round 7 等の実際の公募回と、第7次以降用の加点判定用台帳を使ってください。', file=sys.stderr)
        return 2
    res = judge_bonus_points(ledger)
    print(f'台帳: {args.ledger.name}（従業員 {len(ledger.employees)}名・役員は台帳に含めない運用）')
    _print_report(res)

    if args.fill:
        outdir = args.outdir or args.ledger.parent
        outdir.mkdir(parents=True, exist_ok=True)
        company = args.company or args.ledger.stem
        t1, t2 = form_path(FORM_S1[args.fill], call_round), form_path(FORM_S2, call_round)
        missing = [p for p in (t1, t2) if not p.exists()]
        if missing:
            print(f'❌ 公式様式テンプレが見つかりません: {[str(p) for p in missing]}。'
                  '補助金プラグインを入れ直すか、管理者に連絡してください。', file=sys.stderr)
            return 2
        # 出力ファイル名は公式の様式名のまま（配布版で置き換えた名前にはしない）
        out1 = outdir / f'{company}_{Path(FORM_S1[args.fill]).stem}.xlsx'
        if res.bonus1_status != '未確定' and res.application_ym:
            fill_bonus_sheet_1(t1, out1, res)
            print(f'様式①生成: {out1}（原本: {t1}）')
        else:
            print('⚠ 様式①は未確定のため生成しません。既存出力を今回の確定版として使わないでください。')
        if res.application_ym and res.bonus2_status != '未確定':
            out2 = outdir / f'{company}_{Path(FORM_S2).stem}.xlsx'
            fill_bonus_sheet_2(t2, out2, res)
            print(f'様式②生成: {out2}（原本: {t2}）')
        else:
            print('⚠ 様式②は未確定のため生成しません。申請月・基準月・直近月の要確認事項を解消してください。')
        print('（様式は原本無改変の外科的パッチ。判定式は Excel で開いた時に再計算されます）')

    return 3 if '未確定' in (res.bonus1_status, res.bonus2_status) else 0


if __name__ == '__main__':
    try:
        exit_code = main()
    except (ValueError, TypeError) as exc:
        print(f'❌ 入力又は様式を確認してください: {exc}', file=sys.stderr)
        exit_code = 2
    sys.exit(exit_code)
