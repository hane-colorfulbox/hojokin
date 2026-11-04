# -*- coding: utf-8 -*-
"""
補助金書類自動作成 CLI

使い方:
  # 申請書 + 給与計算を一括実行
  python run.py --company "顧客企業名" --template "通常枠_2026" --folder "./顧客資料/"

  # 申請書のみ
  python run.py --task application --company "顧客企業名" --template "通常枠_2026" --folder "./顧客資料/"

  # 給与計算のみ
  python run.py --task wage --company "顧客企業名" --folder "./顧客資料/"

  # Google Drive連携（管理シートから実行）
  python run.py --drive --sheet-id "1ABC..."

  # ドライラン（APIなしでテスト）
  python run.py --dry-run --company "テスト" --template "通常枠_2026" --folder "./テスト資料/"
"""
from __future__ import annotations

import sys
import argparse
import logging
import json
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

sys.stdout.reconfigure(encoding='utf-8')

# パッケージのインポートパスを追加
sys.path.insert(0, str(Path(__file__).parent))


def setup_logging(verbose: bool = False):
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        datefmt='%H:%M:%S',
        handlers=[logging.StreamHandler(sys.stdout)],
    )


def find_template(base_dir: Path, template_type: str, call_round: int = 6) -> Path | None:
    """テンプレートファイルを検索（v2を優先。app.py の同名関数と同期を保つこと）"""
    if call_round >= 7:
        from hojokin.template_versions import select_application_template
        return select_application_template(base_dir, template_type, call_round)
    keywords = {
        '通常枠_2026': ['原本', '通常枠', '法人', '2026'],
        '通常枠_個人_2026': ['原本', '通常枠', '個人', '2026'],
        'インボイス枠_2026': ['原本', 'インボイス', '法人', '2026'],
        'インボイス枠_個人_2026': ['原本', 'インボイス', '個人', '2026'],
    }
    # 法人テンプレ種別は「個人」を含むファイル名を除外（Drive の通常枠個人正本の DL 名は
    # 【原本_法人】…通常枠_個人2026.xlsx で両方のキーワードを含むため。個人側として扱う）
    exclude = {
        '通常枠_2026': ['個人'],
        'インボイス枠_2026': ['個人'],
    }
    kws = keywords.get(template_type, [])
    ng_kws = exclude.get(template_type, [])
    candidates = []
    search_dirs = [base_dir]
    tool_dir = base_dir / 'ツール'
    if tool_dir.exists():
        search_dirs.append(tool_dir)
    for d in search_dirs:
        for p in d.iterdir():
            if (p.suffix == '.xlsx' and all(kw in p.name for kw in kws)
                    and not any(kw in p.name for kw in ng_kws) and not p.name.startswith('~$')):
                candidates.append(p)
    if not candidates:
        return None
    for c in candidates:
        if 'v2' in c.name:
            return c
    if len({c.name for c in candidates}) > 1:
        print(f'[find_template] {template_type} の候補が複数: '
              f'{[c.name for c in candidates]} → {candidates[0].name} を採用')
    return candidates[0]


def cmd_local(args):
    """ローカル実行"""
    from hojokin.pipeline import run_application_transfer, run_wage_calculation, run_full_pipeline
    from hojokin.ai_extractor import create_extractor
    from hojokin.config import CLAUDE_API_KEY

    folder = Path(args.folder).resolve()
    if not folder.exists():
        print(f'エラー: フォルダが見つかりません: {folder}')
        sys.exit(1)

    base_dir = folder.parent
    company = args.company

    application_context = None
    if args.task in ('all', 'application'):
        from hojokin.application_context import ApplicationContext, validate_application_context
        if not args.application_context:
            print('エラー: 申請書の確定には --application-context で確認済み案件設定JSONを指定してください。')
            return 2
        try:
            values = json.loads(Path(args.application_context).read_text(encoding='utf-8-sig'))
            application_context = ApplicationContext.from_dict(values)
            validate_application_context(application_context, args.template)
        except (OSError, ValueError, TypeError) as exc:
            print(f'申請前提の確認が必要です: {exc}')
            return 2

    # Extractor選択
    api_key = '' if args.dry_run else CLAUDE_API_KEY
    extractor = create_extractor(api_key)

    if args.task in ('all', 'application'):
        template_path = None
        if args.template_path:
            template_path = Path(args.template_path)
        else:
            template_path = find_template(base_dir, args.template, application_context.call_round)
            if template_path is None:
                template_path = find_template(Path(__file__).parent, args.template,
                                              application_context.call_round)

        if template_path is None or not template_path.exists():
            print(f'エラー: テンプレートが見つかりません。--template-path で直接指定してください。')
            sys.exit(1)

        print(f'テンプレート: {template_path.name}')
        output = folder / f'{company}_{args.template}_AI版.xlsx'

        status = run_application_transfer(
            resource_folder=folder,
            template_path=template_path,
            template_type=args.template,
            output_path=output,
            extractor=extractor,
            application_context=application_context,
        )

        print(f'\n=== 申請書作成: {status.status} ===')
        if status.message:
            print(f'  {status.message}')
        if status.status != '完了':
            return 2
        if status.empty_cells:
            print(f'  空欄 {len(status.empty_cells)}件:')
            for e in status.empty_cells[:10]:
                print(f'    {e}')
            if len(status.empty_cells) > 10:
                print(f'    ... 他{len(status.empty_cells) - 10}件')

    if args.task in ('all', 'wage'):
        output_wage = folder / f'{company}_給与支給総額計算.xlsx'

        status = run_wage_calculation(
            resource_folder=folder,
            company_name=company,
            output_path=output_wage,
            extractor=extractor,
        )

        print(f'\n=== 給与計算: {status.status} ===')
        if status.message:
            print(f'  {status.message}')

    if args.task == 'per_employee_wage':
        output_wage = folder / f'{company}_一人当たり給与支給総額.xlsx'

        status = run_wage_calculation(
            resource_folder=folder,
            company_name=company,
            output_path=output_wage,
            extractor=extractor,
            per_employee_only=True,
        )

        print(f'\n=== 一人当たり給与支給総額: {status.status} ===')
        if status.message:
            print(f'  {status.message}')


def cmd_drive(args):
    """Google Drive連携実行"""
    # 管理表の既存列だけでは案件別の制度版・申告事実を確認できない。
    # 確認前に処理中フラグや出力フォルダを外部へ書き込まない。
    print('Drive一括実行は案件別の確認済み申請前提に未対応のため停止しています。'
          'ローカル実行で案件ごとの --application-context を指定してください。')
    return 2




def main():
    parser = argparse.ArgumentParser(
        description='補助金書類自動作成システム',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument('-v', '--verbose', action='store_true', help='詳細ログ')

    sub = parser.add_subparsers(dest='mode')

    # ローカル実行
    local = sub.add_parser('local', help='ローカルフォルダから実行')
    local.add_argument('--company', required=True, help='顧客名')
    local.add_argument('--folder', required=True, help='資料フォルダパス')
    local.add_argument('--template', default='通常枠_2026',
                       choices=['通常枠_2026', '通常枠_個人_2026',
                                'インボイス枠_2026', 'インボイス枠_個人_2026'],
                       help='テンプレートタイプ')
    local.add_argument('--template-path', help='テンプレートExcelの直接パス')
    local.add_argument('--application-context', help='確認済み案件設定のJSON（申請書作成では必須）')
    local.add_argument('--task', default='all',
                       choices=['all', 'application', 'wage', 'per_employee_wage'],
                       help='実行タスク（per_employee_wage は賃金台帳のみで一人当たり給与支給総額を算定）')
    local.add_argument('--dry-run', action='store_true', help='APIなしでテスト実行')

    # Drive連携
    drv = sub.add_parser('drive', help='Google Drive連携で実行')
    drv.add_argument('--sheet-id', help='管理用SpreadsheetのID')
    drv.add_argument('--credentials', help='サービスアカウントJSONパス')

    args = parser.parse_args()
    setup_logging(args.verbose)

    if args.mode == 'local':
        return cmd_local(args) or 0
    elif args.mode == 'drive':
        return cmd_drive(args) or 0
    else:
        parser.print_help()


if __name__ == '__main__':
    sys.exit(main())
