# -*- coding: utf-8 -*-
"""
補助金書類自動作成 CLI

使い方:
  # 申請書 + 給与計算を一括実行
  python run.py --company "京のお肉処弘" --template "通常枠_2026" --folder "./京の食事処資料/"

  # 申請書のみ
  python run.py --task application --company "京のお肉処弘" --template "通常枠_2026" --folder "./京の食事処資料/"

  # 給与計算のみ
  python run.py --task wage --company "京のお肉処弘" --folder "./京の食事処資料/"

  # Google Drive連携（管理シートから実行）
  python run.py --drive --sheet-id "1ABC..."

  # ドライラン（APIなしでテスト）
  python run.py --dry-run --company "テスト" --template "通常枠_2026" --folder "./テスト資料/"
"""
from __future__ import annotations

import sys
import argparse
import logging
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


def find_template(base_dir: Path, template_type: str) -> Path | None:
    """テンプレートファイルを検索（v2を優先）"""
    keywords = {
        '通常枠_2026': ['原本', '通常枠', '2026'],
        'インボイス枠_2026': ['原本', 'インボイス', '2026'],
    }
    kws = keywords.get(template_type, [])
    candidates = []
    search_dirs = [base_dir]
    tool_dir = base_dir / 'ツール'
    if tool_dir.exists():
        search_dirs.append(tool_dir)
    for d in search_dirs:
        for p in d.iterdir():
            if p.suffix == '.xlsx' and all(kw in p.name for kw in kws) and not p.name.startswith('~$'):
                candidates.append(p)
    if not candidates:
        return None
    for c in candidates:
        if 'v2' in c.name:
            return c
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

    # Extractor選択
    api_key = '' if args.dry_run else CLAUDE_API_KEY
    extractor = create_extractor(api_key)

    if args.task in ('all', 'application'):
        template_path = None
        if args.template_path:
            template_path = Path(args.template_path)
        else:
            template_path = find_template(base_dir, args.template)

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
        )

        print(f'\n=== 申請書作成: {status.status} ===')
        if status.message:
            print(f'  {status.message}')
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


def cmd_drive(args):
    """Google Drive連携実行"""
    from hojokin.google_drive import DriveClient
    from hojokin.google_sheets import SheetsClient
    from hojokin.pipeline import run_full_pipeline
    from hojokin.config import GOOGLE_CREDENTIALS_PATH, MANAGEMENT_SHEET_ID
    import tempfile

    creds = args.credentials or GOOGLE_CREDENTIALS_PATH
    sheet_id = args.sheet_id or MANAGEMENT_SHEET_ID

    if not sheet_id:
        print('エラー: --sheet-id または MANAGEMENT_SHEET_ID 環境変数を設定してください')
        sys.exit(1)

    drive = DriveClient(creds)
    sheets = SheetsClient(creds, sheet_id)

    # 管理シートから未処理の会社を取得
    companies = sheets.get_pending_companies()
    if not companies:
        print('処理対象の会社がありません。')
        return

    for company in companies:
        name = company['company_name']
        folder_id = company['folder_id']
        template_type = company['template_type']
        row = company['row_number']

        print(f'\n--- {name} ({template_type}) ---')
        sheets.set_processing(row)

        try:
            # 一時ディレクトリにダウンロード
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp = Path(tmpdir)

                # 入力フォルダの検出
                files = drive.list_files(folder_id)
                input_folder_id = None
                output_folder_id = None
                for f in files:
                    if f['name'] == '入力' and f['mimeType'] == 'application/vnd.google-apps.folder':
                        input_folder_id = f['id']
                    elif f['name'] == '出力' and f['mimeType'] == 'application/vnd.google-apps.folder':
                        output_folder_id = f['id']

                if input_folder_id is None:
                    input_folder_id = folder_id  # 入力フォルダがなければルートを使用

                if output_folder_id is None:
                    output_folder_id = drive.find_or_create_subfolder(folder_id, '出力')

                # ファイルダウンロード
                local_input = tmp / '入力'
                drive.download_folder(input_folder_id, local_input)

                # テンプレートダウンロード（別途テンプレートフォルダから）
                # TODO: テンプレートフォルダIDの設定
                template_local = tmp / 'template.xlsx'
                # 暫定: ローカルのテンプレートを使用
                base = Path('.')
                template_path = find_template(base, template_type)
                if template_path is None:
                    raise FileNotFoundError(f'テンプレートが見つかりません: {template_type}')

                # 実行
                results = run_full_pipeline(
                    resource_folder=local_input,
                    template_path=template_path,
                    template_type=template_type,
                    company_name=name,
                )

                # 出力をDriveにアップロード
                output_names = []
                for result in results:
                    for fname in result.output_files:
                        fpath = local_input / fname
                        if fpath.exists():
                            drive.upload_file(fpath, output_folder_id)
                            output_names.append(fname)

                total_empty = sum(len(r.empty_cells) for r in results)
                sheets.set_completed(row, output_names, total_empty)
                print(f'  完了: {", ".join(output_names)}')

        except Exception as e:
            sheets.set_error(row, str(e))
            print(f'  エラー: {e}')


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
                       choices=['通常枠_2026', 'インボイス枠_2026'], help='テンプレートタイプ')
    local.add_argument('--template-path', help='テンプレートExcelの直接パス')
    local.add_argument('--task', default='all', choices=['all', 'application', 'wage'],
                       help='実行タスク')
    local.add_argument('--dry-run', action='store_true', help='APIなしでテスト実行')

    # Drive連携
    drv = sub.add_parser('drive', help='Google Drive連携で実行')
    drv.add_argument('--sheet-id', help='管理用SpreadsheetのID')
    drv.add_argument('--credentials', help='サービスアカウントJSONパス')

    args = parser.parse_args()
    setup_logging(args.verbose)

    if args.mode == 'local':
        cmd_local(args)
    elif args.mode == 'drive':
        cmd_drive(args)
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
