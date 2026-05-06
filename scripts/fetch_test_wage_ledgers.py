# -*- coding: utf-8 -*-
"""賃金台帳PDF/Excel をDriveから収集してローカルに保存。

Path B (Document AI + Haiku) の品質検証・回帰テスト用データセット作成。
ベンダー名・案件名は引数指定で受け取り、コードにはハードコードしない（機密情報保護）。

使用例:
    # 環境変数 DRIVE_PARENT_FOLDER_ID 配下から、指定ベンダー（部分一致）の案件をDL
    python scripts/fetch_test_wage_ledgers.py --vendor "<ベンダー名一部>"

    # 複数ベンダー指定
    python scripts/fetch_test_wage_ledgers.py --vendor "<ベンダーA>" --vendor "<ベンダーB>"

    # 親フォルダを直接指定 + 保存先カスタム
    python scripts/fetch_test_wage_ledgers.py --parent-id <ID> --vendor <name> --out ../_backups/foo

要件:
    - Google Drive Service Account JSON が `credentials/service_account.json` または
      `GOOGLE_SERVICE_ACCOUNT_JSON` 環境変数で指すパスに存在
    - Service Account に Drive 読取権限 + 対象フォルダへの共有設定
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv

load_dotenv()
from hojokin.drive_client import DriveClient

WAGE_KEYWORDS = ['賃金', '台帳', 'wage', '給与']
DEFAULT_OUTPUT_REL = '../_backups/test_wage_ledgers'


def safe_filename(name: str) -> str:
    """Windows 禁止文字を置換"""
    forbidden = '/\\:*?"<>|\r\n\t'
    return ''.join('-' if c in forbidden else c for c in name)


def find_wage_files_recursive(client, folder_id, depth=0, max_depth=4):
    """フォルダ内を再帰探索し、賃金台帳キーワードを含むファイルを返す。"""
    if depth > max_depth:
        return []
    found = []
    try:
        folders = client.list_folders(folder_id)
        files = client.list_files(folder_id)
        for f in files:
            if any(kw in f['name'] for kw in WAGE_KEYWORDS):
                found.append(f)
        for sub in folders:
            found.extend(find_wage_files_recursive(client, sub['id'], depth + 1, max_depth))
    except Exception:
        pass
    return found


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        '--parent-id', default=os.getenv('DRIVE_PARENT_FOLDER_ID'),
        help='Drive 親フォルダID (省略時 .env の DRIVE_PARENT_FOLDER_ID)',
    )
    parser.add_argument(
        '--vendor', action='append', required=True,
        help='ベンダーフォルダ名の部分一致（複数指定可）',
    )
    parser.add_argument(
        '--out', default=DEFAULT_OUTPUT_REL,
        help=f'保存先ディレクトリ（デフォルト: {DEFAULT_OUTPUT_REL}、補助金リポジトリの相対パス）',
    )
    parser.add_argument(
        '--credentials',
        default=os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON', 'credentials/service_account.json'),
        help='Service Account JSON のパス',
    )
    args = parser.parse_args()

    if not args.parent_id:
        print('エラー: --parent-id または .env の DRIVE_PARENT_FOLDER_ID が必要', file=sys.stderr)
        sys.exit(1)

    client = DriveClient(credentials_path=args.credentials)
    base_dir = (Path(__file__).resolve().parent.parent / args.out).resolve()
    base_dir.mkdir(parents=True, exist_ok=True)
    print(f'保存先: {base_dir}')
    print()

    # 親フォルダ直下のベンダーフォルダ一覧
    vendor_folders = client.list_folders(args.parent_id)
    matched_vendors = []
    for kw in args.vendor:
        for vf in vendor_folders:
            if kw in vf['name']:
                matched_vendors.append(vf)
    if not matched_vendors:
        print(f'エラー: 指定ベンダー {args.vendor} に一致するフォルダが見つかりません', file=sys.stderr)
        sys.exit(2)

    total_size = 0
    summary = []
    for vendor in matched_vendors:
        print(f'■ ベンダー: {vendor["name"]}')
        cases = client.list_folders(vendor['id'])
        for case in cases:
            wages = find_wage_files_recursive(client, case['id'])
            if not wages:
                continue
            # 案件フォルダ名を保存ディレクトリ名にする（先頭の連番は除く）
            case_label = case['name']
            for prefix in ('001.', '002.', '003.', '01.', '02.', '03.', '04.', '05.', '06.', '07.', '08.', '09.'):
                if case_label.startswith(prefix):
                    case_label = case_label[len(prefix):]
                    break
            # サフィックス（_通常枠（5万円～...）など）も削る
            for suffix_marker in ('_通常枠', '_インボイス'):
                if suffix_marker in case_label:
                    case_label = case_label.split(suffix_marker)[0]
            case_label = safe_filename(case_label.strip())
            case_dir = base_dir / case_label
            case_dir.mkdir(exist_ok=True)
            print(f'  📂 {case["name"]} → {case_label}/')
            case_size = 0
            for f in wages:
                try:
                    out_path = case_dir / safe_filename(f['name'])
                    if out_path.exists() and out_path.stat().st_size > 0:
                        size = out_path.stat().st_size
                        print(f'    ⏩ {f["name"]} (キャッシュあり {size//1024}KB)')
                        case_size += size
                        total_size += size
                        continue
                    client.download_file(f['id'], out_path, f.get('mimeType'))
                    size = out_path.stat().st_size
                    case_size += size
                    total_size += size
                    print(f'    ✓ {f["name"]} ({size//1024}KB)')
                except Exception as e:
                    print(f'    ✗ DL失敗 {f["name"]}: {e}')
            summary.append((case_label, len(wages), case_size))
        print()

    print('=' * 60)
    print('集計')
    print('=' * 60)
    for name, count, sz in summary:
        print(f'  {name}: {count}ファイル / {sz / 1_000_000:.2f}MB')
    print(f'  合計: {total_size / 1_000_000:.2f}MB')


if __name__ == '__main__':
    main()
