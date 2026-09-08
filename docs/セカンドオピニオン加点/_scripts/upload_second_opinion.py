# -*- coding: utf-8 -*-
"""セカンドオピニオン加点の配布物を Drive へ差し替える。

使い方:
    python docs/セカンドオピニオン加点/_scripts/upload_second_opinion.py            # dry-run（既定）
    python docs/セカンドオピニオン加点/_scripts/upload_second_opinion.py --apply     # 実行
    python docs/セカンドオピニオン加点/_scripts/upload_second_opinion.py --only 配車頭

方針:
- 既存ファイルは `files.update` で**中身だけ**差し替える。ファイルIDが変わらない＝
  営業に配った共有リンクが切れない。同名が無ければ新規作成する。
- 会社フォルダ名は `build_company_sheets.COMPANY_FOLDERS` を再利用（二重管理しない）。
- サービスアカウントが Drive の「セカンドオピニオン」フォルダにコンテンツ管理者として
  共有されている必要がある（共有ドライブのメンバーではなくフォルダ単位の共有）。
  外れていると `files.get` が 404 になる（`files.list` は例外を出さず0件を返すので切り分け注意）。
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parents[2]          # 補助金/
sys.path.insert(0, str(REPO_ROOT))
sys.path.insert(0, str(SCRIPT_DIR))

from dotenv import load_dotenv                                   # noqa: E402
from hojokin.google_drive import DriveClient                     # noqa: E402
import build_company_sheets as bcs                               # noqa: E402

TEMPLATE_DIR = SCRIPT_DIR.parent / '配布物テンプレ'
DIST_ROOT = TEMPLATE_DIR / '05_会社別' / '_配布'

SHEETS_FOLDER = '05_営業ヒアリングシート各種'
PROMPTS_FOLDER = 'プロンプト各種【メール案内文、電話時ヒアリングメモ、予約フォーム入力文】'
GENERIC_FOLDER = '汎用型'

# 別紙は会社別・汎用のどちらでも同じ紙を使うので、会社フォルダから1階層上に置く
BESSHI = [
    '05_別紙_お客様向け事前案内.pdf',
    '05_別紙_営業向けカンペ.pdf',
]
# 会社別に分けないプロンプト類（置き場所は Drive の現状に合わせる）。
# Drive 側の名前がローカルと違うものがあるので (ローカル名, Drive名) で持つ。
# 名前だけで突き合わせると、同じ中身の新規ファイルを作って二重になる。
PROMPTS = [
    ('07b_お客様確認メール生成プロンプト.md', '07b_お客様確認メール生成プロンプト.md'),
    ('06_入力整理プロンプト.md', '06_予約フォーム入力文生成プロンプト.md'),
]
SPEC = '00_予約フォーム仕様.md'

CREDENTIALS = REPO_ROOT / 'credentials' / 'service_account.json'


def service_account_email(path: Path) -> str:
    """鍵ファイルの client_email。アドレスはリポジトリに書かず、表示するときに読む。"""
    try:
        email = json.loads(path.read_text(encoding='utf-8-sig'))['client_email']
    except (OSError, ValueError, KeyError, TypeError):
        email = None
    if isinstance(email, str) and email:
        return email
    return f'サービスアカウント（{path.name} の client_email）'


class Planner:
    """Drive 上の配置を解決して、update / create / move の計画を組む。"""

    def __init__(self, client: DriveClient, root_id: str):
        self.client = client
        self.root_id = root_id
        self._folders: dict[tuple[str, ...], str] = {(): root_id}
        self.actions: list[tuple] = []

    def folder_id(self, *path: str) -> str | None:
        """セカンドオピニオン起点のフォルダパスを ID に解決（キャッシュ付き）。"""
        if path in self._folders:
            return self._folders[path]
        parent = self.folder_id(*path[:-1])
        if parent is None:
            return None
        found = self.client.find_child_by_name(parent, path[-1])
        fid = found['id'] if found else None
        self._folders[path] = fid
        return fid

    def plan(self, local: Path, *folder_path: str, drive_name: str | None = None) -> None:
        assert local.exists(), f'ローカルに無い: {local}'
        name = drive_name or local.name
        parent = self.folder_id(*folder_path)
        display = '/'.join(folder_path + (name,))
        if parent is None:
            self.actions.append(('SKIP', display, local, None, 'Drive にフォルダが無い'))
            return
        existing = self.client.find_child_by_name(parent, name)
        if existing:
            self.actions.append(('UPDATE', display, local, existing, None))
        else:
            self.actions.append(('CREATE', display, local, parent, None))

    def plan_besshi_move(self) -> None:
        """汎用型/ に取り残されている別紙を、05_営業ヒアリングシート各種 直下へ上げる。"""
        dest = self.folder_id(SHEETS_FOLDER)
        src = self.folder_id(SHEETS_FOLDER, GENERIC_FOLDER)
        if dest is None or src is None:
            return
        for name in BESSHI:
            if self.client.find_child_by_name(dest, name):
                continue                                  # すでに直下にある
            stray = self.client.find_child_by_name(src, name)
            if stray:
                self.actions.append(
                    ('MOVE', f'{SHEETS_FOLDER}/{GENERIC_FOLDER}/{name} → {SHEETS_FOLDER}/{name}',
                     None, stray, (src, dest)))


def build_plan(planner: Planner, only: str | None) -> None:
    planner.plan_besshi_move()

    for name in BESSHI:
        planner.plan(TEMPLATE_DIR / name, SHEETS_FOLDER)
    for local_name, drive_name in PROMPTS:
        planner.plan(TEMPLATE_DIR / local_name, PROMPTS_FOLDER, drive_name=drive_name)
    planner.plan(TEMPLATE_DIR / SPEC)

    for _md, folder in bcs.company_sheets():
        if only and only not in folder:
            continue
        dist = DIST_ROOT / folder
        assert dist.is_dir(), f'先に build_company_sheets.py を回してください: {dist}'
        for f in sorted(dist.iterdir()):
            planner.plan(f, SHEETS_FOLDER, folder)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--apply', action='store_true', help='実際に Drive へ書き込む（既定は dry-run）')
    ap.add_argument('--only', help='会社フォルダ名の部分一致で絞る')
    args = ap.parse_args()

    load_dotenv(REPO_ROOT / '.env')
    root_id = os.getenv('SECOND_OPINION_FOLDER_ID')
    if not root_id:
        print('✗ .env に SECOND_OPINION_FOLDER_ID（Drive の「セカンドオピニオン」フォルダID）がありません')
        return 1
    if not CREDENTIALS.exists():
        print(f'✗ サービスアカウント鍵が無い: {CREDENTIALS}')
        return 1

    client = DriveClient(str(CREDENTIALS))
    if not client.can_access(root_id):
        print('✗ サービスアカウントから Drive のフォルダが見えません（404）')
        print(f'  {service_account_email(CREDENTIALS)} を')
        print('  Drive の「セカンドオピニオン」フォルダにコンテンツ管理者として共有してください')
        return 1

    planner = Planner(client, root_id)
    build_plan(planner, args.only)

    counts: dict[str, int] = {}
    for kind, display, local, target, extra in planner.actions:
        counts[kind] = counts.get(kind, 0) + 1
        if kind == 'UPDATE':
            was = int(target.get('size') or 0)
            now = local.stat().st_size
            mark = '' if was != now else '  ※同サイズ'
            print(f'UPDATE  {display}  ({was:,} → {now:,} bytes){mark}')
        elif kind == 'CREATE':
            print(f'CREATE  {display}  ({local.stat().st_size:,} bytes)')
        elif kind == 'MOVE':
            print(f'MOVE    {display}')
        else:
            print(f'SKIP    {display}  — {extra}')

    print()
    print('計 ' + ' / '.join(f'{k} {v}件' for k, v in sorted(counts.items())))

    if counts.get('SKIP'):
        print('✗ SKIP があります。Drive のフォルダ名を確認してください（実行を中止）')
        return 1
    if not args.apply:
        print('dry-run です。実行するには --apply を付けてください')
        return 0

    print()
    for kind, display, local, target, extra in planner.actions:
        if kind == 'MOVE':
            src, dest = extra
            client.move_file(target['id'], dest, src)
        elif kind == 'UPDATE':
            client.update_file(target['id'], local)
        elif kind == 'CREATE':
            client.upload_file(local, target)
        print(f'done  {kind}  {display}')

    print()
    print(f'完了: {len(planner.actions)}件を Drive へ反映しました')
    return 0


if __name__ == '__main__':
    sys.exit(main())
