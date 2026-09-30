"""case-docs-check（配布版）の check_docs.py を、合成の書類一式で動かして確かめる。

- Drive 経路（manifest）: 必須がそろえば終了コード0、分類と「申請時使用」の除外が効く
- ローカル経路（--local）: 空フォルダは必須欠落で終了コード1。チャットの添付を想定した
  平らなフォルダでも分類でき、入力フォルダには何も書き込まない
"""
import json
import tempfile
from pathlib import Path

import openpyxl

from _common import Checker, run_py, skills_dir

CHECK = skills_dir() / 'case-docs-check' / 'review' / 'check_docs.py'
PDF_STUB = b'%PDF-1.4\n%synthetic\n'
MANIFEST_FILES = [
    {'name': 'ヒアリングシート_テスト社.xlsx', 'mimeType': 'application/x', 'parent_path': ''},
    {'name': '登記事項証明書.pdf', 'mimeType': 'application/pdf', 'parent_path': ''},
    {'name': 'テスト社_第3期決算書.pdf', 'mimeType': 'application/pdf', 'parent_path': '1.交付申請'},
    {'name': '要約決算.pdf', 'mimeType': 'application/pdf', 'parent_path': '申請時使用'},
    {'name': '賃金台帳2025', 'mimeType': 'application/vnd.google-apps.spreadsheet', 'parent_path': ''},
]


def run_check(work: Path, *args):
    out = work / 'result.json'
    if out.exists():
        out.unlink()
    cp = run_py(CHECK, *args, '--json', out, cwd=work)
    data = json.loads(out.read_text(encoding='utf-8')) if out.exists() else None
    return cp, data


def main() -> int:
    c = Checker('case-docs-check（配布版）')
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        manifest = work / 'manifest.json'
        manifest.write_text(json.dumps({'files': MANIFEST_FILES}, ensure_ascii=False), encoding='utf-8')
        cp, data = run_check(work, '--manifest', manifest, '--entity', '法人')
        c.check('Drive 経路: 必須がそろえば終了コード0', cp.returncode == 0, cp.stdout[-600:] + cp.stderr)
        detected = (data or {}).get('detected', {})
        c.check('Drive 経路: ヒアリング・登記・決算書・賃金台帳に分類される',
                {'hearing', 'registry', 'pl', 'wage_ledger'} <= set(detected), str(detected))
        c.check('Drive 経路: 「申請時使用」の中は除外される', '要約決算.pdf' not in json.dumps(detected, ensure_ascii=False))

        empty = work / 'empty'
        empty.mkdir()
        cp, data = run_check(work, '--local', empty, '--entity', '法人', '--task', 'application')
        missing = [f for f in (data or {}).get('findings', []) if f.get('code') == 'MISSING-REQUIRED']
        c.check('ローカル経路: 空フォルダは必須欠落3件で終了コード1', cp.returncode == 1 and len(missing) == 3,
                f'rc={cp.returncode} {missing}')

        uploads = work / 'uploads'
        uploads.mkdir()
        (uploads / '登記事項証明書.pdf').write_bytes(PDF_STUB)
        (uploads / 'テスト社_第3期決算書.pdf').write_bytes(PDF_STUB)
        wb = openpyxl.Workbook()
        wb.active.title = 'Sheet1'
        wb.save(uploads / 'ヒアリングシート_テスト社.xlsx')
        before = sorted(p.name for p in uploads.iterdir())
        cp, data = run_check(work, '--local', uploads, '--entity', '法人', '--task', 'application')
        after = sorted(p.name for p in uploads.iterdir())
        c.check('ローカル経路: 添付の平らなフォルダでも結果の JSON が出る', data is not None, cp.stderr)
        c.check('ローカル経路: 登記と決算書を分類できる',
                {'registry', 'pl'} <= set((data or {}).get('detected', {})), str((data or {}).get('detected')))
        c.check('ローカル経路: 入力フォルダに何も書き込まない', before == after, f'{before} -> {after}')
        c.check('スクリプトエラー（終了コード2）にならない', cp.returncode in (0, 1), f'rc={cp.returncode}')
    return c.finish()


if __name__ == '__main__':
    raise SystemExit(main())
