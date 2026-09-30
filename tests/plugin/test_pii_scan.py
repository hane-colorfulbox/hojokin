"""PII スキャナ（scripts/plugin_pii_scan.py）を合成データだけで確かめる。

実在の名前は使わない（名簿も合成）。検出すべきもの・誤検知してはいけないもの・xlsx の文字参照・
許可リスト（sha256 だけを保存）・名簿の欠落/不足・git の追加行だけの走査を見る。
"""
import subprocess
import tempfile
import zipfile
from pathlib import Path

from _common import Checker

import plugin_pii_scan as ps

ROSTER = ['テスト商事', 'テスト'] + [f'ダミー社{i:02d}' for i in range(ps.ROSTER_MIN_ENTRIES)]
# 検出させたい文字列は実行時に連結して作る（このファイル自体がコミット時の PII 検査に当たらないように）
FAKE_ID = '1AbCdEfGhIjKlMnOp' + 'QrStUvWxYz012345'
SHOULD_HIT = {
    'roster': 'テスト商事の案件を確認した',
    'email-internal': '連絡先 ab.cd@' + 'colorfulbox.co.jp',
    'email': '先方 someone@' + 'sample-company.jp',
    'google-url': '資料 https://drive.' + 'google.com/drive/folders/' + FAKE_ID,
    'long-id': 'ファイルID ' + FAKE_ID + ' を開く',
    'name-spaced': '担当は山田' + '　一美です',
    'name-honorific': '鈴木' + 'さんに確認する',
    'name-gender': '007 甲乙丙丁' + '（女）',
    'amount': '給与総額 4,321,' + '987円',
}
SHOULD_NOT_HIT = [
    'お客様・皆様・担当者様に送る仕様と様式',
    'サンプル建設様の見本',
    '見本は 山田 太郎 さんの形式（ダミー）',
    '上限 1,500,000円 と見本値 12,345,678',
    "モデル 'claude-haiku-4-5-20251001' と色 '00000000'",
    'https://drive.google.com/drive/folders/<ID> の形',
    '連絡先 someone@example.org',
    'コンテストの結果',
]
ENTITY_NAME = ''.join(f'&#{ord(ch)};' for ch in 'テスト商事')


def rules_for(tmp: Path):
    roster = tmp / 'roster.txt'
    roster.write_text('# 合成\n' + '\n'.join(ROSTER) + '\n', encoding='utf-8')
    return ps.build_rules(ps.load_roster(roster), []), roster


def main() -> int:
    c = Checker('PII スキャナ（合成データ）')
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        rules, roster = rules_for(tmp)
        for rule, line in SHOULD_HIT.items():
            hits = ps.scan_text('x.md', line, rules)
            c.check(f'検出する: {rule}', any(h.rule == rule for h in hits), str([(h.rule, h.matched) for h in hits]))
        for line in SHOULD_NOT_HIT:
            hits = ps.scan_text('x.md', line, rules)
            c.check(f'誤検知しない: {line[:24]}', not hits, str([(h.rule, h.matched) for h in hits]))

        hit = ps.scan_text('x.md', SHOULD_HIT['roster'], rules)[0]
        c.check('表示は伏せ字（先頭1文字だけ）', hit.masked().startswith('テ') and 'スト商事' not in hit.masked())

        book = tmp / 'book.xlsx'
        with zipfile.ZipFile(book, 'w') as zf:
            zf.writestr('xl/sharedStrings.xml', f'<sst><si><t>{ENTITY_NAME}</t></si></sst>')
        hits = ps.scan_path(book, rules)
        c.check('xlsx の文字参照（&#...;）で書かれた名前も検出する', any(h.rule == 'roster' for h in hits))

        ps.ALLOWLIST_PATH = tmp / 'allow.json'
        found = ps.scan_text('x.md', SHOULD_HIT['amount'], rules)
        ps.approve(found, '合成テスト')
        c.check('許可リストに入れた行は次から出ない', not ps.filter_allowed(found, ps.load_allowlist()))
        c.check('許可リストには語を保存しない（sha256 だけ）',
                SHOULD_HIT['amount'].split()[-1] not in ps.ALLOWLIST_PATH.read_text(encoding='utf-8'))

        try:
            ps.load_roster(tmp / 'nothing.txt')
            c.check('名簿が無ければエラー', False)
        except FileNotFoundError:
            c.check('名簿が無ければエラー', True)
        small = tmp / 'small.txt'
        small.write_text('テスト商事\n', encoding='utf-8')
        try:
            ps.load_roster(small)
            c.check('名簿が少なすぎればエラー', False)
        except ValueError:
            c.check('名簿が少なすぎればエラー', True)

        repo = tmp / 'git'
        repo.mkdir()
        git = ['git', '-c', 'user.name=t', '-c', 'user.email=t@example.com', '-c', 'commit.gpgsign=false']
        subprocess.run(['git', 'init', '-q'], cwd=repo, check=True)
        (repo / 'a.md').write_text('既存の行 テスト商事\n', encoding='utf-8')
        subprocess.run([*git, 'add', '.'], cwd=repo, check=True)
        subprocess.run([*git, 'commit', '-qm', 'init'], cwd=repo, check=True)
        (repo / 'a.md').write_text('既存の行 テスト商事\n追加 ' + SHOULD_HIT['name-honorific'] + '\n',
                                   encoding='utf-8')
        subprocess.run([*git, 'add', '.'], cwd=repo, check=True)
        hits = ps.scan_git_diff(repo, ['--cached'], rules)
        c.check('git の走査は追加行だけを見る',
                [h.rule for h in hits] == ['name-honorific'] and hits[0].line_no == 2, str([(h.rule, h.line_no) for h in hits]))
    return c.finish()


if __name__ == '__main__':
    raise SystemExit(main())
