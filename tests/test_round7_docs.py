"""第7次の資料一式を合成データ・作業ツリーで確認する。API・通信は使わない。

1. 必要書類チェックの新旧台帳・実申請月・3期またぎ（case-docs-check）
2. 第7次の正本Markdown（対象読者・題名・リンク・確定事実・禁止表現）
3. 配布物の生成（PDF・案内HTML・README・ZIP。作業ツリーから一時フォルダへ）
"""
import importlib.util
import json
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / 'docs'
CHECK = ROOT / '.claude/skills/case-docs-check/review/check_docs.py'
GUIDE = ROOT / 'scripts/build_official_guide.py'
JP_FONT = Path('C:/Windows/Fonts/meiryo.ttc')
EDITION_DATE = '2026-11-05'
STATUS_NOTE = '合成テスト用の状況メモ'


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


mod = load_module('case_docs_rounds_under_test', CHECK)
guide = load_module('official_guide_under_test', GUIDE)
sys.path.insert(0, str(ROOT / 'scripts'))
import distribution_manifest as dm  # noqa: E402

# 正本ごとに必ず書かれているべき確定事実（facts は公式原文で確認済みのもの）
REQUIRED = {
    '第7次申請_担当者向け要点.md': ('2026年12月1日（火）17:00', 'v0.2.99', '2026.11.05', '1,226円以上1,280円未満',
                                '令和8年度改定', '55円', '申請月の前月', '補助金名の年度', '月の見出しに年が入っていない'),
    '第7次申請_変更点と使い方.md': ('労働基準法第20条', '労働基準法第21条', '最低賃金額以上かどうかを確認する方法',
                                '1,031円', '1,032円', '「すべて更新」を実行しません', '年10.95%', '2026年11月30日',
                                '山梨・徳島・愛媛・大分', '通常枠150万円以上＝要件（ト）（ナ）の申請だけ', 'No.40',
                                '1年後要件と年平均要件の両方', '決算がまだ確定していない', 'セキュリティ枠は、本ツール'),
    '第7次申請_公式資料で判断できることと確認の範囲.md': ('No.40', 'No.44', 'p.86', '労働基準法第20条', '2026/11/30',
                                                 '加算金', '最低賃金額以上かどうかを確認する方法'),
    '第7次申請_確認待ち事項.md': ('前事業年度の決算が未確定', '代表者の確定申告書の事業所得欄', '効果報告の比較基準',
                              '照会日と回答の要旨'),
    '第7次申請_ツール入力手順.md': ('v0.2.99', '補助金名の年度', '交付決定日の年（暦年）ではありません',
                                '月の見出しに年が入っていない', 'セキュリティ枠は対象外', '（管理者・開発者向け）'),
    '第7次原本_改訂内容と再取込手順.md': ('hojokin:2026_7_plus', 'docProps/core.xml', 'select_application_template'),
}
FORBIDDEN = ('ユーザー', '前回', '前版', '本番・Drive・全員配信は未実施', '毎年10月に全国一斉', '再照会不要')
LINK = re.compile(r'(?<!!)\[[^\]]*\]\(([^)\s]+)\)')


def make_ledger(path, call_round, app_month='2026/11', mixed_hours=False):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = mod.BWL_SHEET
    ws['C2'] = '東京都'
    ws['C3'] = app_month
    if call_round >= 7:
        ws['AG2'] = '2026_rounds_7_plus'
        ws['AG3'] = call_round
    for col, text in enumerate(mod.bonus_headers(call_round), 2):
        ws.cell(6, col, text)
    hours_round = 6 if mixed_hours else call_round
    for col, text in enumerate(mod.bonus_headers(hours_round)[4:-1], 19):
        ws.cell(6, col, text.replace('基本給', '労働時間'))
    ws['AE6'] = '交付申請直近月\n労働時間'
    wb.save(path)
    wb.close()


def check_ledger(path, call_round, application_ym=(2026, 11), level='FAIL'):
    report = mod.Report()
    detected = {'wage_ledger': [{'name': path.name, 'local_path': str(path), 'ext': '.xlsx'}]}
    mod.check_excel_structures(report, detected, 'katen', '不明', '不明',
                               call_round, application_ym)
    return {f['code'] for f in report.findings if f['level'] == level}


def check_case_docs_rounds():
    assert mod.bonus_headers(6)[4] == '令和6年10月\n基本給'
    assert mod.bonus_headers(7)[4] == '令和7年10月\n基本給'
    assert mod.bonus_headers(8)[-2] == '令和8年9月\n基本給'
    assert mod.parse_application_month('2026/11') == (2026, 11)
    assert mod.parse_application_month('2026/13') is None
    assert mod.parse_application_month('11月') is None
    months = mod.required_wage_months(7, (2026, 12))
    assert (2025, 10) in months and (2026, 6) in months and (2026, 11) in months
    assert len(months) == 13
    # 10月末決算：2025/10・2026/6・2026/11の期末年が3種類になる。
    fiscal_ends = {y if m <= 10 else y + 1 for y, m in months}
    assert fiscal_ends == {2025, 2026, 2027}
    assert (2026, 12) in mod.required_wage_months(8, (2027, 1))
    try:
        mod.bonus_headers(0)
        raise AssertionError('invalid call round was accepted')
    except ValueError:
        pass
    with tempfile.TemporaryDirectory(prefix='docs-rounds-') as tmp:
        folder = Path(tmp)
        path = folder / '合成_加点判定用賃金台帳.xlsx'
        make_ledger(path, 6)
        assert not check_ledger(path, 6)
        assert 'TPL-BONUS-HEADER' in check_ledger(path, 7)
        make_ledger(path, 7)
        assert not check_ledger(path, 7)
        # AG3=7の台帳を第8次で照合：制度版（2026_rounds_7_plus）が同じなので FAIL ではなく WARN
        assert 'BONUS-ROUND-MISMATCH' not in check_ledger(path, 8)
        assert 'BONUS-ROUND-MISMATCH' in check_ledger(path, 8, level='WARN')
        # 制度版が違う第6次で照合すると FAIL のまま
        assert 'BONUS-ROUND-MISMATCH' in check_ledger(path, 6)
        assert 'TPL-BONUS-HEADER' in check_ledger(path, 6)
        make_ledger(path, 7, mixed_hours=True)
        assert 'TPL-BONUS-HOURS-HEADER' in check_ledger(path, 7)
        make_ledger(path, 7, app_month='2026/12')
        assert 'BONUS-APPYM-MISMATCH' in check_ledger(path, 7)
        make_ledger(path, 7, app_month='2026/13')
        assert 'BONUS-APPYM-INVALID' in check_ledger(path, 7)
        make_ledger(path, 7)
        out = folder / 'check.json'
        cp = subprocess.run([sys.executable, '-B', str(CHECK), '--local', str(folder),
                             '--task', 'katen', '--call-round', '7', '--app-month', '2026/11',
                             '--json', str(out)], capture_output=True, text=True, encoding='utf-8')
        assert cp.returncode == 0, cp.stdout + cp.stderr
        result = json.loads(out.read_text(encoding='utf-8'))
        assert result['context']['call_round'] == 7
        assert result['context']['application_ym'] == [2026, 11]
        assert [2026, 10] in result['context']['bonus_required_months']
        cp = subprocess.run([sys.executable, '-B', str(CHECK), '--local', str(folder),
                             '--task', 'katen', '--json', str(out)],
                            capture_output=True, text=True, encoding='utf-8')
        result = json.loads(out.read_text(encoding='utf-8'))
        assert any(f['code'] == 'CALL-ROUND-UNCONFIRMED' for f in result['findings'])
    print('PASS: 新旧ヘッダー、時間列、実申請月、必要月、3期またぎ、年またぎ、未確認回の報告')
    return 0


def check_round7_guides():
    names = [p.split('/', 1)[1] for p in dm.ROUND7_GUIDES]
    assert set(names) == set(guide.STAFF_GUIDES + guide.ADMIN_GUIDES) == set(REQUIRED), names
    assert names[0] == guide.YOTEN, '担当者向け要点を最初に置く'
    for name in names:
        text = (DOCS / name).read_text(encoding='utf-8')
        lines = [s for s in text.splitlines() if s.strip()]
        title_head = name[:-3].replace('_', '：', 1)
        assert lines[0].startswith('# ' + title_head), (name, lines[0])
        assert lines[1].startswith('対象読者：'), (name, lines[1])
        for word in REQUIRED[name]:
            assert word in text, (name, word)
        for word in FORBIDDEN + dm.STALE_TERMS:
            assert word not in text, (name, word)
        for line in text.splitlines():
            if '63円' in line:   # 旧回の値は、新旧を並べた行か「第1〜6次」と明示した行でだけ書く
                assert '55円' in line or '第1〜6次' in line, (name, line)
        for target in LINK.findall(text):
            if not re.match(r'https?://', target):
                assert target in names, (name, target)
    print(f'PASS: 第7次の正本{len(names)}本（題名・対象読者・確定事実・禁止表現・リンク）')
    return 0


def check_template_command():
    """原本再取込手順に書いた確認コマンドが、そのまま動いて4種類とも第7次版を選ぶこと。"""
    text = (DOCS / '第7次原本_改訂内容と再取込手順.md').read_text(encoding='utf-8')
    line = next(s.strip() for s in text.splitlines() if 'select_application_template' in s and '-c "' in s)
    code = line.split('-c "', 1)[1].rsplit('"', 1)[0]
    cp = subprocess.run([sys.executable, '-X', 'utf8', '-B', '-c', code], cwd=ROOT, capture_output=True,
                        text=True, encoding='utf-8')
    assert cp.returncode == 0, cp.stderr
    rows = cp.stdout.strip().splitlines()
    assert len(rows) == 4 and all('第7次以降' in r and 'None' not in r for r in rows), rows
    print('PASS: 原本再取込手順の確認コマンドで4原本とも第7次版が選ばれる')
    return 0


def check_unsupported_markdown():
    styles = guide.make_styles('Helvetica', 'Helvetica')
    for sample in ('> 引用', '#### 見出し4', '---', '```\nコード\n```', '- 親\n    - 子', '![図](x.png)',
                   '- 項目\n\n        コード\n'):
        try:
            guide.markdown_to_story('合成', sample, styles)
        except guide.GuideBuildError:
            continue
        raise AssertionError(f'未対応の書式が黙って通りました: {sample!r}')
    ok = '## 見出し\n\n本文 **太字** [公式](https://example.com) `code`\n\n| a | b |\n|---|---|\n| 1 | 2 |\n\n1. 一\n2. 二\n'
    assert len(guide.markdown_to_story('合成', ok, styles)) >= 5
    print('PASS: PDFに未対応のMarkdown要素は停止する')
    return 0


def check_guide_build():
    if not JP_FONT.is_file():
        print(f'SKIP: 日本語フォントがないため配布物の生成は確認しない（{JP_FONT}）')
        return 0
    from pypdf import PdfReader
    with tempfile.TemporaryDirectory(prefix='official-guide-') as tmp:
        out = Path(tmp) / 'out'
        args = ['--from-worktree', '--edition-date', EDITION_DATE, '--out-dir', str(out), '--status-note', STATUS_NOTE]
        assert guide.main(args) == 0
        pdf = out / guide.PDF_NAME.format(date=EDITION_DATE)
        archive = out / guide.ZIP_NAME.format(date=EDITION_DATE)
        first = re.sub(r'\s+', '', PdfReader(pdf).pages[0].extract_text())
        assert guide.FIRST_PAGE_MARK in first and STATUS_NOTE in first and '本番・Drive' not in first
        with zipfile.ZipFile(archive) as zf:
            members = set(zf.namelist())
            readme = zf.read(guide.README_NAME).decode('utf-8')
            pages = {n: zf.read(n).decode('utf-8') for n in members if n.endswith('.html')}
        expected = {guide.README_NAME, pdf.name} | {guide.html_path(n) for n in guide.STAFF_GUIDES + guide.ADMIN_GUIDES}
        assert members == expected, members ^ expected
        assert all(guide.html_path(n).startswith(guide.ADMIN_DIR + '/') for n in guide.ADMIN_GUIDES)
        for word in ('読む順番', '対象読者', guide.DEFAULT_TOOL_MIN, guide.DEFAULT_PLUGIN_MIN, '管理者・開発者向け', STATUS_NOTE):
            assert word in readme, word
        for name, page in pages.items():
            assert page.count('<h1>') == 1 and '\r' not in page, name
            assert not re.search(r'href="(?!https?://)[^"]*\.md"', page), name
        summary = json.loads((out / guide.SUMMARY_NAME).read_text(encoding='utf-8'))
        assert summary['source']['mode'] == 'worktree' and summary['pdf']['pages'] >= guide.PDF_PAGE_RANGE[0]
        try:
            guide.main(args)
            raise AssertionError('既存の出力先が上書きされました')
        except SystemExit as e:
            assert e.code == 2
    print(f'PASS: 配布物の生成（PDF {summary["pdf"]["pages"]}ページ・HTML{len(pages)}本・README・ZIP、上書き拒否）')
    return 0


def test_document_rounds():
    assert check_case_docs_rounds() == 0


def test_round7_guides():
    assert check_round7_guides() == 0


def test_template_command():
    assert check_template_command() == 0


def test_unsupported_markdown():
    assert check_unsupported_markdown() == 0


def test_guide_build():
    assert check_guide_build() == 0


def main():
    for check in (check_case_docs_rounds, check_round7_guides, check_template_command,
                  check_unsupported_markdown, check_guide_build):
        assert check() == 0
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
