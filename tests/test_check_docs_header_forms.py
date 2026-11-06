# -*- coding: utf-8 -*-
"""標準賃金台帳の年付きの月見出しを、3か所が同じ形で受け付けるかの回帰テスト。合成データだけで API は呼ばない。

- case-docs-check の check_docs.py（_year_month_of と TPL-WAGE-HEADER）
- wagebook-convert の verify_wagebook.py（_header_is_clean_month＝FLAG-H）
- ツールの hojokin/wage_reader.py（parse_ym_header と read_wage_ledgers の対象年月）

check_docs が年付きとして受け付けるのは、verify が規格どおりとする形のうち年があるものと同じ集合で、
読み取る年月はツールと一致する。注記入り（「4月\\n(R7.4)」など）は受け付けない。
"""
import datetime
import importlib.util
import os
import shutil
import sys
import tempfile
import unittest
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['CLAUDE_API_KEY'] = ''
warnings.filterwarnings('ignore', category=UserWarning)

import openpyxl

from hojokin.wage_reader import parse_ym_header, read_wage_ledgers

CHECK_DOCS = ROOT / '.claude' / 'skills' / 'case-docs-check' / 'review' / 'check_docs.py'
VERIFY = ROOT / '.claude' / 'skills' / 'wagebook-convert' / 'review' / 'verify_wagebook.py'
STD_TEMPLATE = ROOT / 'ツール' / '賃金台帳テンプレート.xlsx'
SHEET = '従業員別明細'
HEADER_ROW, FIRST_ROW, COL_MONTH = 5, 6, 7
REIWA_BASE = 2018
MONTHLY_PAY, MONTHLY_HOURS = 250_000, 160
# 3月決算の前事業年度（令和7年4月〜令和8年3月）
FISCAL_YMS = [(2025, m) for m in range(4, 13)] + [(2026, m) for m in range(1, 4)]
FULLWIDTH_DIGITS = str.maketrans('0123456789', '０１２３４５６７８９')

# 見出しの書き方の見本。年付きで受け付けるものと、受け付けないもの（年なし・注記入り・未対応の形）を混ぜる。
HEADER_SPELLINGS = [
    '令和7年4月', '令和7年 4月', '令和7年\n4月', ' 令和7年4月 ', '令和7年　4月', '令和７年４月', '平成31年4月',
    'R7.4', 'r7.4', 'R7/4', 'R7-4', 'R7.4月', 'H31.4', 'Ｒ７．４',
    '2025年4月', '2025/04', '2025/4', '2025-04', '2025.4', '2025/4月', '２０２５年４月', '2025年\n4月',
    '4月', '１２月', '4月\n(R7.4)', '1月(R7.1)', '令和7年4月(注)', '令和7年4月分', '2025年4月度', 'R7.4(4月分)',
    'R7年4月', '令和元年4月', '202504', '4月度', '2025年', '', None,
]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def era_newline(y, m):
    return f'令和{y - REIWA_BASE}年\n{m}月'


def era_space(y, m):
    return f'令和{y - REIWA_BASE}年 {m}月'


def r_dot(y, m):
    return f'R{y - REIWA_BASE}.{m}'


def west_slash(y, m):
    return f'{y}/{m:02d}'


def west_dash(y, m):
    return f'{y}-{m:02d}'


def era_fullwidth(y, m):
    return f'令和{y - REIWA_BASE}年{m}月'.translate(FULLWIDTH_DIGITS)


ACCEPTED_FORMS = (era_newline, era_space, r_dot, west_slash, west_dash, era_fullwidth)


class HeaderFormParityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.check_docs = load_module('check_docs_for_header_forms', CHECK_DOCS)
        cls.verify = load_module('verify_wagebook_for_header_forms', VERIFY)

    def test_accepted_set_matches_verify_and_tool(self):
        for text in HEADER_SPELLINGS:
            with self.subTest(header=text):
                got = self.check_docs._year_month_of(text)
                tool = parse_ym_header(text)
                year_stamped_clean = (self.verify._header_is_clean_month(text)
                                      and tool is not None and tool[0] is not None)
                self.assertEqual(got is not None, year_stamped_clean)
                if got is not None:
                    self.assertEqual(got, tool)

    def test_annotated_and_yearless_headers_are_not_year_stamped(self):
        for text in ('4月\n(R7.4)', '1月(R7.1)', '令和7年4月(注)', '4月'):
            with self.subTest(header=text):
                self.assertIsNone(self.check_docs._year_month_of(text))


class LedgerHeaderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.check_docs = load_module('check_docs_for_ledger_headers', CHECK_DOCS)

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def ledger(self, headers, name='合成_賃金台帳一覧.xlsx'):
        path = self.path / name
        shutil.copyfile(STD_TEMPLATE, path)
        wb = openpyxl.load_workbook(path)
        ws = wb[SHEET]
        for i, text in enumerate(headers):
            ws.cell(HEADER_ROW, COL_MONTH + i).value = text
        for c in range(2, 21):
            ws.cell(FIRST_ROW, c).value = None
        for col, value in ((2, 1), (3, '合成 太郎'), (4, '正社員'), (5, MONTHLY_HOURS)):
            ws.cell(FIRST_ROW, col).value = value
        for i in range(len(headers)):
            ws.cell(FIRST_ROW, COL_MONTH + i).value = MONTHLY_PAY
        wb.save(path)
        wb.close()
        return path

    def header_findings(self, path, frame='通常枠', call_round=7):
        report = self.check_docs.Report()
        self.check_docs.check_excel_structures(
            report, {'wage_ledger': [{'name': path.name, 'local_path': str(path), 'ext': '.xlsx'}]},
            'application', frame, '法人', call_round, None)
        return [f for f in report.findings if 'WAGE-HEADER' in f['code']]

    def test_forms_the_tool_reads_pass_case_docs_check(self):
        for form in ACCEPTED_FORMS:
            with self.subTest(form=form.__name__):
                path = self.ledger([form(*ym) for ym in FISCAL_YMS], f'{form.__name__}.xlsx')
                self.assertEqual(self.header_findings(path), [])
                employee = read_wage_ledgers([path])[0]
                self.assertEqual(employee.fiscal_window, FISCAL_YMS)
                self.assertEqual(employee.fiscal_window_issue, '')

    def test_date_typed_headers_pass_like_the_tool(self):
        path = self.ledger([datetime.datetime(y, m, 1) for y, m in FISCAL_YMS], 'date_typed.xlsx')
        self.assertEqual(self.header_findings(path), [])
        self.assertEqual(read_wage_ledgers([path])[0].fiscal_window, FISCAL_YMS)

    def test_calendar_position_layout_passes(self):
        headers = [r_dot(*ym) for ym in FISCAL_YMS[9:] + FISCAL_YMS[:9]]
        self.assertEqual(self.header_findings(self.ledger(headers)), [])

    def test_annotated_mixed_or_gapped_headers_fail_with_reason(self):
        plain = [f'{m}月' for m in (4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3)]
        cases = {
            'annotated': [f'{m}月\n(R{y - REIWA_BASE}.{m})' for y, m in FISCAL_YMS],
            'mixed': plain[:6] + [r_dot(*ym) for ym in FISCAL_YMS[6:]],
            'gapped': [west_slash(*ym) for ym in FISCAL_YMS[:11]] + [west_slash(2026, 4)],
        }
        for label, headers in cases.items():
            with self.subTest(case=label):
                findings = self.header_findings(self.ledger(headers, f'{label}.xlsx'))
                self.assertEqual([(f['level'], f['code']) for f in findings], [('FAIL', 'TPL-WAGE-HEADER')])
                self.assertIn('注記', findings[0]['message'])
                self.assertIn('連続', findings[0]['message'])

    def test_plain_headers_warn_only_for_round7_normal_frame(self):
        path = self.ledger([f'{m}月' for m in range(1, 13)])
        self.assertEqual([(f['level'], f['code']) for f in self.header_findings(path)],
                         [('WARN', 'WAGE-HEADER-NO-YEAR')])
        self.assertEqual(self.header_findings(path, frame='インボイス枠'), [])


if __name__ == '__main__':
    unittest.main()
