# -*- coding: utf-8 -*-
"""第7次の残り制約2点の回帰テスト。合成データ・スタブだけで、Anthropic API は呼ばない。

1. 標準台帳の「非在籍」: 前事業年度の途中に入社・退職した非役員がいても、在籍していない月を
   明示すれば生産性の総労働時間を確定できる。空欄・0円は非在籍と推測しない。
   リーダー・ライター・AI変換（印字された入社日・退職日）・wagebook-convert の検算が同じ印を使う。
2. 加点台帳の「基本給」と「時間換算の対象月額」: 第7次以降の台帳は AH〜AT列に対象月額を持ち、
   公式様式の【基本給】欄と【時間換算給与】欄を別の値から作る。列のない旧形式は従来どおり。
"""
import datetime
import importlib.util
import os
import shutil
import subprocess
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

from hojokin import ai_extractor, pipeline
from hojokin.bonus_wage_ledger_writer import template_has_hourly_basis, write_bonus_wage_ledger
from hojokin.pipeline import ABSENT_MONTH_GUIDE, ApplicationValidationError, _validated_productivity_annual_hours
from hojokin.subsidy_rules import rules_for_round
from hojokin.wage_ledger_writer import write_wage_ledger_to_template
from hojokin.wage_reader import (
    ABSENT_LABELS, ABSENT_MARKER, BWL_COL_BASIS_LATEST, BWL_COL_BASIS_WINDOW_START, BWL_HEADER_ROW,
    BonusWageEmployee, BonusWageLedger, WageEmployee, _adopt_explicit_absence, _ai_data_to_bonus_employees,
    _ai_data_to_wage_employees, _merge_two_employees, fill_bonus_sheet_1, fill_bonus_sheet_2,
    is_absent_marker, judge_bonus_points, read_bonus_wage_ledger, read_wage_ledgers, wareki_label, ym_label,
)

STD_TEMPLATE = ROOT / 'ツール' / '賃金台帳テンプレート.xlsx'
SKILL_TEMPLATE = ROOT / '.claude' / 'skills' / 'wagebook-convert' / 'templates' / '賃金台帳テンプレート.xlsx'
VERIFY = ROOT / '.claude' / 'skills' / 'wagebook-convert' / 'review' / 'verify_wagebook.py'
KATEN_DRIVER = ROOT / '.claude' / 'skills' / 'katen-judge' / 'review' / 'run_katen_judge.py'
LEDGER7 = ROOT / 'ツール' / '第7次以降' / '加点判定用賃金台帳テンプレート.xlsx'
LEDGER6 = ROOT / 'ツール' / '加点判定用賃金台帳テンプレート.xlsx'
FORMS7 = ROOT / '補助金加点' / '第7次以降'

SHEET = '従業員別明細'
HEADER_ROW, FIRST_ROW, COL_NAME, COL_TYPE, COL_HOURS, COL_MONTH, COL_TRANSPORT = 5, 6, 3, 4, 5, 7, 19
FY2025 = [(2025, m) for m in range(1, 13)]
PERIOD = (datetime.date(2025, 1, 1), datetime.date(2025, 12, 31))
FULL_TIME_HOURS, NEW_HIRE_HOURS, PART_HOURS = 160, 150, 80
BASE, ALLOWANCE, BONUS_HOURS = 120_000, 10_000, 100
APPLICATION_YM, LATEST_YM = (2026, 11), (2026, 10)


def std_ledger(path, rows, headers=None):
    """標準テンプレートに合成の従業員を入れる。rows = [(氏名, 雇用形態, E列, 12か月の値)]。"""
    shutil.copyfile(STD_TEMPLATE, path)
    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET]
    for i, text in enumerate(headers or [wareki_label(*ym) for ym in FY2025]):
        ws.cell(HEADER_ROW, COL_MONTH + i).value = text
    for r in range(FIRST_ROW, FIRST_ROW + 10):
        for c in range(2, 21):
            ws.cell(r, c).value = None
    for offset, (name, emp_type, hours, months) in enumerate(rows):
        row = FIRST_ROW + offset
        ws.cell(row, 2).value = offset + 1
        ws.cell(row, COL_NAME).value = name
        ws.cell(row, COL_TYPE).value = emp_type
        ws.cell(row, COL_HOURS).value = hours
        for i, value in enumerate(months):
            ws.cell(row, COL_MONTH + i).value = value
    wb.save(path)
    wb.close()
    return path


def mid_year_rows():
    return [
        ('合成 正社員', '正社員', FULL_TIME_HOURS, [300_000] * 12),
        ('合成 四月入社', '正社員', NEW_HIRE_HOURS, [ABSENT_MARKER] * 3 + [250_000] * 9),
        ('合成 八月退職', 'パート・アルバイト', PART_HOURS, [100_000] * 8 + ['在籍なし'] * 4),
    ]


def expected_hours():
    return FULL_TIME_HOURS * 12 + NEW_HIRE_HOURS * 9 + PART_HOURS * 8


def load_verify_module():
    spec = importlib.util.spec_from_file_location('verify_wagebook_for_test', VERIFY)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class StandardLedgerPresenceTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_marker_spellings(self):
        for text in ('非在籍', ' 非在籍 ', '非在籍　', '在籍なし', '対象外（非在籍）', '対象外(非在籍)'):
            self.assertTrue(is_absent_marker(text), text)
        for value in (None, '', 0, '退職', '非在籍者', 100_000):
            self.assertFalse(is_absent_marker(value), value)

    def test_reader_keeps_absent_months_as_explicit_false(self):
        employees = read_wage_ledgers([std_ledger(self.path / '賃金台帳.xlsx', mid_year_rows())])
        by_name = {e.name: e for e in employees}
        hired = by_name['合成 四月入社']
        self.assertEqual(hired.monthly_presence, {(2025, 1): False, (2025, 2): False, (2025, 3): False})
        self.assertEqual(hired.monthly_wages[:4], [None, None, None, 250_000])
        self.assertFalse(hired.is_full_year)
        left = by_name['合成 八月退職']
        self.assertEqual(sorted(left.monthly_presence), [(2025, m) for m in range(9, 13)])
        self.assertEqual(by_name['合成 正社員'].monthly_presence, {})

    def test_productivity_hours_exclude_only_marked_months(self):
        employees = read_wage_ledgers([std_ledger(self.path / '賃金台帳.xlsx', mid_year_rows())])
        self.assertEqual(_validated_productivity_annual_hours(employees, PERIOD), expected_hours())

    def test_blank_or_zero_month_is_not_treated_as_absent(self):
        for filler in (None, 0):
            rows = mid_year_rows()
            rows[1] = ('合成 四月入社', '正社員', NEW_HIRE_HOURS, [filler] * 3 + [250_000] * 9)
            employees = read_wage_ledgers([std_ledger(self.path / f'台帳_{filler}.xlsx', rows)])
            self.assertEqual(next(e for e in employees if e.name == '合成 四月入社').monthly_presence, {})
            with self.assertRaises(ApplicationValidationError) as raised:
                _validated_productivity_annual_hours(employees, PERIOD)
            self.assertIn(ABSENT_MONTH_GUIDE, str(raised.exception))

    def test_absent_month_with_wage_in_another_ledger_stops(self):
        first = std_ledger(self.path / '賃金台帳_1.xlsx', mid_year_rows())
        second = std_ledger(self.path / '賃金台帳_2.xlsx',
                            [('合成 四月入社', '正社員', NEW_HIRE_HOURS, [250_000] + [None] * 11)])
        employees = read_wage_ledgers([first, second])
        with self.assertRaises(ApplicationValidationError) as raised:
            _validated_productivity_annual_hours(employees, PERIOD)
        self.assertIn('非在籍の明示と給与・労働時間が矛盾', str(raised.exception))

    def test_yearless_headers_keep_month_index_keys(self):
        rows = [('合成 四月入社', '正社員', NEW_HIRE_HOURS, [ABSENT_MARKER] * 3 + [250_000] * 9)]
        path = std_ledger(self.path / '賃金台帳.xlsx', rows, headers=[f'{m}月' for m in range(1, 13)])
        employee = read_wage_ledgers([path])[0]
        self.assertEqual(employee.monthly_presence, {0: False, 1: False, 2: False})

    def test_transport_deduction_uses_only_present_months(self):
        path = std_ledger(self.path / '賃金台帳.xlsx',
                          [('合成 四月入社', '正社員', NEW_HIRE_HOURS, [ABSENT_MARKER] * 3 + [250_000] * 9)])
        wb = openpyxl.load_workbook(path)
        wb[SHEET].cell(FIRST_ROW, COL_TRANSPORT).value = 90_000
        wb.save(path)
        wb.close()
        employee = read_wage_ledgers([path])[0]
        self.assertEqual(employee.monthly_wages[3:], [240_000.0] * 9)
        self.assertEqual(employee.monthly_wages[:3], [None] * 3)

    def test_merge_keeps_presence_of_both_records(self):
        a = WageEmployee(1, '合成 甲', '正社員', 160, 0, [None] * 3 + [1.0] * 9, fiscal_window=FY2025,
                         monthly_presence={(2025, 1): False})
        b = WageEmployee(2, '合成 甲', '正社員', 160, 0, [None] * 3 + [1.0] * 9, fiscal_window=FY2025,
                         monthly_presence={(2025, 2): False}, presence_note='根拠')
        merged = _merge_two_employees(a, b)
        self.assertEqual(merged.monthly_presence, {(2025, 1): False, (2025, 2): False})
        self.assertEqual(merged.presence_note, '根拠')

    def test_writer_writes_marker_and_reports_conflicts(self):
        wages = [None, 0.0, 200_000.0] + [250_000.0] * 9
        emp = WageEmployee(1, '合成 四月入社', '正社員', 160, 0, wages, fiscal_window=FY2025,
                           monthly_presence={(2025, 1): False, (2025, 2): False, (2025, 3): False},
                           presence_note='台帳の入社日 2025-03-20から判定')
        out = self.path / 'out.xlsx'
        result = write_wage_ledger_to_template([emp], STD_TEMPLATE, out, company_name='合成')
        wb = openpyxl.load_workbook(out)
        ws = wb[SHEET]
        self.assertEqual([ws.cell(FIRST_ROW, COL_MONTH + i).value for i in range(3)],
                         [ABSENT_MARKER, ABSENT_MARKER, 200_000])
        memo = [str(c.value) for row in wb['変換メモ'].iter_rows() for c in row if c.value]
        wb.close()
        self.assertEqual(result.absent_rows, [('合成 四月入社', ['令和7年1月', '令和7年2月'],
                                               '台帳の入社日 2025-03-20から判定')])
        self.assertEqual(result.presence_conflicts, [('合成 四月入社', ['令和7年3月'])])
        self.assertTrue(any(text.startswith('非在籍：令和7年1月') for text in memo))
        again = read_wage_ledgers([out])
        self.assertEqual(len(again), 1, '変換メモが賃金の表として読まれていないこと')
        self.assertEqual(again[0].monthly_presence, {(2025, 1): False, (2025, 2): False})

    def test_writer_without_presence_keeps_blank_cells(self):
        emp = WageEmployee(1, '合成 中途', '正社員', 160, 0, [None] * 3 + [1.0] * 9, fiscal_window=FY2025)
        out = self.path / 'out.xlsx'
        result = write_wage_ledger_to_template([emp], STD_TEMPLATE, out)
        wb = openpyxl.load_workbook(out)
        self.assertEqual([wb[SHEET].cell(FIRST_ROW, COL_MONTH + i).value for i in range(3)], [None] * 3)
        wb.close()
        self.assertEqual(result.absent_rows, [])

    def test_template_copies_explain_marker_and_match_bytes(self):
        self.assertEqual(STD_TEMPLATE.read_bytes(), SKILL_TEMPLATE.read_bytes())
        wb = openpyxl.load_workbook(STD_TEMPLATE)
        self.assertIn('「非在籍」', wb[SHEET]['B4'].value)
        self.assertIn('「非在籍」', wb['記入ルール']['B8'].value)
        self.assertIn('空欄・0円を非在籍とは扱', wb['記入ルール']['B8'].value)
        wb.close()


def ai_employee(name, monthly, **dates):
    data = {'name': name, 'employment_type': '正社員', 'monthly': monthly, 'bonuses': []}
    data.update(dates)
    return data


def months(values, start=1):
    return [{'ym': f'2025-{start + i:02d}', 'taxable': v, 'base': v, 'hours': None, 'work_days': 20}
            for i, v in enumerate(values) if v is not None]


class AiPresenceTests(unittest.TestCase):
    def convert(self, data):
        return _ai_data_to_wage_employees(data, '2025-01〜2025-12')

    def test_printed_hire_and_leave_dates_mark_outside_months(self):
        hired, left = self.convert([
            ai_employee('合成 入社', months([0, 0, 0] + [250_000] * 9), hire_date='2025-04-15'),
            ai_employee('合成 退職', months([100_000] * 8), leave_date='2025年8月31日'),
        ])
        self.assertEqual(hired.monthly_presence, {(2025, m): False for m in (1, 2, 3)})
        self.assertEqual(hired.monthly_wages[:3], [None] * 3)
        self.assertIn('入社日 2025-04-15', hired.presence_note)
        self.assertEqual(left.monthly_presence, {(2025, m): False for m in range(9, 13)})

    def test_months_with_pay_outside_dates_are_kept_and_reported(self):
        emp = self.convert([ai_employee('合成 退職', months([100_000] * 9), leave_date='2025-08-31')])[0]
        self.assertEqual(sorted(emp.monthly_presence), [(2025, m) for m in range(10, 13)])
        self.assertEqual(emp.monthly_wages[8], 100_000)
        self.assertIn('令和7年9月', emp.presence_note)

    def test_no_dates_or_unreadable_or_conflicting_dates_mark_nothing(self):
        cases = [
            ai_employee('合成 日付なし', months([None, None] + [1] * 10)),
            ai_employee('合成 読めない', months([1] * 12), hire_date='入社：春頃'),
            ai_employee('合成 逆転', months([1] * 12), hire_date='2025-10-01', leave_date='2025-03-31'),
            dict(ai_employee('合成 食い違い', months([1] * 12), hire_date='2025-04-01'),
                 _date_conflicts=['入社日 2025-04-01 と 2025-05-01']),
        ]
        for emp in self.convert(cases):
            self.assertEqual(emp.monthly_presence, {}, emp.name)
        self.assertEqual(self.convert(cases[:1])[0].presence_note, '')

    def test_chunk_merge_carries_and_flags_dates(self):
        merged = ai_extractor._merge_wage_employees_by_month([
            [ai_employee('合成 甲', months([1]), hire_date='2025-04-01')],
            [ai_employee('合成 甲', months([2], start=2)), ai_employee('合成 乙', months([3]), leave_date='2025-06-30')],
            [ai_employee('合成 乙', months([3]), leave_date='2025-07-31')],
        ])
        by_name = {e['name']: e for e in merged}
        self.assertEqual(by_name['合成 甲']['hire_date'], '2025-04-01')
        self.assertEqual(len(by_name['合成 甲']['monthly']), 2)
        self.assertEqual(by_name['合成 乙']['_date_conflicts'], ['退職日 2025-06-30 と 2025-07-31'])

    def test_excel_absence_is_adopted_into_ai_result(self):
        det = WageEmployee(1, '合成 甲', '正社員', 160, 0, [None] * 3 + [1.0] * 9, fiscal_window=FY2025,
                           monthly_presence={(2025, 1): False, (2025, 2): False, (2025, 3): False, 4: False})
        ai = WageEmployee(1, '合成 甲', '正社員', 160, 0, [0.0, None, 5.0] + [1.0] * 9, fiscal_window=FY2025)
        _adopt_explicit_absence(ai, det)
        self.assertEqual(ai.monthly_presence, {(2025, 1): False, (2025, 2): False})
        self.assertEqual(ai.monthly_wages[:3], [None, None, 5.0])
        self.assertIn('令和7年3月', ai.presence_note)

    def test_both_prompt_versions_ask_for_the_same_fields(self):
        full = ai_extractor.PROMPT_WAGE_LEDGER.format(fiscal_period_section='<期間>', tsv_data='<表>')
        static = ai_extractor.PROMPT_WAGE_LEDGER_STATIC
        for text in ('"hourly_basis": 整数 or null', 'hire_date / leave_date', '最低賃金額以上か',
                     '空欄の月・0円の月・支給額の有無から入社日や退職日を推測しない',
                     '「基準内賃金」は手当を含むことがあるので base に使わない'):
            self.assertEqual(static.count(text), 1, text)
            self.assertEqual(full.count(text), 1, text)


class VerifyWagebookTests(unittest.TestCase):
    def test_absent_labels_match_the_tool(self):
        self.assertEqual(load_verify_module().ABSENT_LABELS, ABSENT_LABELS)

    def test_verify_accepts_marker_and_fails_other_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            good = std_ledger(Path(tmp) / 'good.xlsx', mid_year_rows())
            done = subprocess.run([sys.executable, '-B', str(VERIFY), str(good)],
                                  capture_output=True, text=True, encoding='utf-8')
            self.assertNotIn('修正必須 FAIL', done.stdout, done.stdout)
            self.assertIn('[非在籍の月]（2名', done.stdout)
            rows = mid_year_rows()
            rows[2] = ('合成 八月退職', 'パート・アルバイト', PART_HOURS, [100_000] * 8 + ['退職'] * 4)
            bad = std_ledger(Path(tmp) / 'bad.xlsx', rows)
            done = subprocess.run([sys.executable, '-B', str(VERIFY), str(bad)],
                                  capture_output=True, text=True, encoding='utf-8')
            self.assertEqual(done.returncode, 1)
            self.assertIn('月の欄の不正な値 1名', done.stdout)
            rows = mid_year_rows()
            rows[1] = ('合成 再雇用', '正社員', NEW_HIRE_HOURS, [1] * 3 + [ABSENT_MARKER] * 2 + [1] * 7)
            sandwich = std_ledger(Path(tmp) / 'sandwich.xlsx', rows)
            done = subprocess.run([sys.executable, '-B', str(VERIFY), str(sandwich)],
                                  capture_output=True, text=True, encoding='utf-8')
            self.assertIn('在籍月の間の非在籍 1名', done.stdout)


def bonus_employee(no, base, basis, hours=BONUS_HOURS, latest_base=None, latest_basis=None, name=None):
    window = list(rules_for_round(7).bonus1_window)
    base_by = {ym: base for ym in window}
    basis_by = {ym: basis for ym in window}
    base_by[LATEST_YM] = latest_base if latest_base is not None else base
    basis_by[LATEST_YM] = latest_basis if latest_basis is not None else basis
    return BonusWageEmployee(no, name or f'合成従業員{no}', '正社員', hours, base_by,
                             monthly_hourly_basis=basis_by, hourly_basis_separate=True)


def bonus_ledger(employees):
    return BonusWageLedger(prefecture='東京都', application_ym=APPLICATION_YM, call_round=7, employees=employees)


class HourlyBasisTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_round7_template_has_block_and_round6_does_not(self):
        self.assertTrue(template_has_hourly_basis(LEDGER7, 7))
        self.assertFalse(template_has_hourly_basis(LEDGER6, 6))
        wb = openpyxl.load_workbook(LEDGER7)
        ws = wb['加点判定用明細']
        window = rules_for_round(7).bonus1_window
        for i, ym in enumerate(window):
            self.assertEqual(ws.cell(BWL_HEADER_ROW, BWL_COL_BASIS_WINDOW_START + i).value,
                             f'{ym_label(ym)}\n時間換算の対象月額')
        self.assertEqual(ws.cell(BWL_HEADER_ROW, BWL_COL_BASIS_LATEST).value, '交付申請直近月\n時間換算の対象月額')
        self.assertIn('AH5:AT5', {str(m) for m in ws.merged_cells.ranges})
        self.assertIn('AH〜AT列に時間換算の対象月額', ws['B4'].value)
        rules = wb['記入ルール']
        self.assertIn('職務手当などの手当は含めない', rules['B7'].value)
        self.assertIn('最低賃金額以上かどうかを確認する方法', rules['B8'].value)
        wb.close()

    def test_forms_take_base_and_hourly_from_different_values(self):
        worker = bonus_employee(1, BASE, BASE + ALLOWANCE, latest_base=BASE + 10_000,
                                latest_basis=BASE + ALLOWANCE + 6_000)
        result = judge_bonus_points(bonus_ledger([worker, bonus_employee(2, 200_000, 210_000)]))
        self.assertEqual((result.bonus1_status, result.bonus2_status), ('対象外', '対象'))
        detail = result.bonus1_details[0]['employees'][0]
        self.assertEqual((detail['base'], detail['basis'], detail['hourly']), (BASE, BASE + ALLOWANCE, 1300))
        self.assertEqual(result.bonus2_diff, 1360 - 1300)
        form1 = fill_bonus_sheet_1(FORMS7 / '加点措置①用.xlsx', self.path / 'f1.xlsx', result)
        form2 = fill_bonus_sheet_2(FORMS7 / '加点措置②用.xlsx', self.path / 'f2.xlsx', result)
        ws = openpyxl.load_workbook(form1).worksheets[0]
        self.assertEqual((ws['H18'].value, ws['I18'].value), (BASE, 1300))
        ws = openpyxl.load_workbook(form2).worksheets[0]
        self.assertEqual((ws['F17'].value, ws['G17'].value, ws['N17'].value, ws['O17'].value),
                         (BASE, 1300, BASE + 10_000, 1360))

    def test_allowance_changes_the_decision(self):
        # 基本給だけだと1,200円で最賃未満の人が、通常の手当を含めると1,300円で対象外になる。
        with_allowance = judge_bonus_points(bonus_ledger([bonus_employee(1, BASE, BASE + ALLOWANCE)]))
        self.assertFalse(with_allowance.bonus1_details[5]['employees'][0]['is_target'])

    def test_missing_basis_or_base_or_reversed_amounts_stay_unconfirmed(self):
        missing = bonus_employee(1, BASE, BASE)
        missing.monthly_hourly_basis.clear()
        no_base = bonus_employee(2, BASE, BASE + ALLOWANCE)
        no_base.monthly_base.clear()
        reversed_ = bonus_employee(3, BASE + ALLOWANCE, BASE)
        result = judge_bonus_points(bonus_ledger([missing, no_base, reversed_]))
        self.assertEqual(result.bonus1_status, '未確定')
        issues = ' '.join(result.bonus1_details[5]['issues'])
        self.assertIn('合成従業員1: 時間換算の対象月額（基本給＋通常の手当）が未確認', issues)
        self.assertIn('合成従業員2: 基本給（様式の【基本給】欄）が未確認', issues)
        self.assertIn('合成従業員3: 時間換算の対象月額が基本給より少ない', issues)
        with self.assertRaises(ValueError):
            fill_bonus_sheet_2(FORMS7 / '加点措置②用.xlsx', self.path / 'f2.xlsx', result)

    def test_legacy_single_amount_still_drives_both_columns(self):
        legacy = BonusWageEmployee(1, '合成従業員1', '正社員', BONUS_HOURS,
                                   {ym: BASE + ALLOWANCE for ym in rules_for_round(7).bonus1_window})
        self.assertFalse(legacy.hourly_basis_separate)
        detail = judge_bonus_points(bonus_ledger([legacy])).bonus1_details[0]['employees'][0]
        self.assertEqual((detail['base'], detail['hourly']), (BASE + ALLOWANCE, 1300))

    def test_roundtrip_through_round7_template(self):
        worker = bonus_employee(1, BASE, BASE + ALLOWANCE)
        worker.monthly_presence[(2025, 10)] = False
        del worker.monthly_base[(2025, 10)], worker.monthly_hourly_basis[(2025, 10)]
        out = write_bonus_wage_ledger([worker], LEDGER7, self.path / 'ledger.xlsx', prefecture='東京都',
                                      application_ym=APPLICATION_YM, call_round=7)
        wb = openpyxl.load_workbook(out)
        ws = wb['加点判定用明細']
        self.assertEqual((ws['F7'].value, ws['AH7'].value), (ABSENT_MARKER, ABSENT_MARKER))
        self.assertEqual((ws['G7'].value, ws['AI7'].value, ws['AT7'].value), (BASE, BASE + ALLOWANCE, BASE + ALLOWANCE))
        wb.close()
        read = read_bonus_wage_ledger(out, call_round=7)
        self.assertTrue(read.hourly_basis_separate)
        self.assertFalse(any('AH〜AT列）がない形式' in n for n in read.notes))
        emp = read.employees[0]
        self.assertEqual((emp.monthly_base[(2026, 1)], emp.monthly_hourly_basis[(2026, 1)]), (BASE, BASE + ALLOWANCE))
        self.assertIs(emp.monthly_presence[(2025, 10)], False)

    def test_blank_basis_column_is_not_filled_from_base(self):
        legacy = BonusWageEmployee(1, '合成従業員1', '正社員', BONUS_HOURS,
                                   {ym: BASE for ym in rules_for_round(7).bonus1_window})
        out = write_bonus_wage_ledger([legacy], LEDGER7, self.path / 'ledger.xlsx', prefecture='東京都',
                                      application_ym=APPLICATION_YM, call_round=7)
        result = judge_bonus_points(read_bonus_wage_ledger(out, call_round=7))
        self.assertEqual(result.bonus1_status, '未確定')
        self.assertIn('時間換算の対象月額（基本給＋通常の手当）が未確認', ' '.join(result.bonus1_issues))

    def test_marker_conflict_and_partial_headers_are_rejected(self):
        out = write_bonus_wage_ledger([bonus_employee(1, BASE, BASE + ALLOWANCE)], LEDGER7, self.path / 'l.xlsx',
                                      prefecture='東京都', application_ym=APPLICATION_YM, call_round=7)
        wb = openpyxl.load_workbook(out)
        wb['加点判定用明細']['AH7'] = ABSENT_MARKER
        wb.save(out)
        wb.close()
        issues = read_bonus_wage_ledger(out, call_round=7).input_issues
        self.assertTrue(any('非在籍指定と金額が矛盾' in issue for issue in issues), issues)
        wb = openpyxl.load_workbook(out)
        wb['加点判定用明細']['AJ6'] = None
        wb.save(out)
        wb.close()
        with self.assertRaises(ValueError):
            read_bonus_wage_ledger(out, call_round=7)

    def test_round6_template_receives_the_hourly_basis_in_its_single_column(self):
        window = list(rules_for_round(6).bonus1_window)
        worker = BonusWageEmployee(1, '合成従業員1', '正社員', BONUS_HOURS, {ym: BASE for ym in window},
                                   monthly_hourly_basis={ym: BASE + ALLOWANCE for ym in window},
                                   hourly_basis_separate=True)
        out = write_bonus_wage_ledger([worker], LEDGER6, self.path / 'l6.xlsx', prefecture='東京都',
                                      application_ym=(2026, 6), call_round=6)
        read = read_bonus_wage_ledger(out, call_round=6)
        self.assertFalse(read.hourly_basis_separate)
        self.assertEqual(read.employees[0].monthly_base[window[0]], BASE + ALLOWANCE)
        self.assertTrue(any('様式の【基本給】欄を原本の基本給に直して' in n for n in read.notes))

    def test_ai_output_with_and_without_hourly_basis(self):
        target = list(rules_for_round(7).bonus1_window) + [LATEST_YM]
        new = [{'name': '合成 新', 'employment_type': '正社員', 'hire_date': '2025-12-01',
                'monthly': [{'ym': '2026-01', 'base': BASE, 'hourly_basis': BASE + ALLOWANCE, 'hours': 100},
                            {'ym': '2026-02', 'base': BASE, 'hourly_basis': None, 'hours': 100}]}]
        emp = _ai_data_to_bonus_employees(new, target_yms=target)[0]
        self.assertTrue(emp.hourly_basis_separate)
        self.assertEqual(emp.monthly_hourly_basis, {(2026, 1): BASE + ALLOWANCE})
        self.assertEqual(sorted(k for k, v in emp.monthly_presence.items() if v is False),
                         [(2025, 10), (2025, 11)])
        old = [{'name': '合成 旧', 'employment_type': '正社員', 'monthly': [{'ym': '2026-01', 'base': BASE}]}]
        legacy = _ai_data_to_bonus_employees(old)[0]
        self.assertFalse(legacy.hourly_basis_separate)
        self.assertEqual(legacy.monthly_presence, {})

    def test_bonus_ledger_creation_task_reports_coverage(self):
        window = list(rules_for_round(7).bonus1_window)

        class StubExtractor:
            def extract_wage_ledger(self, tsv, hint, pdf_files=None, disable_image_fallback=False):
                return [{'name': '合成 手当あり', 'employment_type': '正社員',
                         'monthly': [{'ym': f'{y}-{m:02d}', 'base': BASE, 'hourly_basis': BASE + ALLOWANCE,
                                      'hours': BONUS_HOURS} for y, m in window]},
                        {'name': '合成 一部不明', 'employment_type': '正社員', 'leave_date': '2026-08-31',
                         'monthly': [{'ym': f'{y}-{m:02d}', 'base': BASE, 'hourly_basis': None,
                                      'hours': BONUS_HOURS} for y, m in window[:-1]]}]

        source = self.path / 'source'
        source.mkdir()
        std_ledger(source / '賃金台帳_合成.xlsx', [('合成 手当あり', '正社員', 160, [1] * 12)])
        out = self.path / '加点判定用賃金台帳.xlsx'
        status = pipeline.run_bonus_wage_ledger_creation(
            source, '合成', LEDGER7, out, extractor=StubExtractor(), prefecture='東京都',
            application_ym=APPLICATION_YM, call_round=7)
        self.assertEqual(status.status, '完了', status.message)
        self.assertIn('通常の手当を時間換算に含めた従業員: 1名', status.message)
        self.assertIn('時間換算の対象月額が読み取れなかった月がある従業員: 1名', status.message)
        self.assertIn('「非在籍」と書いた月がある従業員: 1名', status.message)
        ws = openpyxl.load_workbook(out)['加点判定用明細']
        self.assertEqual((ws['F7'].value, ws['AH7'].value), (BASE, BASE + ALLOWANCE))
        self.assertEqual((ws['Q8'].value, ws['AS8'].value), (ABSENT_MARKER, ABSENT_MARKER))

    def fill_ledger(self, basis):
        path = self.path / 'filled.xlsx'
        shutil.copyfile(LEDGER7, path)
        wb = openpyxl.load_workbook(path)
        ws = wb['加点判定用明細']
        ws['C2'], ws['C3'] = '東京都', datetime.datetime(*APPLICATION_YM, 1)
        ws['B7'], ws['C7'], ws['D7'], ws['E7'] = 1, '合成従業員', '正社員', BONUS_HOURS
        for col in range(6, 19):
            ws.cell(7, col).value = 120_000 if col < 18 else 125_000
            ws.cell(7, BWL_COL_BASIS_WINDOW_START + col - 6).value = (
                None if basis is None else (122_600 if col < 18 else 128_100))
        wb.save(path)
        wb.close()
        return path

    def test_driver_uses_basis_columns(self):
        done = subprocess.run([sys.executable, '-B', str(KATEN_DRIVER), str(self.fill_ledger(True)),
                               '--call-round', '7', '--fill', '通常', '--outdir', str(self.path / 'forms')],
                              capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(done.returncode, 0, done.stdout + done.stderr)
        self.assertIn('差: +55円', done.stdout)
        form = next((self.path / 'forms').glob('*加点措置②*.xlsx'))
        ws = openpyxl.load_workbook(form).worksheets[0]
        self.assertEqual((ws['F17'].value, ws['G17'].value), (120_000, 1226))
        blank = subprocess.run([sys.executable, '-B', str(KATEN_DRIVER), str(self.fill_ledger(None)),
                                '--call-round', '7'], capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(blank.returncode, 3, blank.stdout + blank.stderr)


if __name__ == '__main__':
    unittest.main()
