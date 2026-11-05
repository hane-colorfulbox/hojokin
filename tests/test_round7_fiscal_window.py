# -*- coding: utf-8 -*-
"""第7次: 決定論リーダーの対象年月と、台帳だけで分かる停止をAPI抽出前に出す順序。

- 年付きの月見出し（令和7年4月 / R7.4 / 2025年4月 / 2025/04 / 日付）から fiscal_window を確定する
- 年のない見出し・2年度分の混在・連続しない年月は確定せず、理由を残す
- 第7次通常枠は、台帳だけで判定できる停止を登記・決算書等の抽出と事業内容生成より前に出す
- 第1〜6次は従来どおり（抽出の後に台帳を読む・算定値は変わらない）
- 加点判定用台帳 AG3=7 は、制度版が同じ第8次以降でも読める

合成データ・StubExtractor のみ。Anthropic API は呼ばない。
"""
import datetime
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['CLAUDE_API_KEY'] = ''

import openpyxl

from hojokin import pipeline
from hojokin.ai_extractor import create_extractor
from hojokin.application_context import ApplicationContext
from hojokin.bonus_wage_ledger_writer import write_bonus_wage_ledger
from hojokin.models import FinancialData
from hojokin.subsidy_rules import MIN_WAGE_R6 as RULES_MIN_WAGE_R6, rules_for_round
from hojokin.template_versions import select_application_template
from hojokin.wage_reader import (
    BONUS1_WINDOW, BONUS2_BASE_YM, BONUS_THRESHOLD_YEN, MIN_WAGE_R6,
    BonusWageEmployee, WageEmployee, _check_stamped_round, _merge_fiscal_windows,
    _parse_stamped_round, read_bonus_wage_ledger, read_wage_ledgers, wareki_label, ym_label,
)

LEDGER_TEMPLATE = ROOT / 'ツール' / '賃金台帳テンプレート.xlsx'
SHEET = '従業員別明細'
HEADER_ROW = 5
FIRST_ROW = 6
COL_NO, COL_NAME, COL_TYPE, COL_HOURS, COL_MONTH_START = 2, 3, 4, 5, 7
MONTHLY_HOURS = 160
FY_APRIL = (2025, 4)          # 第7次・3月決算・申請日2026/11 の前事業年度の期首
EMPLOYEE_COUNT = 4
API_METHODS = ('extract_registry', 'extract_pl', 'extract_tax', 'extract_estimate',
               'generate_ai_judgment')


def window_from(start, months=12):
    out, (y, m) = [], start
    for _ in range(months):
        out.append((y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def header_text(ym, style):
    y, m = ym
    return {
        'wareki': wareki_label(y, m),
        'r_dot': f'R{y - 2018}.{m}',
        'seireki': f'{y}年{m}月',
        'slash': f'{y}/{m:02d}',
        'date': datetime.datetime(y, m, 1),
        'plain': f'{m}月',
    }[style]


def wage_for(row, ym):
    """列ごとに違う額にして、暦月への割付けを確かめる。"""
    return 200_000 + row * 1_000 + ym[1]


def make_ledger(path, start=FY_APRIL, style='wareki', yms=None, names=None):
    """標準テンプレート（G〜R列が月）に合成値を入れる。列は事業年度順。"""
    yms = yms or window_from(start)
    names = names or [f'合成 従業員{i}' for i in range(1, EMPLOYEE_COUNT + 1)]
    shutil.copy(LEDGER_TEMPLATE, path)
    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET]
    for i, ym in enumerate(yms):
        ws.cell(HEADER_ROW, COL_MONTH_START + i).value = header_text(ym, style)
    for r_off, name in enumerate(names):
        row = FIRST_ROW + r_off
        ws.cell(row, COL_NO).value = r_off + 1
        ws.cell(row, COL_NAME).value = name
        ws.cell(row, COL_TYPE).value = '正社員'
        ws.cell(row, COL_HOURS).value = MONTHLY_HOURS
        for i, ym in enumerate(yms):
            ws.cell(row, COL_MONTH_START + i).value = wage_for(row, ym)
    wb.save(path)
    wb.close()
    return path


def make_rowwise_ledger(path, month_texts_by_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '月別'
    ws.append(['氏名', '雇用形態', '対象年月', '課税支給合計'])
    for name, months in month_texts_by_name.items():
        for text in months:
            ws.append([name, '正社員', text, 250_000])
    wb.save(path)
    wb.close()
    return path


def make_individual_ledger(path, header_cells):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws['A1'] = '賃金台帳'
    ws['A2'] = '001\n合成花子 (女)'
    for i, text in enumerate(header_cells):
        ws.cell(3, 2 + i).value = text
    for label_row, label in ((4, '基本給'), (5, '所定労働時間'), (6, '課税支給合計')):
        ws.cell(label_row, 1).value = label
        for i in range(len(header_cells)):
            ws.cell(label_row, 2 + i).value = {4: 200_000, 5: MONTHLY_HOURS, 6: 230_000}[label_row]
    wb.save(path)
    wb.close()
    return path


class DeterministicWindowTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def read(self, *files):
        return read_wage_ledgers(list(files))

    def test_year_headers_in_each_format_set_the_window(self):
        expected = window_from(FY_APRIL)
        for style in ('wareki', 'r_dot', 'seireki', 'slash', 'date'):
            with self.subTest(style=style):
                emps = self.read(make_ledger(self.path / f'{style}.xlsx', style=style))
                self.assertEqual(len(emps), EMPLOYEE_COUNT)
                for emp in emps:
                    self.assertEqual(emp.fiscal_window, expected)
                    self.assertEqual(emp.fiscal_window_issue, '')
                # 暦月Index固定の配列へ正しく割り付く（4月=Index3、3月=Index2）
                first = emps[0]
                self.assertEqual(first.monthly_wages[3], wage_for(FIRST_ROW, (2025, 4)))
                self.assertEqual(first.monthly_wages[2], wage_for(FIRST_ROW, (2026, 3)))

    def test_plain_month_headers_stay_unset_with_reason_and_same_amounts(self):
        plain = self.read(make_ledger(self.path / 'plain.xlsx', style='plain'))
        stamped = self.read(make_ledger(self.path / 'stamped.xlsx', style='wareki'))
        for emp in plain:
            self.assertIsNone(emp.fiscal_window)
            self.assertIn('月見出しに年がありません', emp.fiscal_window_issue)
        # 年の有無で金額・人数は変わらない（R215/R216 の従来動作を保つ）
        self.assertEqual([e.monthly_wages for e in plain], [e.monthly_wages for e in stamped])
        self.assertEqual([e.monthly_avg_hours for e in plain], [e.monthly_avg_hours for e in stamped])

    def test_two_fiscal_years_in_the_folder_are_not_a_window(self):
        older = make_ledger(self.path / 'r6.xlsx', start=(2024, 4))
        newer = make_ledger(self.path / 'r7.xlsx', start=(2025, 4))
        for emp in self.read(older, newer):
            self.assertIsNone(emp.fiscal_window)
            self.assertIn('24か月分', emp.fiscal_window_issue)
            self.assertIn('令和6年4月〜令和8年3月', emp.fiscal_window_issue)

    def test_non_consecutive_or_duplicate_headers_are_not_a_window(self):
        gap = window_from(FY_APRIL, 11) + [(2027, 3)]
        emps = self.read(make_ledger(self.path / 'gap.xlsx', yms=gap))
        self.assertIsNone(emps[0].fiscal_window)
        self.assertIn('連続した12か月', emps[0].fiscal_window_issue)
        dup = window_from(FY_APRIL, 11) + [(2025, 4)]
        emps = self.read(make_ledger(self.path / 'dup.xlsx', yms=dup))
        self.assertIsNone(emps[0].fiscal_window)
        self.assertIn('2つ以上', emps[0].fiscal_window_issue)

    def test_rowwise_ledger_uses_the_whole_table_for_midyear_joiners(self):
        full = [f'{y}年{m}月' for (y, m) in window_from(FY_APRIL)]
        path = make_rowwise_ledger(self.path / 'rows.xlsx', {'合成A': full, '合成B': full[6:]})
        emps = {e.name: e for e in self.read(path)}
        self.assertEqual(emps['合成A'].fiscal_window, window_from(FY_APRIL))
        self.assertEqual(emps['合成B'].fiscal_window, window_from(FY_APRIL))
        plain = make_rowwise_ledger(self.path / 'rows_plain.xlsx',
                                    {'合成A': [f'{m}月' for (_y, m) in window_from(FY_APRIL)]})
        emp = self.read(plain)[0]
        self.assertIsNone(emp.fiscal_window)
        self.assertIn('年のない行', emp.fiscal_window_issue)

    def test_individual_ledger_reads_the_year_in_the_month_header(self):
        stamped = [f'令和 {y - 2018}年\n{m}月度給与' for (y, m) in window_from(FY_APRIL)]
        emp = self.read(make_individual_ledger(self.path / 'ind.xlsx', stamped))[0]
        self.assertEqual(emp.name, '合成花子')
        self.assertEqual(emp.fiscal_window, window_from(FY_APRIL))
        plain = [f'{m}月度給与' for (_y, m) in window_from(FY_APRIL)]
        emp = self.read(make_individual_ledger(self.path / 'ind_plain.xlsx', plain))[0]
        self.assertIsNone(emp.fiscal_window)
        self.assertIn('月見出しに年がありません', emp.fiscal_window_issue)

    def test_name_merge_with_different_windows_is_not_resolved(self):
        a = WageEmployee(1, '合成 花子', '正社員', 160, 0, [1.0] * 12,
                         fiscal_window=window_from((2024, 4)))
        b = WageEmployee(2, '合成花子', '正社員', 160, 0, [1.0] * 12,
                         fiscal_window=window_from((2025, 4)))
        window, issue = _merge_fiscal_windows(a, b)
        self.assertIsNone(window)
        self.assertIn('令和6年4月〜令和7年3月', issue)
        self.assertIn('令和7年4月〜令和8年3月', issue)
        # AI経路のように同じ窓なら、そのまま残す
        b.fiscal_window = window_from((2024, 4))
        self.assertEqual(_merge_fiscal_windows(a, b), (window_from((2024, 4)), ''))


def round7_context(frame='通常'):
    return ApplicationContext.from_dict({
        'call_round': 7, 'application_date': '2026-11-20',
        'fiscal_year_start': '2026-04-01', 'fiscal_year_end': '2027-03-31',
        'frame': frame, 'subsidy_amount': 1_000_000, 'past_grants': [],
        'optional_wage_bonus': False, 'wage_exempt': False, 'payroll_scope_confirmed': True,
        'workplace_prefecture': '東京都', 'workplace_confirmed': True, 'regional_min_wage_yen': 1280,
        'forecast_years': [dict(total_salary=10_000_000 + i * 400_000, employee_count_fte=4,
                                source='合成の確認済み見込み', confirmed=True) for i in range(4)],
        'productivity_forecast_years': (
            [dict(operating_profit=3_000_000, personnel_cost=10_000_000, depreciation=500_000,
                  total_annual_hours=8000, source='合成の確認済み予算', confirmed=True)] * 3
            if frame == '通常' else []),
    })


def round6_context():
    return ApplicationContext(call_round=6, application_date=datetime.date(2026, 10, 1),
                              fiscal_year_start=datetime.date(2026, 1, 1),
                              fiscal_year_end=datetime.date(2026, 12, 31),
                              frame='通常', subsidy_amount=500_000, past_grants=[],
                              optional_wage_bonus=False, wage_exempt=False,
                              declaration_confirmed=False)


def make_pl_pdf(path):
    import fitz
    doc = fitz.open()
    doc.new_page().insert_text((72, 72), 'synthetic')
    doc.save(str(path))
    doc.close()


class ApplicationOrderTests(unittest.TestCase):
    """申請書作成で、台帳だけで分かる停止が API 抽出より前に出ること。"""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name)
        self.calls: list[str] = []

    def tearDown(self):
        self.tmp.cleanup()

    def spy_extractor(self, financial):
        stub = create_extractor('')
        for name in API_METHODS:
            original = getattr(stub, name)

            def wrapper(*args, _name=name, _original=original, **kwargs):
                self.calls.append(_name)
                return financial if _name == 'extract_pl' else _original(*args, **kwargs)
            setattr(stub, name, wrapper)
        return stub

    def run_case(self, context, template_type='通常枠_2026', ledger_style='wareki',
                 ledger_start=FY_APRIL, with_ledger=True, pl_period=('2025-04-01', '2026-03-31')):
        work = self.path / f'case{len(list(self.path.iterdir()))}'
        work.mkdir()
        if with_ledger:
            make_ledger(work / '賃金台帳_合成.xlsx', start=ledger_start, style=ledger_style)
        make_pl_pdf(work / '決算書_合成.pdf')
        financial = FinancialData(fiscal_year_start=pl_period[0], fiscal_year_end=pl_period[1],
                                  revenue=50_000_000, operating_profit=1_000_000,
                                  salary=4_000_000, depreciation=300_000)
        template = select_application_template(ROOT, template_type, context.call_round)
        output = work / 'out.xlsx'
        original_read = pipeline._read_application_ledger

        def recording_read(*args, **kwargs):
            self.calls.append('ledger')
            return original_read(*args, **kwargs)
        with patch('hojokin.pipeline._read_application_ledger', side_effect=recording_read):
            result = pipeline.run_application_transfer(
                work, template, template_type, output, self.spy_extractor(financial),
                application_context=context)
        return result, output

    def assert_stopped_before_api(self, result, *phrases):
        self.assertEqual(result.status, 'エラー', result.message)
        for phrase in phrases:
            self.assertIn(phrase, result.message)
        self.assertEqual([c for c in self.calls if c != 'ledger'], [], self.calls)

    def test_round7_normal_frame_with_year_headers_is_generated(self):
        result, output = self.run_case(round7_context())
        self.assertNotEqual(result.status, 'エラー', result.message)
        self.assertTrue(output.exists())
        self.assertLess(self.calls.index('ledger'), self.calls.index('extract_pl'))

    def test_round7_plain_month_headers_stop_before_api_with_reason(self):
        result, output = self.run_case(round7_context(), ledger_style='plain')
        self.assert_stopped_before_api(result, '月見出しに年がありません', '令和7年4月〜令和8年3月')
        self.assertFalse(output.exists())

    def test_round7_older_ledger_stops_before_api_with_both_periods(self):
        result, _ = self.run_case(round7_context(), ledger_start=(2024, 4))
        self.assert_stopped_before_api(result, '対象年月が前事業年度と一致しません',
                                       '令和6年4月〜令和7年3月', '令和7年4月〜令和8年3月')

    def test_round7_missing_ledger_stops_before_api_for_normal_frame(self):
        result, _ = self.run_case(round7_context(), with_ledger=False)
        self.assert_stopped_before_api(result, '賃金台帳が見つかりません')

    def test_round7_stop_does_not_create_extractor(self):
        work = self.path / 'no_extractor'
        work.mkdir()
        make_ledger(work / '賃金台帳_合成.xlsx', style='plain')
        template = select_application_template(ROOT, '通常枠_2026', 7)
        with patch('hojokin.pipeline.create_extractor',
                   side_effect=AssertionError('must not create extractor')):
            result = pipeline.run_application_transfer(
                work, template, '通常枠_2026', work / 'out.xlsx',
                application_context=round7_context())
        self.assertEqual(result.status, 'エラー')
        self.assertIn('月見出しに年がありません', result.message)
        self.assertNotIn('must not create extractor', result.message)

    def test_round7_invoice_frame_keeps_previous_behavior_without_year_headers(self):
        result, output = self.run_case(round7_context('インボイス'), template_type='インボイス枠_2026',
                                       ledger_style='plain')
        self.assertNotEqual(result.status, 'エラー', result.message)
        self.assertTrue(output.exists())

    def test_round6_still_reads_ledger_after_extraction(self):
        result, output = self.run_case(round6_context(), ledger_style='plain', ledger_start=(2025, 1),
                                       pl_period=('2025-01-01', '2025-12-31'))
        self.assertNotEqual(result.status, 'エラー', result.message)
        self.assertTrue(output.exists())
        self.assertGreater(self.calls.index('ledger'), self.calls.index('extract_pl'))


class BonusLedgerRoundTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def make_round7_ledger(self):
        template = self.path / 'template.xlsx'
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '加点判定用明細'
        ws['B6'], ws['C6'] = 'No', '氏名'
        for i, ym in enumerate(rules_for_round(7).bonus1_window):
            ws.cell(6, 6 + i, ym_label(ym) + '\n基本給')
            ws.cell(6, 19 + i, ym_label(ym) + '\n労働時間')
        ws['R6'] = '交付申請直近月\n基本給'
        wb.save(template)
        wb.close()
        base = {ym: 123_000 for ym in rules_for_round(7).bonus1_window}
        out = self.path / 'ledger.xlsx'
        write_bonus_wage_ledger([BonusWageEmployee(1, '合成従業員', '正社員', 100, base)],
                                template, out, prefecture='東京都',
                                application_ym=(2026, 12), call_round=7)
        return out

    def test_round7_stamp_is_read_in_round8_with_a_note(self):
        ledger_path = self.make_round7_ledger()
        same = read_bonus_wage_ledger(ledger_path, call_round=7)
        self.assertFalse(any('AG3' in note for note in same.notes))
        later = read_bonus_wage_ledger(ledger_path, call_round=8)
        self.assertEqual(later.call_round, 8)
        self.assertTrue(any('第7次' in n and '第8次' in n for n in later.notes), later.notes)

    def test_round_stamp_with_another_rule_or_unreadable_value_stops(self):
        rules7 = rules_for_round(7)
        with self.assertRaisesRegex(ValueError, '一致しません'):
            _check_stamped_round(6, 7, rules7)
        with self.assertRaisesRegex(ValueError, 'AG3の公募回を読み取れません'):
            _check_stamped_round('未定', 7, rules7)
        self.assertEqual(_check_stamped_round(None, 7, rules7), '')
        self.assertEqual(_check_stamped_round(7.0, 7, rules7), '')
        self.assertEqual(_parse_stamped_round('第７次'), 7)
        self.assertIsNone(_parse_stamped_round(True))
        self.assertIsNone(_parse_stamped_round(0))


class LegacyConstantTests(unittest.TestCase):
    def test_legacy_constants_are_round_1_6_values_from_subsidy_rules(self):
        legacy = rules_for_round(6)
        self.assertEqual(BONUS1_WINDOW, list(legacy.bonus1_window))
        self.assertEqual(BONUS1_WINDOW[0], (2024, 10))
        self.assertEqual(BONUS2_BASE_YM, (2025, 7))
        self.assertEqual(BONUS_THRESHOLD_YEN, 63)
        self.assertEqual(MIN_WAGE_R6, dict(RULES_MIN_WAGE_R6))
        self.assertIsNot(MIN_WAGE_R6, RULES_MIN_WAGE_R6)
        # 第7次以降の現役値は rules_for_round 側にある
        self.assertEqual(rules_for_round(7).bonus2_threshold_yen, 55)
        self.assertEqual(rules_for_round(7).bonus2_base_ym, (2026, 6))


if __name__ == '__main__':
    unittest.main()
