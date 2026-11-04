# -*- coding: utf-8 -*-
"""公募回別加点・欠損停止・公式様式保持の合成データテスト。API不使用。"""
import datetime
import sys
import subprocess
import tempfile
import unittest
import warnings
import zipfile
from xml.etree import ElementTree
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding='utf-8')

import openpyxl
from hojokin.subsidy_rules import rules_for_round, legal_minimum_bounds, EFFECTIVE_R7, EFFECTIVE_R8
from hojokin.wage_reader import (
    BonusWageEmployee, BonusWageLedger, judge_bonus_points, read_bonus_wage_ledger,
    fill_bonus_sheet_1, fill_bonus_sheet_2, ym_label,
    _ai_data_to_bonus_employees,
)
from hojokin.bonus_wage_ledger_writer import write_bonus_wage_ledger
from hojokin.xlsx_surgical import patch_xlsx


def employee(no=1, hourly=1230, latest_hourly=1285, call_round=7):
    rules = rules_for_round(call_round)
    data = {ym: hourly * 100 for ym in rules.bonus1_window}
    data[(2026, 11)] = latest_hourly * 100
    return BonusWageEmployee(no, f'合成従業員{no}', '正社員', 100, data)


def ledger(employees=None, call_round=7):
    return BonusWageLedger(prefecture='東京都', application_ym=(2026, 12),
                           employees=employees or [employee(call_round=call_round)], call_round=call_round)


def blank_template(path, call_round):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '加点判定用明細'
    ws['B6'] = 'No'
    ws['C6'] = '氏名'
    for i, ym in enumerate(rules_for_round(call_round).bonus1_window):
        ws.cell(6, 6 + i, ym_label(ym) + '\n基本給')
        ws.cell(6, 19 + i, ym_label(ym) + '\n労働時間')
    ws['R6'] = '交付申請直近月\n基本給'
    wb.save(path)
    wb.close()


class BonusRound7Tests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_rules_and_effective_dates(self):
        self.assertEqual(rules_for_round(6).bonus1_window[0], (2024, 10))
        new = rules_for_round(7)
        self.assertEqual(new.bonus1_window[0], (2025, 10))
        self.assertEqual(new.bonus1_window[-1], (2026, 9))
        self.assertEqual(new.bonus2_base_ym, (2026, 6))
        self.assertEqual(len(new.min_wage_after), 47)
        self.assertEqual(new.min_wage_after['東京都'], 1280)
        self.assertEqual(len(EFFECTIVE_R7), 47)
        self.assertEqual(len(EFFECTIVE_R8), 47)
        self.assertEqual(EFFECTIVE_R7['秋田県'], '2026-03-31')
        self.assertEqual(EFFECTIVE_R7['広島県'], '2025-11-01')
        self.assertEqual(EFFECTIVE_R7['福島県'], '2026-01-01')
        self.assertEqual(legal_minimum_bounds('東京都', (2026, 9)), (1226, 1226, False))
        self.assertEqual(legal_minimum_bounds('東京都', (2026, 10)), (1280, 1280, False))
        self.assertEqual(legal_minimum_bounds('京都府', (2026, 11)), (1122, 1180, True))
        self.assertEqual(legal_minimum_bounds('京都府', (2026, 12)), (1180, 1180, False))
        self.assertEqual(legal_minimum_bounds('福島県', (2025, 12)), (955, 955, False))

    def test_difference_54_55_56(self):
        for diff in (54, 55, 56):
            result = judge_bonus_points(ledger([employee(hourly=1226, latest_hourly=1226 + diff)]))
            self.assertEqual(result.bonus2_diff, diff)
            self.assertEqual(result.bonus2_status, '対象' if diff >= 55 else '対象外')

    def test_30percent_and_three_nonconsecutive_months(self):
        workers = [employee(no=i, hourly=1400, latest_hourly=1500) for i in range(1, 11)]
        yms = rules_for_round(7).bonus1_window
        for worker in workers[:3]:
            for ym in (yms[0], yms[3], yms[8]):
                worker.monthly_base[ym] = 123000
        result = judge_bonus_points(ledger(workers))
        self.assertEqual(result.bonus1_status, '対象')
        self.assertEqual(len(result.bonus1_months_met), 3)
        self.assertEqual(result.bonus1_details[0]['ratio'], 0.30)
        workers[0].monthly_base[yms[8]] = 140000
        self.assertEqual(judge_bonus_points(ledger(workers)).bonus1_status, '対象外')

    def test_upper_threshold_is_exclusive(self):
        result = judge_bonus_points(ledger([employee(hourly=1280)]))
        self.assertEqual(result.bonus1_status, '対象外')

    def test_missing_hours_do_not_shrink_denominator(self):
        workers = [employee(no=i) for i in range(1, 11)]
        for worker in workers[1:]:
            worker.scheduled_hours = None
        result = judge_bonus_points(ledger(workers))
        self.assertEqual(result.bonus1_details[0]['total'], 10)
        self.assertEqual(result.bonus1_status, '未確定')
        self.assertEqual(result.bonus2_status, '未確定')
        self.assertFalse(result.bonus1_eligible)

    def test_ai_does_not_fill_unknown_months_from_one_hours_sample(self):
        yms = list(rules_for_round(7).bonus1_window) + [(2026, 11)]
        extracted = [{
            'name': '合成従業員', 'employment_type': '正社員',
            'scheduled_hours': 100,  # AI応答の代表値を確認済み情報として採用しない。
            'monthly': [{'ym': f'{y}-{m:02d}', 'base': 123000,
                         'hours': 100 if index == 0 else None}
                        for index, (y, m) in enumerate(yms)],
        }]
        converted = _ai_data_to_bonus_employees(extracted)
        self.assertIsNone(converted[0].scheduled_hours)
        self.assertEqual(converted[0].hours_for(yms[0]), 100)
        self.assertIsNone(converted[0].hours_for(yms[1]))
        self.assertIsNone(converted[0].hours_for((2026, 11)))
        result = judge_bonus_points(ledger(converted))
        self.assertEqual(result.bonus1_status, '未確定')
        self.assertEqual(result.bonus2_status, '未確定')
        self.assertEqual(sum(d['complete'] for d in result.bonus1_details), 1)
        template, out = self.path / 'ai_template.xlsx', self.path / 'ai.xlsx'
        blank_template(template, 7)
        # 古いテンプレ由来の代表時間があっても、出力で清掃する。
        wb = openpyxl.load_workbook(template)
        wb.active['E7'] = 160
        wb.save(template)
        wb.close()
        write_bonus_wage_ledger(converted, template, out, prefecture='東京都',
                               application_ym=(2026, 12), call_round=7)
        reread = read_bonus_wage_ledger(out, call_round=7)
        self.assertIsNone(reread.employees[0].scheduled_hours)
        self.assertIsNone(reread.employees[0].hours_for((2026, 11)))

    def test_missing_salary_and_absence_are_different(self):
        absent = employee(no=2)
        yms = rules_for_round(7).bonus1_window
        absent.monthly_base = {}
        result = judge_bonus_points(ledger([employee(), absent]))
        self.assertEqual(result.bonus1_status, '未確定')
        absent.monthly_presence = {ym: False for ym in yms}
        absent.monthly_presence[(2026, 11)] = False
        result = judge_bonus_points(ledger([employee(), absent]))
        self.assertEqual(result.bonus1_status, '対象')
        self.assertEqual(result.bonus2_status, '対象')

    def test_three_complete_months_are_sufficient(self):
        worker = employee()
        allowed = rules_for_round(7).bonus1_window[:3]
        worker.monthly_base = {ym: 123000 for ym in allowed}
        result = judge_bonus_points(ledger([worker]))
        self.assertEqual(result.bonus1_status, '対象')
        self.assertEqual(result.bonus2_status, '未確定')

    def test_mixed_workplaces_are_not_auto_confirmed(self):
        book = ledger()
        book.workplace_prefectures = ['東京都', '大阪府']
        self.assertEqual(judge_bonus_points(book).bonus1_status, '未確定')

    def test_unknown_employment_and_unregistered_future_wage_stop(self):
        worker = employee()
        worker.employment_type = ''
        self.assertEqual(judge_bonus_points(ledger([worker])).bonus1_status, '未確定')
        book = ledger()
        book.application_ym = (2027, 11)
        book.employees[0].monthly_base[(2027, 10)] = 150000
        self.assertEqual(judge_bonus_points(book).bonus2_status, '未確定')

    def test_roundtrip_metadata_and_mismatch_stop(self):
        template = self.path / 'template.xlsx'
        out = self.path / 'ledger.xlsx'
        blank_template(template, 7)
        write_bonus_wage_ledger([employee()], template, out, prefecture='東京都', application_ym=(2026, 12), call_round=7)
        reread = read_bonus_wage_ledger(out, call_round=7)
        self.assertEqual(reread.rule_id, rules_for_round(7).rule_id)
        self.assertEqual(judge_bonus_points(reread).bonus2_status, '対象')
        with self.assertRaises(ValueError):
            read_bonus_wage_ledger(out, call_round=6)
        with self.assertRaises(ValueError):
            read_bonus_wage_ledger(out, application_ym_fallback=(2026, 11), call_round=7)
        old = self.path / 'old.xlsx'
        blank_template(old, 6)
        with self.assertRaises(ValueError):
            write_bonus_wage_ledger([employee()], old, out, call_round=7)

    def test_unknown_sheets_and_conflicting_overlap_are_rejected(self):
        template = self.path / 'template.xlsx'
        blank_template(template, 7)
        wb = openpyxl.load_workbook(template)
        ws = wb.active
        ws['C2'] = '東京都'
        ws['C3'] = datetime.datetime(2026, 7, 1)
        ws['C7'], ws['D7'], ws['E7'] = '合成従業員', '正社員', 100
        ws['N7'] = 123000  # 2026/6 in the fixed window
        ws['R7'] = 124000  # same 2026/6 as latest month, but conflicting
        wb.save(template)
        wb.close()
        book = read_bonus_wage_ledger(template, call_round=7)
        self.assertTrue(book.input_issues)
        self.assertEqual(judge_bonus_points(book).bonus1_status, '未確定')
        wb = openpyxl.Workbook()
        wb.save(template)
        wb.close()
        with self.assertRaises(ValueError):
            read_bonus_wage_ledger(template, call_round=7)

    def test_true_blank_and_formula_protection(self):
        template = self.path / 'source.xlsx'
        out = self.path / 'clear.xlsx'
        wb = openpyxl.Workbook()
        wb.active['C25'], wb.active['D7'] = '6', '=1+2'
        wb.save(template)
        wb.close()
        patch_xlsx(template, out, {'C25': None})
        wb = openpyxl.load_workbook(out)
        self.assertIsNone(wb.active['C25'].value)
        self.assertEqual(wb.active['D7'].value, '=1+2')
        wb.close()
        with self.assertRaises(ValueError):
            patch_xlsx(template, self.path / 'bad.xlsx', {'D7': None})
        before = template.read_bytes()
        with self.assertRaises(ValueError):
            patch_xlsx(template, template, {'C25': None})
        self.assertEqual(template.read_bytes(), before)

    def test_missing_application_month_driver_does_not_create_final_forms(self):
        template = self.path / 'template.xlsx'
        out = self.path / 'ledger.xlsx'
        blank_template(template, 7)
        write_bonus_wage_ledger([employee()], template, out, prefecture='東京都', call_round=7)
        result = subprocess.run([
            sys.executable, '-B', str(ROOT / '.claude/skills/katen-judge/review/run_katen_judge.py'),
            str(out), '--call-round', '7', '--fill', '通常', '--outdir', str(self.path / 'final')
        ], capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(result.returncode, 3, result.stderr)
        self.assertFalse(list((self.path / 'final').glob('*.xlsx')))
        self.assertNotIn('Traceback', result.stderr)

    def test_legacy_normal_calculation(self):
        old = employee(hourly=1163, latest_hourly=1226, call_round=6)
        old.monthly_base[(2026, 5)] = 122600
        # Legacy bonus2 baseline is July 2025, which is included in the fixed window.
        book = ledger([old], call_round=6)
        book.application_ym = (2026, 6)
        result = judge_bonus_points(book)
        self.assertEqual(result.bonus1_status, '対象')
        self.assertEqual(result.bonus2_diff, 63)
        self.assertEqual(result.bonus2_status, '対象')

    def test_new_official_forms_preserve_parts_and_fill_dates(self):
        forms = ROOT / '補助金加点/第7次以降'
        if not forms.exists():
            self.skipTest('第7次公式様式はrootの原本準備後に検証')
        result = judge_bonus_points(ledger())
        for name, first_row in [('補助率引き上げ・加点措置①用.xlsx', 19), ('加点措置①用.xlsx', 18)]:
            original = forms / name
            out = self.path / name
            fill_bonus_sheet_1(original, out, result)
            with warnings.catch_warnings():
                warnings.simplefilter('ignore')
                wb = openpyxl.load_workbook(out)
            ws = wb.worksheets[0]
            self.assertEqual(ws['D5'].value, datetime.datetime(2026, 12, 1))
            for col in ('H', 'R', 'AB'):
                self.assertIsInstance(ws[f'{col}{first_row - 3}'].value, datetime.datetime)
            self.assertEqual(ws[f'C{first_row}'].value, '合成従業員1')
            wb.close()
            with zipfile.ZipFile(original) as before, zipfile.ZipFile(out) as after:
                for name_in_zip in before.namelist():
                    if name_in_zip not in ('xl/worksheets/sheet1.xml', 'xl/workbook.xml'):
                        self.assertEqual(before.read(name_in_zip), after.read(name_in_zip))
                ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
                old_formulas = [ElementTree.tostring(f) for f in ElementTree.fromstring(before.read('xl/worksheets/sheet1.xml')).iter(ns + 'f')]
                new_formulas = [ElementTree.tostring(f) for f in ElementTree.fromstring(after.read('xl/worksheets/sheet1.xml')).iter(ns + 'f')]
                self.assertEqual(old_formulas, new_formulas)
        original = forms / '加点措置②用.xlsx'
        out = self.path / 'second.xlsx'
        fill_bonus_sheet_2(original, out, result)
        wb = openpyxl.load_workbook(out)
        self.assertIsNone(wb.worksheets[0]['C25'].value)
        self.assertEqual(wb.worksheets[0]['C1017'].value, '*')
        self.assertEqual(wb.worksheets[0]['C17'].value, '合成従業員1')
        self.assertIn('<55', wb.worksheets[0]['D10'].value)
        wb.close()
        with self.assertRaises(ValueError):
            fill_bonus_sheet_2(ROOT / '補助金加点/加点措置②用.xlsx', self.path / 'wrong.xlsx', result)


if __name__ == '__main__':
    unittest.main(verbosity=2)
