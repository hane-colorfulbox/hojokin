# -*- coding: utf-8 -*-
"""加点の「時間換算の対象月額」と労働時間の定義の回帰テスト。合成データ・スタブだけで、API は呼ばない。

1. 対象月額＝その月に支払われた賃金から、厚労省「最低賃金額以上かどうかを確認する方法」で除く6種類を
   除いた額（手当を含み、基本給に含まれる固定残業代は除く）。「毎月決まって払う手当」に限らない。
   歩合給・出来高給がある月は AI では null（総労働時間での換算が要るため推測しない）。
2. 対象月額が基本給より少ない月は、判定を止めずに注意を出す（固定残業代込みの基本給では正しい入力）。
3. 加点の労働時間（AI の basis_hours）は時給・日給制の所定内の実労働時間。月給制は null（E列の1か月平均所定労働時間）。
   時間外・休日労働を含む hours は加点に使わない。
"""
import importlib.util
import os
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

from hojokin import ai_extractor
from hojokin.subsidy_rules import rules_for_round
from hojokin.wage_reader import (
    BonusWageEmployee, BonusWageLedger, _ai_data_to_bonus_employees, _ai_data_to_wage_employees,
    fill_bonus_sheet_1, fill_bonus_sheet_2, judge_bonus_points, ym_label,
)

LEDGER7 = ROOT / 'ツール' / '第7次以降' / '加点判定用賃金台帳テンプレート.xlsx'
FORMS7 = ROOT / '補助金加点' / '第7次以降'
BUILD_SCRIPT = ROOT / 'scripts' / 'build_round7_templates.py'

APPLICATION_YM, LATEST_YM = (2026, 11), (2026, 10)
WINDOW = list(rules_for_round(7).bonus1_window)
# 基本給に固定残業代40,000円を含む月給制。対象月額＝260,000−40,000、所定160時間 → 1,375円。
FIXED_OT_BASE, FIXED_OT_BASIS, FIXED_OT_HOURS, FIXED_OT_HOURLY = 260_000, 220_000, 160, 1375
# 時給制パート。総労働時間130時間のうち時間外10時間 → 所定内120時間。
PART_TOTAL_HOURS, PART_BASIS_HOURS = 130, 120
PART_PAY = 156_000

D1_PHRASES = (
    'その月に支払われた賃金から、厚生労働省「最低賃金額以上かどうかを確認する方法」で除く6種類',
    '基本給と、職務手当・役職手当・資格手当などの手当を含む',
    '基本給に固定残業代（みなし残業代）が含まれる場合はその分を除く',
)
D2_PHRASE = '**歩合給・出来高給・インセンティブがある月は null**'
D4_PHRASES = (
    '"basis_hours": 数値 or null',
    '**時間外・休日労働の時間は含めない**',
    '**月給制は null**',
    '「総労働時間 − 時間外・休日労働時間」',
    '時間外を含む合計しか無く時間外の時間を分けられない月',
)
STANDARD_HOURS_LINE = '- **hours**: 総労働時間/実労働時間/所定労働時間。下記ルール6に従い、無ければ null。'
OLD_PHRASE = '毎月決まって'


def prompt_copies():
    full = ai_extractor.PROMPT_WAGE_LEDGER.format(fiscal_period_section='<期間>', tsv_data='<表>')
    return {'static': ai_extractor.PROMPT_WAGE_LEDGER_STATIC, 'full': full}


def load_build_script():
    spec = importlib.util.spec_from_file_location('build_round7_templates_for_test', BUILD_SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def worker(name, base, basis, hours, months=None):
    yms = months or WINDOW + [LATEST_YM]
    return BonusWageEmployee(1, name, '正社員', hours, {ym: base for ym in yms},
                             monthly_hourly_basis={ym: basis for ym in yms}, hourly_basis_separate=True)


def ledger(employees):
    return BonusWageLedger(prefecture='東京都', application_ym=APPLICATION_YM, call_round=7, employees=employees)


def below_base_notes(result, name):
    return [n for n in result.notes if n.startswith(f'{name}: 時間換算の対象月額が基本給より少ない月があります')]


class PromptDefinitionTests(unittest.TestCase):
    def test_both_copies_define_the_basis_by_the_official_method(self):
        for label, text in prompt_copies().items():
            for phrase in D1_PHRASES + (D2_PHRASE,):
                self.assertEqual(text.count(phrase), 1, (label, phrase))
            self.assertNotIn(OLD_PHRASE, text, label)

    def test_both_copies_define_bonus_hours_without_changing_standard_hours(self):
        for label, text in prompt_copies().items():
            for phrase in D4_PHRASES:
                self.assertEqual(text.count(phrase), 1, (label, phrase))
            self.assertEqual(text.count(STANDARD_HOURS_LINE), 1, label)


class AiConversionTests(unittest.TestCase):
    def test_bonus_path_uses_basis_hours_and_ignores_overtime_hours(self):
        months = [{'ym': '2025-10', 'taxable': PART_PAY, 'base': PART_PAY, 'hourly_basis': PART_PAY,
                   'hours': PART_TOTAL_HOURS, 'basis_hours': PART_BASIS_HOURS},
                  {'ym': '2025-11', 'taxable': PART_PAY, 'base': PART_PAY, 'hourly_basis': PART_PAY,
                   'hours': PART_TOTAL_HOURS, 'basis_hours': None}]
        emp = _ai_data_to_bonus_employees([{'name': '合成 花子', 'employment_type': 'パート・アルバイト',
                                            'monthly': months}])[0]
        self.assertEqual(emp.monthly_hours_override, {(2025, 10): PART_BASIS_HOURS})
        self.assertIsNone(emp.hours_for((2025, 11)))

    def test_old_output_without_basis_hours_keeps_hours(self):
        old = [{'name': '合成 花子', 'employment_type': 'パート・アルバイト',
                'monthly': [{'ym': '2025-10', 'base': PART_PAY, 'hours': PART_TOTAL_HOURS}]}]
        emp = _ai_data_to_bonus_employees(old)[0]
        self.assertEqual(emp.monthly_hours_override, {(2025, 10): PART_TOTAL_HOURS})

    def test_standard_ledger_path_still_reads_hours(self):
        data = [{'name': '合成 花子', 'employment_type': 'パート・アルバイト', 'bonuses': [],
                 'monthly': [{'ym': f'2025-{m:02d}', 'taxable': PART_PAY, 'base': PART_PAY,
                              'hours': PART_TOTAL_HOURS, 'basis_hours': PART_BASIS_HOURS, 'work_days': 20}
                             for m in range(1, 13)]}]
        emp = _ai_data_to_wage_employees(data, '2025-01〜2025-12')[0]
        self.assertEqual(emp.monthly_hours, [PART_TOTAL_HOURS] * 12)

    def test_commission_month_returned_as_null_stays_unconfirmed(self):
        target = WINDOW + [LATEST_YM]
        monthly = [{'ym': f'{y}-{m:02d}', 'taxable': FIXED_OT_BASE, 'base': FIXED_OT_BASE,
                    'hourly_basis': None if (y, m) == WINDOW[0] else FIXED_OT_BASIS,
                    'hours': None, 'basis_hours': FIXED_OT_HOURS} for y, m in target]
        emp = _ai_data_to_bonus_employees([{'name': '合成 太郎', 'employment_type': '正社員',
                                            'monthly': monthly}], target_yms=target)[0]
        result = judge_bonus_points(ledger([emp]))
        self.assertEqual(result.bonus1_status, '未確定')
        self.assertIn('合成 太郎: 時間換算の対象月額が未確認', ' '.join(result.bonus1_issues))
        self.assertFalse(result.bonus1_details[0]['complete'])
        self.assertTrue(all(d['complete'] for d in result.bonus1_details[1:]))


class BasisBelowBaseTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_fixed_overtime_inside_base_is_confirmed_with_one_warning(self):
        result = judge_bonus_points(ledger([worker('合成 太郎', FIXED_OT_BASE, FIXED_OT_BASIS, FIXED_OT_HOURS)]))
        self.assertEqual((result.bonus1_status, result.bonus2_status), ('対象外', '対象外'))
        self.assertEqual((result.bonus1_issues, result.bonus2_issues), ([], []))
        notes = below_base_notes(result, '合成 太郎')
        self.assertEqual(len(notes), 1, result.notes)
        self.assertIn(f'（{ym_label(WINDOW[0])}、', notes[0])
        self.assertIn(f'{ym_label(LATEST_YM)}）', notes[0])
        self.assertIn('固定残業代（みなし残業代）', notes[0])
        self.assertIn('対象月額を基本給に合わせて引き上げない', notes[0])
        detail = result.bonus1_details[0]['employees'][0]
        self.assertEqual((detail['base'], detail['basis'], detail['hourly']),
                         (FIXED_OT_BASE, FIXED_OT_BASIS, FIXED_OT_HOURLY))

        form1 = fill_bonus_sheet_1(FORMS7 / '加点措置①用.xlsx', self.path / 'f1.xlsx', result)
        ws = openpyxl.load_workbook(form1).worksheets[0]
        self.assertEqual((ws['H18'].value, ws['I18'].value), (FIXED_OT_BASE, FIXED_OT_HOURLY))
        form2 = fill_bonus_sheet_2(FORMS7 / '加点措置②用.xlsx', self.path / 'f2.xlsx', result)
        ws = openpyxl.load_workbook(form2).worksheets[0]
        self.assertEqual((ws['F17'].value, ws['G17'].value, ws['N17'].value, ws['O17'].value),
                         (FIXED_OT_BASE, FIXED_OT_HOURLY, FIXED_OT_BASE, FIXED_OT_HOURLY))

    def test_warning_lists_only_the_months_below_base(self):
        emp = worker('合成 次郎', FIXED_OT_BASE, FIXED_OT_BASE, FIXED_OT_HOURS)
        emp.monthly_hourly_basis[WINDOW[2]] = FIXED_OT_BASIS
        result = judge_bonus_points(ledger([emp, worker('合成 三郎', FIXED_OT_BASE, FIXED_OT_BASE, FIXED_OT_HOURS)]))
        notes = below_base_notes(result, '合成 次郎')
        self.assertEqual(len(notes), 1, result.notes)
        self.assertIn(f'（{ym_label(WINDOW[2])}）', notes[0])
        self.assertEqual(below_base_notes(result, '合成 三郎'), [])
        self.assertNotEqual(result.bonus1_status, '未確定')

    def test_blank_or_zero_base_still_blocks(self):
        for blank in (None, 0):
            emp = worker('合成 太郎', FIXED_OT_BASE, FIXED_OT_BASIS, FIXED_OT_HOURS)
            if blank is None:
                del emp.monthly_base[WINDOW[0]]
            else:
                emp.monthly_base[WINDOW[0]] = blank
            result = judge_bonus_points(ledger([emp]))
            self.assertEqual(result.bonus1_status, '未確定', blank)
            self.assertIn('合成 太郎: 基本給（様式の【基本給】欄）が未確認', ' '.join(result.bonus1_issues))


class LedgerTemplateTests(unittest.TestCase):
    def test_template_texts_follow_the_official_definition(self):
        build = load_build_script()
        wb = openpyxl.load_workbook(LEDGER7)
        try:
            sheet, rules = wb['加点判定用明細'], wb['記入ルール']
            b8, b4, banner = rules['B8'].value, sheet['B4'].value, sheet['AH5'].value
        finally:
            wb.close()
        self.assertEqual((b8, b4, banner), (build.LEDGER_RULES['B8'], build.LEDGER_B4, build.LEDGER_BASIS_BANNER))
        for phrase in ('各暦月に支払われた賃金から', '基本給と、職務手当・役職手当・資格手当などの手当を含む',
                       '基本給に固定残業代（みなし残業代）が含まれる場合はその分を除く', '基本給より少なくなることがある',
                       '基本給に合わせて引き上げない', '歩合・出来高の額×（その月の労働時間',
                       '÷その月の総労働時間）', '総労働時間を原本で確かめられない月は空欄にする'):
            self.assertIn(phrase, b8)
        for phrase in ('その月の賃金から厚労省方式で除く6種類を除いた額', '基本給に含まれる固定残業代は除く',
                       '歩合給・出来高給は記入ルールのとおり換算', 'AH〜AT列に時間換算の対象月額'):
            self.assertIn(phrase, b4)
        self.assertIn('基本給に含む固定残業代を除いた額', banner)
        self.assertIn('歩合給は記入ルール参照', banner)
        for text in (b8, b4, banner):
            self.assertNotIn(OLD_PHRASE, text)
        self.assertIn(r'毎月決まって払う', build.STALE_PATTERNS)


if __name__ == '__main__':
    unittest.main()
