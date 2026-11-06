# -*- coding: utf-8 -*-
"""在籍の注意と読み取れなかった月の知らせ方の回帰テスト。合成データ・スタブだけで、Anthropic API は呼ばない。

1. 生産性の停止の案内は、「非在籍」を読む標準テンプレートの形式を名指しする。
2. 入社日・退職日の外に給与がある月を非在籍にしなかった注意と、日付を読めなかった注意は、
   賃金台帳作成の戻り値・B1〜B3・完了メッセージに出る。確認の残らない注意だけの人は出さない。
3. 加点台帳作成で、基本給も時間換算の対象月額も読めなかった月（非在籍以外）がある人は、
   完了メッセージに名前が出る。すべての対象月が非在籍の人には基本給の入力を求めない。
"""
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

from hojokin import pipeline
from hojokin.ai_extractor import StubExtractor
from hojokin.pipeline import ABSENT_MONTH_GUIDE
from hojokin.subsidy_rules import rules_for_round
from hojokin.wage_ledger_writer import is_unresolved_presence_note, write_wage_ledger_to_template
from hojokin.wage_reader import _ai_data_to_bonus_employees, _ai_data_to_wage_employees, read_wage_ledgers

STD_TEMPLATE = ROOT / 'ツール' / '賃金台帳テンプレート.xlsx'
LEDGER7 = ROOT / 'ツール' / '第7次以降' / '加点判定用賃金台帳テンプレート.xlsx'
SHEET = '従業員別明細'
MEMO_SHEET = '変換メモ'
DUMMY_PDF = b'%PDF-1.4\n%dummy\n'
FISCAL_PERIOD = '2025-01〜2025-12'
FISCAL_MONTH = 12
WAGE, BASE, BASIS, HOURS = 200_000, 180_000, 190_000, 160
APPLICATION_YM, LATEST_YM = (2026, 11), (2026, 10)
UNREAD_YM = (2026, 1)

RETIRED = '合成 退職'
UNREADABLE = '合成 日付不明'
VETERAN = '合成 古参'
APRIL_HIRE = '合成 四月入社'


def wage_ai_employee(name, first_month=1, last_month=12, **dates):
    monthly = [{'ym': f'2025-{m:02d}', 'taxable': WAGE, 'base': WAGE, 'hours': HOURS, 'work_days': 20}
               for m in range(first_month, last_month + 1)]
    data = {'name': name, 'employment_type': '正社員', 'monthly': monthly, 'bonuses': []}
    data.update(dates)
    return data


def presence_cases(resolved_only=False):
    resolved = [wage_ai_employee(VETERAN, hire_date='2020-04-01'),
                wage_ai_employee(APRIL_HIRE, first_month=4, hire_date='2025-04-01')]
    if resolved_only:
        return resolved
    return [wage_ai_employee(RETIRED, leave_date='2025-11-30'),
            wage_ai_employee(UNREADABLE, hire_date='入社：春頃')] + resolved


class FixedWageExtractor(StubExtractor):
    def __init__(self, data):
        super().__init__()
        self.data = data

    def extract_wage_ledger(self, tsv_data, fiscal_period_hint=None, pdf_files=None, **_kwargs):
        return self.data


class AbsentGuideTests(unittest.TestCase):
    def test_guide_names_the_standard_template(self):
        self.assertIn('標準テンプレート', ABSENT_MONTH_GUIDE)
        self.assertIn('G〜R列', ABSENT_MONTH_GUIDE)
        self.assertIn('「非在籍」を読まない', ABSENT_MONTH_GUIDE)
        self.assertIn('月度給与', ABSENT_MONTH_GUIDE)


class UnresolvedPresenceTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def convert(self, data):
        return _ai_data_to_wage_employees(data, FISCAL_PERIOD)

    def test_predicate_matches_the_notes_the_reader_makes(self):
        by_name = {e.name: e for e in self.convert(presence_cases())}
        self.assertIn('非在籍にしなかった月：令和7年12月', by_name[RETIRED].presence_note)
        self.assertIn('決めていません', by_name[UNREADABLE].presence_note)
        self.assertTrue(is_unresolved_presence_note(by_name[RETIRED].presence_note))
        self.assertTrue(is_unresolved_presence_note(by_name[UNREADABLE].presence_note))
        self.assertEqual(by_name[VETERAN].presence_note, '台帳の入社日 2020-04-01から判定')
        for name in (VETERAN, APRIL_HIRE):
            self.assertFalse(is_unresolved_presence_note(by_name[name].presence_note), name)
        for note in ('', None):
            self.assertFalse(is_unresolved_presence_note(note))

    def test_writer_returns_and_shows_unresolved_notes(self):
        out = self.path / 'out.xlsx'
        result = write_wage_ledger_to_template(self.convert(presence_cases()), STD_TEMPLATE, out,
                                               company_name='合成')
        self.assertEqual([name for name, _note in result.presence_unresolved], [RETIRED, UNREADABLE])
        self.assertTrue(any('日付を読めない 2名' in w for w in result.warnings), result.warnings)
        self.assertEqual([name for name, _labels, _basis in result.absent_rows], [APRIL_HIRE])
        wb = openpyxl.load_workbook(out)
        self.assertIn('日付を読めない 2名', wb[SHEET]['B3'].value)
        memo = [str(c.value) for row in wb[MEMO_SHEET].iter_rows() for c in row if c.value]
        wb.close()
        self.assertTrue(any(text.startswith('⚠ ') and '令和7年12月' in text for text in memo), memo)
        self.assertIn('台帳の入社日 2020-04-01から判定', memo, '確認の残らない注意には⚠を付けない')
        self.assertEqual(len(read_wage_ledgers([out])), 4, '変換メモが賃金の表として読まれていないこと')

    def test_resolved_note_alone_does_not_warn(self):
        out = self.path / 'out.xlsx'
        veteran = self.convert([wage_ai_employee(VETERAN, hire_date='2020-04-01')])
        result = write_wage_ledger_to_template(veteran, STD_TEMPLATE, out, company_name='合成')
        self.assertEqual(result.presence_unresolved, [])
        self.assertEqual(result.warnings, [])

    def run_conversion(self, data):
        source = self.path / 'source'
        source.mkdir()
        (source / '賃金台帳_合成.pdf').write_bytes(DUMMY_PDF)
        return pipeline.run_wage_ledger_conversion(
            source, '合成', STD_TEMPLATE, self.path / '合成_賃金台帳一覧.xlsx',
            extractor=FixedWageExtractor(data), fiscal_month_override=FISCAL_MONTH)

    def test_completion_message_lists_unresolved_people(self):
        status = self.run_conversion(presence_cases())
        self.assertEqual(status.status, '完了', status.message)
        self.assertIn(f'日付を読めなかった従業員 2名（{RETIRED}、{UNREADABLE}）', status.message)
        self.assertIn('変換メモシートで原本と照合', status.message)

    def test_completion_message_is_quiet_for_resolved_notes(self):
        status = self.run_conversion(presence_cases(resolved_only=True))
        self.assertEqual(status.status, '完了', status.message)
        self.assertNotIn('日付を読めなかった', status.message)
        self.assertIn('非在籍の月あり1名', status.message)


def bonus_target_yms():
    return list(rules_for_round(7).bonus1_window) + [LATEST_YM]


def bonus_ai_employee(name, skip=(), **dates):
    monthly = []
    for y, m in bonus_target_yms():
        if (y, m) in skip:
            continue
        unread = (y, m) == UNREAD_YM
        monthly.append({'ym': f'{y}-{m:02d}', 'taxable': WAGE, 'base': None if unread else BASE,
                        'hourly_basis': None if unread else BASIS, 'hours': HOURS})
    data = {'name': name, 'employment_type': '正社員', 'monthly': monthly}
    data.update(dates)
    return data


def complete_bonus_employee(name):
    data = bonus_ai_employee(name)
    for entry in data['monthly']:
        entry['base'], entry['hourly_basis'] = BASE, BASIS
    return data


class BonusUnreadMonthTests(unittest.TestCase):
    UNREAD = '合成 一月欠け'
    COMPLETE = '合成 全月あり'
    LEFT = '合成 窓前退職'

    def employees(self):
        left_months = set(bonus_target_yms())
        return _ai_data_to_bonus_employees([
            bonus_ai_employee(self.UNREAD),
            complete_bonus_employee(self.COMPLETE),
            dict(bonus_ai_employee(self.LEFT, skip=left_months), leave_date='2025-06-30'),
        ], target_yms=bonus_target_yms())

    def test_month_without_base_and_basis_is_listed(self):
        employees = self.employees()
        unread = next(e for e in employees if e.name == self.UNREAD)
        self.assertNotIn(UNREAD_YM, unread.monthly_base)
        self.assertNotIn(UNREAD_YM, unread.monthly_hourly_basis)
        text = ' '.join(pipeline._bonus_base_coverage_messages(employees, bonus_target_yms()))
        self.assertIn(f'基本給を読み取れなかった月がある従業員: 1名（{self.UNREAD}）', text)
        self.assertIn('AH〜AT列', text)
        self.assertNotIn(self.COMPLETE, text)
        legacy = ' '.join(pipeline._bonus_base_coverage_messages(employees, bonus_target_yms(), False))
        self.assertIn(f'基本給も時間換算の対象月額も読み取れなかった月がある従業員: 1名（{self.UNREAD}）', legacy)
        self.assertNotIn('AH〜AT列', legacy)

    def test_month_with_basis_but_no_base_depends_on_the_ledger_columns(self):
        data = bonus_ai_employee(self.UNREAD)
        for entry in data['monthly']:
            entry['hourly_basis'] = BASIS
        employees = _ai_data_to_bonus_employees([data], target_yms=bonus_target_yms())
        with_block = ' '.join(pipeline._bonus_base_coverage_messages(employees, bonus_target_yms(), True))
        self.assertIn(self.UNREAD, with_block)
        # 旧回用の台帳は F〜R列に対象月額を書くので、その月は空欄にならない
        self.assertEqual(pipeline._bonus_base_coverage_messages(employees, bonus_target_yms(), False), [])

    def test_all_absent_person_is_not_asked_for_base(self):
        employees = self.employees()
        left = next(e for e in employees if e.name == self.LEFT)
        self.assertTrue(all(left.monthly_presence.get(ym) is False for ym in bonus_target_yms()))
        text = ' '.join(pipeline._bonus_base_coverage_messages(employees, bonus_target_yms()))
        self.assertNotIn(self.LEFT, text)
        basis_text = ' '.join(pipeline._bonus_basis_coverage_messages(employees, bonus_target_yms()))
        self.assertNotIn('時間換算の対象月額が読み取れなかった月がある従業員', basis_text)

    def test_bonus_ledger_creation_message_lists_unread_month(self):
        data = [bonus_ai_employee(self.UNREAD), complete_bonus_employee(self.COMPLETE)]
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / 'source'
            source.mkdir()
            (source / '賃金台帳_合成.pdf').write_bytes(DUMMY_PDF)
            status = pipeline.run_bonus_wage_ledger_creation(
                source, '合成', LEDGER7, Path(tmp) / '加点判定用賃金台帳.xlsx',
                extractor=FixedWageExtractor(data), prefecture='東京都',
                application_ym=APPLICATION_YM, call_round=7)
        self.assertEqual(status.status, '完了', status.message)
        self.assertIn(f'基本給を読み取れなかった月がある従業員: 1名（{self.UNREAD}）', status.message)
        self.assertNotIn('加点対象月の基本給が取得できなかった従業員', status.message)
        self.assertNotIn('時間換算の対象月額が読み取れなかった月がある従業員', status.message)


if __name__ == '__main__':
    unittest.main()
