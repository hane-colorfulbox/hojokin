"""第7次：表明日・賃上げ幅・過去交付の年度・原本の受け付け・加点ドライバの公募回（合成データ、API不使用）。"""
from __future__ import annotations

import datetime
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

from hojokin.application_context import (
    ApplicationContext, ApplicationValidationError, validate_application_context, wage_requirements,
)
from hojokin.config import SHINSEI_ANCHOR_LABELS, get_mapping
from hojokin.models import ExtractionResult, FinancialData
from hojokin.template_filler import (
    DECLARATION_DATE_FORMAT, DECLARATION_GUIDANCE, STALE_DECLARATION_GUIDANCE,
    WAGE_RAISE_30, WAGE_RAISE_50, WAGE_RAISE_NONE, fill_shinsei_sheet, fill_template, merge_review_notes,
)
from hojokin.template_versions import (
    APPLICATION_FILENAMES, application_template_issues, find_application_template,
    select_application_template, template_rule_marker, validate_application_template_version,
)

ROUND7_DIR = ROOT / 'ツール' / '第7次以降'
LEGACY_DIR = ROOT / 'ツール'
DRIVER = ROOT / '.claude' / 'skills' / 'katen-judge' / 'review' / 'run_katen_judge.py'
LEGACY_LEDGER_TEMPLATE = LEGACY_DIR / '加点判定用賃金台帳テンプレート.xlsx'
DECLARED_ON = datetime.datetime(2026, 11, 1)
HEARING_50 = '❸＋５０円以上'
PAST_GRANT_RATE_PLAN = (10_000_000, 10_350_000, 10_712_250, 11_087_179)
PRODUCTIVITY_HOURS = 8000
EXIT_INPUT_ERROR = 2


def values(**changes):
    v = {
        'call_round': 7, 'application_date': '2026-11-10',
        'fiscal_year_start': '2026-04-01', 'fiscal_year_end': '2027-03-31',
        'frame': '通常', 'subsidy_amount': 1_500_000, 'past_grants': [],
        'optional_wage_bonus': False, 'wage_exempt': False,
        'payroll_scope_confirmed': True,
        'declaration_confirmed': True, 'declaration_method': '書面',
        'declaration_date': DECLARED_ON.date().isoformat(), 'planned_min_wage_increase': 30,
        'workplace_prefecture': '東京都', 'workplace_confirmed': True,
        'regional_min_wage_yen': 1280,
        'forecast_years': [
            {'total_salary': amount, 'employee_count_fte': 10,
             'source': '担当者確認の合成見込表', 'confirmed': True}
            for amount in (10_000_000, 10_300_000, 10_609_000, 10_927_270)
        ],
        'productivity_forecast_years': [
            {'operating_profit': 2_000_000, 'personnel_cost': amount,
             'depreciation': 1_000_000, 'total_annual_hours': 1000,
             'source': '担当者確認の合成生産性予算', 'confirmed': True}
            for amount in (15_000_000, 16_000_000, 17_000_000)
        ],
    }
    v.update(changes)
    return v


def settings_for(kind: str, **changes) -> dict:
    v = values(frame='通常' if kind.startswith('通常') else 'インボイス')
    if '個人' in kind:
        v.update(fiscal_year_start='2026-01-01', fiscal_year_end='2026-12-31')
    v.update(changes)
    return v


def prior_financial(kind: str) -> FinancialData:
    if '個人' in kind:
        return FinancialData(fiscal_year_start='2025-01-01', fiscal_year_end='2025-12-31',
                             operating_profit=1_000_000, salary=4_000_000)
    return FinancialData(fiscal_year_start='2025-04-01', fiscal_year_end='2026-03-31',
                         operating_profit=1_000_000, salary=4_000_000)


def strip_core_properties(source: Path, dest: Path) -> None:
    """Drive から書き出した原本と同じく docProps/core.xml（版マーカー）のない複製を作る。"""
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(dest, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename == 'docProps/core.xml':
                continue
            data = zin.read(info.filename)
            if info.filename == '[Content_Types].xml':
                data = re.sub(rb'<Override[^>]*docProps/core\.xml[^>]*/>', b'', data)
            if info.filename == '_rels/.rels':
                data = re.sub(rb'<Relationship[^>]*docProps/core\.xml[^>]*/>', b'', data)
            zout.writestr(info, data)


class Round7OutputCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self._count = 0

    def tearDown(self):
        self._tmp.cleanup()

    def fill(self, kind: str, hearing_value: str | None = None, template: Path | None = None, **changes):
        mapping = get_mapping(kind)
        self._count += 1
        out = self.tmp / f'{kind}_{self._count}.xlsx'
        hearing = ({mapping.hearing_wage_raise_row: {'label': '賃金引上げ幅', 'value': hearing_value}}
                   if hearing_value else {})
        empty = fill_template(template or ROUND7_DIR / APPLICATION_FILENAMES[kind], out, mapping, hearing,
                              ExtractionResult(financial=prior_financial(kind)),
                              application_context=ApplicationContext.from_dict(settings_for(kind, **changes)),
                              productivity_total_annual_hours=PRODUCTIVITY_HOURS)
        wb = openpyxl.load_workbook(out)
        self.addCleanup(wb.close)
        return wb['申請内容'], mapping, empty, out


class DeclarationOutputTests(Round7OutputCase):
    """#4 確認済みの表明日を書き、未確認なら空欄にして理由を出す。"""

    def test_confirmed_declaration_date_and_guidance_in_all_four_templates(self):
        for kind in APPLICATION_FILENAMES:
            with self.subTest(kind=kind):
                ws, mapping, _, _ = self.fill(kind)
                row = mapping.shinsei['wage_raise_date']
                self.assertEqual(ws.cell(row, 3).value, DECLARED_ON)
                self.assertEqual(ws.cell(row, 3).number_format, DECLARATION_DATE_FORMAT)
                self.assertEqual(ws.cell(mapping.shinsei['wage_raise_method'], 3).value, '書面')
                self.assertEqual(ws.cell(mapping.shinsei['wage_raise_declaration'], 3).value, '■はい\n□いいえ')
                for field in DECLARATION_GUIDANCE:
                    text = ws.cell(mapping.shinsei[field], 4).value or ''
                    self.assertFalse(any(s in text for s in STALE_DECLARATION_GUIDANCE), text)

    def test_stale_guidance_in_any_original_is_replaced_and_new_guidance_kept(self):
        mapping = get_mapping('通常枠_2026')
        wb = openpyxl.Workbook()
        ws = wb.active
        kept = '顧客に確認した表明の実施日を入力'
        ws.cell(mapping.shinsei['wage_raise_method'], 4, '特に指定がなければ、「口頭によって」を選択する')
        ws.cell(mapping.shinsei['wage_raise_date'], 4, '申請日より前の日付を設定（シート作成日）')
        ws.cell(mapping.shinsei['wage_raise_date'], 3, '2026/10/01')
        legacy = ApplicationContext.from_dict(values(call_round=6))
        fill_shinsei_sheet(ws, mapping, ExtractionResult(), application_context=legacy)
        self.assertIn('シート作成日', ws.cell(mapping.shinsei['wage_raise_date'], 4).value)
        round7 = ApplicationContext.from_dict(values(declaration_confirmed=None, declaration_method='',
                                                     declaration_date=None))
        fill_shinsei_sheet(ws, mapping, ExtractionResult(), application_context=round7)
        self.assertEqual(ws.cell(mapping.shinsei['wage_raise_method'], 4).value,
                         DECLARATION_GUIDANCE['wage_raise_method'])
        self.assertEqual(ws.cell(mapping.shinsei['wage_raise_date'], 4).value,
                         DECLARATION_GUIDANCE['wage_raise_date'])
        self.assertIsNone(ws.cell(mapping.shinsei['wage_raise_date'], 3).value)
        ws.cell(mapping.shinsei['wage_raise_date'], 4, kept)
        fill_shinsei_sheet(ws, mapping, ExtractionResult(), application_context=round7)
        self.assertEqual(ws.cell(mapping.shinsei['wage_raise_date'], 4).value, kept)
        wb.close()

    def test_unconfirmed_declaration_stays_blank_with_reason(self):
        ws, mapping, empty, _ = self.fill('通常枠_2026', subsidy_amount=1_000_000, declaration_confirmed=None,
                                          declaration_method='', declaration_date=None)
        for field in ('wage_raise_declaration', 'wage_raise_method', 'wage_raise_date'):
            self.assertIsNone(ws.cell(mapping.shinsei[field], 3).value, field)
        for field in ('wage_raise_method', 'wage_raise_date'):
            row = mapping.shinsei[field]
            notes = [e for e in empty if e.startswith(f'⚠ 行{row:3d} ')]
            self.assertEqual(len(notes), 1, empty)
            self.assertIn('推測の方法・日付は入れない', notes[0])
            self.assertFalse([e for e in empty if e.startswith(f'行{row:3d} ')], empty)

    def test_declared_not_done_is_recorded_as_no(self):
        ws, mapping, empty, _ = self.fill('インボイス枠_2026', subsidy_amount=500_000,
                                          declaration_confirmed=False, declaration_method='',
                                          declaration_date=None)
        self.assertEqual(ws.cell(mapping.shinsei['wage_raise_declaration'], 3).value, '□はい\n■いいえ')
        self.assertIsNone(ws.cell(mapping.shinsei['wage_raise_date'], 3).value)
        self.assertTrue(any('未実施' in e for e in empty), empty)


class WageRaiseChoiceTests(Round7OutputCase):
    """#10 賃上げ幅は案件設定で決め、加点を希望しない案件にヒアリングの＋50円を書かない。"""

    def amount(self, ws, mapping):
        return ws.cell(mapping.shinsei['wage_raise_amount'], 3).value

    def test_no_bonus_case_writes_no_raise_even_if_hearing_says_50(self):
        for kind in ('インボイス枠_2026', 'インボイス枠_個人_2026', '通常枠_2026'):
            with self.subTest(kind=kind):
                ws, mapping, empty, _ = self.fill(kind, HEARING_50, subsidy_amount=500_000)
                self.assertFalse(wage_requirements(ApplicationContext.from_dict(
                    settings_for(kind, subsidy_amount=500_000))).selected)
                self.assertEqual(self.amount(ws, mapping), WAGE_RAISE_NONE)
                row = mapping.shinsei['wage_raise_amount']
                self.assertTrue(any(e.startswith(f'⚠ 行{row:3d} [賃上げ幅]') and 'ヒアリング' in e
                                    for e in empty), empty)

    def test_optional_bonus_uses_confirmed_amount(self):
        for planned, expected in ((30, WAGE_RAISE_30), (50, WAGE_RAISE_50)):
            with self.subTest(planned=planned):
                ws, mapping, empty, _ = self.fill('インボイス枠_2026', subsidy_amount=500_000,
                                                  optional_wage_bonus=True, planned_min_wage_increase=planned)
                self.assertEqual(self.amount(ws, mapping), expected)
                self.assertFalse(any('[賃上げ幅]' in e for e in empty), empty)

    def test_large_normal_without_bonus_is_requirement_plus_30(self):
        ws, mapping, _, _ = self.fill('通常枠_2026', planned_min_wage_increase=50)
        self.assertEqual(self.amount(ws, mapping), WAGE_RAISE_30)

    def test_claimed_level_conflicting_with_hearing_stops_without_output(self):
        mapping = get_mapping('通常枠_2026')
        out = self.tmp / 'conflict.xlsx'
        with self.assertRaisesRegex(ApplicationValidationError, 'ヒアリングの賃金引上げ幅'):
            fill_template(ROUND7_DIR / APPLICATION_FILENAMES['通常枠_2026'], out, mapping,
                          {mapping.hearing_wage_raise_row: {'label': '賃金引上げ幅', 'value': HEARING_50}},
                          ExtractionResult(financial=prior_financial('通常枠_2026')),
                          application_context=ApplicationContext.from_dict(settings_for('通常枠_2026')),
                          productivity_total_annual_hours=PRODUCTIVITY_HOURS)
        self.assertFalse(out.exists())

    def test_past_grant_mandatory_without_bonus_does_not_claim_bonus_level(self):
        forecasts = [{'total_salary': amount, 'employee_count_fte': 10,
                      'source': '担当者確認の合成見込表', 'confirmed': True} for amount in PAST_GRANT_RATE_PLAN]
        ws, mapping, _, _ = self.fill(
            'インボイス枠_2026', HEARING_50, subsidy_amount=500_000, planned_min_wage_increase=None,
            past_grants=[{'year': 2024, 'frame': 'インボイス枠（インボイス対応類型）', 'decided': True}],
            forecast_years=forecasts)
        self.assertEqual(self.amount(ws, mapping), WAGE_RAISE_NONE)
        self.assertEqual(ws.cell(mapping.shinsei['past_subsidies'], 3).value, 'あり')

    def test_merge_review_notes_replaces_blank_rows_and_appends_others(self):
        merged = merge_review_notes(['行225 [表明を行った日付]', '行 93 [過去年度交付決定]'],
                                    [(225, '注意225'), (300, '注意300')])
        self.assertEqual(merged, ['注意225', '行 93 [過去年度交付決定]', '注意300'])


class PastGrantYearTests(unittest.TestCase):
    """#14 年度は補助金名の年度。2025より後（暦年の誤入力）や範囲外は止める。"""

    def context(self, year: int, frame: str = '通常枠'):
        return ApplicationContext.from_dict(values(past_grants=[{'year': year, 'frame': frame, 'decided': True}]))

    def test_calendar_year_or_current_program_year_stops(self):
        for year in (2026, 2027, 2015):
            with self.subTest(year=year):
                with self.assertRaisesRegex(ApplicationValidationError, '補助金名の年度'):
                    validate_application_context(self.context(year))

    def test_program_years_keep_requirement_bands(self):
        self.assertEqual(wage_requirements(self.context(2025, 'インボイス枠')).required_rate * 1000, 35)
        self.assertEqual(wage_requirements(self.context(2022, 'インボイス枠')).required_rate * 1000, 35)
        older = self.context(2021, 'インボイス枠')
        self.assertEqual(wage_requirements(older).required_rate * 100, 3)
        self.assertIn('IT導入補助金2021', older.past_grant_text())


class TemplateAcceptanceTests(Round7OutputCase):
    """#18 版マーカーのない原本（Drive書き出し・改名）も中身で確認して受け付け、通らない理由を返す。"""

    def setUp(self):
        super().setUp()
        self.dir = self.tmp / 'work'
        self.dir.mkdir()

    def test_markerless_renamed_round7_templates_are_selected_and_filled(self):
        for kind, name in APPLICATION_FILENAMES.items():
            with self.subTest(kind=kind):
                work = self.dir / kind
                work.mkdir()
                exported = work / f'顧客企業A_{kind}_Drive書き出し.xlsx'
                strip_core_properties(ROUND7_DIR / name, exported)
                self.assertEqual(template_rule_marker(exported), '')
                selection = find_application_template(work, kind, 7)
                self.assertEqual(selection.path, exported, selection.rejected)
                validate_application_template_version(exported, 7, kind)
                validate_application_template_version(exported, 7)
                with self.assertRaisesRegex(ValueError, '旧回'):
                    validate_application_template_version(exported, 6)

    def test_markerless_template_can_be_filled(self):
        exported = self.dir / '顧客企業A_通常枠_個人.xlsx'
        strip_core_properties(ROUND7_DIR / APPLICATION_FILENAMES['通常枠_個人_2026'], exported)
        ws, mapping, _, _ = self.fill('通常枠_個人_2026', template=exported)
        self.assertEqual(ws.cell(mapping.shinsei['wage_raise_date'], 3).value, DECLARED_ON)

    def test_rejections_explain_the_reason(self):
        legacy = self.dir / '顧客企業A_旧版.xlsx'
        strip_core_properties(LEGACY_DIR / APPLICATION_FILENAMES['インボイス枠_2026'], legacy)
        output = self.dir / '顧客企業A_インボイス枠_2026_AI版.xlsx'
        strip_core_properties(ROUND7_DIR / APPLICATION_FILENAMES['インボイス枠_2026'], output)
        selection = find_application_template(self.dir, 'インボイス枠_2026', 7)
        self.assertIsNone(selection.path)
        text = '\n'.join(selection.rejected)
        self.assertIn('顧客企業A_旧版.xlsx', text)
        self.assertIn('第7次以降用と確認できない', text)
        self.assertIn('ツールが作成した申請書', text)

    def test_wrong_kind_and_ambiguous_copies_are_not_used(self):
        exported = self.dir / '原本A.xlsx'
        strip_core_properties(ROUND7_DIR / APPLICATION_FILENAMES['インボイス枠_2026'], exported)
        self.assertTrue(application_template_issues(exported, 7, '通常枠_2026'))
        shutil.copyfile(exported, self.dir / '原本B.xlsx')
        selection = find_application_template(self.dir, 'インボイス枠_2026', 7)
        self.assertIsNone(selection.path)
        self.assertIn('複数', selection.rejected[-1])

    def synthetic(self, path: Path, extra_text: str) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '申請内容'
        mapping = get_mapping('通常枠_2026')
        for key, expected in SHINSEI_ANCHOR_LABELS.items():
            if key in mapping.shinsei:
                label = expected[0] if isinstance(expected, tuple) else expected
                if key == 'employee_count_fte':
                    label = '0年度目の対象従業員数（確認済みの正社員換算人数）'
                ws.cell(mapping.shinsei[key], 2, label)
        ws.cell(300, 2, extra_text)
        wb.save(path)
        wb.close()

    def test_legacy_rule_text_is_rejected_but_wage_amounts_are_not(self):
        stale = self.dir / 'stale.xlsx'
        self.synthetic(stale, '令和7年7月の事業場内最低賃金+63円以上')
        issues = application_template_issues(stale, 7, '通常枠_2026')
        self.assertTrue(issues and '旧回' in issues[0], issues)
        fine = self.dir / 'fine.xlsx'
        self.synthetic(fine, '群馬/1063円')
        self.assertEqual(application_template_issues(fine, 7, '通常枠_2026'), [])

    def test_bundled_templates_still_selected_by_round(self):
        for kind in APPLICATION_FILENAMES:
            with self.subTest(kind=kind):
                self.assertEqual(select_application_template(ROOT, kind, 7).parent.name, '第7次以降')
                self.assertEqual(select_application_template(ROOT, kind, 6).parent.name, 'ツール')


class KatenDriverRoundTests(unittest.TestCase):
    """#8 判定だけの実行でも公募回を必須にし、2026年11月以降の申請を旧回で判定しない。"""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def legacy_ledger(self, application_month: datetime.datetime) -> Path:
        path = self.dir / f'台帳_{application_month:%Y%m}.xlsx'
        shutil.copyfile(LEGACY_LEDGER_TEMPLATE, path)
        wb = openpyxl.load_workbook(path)
        ws = wb['加点判定用明細']
        ws['C2'] = '東京都'
        ws['C3'] = application_month
        for i, (name, kind, hours, base, latest) in enumerate(
                [('合成甲', 'パート', 100, 120000, 127000), ('合成乙', '正社員', 173.33, 320000, 330000)]):
            row = 7 + i
            ws.cell(row, 2, i + 1)
            ws.cell(row, 3, name)
            ws.cell(row, 4, kind)
            ws.cell(row, 5, hours)
            for col in range(6, 18):
                ws.cell(row, col, base)
            ws.cell(row, 18, latest)
        wb.save(path)
        wb.close()
        return path

    def run_driver(self, *args):
        return subprocess.run([sys.executable, '-B', str(DRIVER), *map(str, args)], capture_output=True,
                              text=True, encoding='utf-8', errors='replace', cwd=self.dir)

    def test_call_round_is_required_even_for_judgment_only(self):
        result = self.run_driver(self.legacy_ledger(datetime.datetime(2026, 8, 1)))
        self.assertEqual(result.returncode, EXIT_INPUT_ERROR, result.stdout)
        self.assertIn('--call-round', result.stderr)
        self.assertNotIn('加点判定レポート', result.stdout)

    def test_november_application_is_not_judged_with_legacy_round(self):
        result = self.run_driver(self.legacy_ledger(datetime.datetime(2026, 11, 1)), '--call-round', 6)
        self.assertEqual(result.returncode, EXIT_INPUT_ERROR, result.stdout)
        self.assertIn('第7次以降', result.stderr)
        self.assertNotIn('判定: 対象', result.stdout)

    def test_legacy_correction_before_november_still_runs(self):
        result = self.run_driver(self.legacy_ledger(datetime.datetime(2026, 8, 1)), '--call-round', 6)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn('差: +70円', result.stdout)

    def test_docstring_examples_name_round_7(self):
        text = DRIVER.read_text(encoding='utf-8')
        doc = text.split('"""')[1]
        examples = [line for line in doc.splitlines() if 'run_katen_judge.py 台帳.xlsx' in line]
        self.assertTrue(examples)
        self.assertTrue(all('--call-round 7' in line for line in examples), examples)


class RunHelpTests(unittest.TestCase):
    """#47 run.py --help の例は local サブコマンド＋--application-context。"""

    def test_help_examples_use_local_and_application_context(self):
        for args in (['--help'], ['local', '--help']):
            with self.subTest(args=args):
                result = subprocess.run([sys.executable, '-B', str(ROOT / 'run.py'), *args],
                                        capture_output=True, text=True, encoding='utf-8',
                                        errors='replace', cwd=ROOT)
                self.assertEqual(result.returncode, 0, result.stderr)
                examples = [line.strip() for line in result.stdout.splitlines()
                            if line.strip().startswith('python')]
                self.assertTrue(examples)
                self.assertTrue(all(' run.py local ' in line for line in examples), examples)
                self.assertTrue(any('--application-context' in line for line in examples))
                self.assertNotIn('--sheet-id "1ABC', result.stdout)
                self.assertIn('停止', result.stdout)


if __name__ == '__main__':
    unittest.main()
