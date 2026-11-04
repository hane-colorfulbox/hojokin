"""第7次の申請前提・給与見込・申告事実の安全ゲート（合成データ、API不使用）。"""
import copy
import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

import openpyxl

from hojokin.application_context import (
    ApplicationContext, ApplicationValidationError, WageForecastYear,
    build_application_wage_plan, propose_constant_fte_plan,
    validate_application_context, validate_productivity_inputs, wage_requirements,
)
from hojokin.config import get_mapping, SHINSEI_ANCHOR_LABELS
from hojokin.models import ExtractionResult, FinancialData, FieldConfidence, ProcessingStatus
from hojokin.template_filler import fill_template, fill_shinsei_sheet
from hojokin.template_versions import ROUND7_SUBJECT
from hojokin.template_versions import APPLICATION_FILENAMES


def values(**changes):
    v = {
        'call_round': 7, 'application_date': '2026-11-10',
        'fiscal_year_start': '2026-04-01', 'fiscal_year_end': '2027-03-31',
        'frame': '通常', 'subsidy_amount': 1_500_000, 'past_grants': [],
        'optional_wage_bonus': False, 'wage_exempt': False,
        'payroll_scope_confirmed': True,
        'declaration_confirmed': True, 'declaration_method': '書面',
        'declaration_date': '2026-11-01', 'planned_min_wage_increase': 30,
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


def context(**changes):
    return ApplicationContext.from_dict(values(**changes))


def actual_financial(**changes):
    fields = dict(fiscal_year_start='2025-04-01', fiscal_year_end='2026-03-31',
                  operating_profit=1_000_000, salary=4_000_000)
    fields.update(changes)
    return FinancialData(**fields)


class ApplicationContextTests(unittest.TestCase):
    def test_context_required(self):
        with self.assertRaises(ApplicationValidationError):
            validate_application_context(None)

    def test_unselected_wage_requirements_are_not_mandatory(self):
        c = context(subsidy_amount=1_499_999, declaration_confirmed=None,
                    declaration_method='', declaration_date=None,
                    planned_min_wage_increase=None)
        validate_application_context(c, '通常枠_2026')
        self.assertFalse(wage_requirements(c).selected)
        self.assertEqual(build_application_wage_plan(c)['wage_total_base'], 10_000_000)

    def test_round7_plan_inputs_are_required_without_growth_obligation(self):
        for frame in ('通常', 'インボイス', 'セキュリティ'):
            with self.subTest(frame=frame), self.assertRaisesRegex(ApplicationValidationError, '4年分'):
                validate_application_context(context(frame=frame, subsidy_amount=1_499_999,
                    forecast_years=[], payroll_scope_confirmed=False, declaration_confirmed=None,
                    declaration_method='', declaration_date=None, planned_min_wage_increase=None))
        # 旧制度の非選択案件の正常な計算条件は維持する。
        c = context(call_round=6, subsidy_amount=1_499_999, forecast_years=[],
                    payroll_scope_confirmed=False, declaration_confirmed=None,
                    declaration_method='', declaration_date=None, planned_min_wage_increase=None)
        validate_application_context(c)
        self.assertIsNone(build_application_wage_plan(c))

    def test_no_growth_obligation_does_not_require_positive_growth_or_declaration(self):
        v = values(subsidy_amount=1_499_999, declaration_confirmed=None,
                   declaration_method='', declaration_date=None, planned_min_wage_increase=None)
        for year, amount in zip(v['forecast_years'], (10_000_000, 9_000_000, 8_000_000, 7_000_000)):
            year['total_salary'] = amount
        c = ApplicationContext.from_dict(v)
        validate_application_context(c)
        self.assertFalse(wage_requirements(c).selected)
        self.assertEqual(build_application_wage_plan(c)['wage_total_y3'], 7_000_000)

    def test_150_man_boundary(self):
        c = context()
        validate_application_context(c, '通常枠_2026')
        req = wage_requirements(c)
        self.assertTrue(req.mandatory)
        self.assertTrue(req.first_year_required)
        self.assertEqual(req.required_rate, Decimal('0.03'))
        self.assertEqual(req.min_wage_increase, 30)

    def test_zero_or_unknown_subsidy_amount_cannot_select_smaller_band(self):
        for amount in (None, 0, -1, float('nan')):
            with self.subTest(amount=amount), self.assertRaises(ApplicationValidationError):
                validate_application_context(context(subsidy_amount=amount))

    def test_grant_years_and_frames_are_independent(self):
        c = context(past_grants=[{'year': 2022, 'frame': '通常枠', 'decided': True}])
        self.assertEqual(wage_requirements(c).required_rate, Decimal('0.035'))
        self.assertEqual(wage_requirements(c).productivity_rate, Decimal('0.03'))
        c = context(past_grants=[{'year': 2024, 'frame': '通常枠', 'decided': True}])
        self.assertEqual(wage_requirements(c).productivity_rate, Decimal('0.04'))
        c = context(past_grants=[{'year': 2023, 'frame': 'インボイス枠', 'decided': True}])
        self.assertEqual(wage_requirements(c).productivity_rate, Decimal('0.03'))
        c = context(past_grants=[{'year': 2023, 'frame': '通常枠A類型', 'decided': True}])
        self.assertEqual(wage_requirements(c).productivity_rate, Decimal('0.04'))
        c = context(past_grants=[{'year': 2023, 'frame': '通常枠', 'decided': True}])
        self.assertEqual(wage_requirements(c).productivity_rate, Decimal('0.04'))

    def test_ambiguous_prior_grant_frame_is_not_silently_three_percent(self):
        for frame in ('デジタル化基盤導入枠', 'foo', 'fooインボイス', '通常枠foo'):
            with self.subTest(frame=frame), self.assertRaisesRegex(ApplicationValidationError, '枠・類型が不明'):
                validate_application_context(context(past_grants=[
                    {'year': 2023, 'frame': frame, 'decided': True}
                ]))

    def test_unknown_grants_or_declaration_cannot_export(self):
        for updates in ({'past_grants': None}, {'declaration_confirmed': None},
                        {'declaration_confirmed': False}, {'declaration_method': ''},
                        {'declaration_date': '2026-11-11'}):
            with self.subTest(updates=updates), self.assertRaises(ApplicationValidationError):
                validate_application_context(context(**updates))

    def test_literal_booleans_required(self):
        for updates in ({'optional_wage_bonus': 'false'}, {'wage_exempt': 'false'},
                        {'past_grants': [{'year': 2024, 'frame': '通常枠', 'decided': 'false'}]}):
            with self.subTest(updates=updates), self.assertRaises(ApplicationValidationError):
                validate_application_context(context(**updates))

    def test_large_normal_optional_bonus_requires_50_yen(self):
        with self.assertRaises(ApplicationValidationError):
            validate_application_context(context(optional_wage_bonus=True, planned_min_wage_increase=30))
        validate_application_context(context(optional_wage_bonus=True, planned_min_wage_increase=50))

    def test_optional_basic_bonus_30_and_further_bonus_50_have_correct_frame_scope(self):
        for frame, amount in (('通常', 1_499_999), ('インボイス', 1_499_999),
                              ('インボイス', 1_500_000), ('セキュリティ', 1_500_000)):
            for increase in (30, 49, 50):
                with self.subTest(frame=frame, amount=amount, increase=increase):
                    c = context(frame=frame, subsidy_amount=amount,
                                optional_wage_bonus=True, planned_min_wage_increase=increase)
                    validate_application_context(c)
                    self.assertEqual(wage_requirements(c).min_wage_increase, 30)
                    self.assertEqual(wage_requirements(c).required_rate, Decimal('0.03'))
            with self.subTest(frame=frame), self.assertRaisesRegex(ApplicationValidationError, '30円'):
                validate_application_context(context(frame=frame, subsidy_amount=amount,
                    optional_wage_bonus=True, planned_min_wage_increase=29))
        # 150万円の境界を越えた通常枠では基本加点+30円を誤って採用しない。
        for increase in (30, 49):
            with self.subTest(increase=increase), self.assertRaisesRegex(ApplicationValidationError, '50円'):
                validate_application_context(context(subsidy_amount=1_500_000,
                    optional_wage_bonus=True, planned_min_wage_increase=increase))

    def test_optional_basic_bonus_growth_and_prior_grant_conditions_remain_independent(self):
        v = values(subsidy_amount=1_499_999, optional_wage_bonus=True, planned_min_wage_increase=30)
        v['forecast_years'][1]['total_salary'] = 10_000_000
        c = ApplicationContext.from_dict(v)
        validate_application_context(c)  # 任意のみは1年後要件を追加せず、3年平均を検証。
        self.assertFalse(wage_requirements(c).first_year_required)
        v['forecast_years'][3]['total_salary'] = 10_500_000
        with self.assertRaisesRegex(ApplicationValidationError, '3年間'):
            validate_application_context(ApplicationContext.from_dict(v))
        v = values(subsidy_amount=1_499_999, optional_wage_bonus=True, planned_min_wage_increase=30,
                   past_grants=[{'year': 2024, 'frame': '通常枠', 'decided': True}])
        for year, amount in zip(v['forecast_years'], (10_000_000, 10_400_000, 10_900_000, 11_500_000)):
            year['total_salary'] = amount
        c = ApplicationContext.from_dict(v)
        validate_application_context(c)
        self.assertEqual(wage_requirements(c).required_rate, Decimal('0.035'))
        self.assertTrue(wage_requirements(c).first_year_required)
        v['forecast_years'][1]['total_salary'] = 10_300_000
        with self.assertRaisesRegex(ApplicationValidationError, '1年後'):
            validate_application_context(ApplicationContext.from_dict(v))

    def test_forecast_not_inferred_from_prior_actuals(self):
        with self.assertRaises(ApplicationValidationError):
            build_application_wage_plan(context(forecast_years=[]))

    def test_fte_increase_checks_per_capita(self):
        v = values()
        for yr in v['forecast_years'][1:]:
            yr['employee_count_fte'] = 11
        with self.assertRaisesRegex(ApplicationValidationError, '1年後'):
            validate_application_context(ApplicationContext.from_dict(v))

    def test_explicit_changing_fte_plan_can_pass(self):
        v = values()
        for yr, amount, fte in zip(v['forecast_years'], (10000, 11330, 11670, 12021), (10, 11, 11, 11)):
            yr.update(total_salary=amount, employee_count_fte=fte)
        plan = build_application_wage_plan(ApplicationContext.from_dict(v))
        self.assertEqual(plan['employee_count_fte_by_year'], [10, 11, 11, 11])

    def test_third_year_growth_independent_of_first_year(self):
        v = values()
        v['forecast_years'][3]['total_salary'] = 10_500_000
        with self.assertRaisesRegex(ApplicationValidationError, '3年間'):
            validate_application_context(ApplicationContext.from_dict(v))

    def test_forecast_confirmation_source_and_numbers_required(self):
        for field, bad in (('total_salary', float('nan')), ('total_salary', True),
                           ('total_salary', 1.5), ('employee_count_fte', 0),
                           ('employee_count_fte', float('inf')), ('confirmed', False),
                           ('source', '')):
            v = values()
            v['forecast_years'][0][field] = bad
            with self.subTest(field=field, bad=bad), self.assertRaises(ApplicationValidationError):
                validate_application_context(ApplicationContext.from_dict(v))

    def test_period_and_frame_mismatch(self):
        for update in ({'fiscal_year_end': '2026-12-31'},
                       {'application_date': '2027-04-01'}, {'frame': 'インボイス枠（電子取引類型）'}):
            with self.subTest(update=update), self.assertRaises(ApplicationValidationError):
                validate_application_context(context(**update))
        with self.assertRaises(ApplicationValidationError):
            validate_application_context(context(), 'インボイス枠_2026')

    def test_workplace_is_not_inferred(self):
        with self.assertRaises(ApplicationValidationError):
            validate_application_context(context(workplace_confirmed=False))
        with self.assertRaises(ApplicationValidationError):
            validate_application_context(context(regional_min_wage_yen=1226))
        with self.assertRaises(ApplicationValidationError):
            validate_application_context(context(regional_min_wage_yen=1230))
        with self.assertRaises(ApplicationValidationError):
            validate_application_context(context(workplace_prefecture='未確認県'))

    def test_future_report_basis_is_not_invented(self):
        plan = build_application_wage_plan(context())
        self.assertEqual(plan['basis_kind'], '当期見込み')
        self.assertNotIn('total_annual_hours', plan)
        self.assertNotIn('effect_report_basis', plan)

    def test_constant_fte_proposal_requires_confirmation_and_stays_draft(self):
        base = WageForecastYear(143, 1, '確認済みの合成例', True)
        with self.assertRaises(ApplicationValidationError):
            propose_constant_fte_plan(base, Decimal('0.035'))
        years = propose_constant_fte_plan(base, Decimal('0.035'), constant_fte_confirmed=True)
        self.assertEqual(years[1].total_salary, 149)  # 148.005を148へ丸めない
        self.assertFalse(years[1].confirmed)

    def test_productivity_period_and_hours_are_separate(self):
        c = context()
        good = actual_financial()
        validate_productivity_inputs(c, good, 8000)
        for data, hours in ((good, None), (good, 0),
                            (FinancialData(fiscal_year_start='2026-04-01', fiscal_year_end='2027-03-31'), 8000)):
            with self.assertRaises(ApplicationValidationError):
                validate_productivity_inputs(c, data, hours)

    def test_productivity_forecasts_require_confirmation_and_finite_values(self):
        for field, bad in (('operating_profit', float('nan')), ('personnel_cost', float('inf')),
                           ('depreciation', -1), ('total_annual_hours', 0),
                           ('total_annual_hours', True), ('source', ''), ('confirmed', False)):
            v = values()
            v['productivity_forecast_years'][1][field] = bad
            with self.subTest(field=field, bad=bad), self.assertRaisesRegex(ApplicationValidationError, '生産性2年度目'):
                validate_application_context(ApplicationContext.from_dict(v))
        with self.assertRaisesRegex(ApplicationValidationError, '3年分'):
            validate_application_context(context(productivity_forecast_years=[]))

    def test_productivity_forecasts_are_not_mandatory_for_invoice_or_legacy(self):
        for changes in ({'frame': 'インボイス'}, {'call_round': 6}):
            with self.subTest(changes=changes):
                c = context(productivity_forecast_years=[], **changes)
                validate_application_context(c)
                self.assertIsNone(validate_productivity_inputs(c, FinancialData(), None))

    def test_productivity_first_and_three_year_average_are_independent(self):
        fi = actual_financial(operating_profit=100, salary=900)
        for past, rate, amounts in (([], '0.03', (1030, 1060.9, 1092.727)),
                                   ([{'year': 2024, 'frame': '通常枠', 'decided': True}],
                                    '0.04', (1040, 1081.6, 1124.864))):
            v = values(past_grants=past)
            # 給与の3.5%要件は独立して満たす。生産性の予算から給与を導かない。
            for year, amount in zip(v['forecast_years'], (10_000_000, 10_400_000, 10_900_000, 11_500_000)):
                year['total_salary'] = amount
            for year, amount in zip(v['productivity_forecast_years'], amounts):
                year.update(operating_profit=0, personnel_cost=amount, depreciation=0,
                            total_annual_hours=10)
            with self.subTest(rate=rate):
                result = validate_productivity_inputs(ApplicationContext.from_dict(v), fi, 10)
                self.assertAlmostEqual(result['required_rate'], float(rate))
                self.assertAlmostEqual(result['first_year_growth'], float(rate))
                self.assertAlmostEqual(result['third_year_average'], float(rate))
                first_low = copy.deepcopy(v)
                first_low['productivity_forecast_years'][0]['personnel_cost'] -= 1
                with self.assertRaisesRegex(ApplicationValidationError, '1年後'):
                    validate_productivity_inputs(ApplicationContext.from_dict(first_low), fi, 10)
                third_low = copy.deepcopy(v)
                third_low['productivity_forecast_years'][2]['personnel_cost'] -= 1
                with self.assertRaisesRegex(ApplicationValidationError, '3年間の年平均'):
                    validate_productivity_inputs(ApplicationContext.from_dict(third_low), fi, 10)

    def test_productivity_nonpositive_or_unconfirmed_baseline_cannot_pass(self):
        for fi in (actual_financial(operating_profit=-4_000_000),
                   actual_financial(operating_profit=-5_000_000),
                   actual_financial(salary=float('nan')),
                   actual_financial(confidence={'salary': FieldConfidence(level='low')})):
            with self.subTest(fi=fi), self.assertRaises(ApplicationValidationError):
                validate_productivity_inputs(context(), fi, 1000)

    def test_productivity_does_not_use_salary_fte_or_current_salary_totals(self):
        fi = actual_financial(operating_profit=100, salary=900)
        v = values()
        result = validate_productivity_inputs(ApplicationContext.from_dict(v), fi, 10)
        for year in v['forecast_years']:
            year['employee_count_fte'] *= 7
            year['total_salary'] *= 11
        result2 = validate_productivity_inputs(ApplicationContext.from_dict(v), fi, 10)
        self.assertEqual(result, result2)

    def test_missing_productivity_forecasts_stop_before_api(self):
        from hojokin.pipeline import run_application_transfer
        with tempfile.TemporaryDirectory() as td:
            with patch('hojokin.pipeline.create_extractor', side_effect=AssertionError('must not create extractor')):
                result = run_application_transfer(Path(td), Path(td) / 'missing.xlsx', '通常枠_2026',
                                                  Path(td) / 'out.xlsx',
                                                  application_context=context(productivity_forecast_years=[]))
            self.assertEqual(result.status, 'エラー')
            self.assertIn('生産性', result.message)

    def test_productivity_hours_do_not_ignore_another_employee_or_missing_month(self):
        from hojokin.pipeline import _validated_productivity_annual_hours
        window = [(2025, m) for m in range(4, 13)] + [(2026, m) for m in range(1, 4)]
        good = SimpleNamespace(name='合成A', employment_type='正社員', monthly_wages=[100_000] * 12,
                               monthly_hours=[160] * 12, monthly_avg_hours=160, fiscal_window=window)
        missing_employee = copy.deepcopy(good)
        missing_employee.name = '合成B'
        missing_employee.monthly_hours = [None] * 12
        missing_employee.monthly_avg_hours = 0
        with self.assertRaisesRegex(ApplicationValidationError, '合成B'):
            _validated_productivity_annual_hours([good, missing_employee], context().productivity_period)
        missing_month = copy.deepcopy(good)
        missing_month.monthly_hours[5] = None
        # 正の月平均があっても、部分的な欠損を予測で埋めない。
        with self.assertRaisesRegex(ApplicationValidationError, '6月'):
            _validated_productivity_annual_hours([missing_month], context().productivity_period)
        self.assertEqual(_validated_productivity_annual_hours([good], context().productivity_period), 1920)

    def test_productivity_hours_require_explicit_absence_and_preserve_raw_calculation(self):
        from hojokin.pipeline import _validated_productivity_annual_hours
        from hojokin.wage_calculator import wage_employees_to_payroll
        from hojokin.wage_reader import WageEmployee
        wages = [100_000] * 12
        hours = [160] * 12
        wages[0], hours[0] = None, None
        window = [(2025, m) for m in range(4, 13)] + [(2026, m) for m in range(1, 4)]
        period = context().productivity_period
        emp = WageEmployee(1, '合成短期', '正社員', 160, 0, wages,
                           monthly_hours=hours, fiscal_window=window)
        raw_before = wage_employees_to_payroll([emp])[1]
        self.assertEqual(raw_before, 1760)
        with self.assertRaisesRegex(ApplicationValidationError, '1月'):
            _validated_productivity_annual_hours([emp], period)
        emp.monthly_presence = {0: False}
        self.assertEqual(_validated_productivity_annual_hours([emp], period), 1760)
        self.assertEqual(wage_employees_to_payroll([emp])[1], raw_before)
        # 0円だけでは非在籍とみなさない。
        emp.monthly_presence = {}
        emp.monthly_wages[0] = 0
        with self.assertRaisesRegex(ApplicationValidationError, '1月'):
            _validated_productivity_annual_hours([emp], period)

    def test_confirmed_fixed_hours_require_every_months_presence(self):
        from hojokin.pipeline import _validated_productivity_annual_hours
        window = [(2025, m) for m in range(4, 13)] + [(2026, m) for m in range(1, 4)]
        period = context().productivity_period
        emp = SimpleNamespace(name='合成所定', employment_type='正社員', monthly_wages=[100_000] * 12,
                              monthly_hours=[None] * 12, monthly_avg_hours=160, fiscal_window=window)
        self.assertEqual(_validated_productivity_annual_hours([emp], period), 1920)
        emp.monthly_wages[0] = None
        with self.assertRaisesRegex(ApplicationValidationError, '1月'):
            _validated_productivity_annual_hours([emp], period)

    def test_productivity_complete_hours_still_require_confirmed_prior_year(self):
        from hojokin.pipeline import _validated_productivity_annual_hours
        period = context().productivity_period
        emp = SimpleNamespace(name='合成年次', employment_type='正社員', monthly_wages=[100_000] * 12,
                              monthly_hours=[160] * 12, monthly_avg_hours=160, fiscal_window=None)
        with self.assertRaisesRegex(ApplicationValidationError, '対象年度を確認できません'):
            _validated_productivity_annual_hours([emp], period)
        emp.fiscal_window = [(2024, m) for m in range(4, 13)] + [(2025, m) for m in range(1, 4)]
        with self.assertRaisesRegex(ApplicationValidationError, '対象年月が前事業年度と一致しません'):
            _validated_productivity_annual_hours([emp], period)
        emp.fiscal_window = [(2025, m) for m in range(4, 13)] + [(2026, m) for m in range(1, 4)]
        self.assertEqual(_validated_productivity_annual_hours([emp], period), 1920)

    def test_missing_context_stops_before_extractor_and_keeps_existing_output(self):
        from hojokin.pipeline import run_application_transfer
        with tempfile.TemporaryDirectory() as td:
            output = Path(td) / 'existing.xlsx'
            output.write_bytes(b'keep previous data')
            with patch('hojokin.pipeline.create_extractor', side_effect=AssertionError('must not create extractor')):
                result = run_application_transfer(Path(td), Path(td) / 'missing.xlsx',
                                                  '通常枠_2026', output)
            self.assertEqual(result.status, 'エラー')
            self.assertIn('案件設定', result.message)
            self.assertEqual(output.read_bytes(), b'keep previous data')


class TemplateOutputTests(unittest.TestCase):
    def template(self, path, round7=True):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '申請内容'
        mapping = get_mapping('通常枠_2026')
        for key, expected in SHINSEI_ANCHOR_LABELS.items():
            if key in mapping.shinsei:
                if round7 and key == 'employee_count_fte':
                    expected = '0年度目の対象従業員数（確認済みの正社員換算人数）'
                ws.cell(mapping.shinsei[key], 2, expected[0] if isinstance(expected, tuple) else expected)
        for field in ('past_subsidies', 'wage_raise_declaration', 'wage_raise_method'):
            ws.cell(mapping.shinsei[field], 3, '危険な原本初期値')
        for field in ('employee_count_fte', 'wage_total_base', 'wage_total_y1', 'wage_total_y2', 'wage_total_y3'):
            ws.cell(mapping.shinsei[field], 3, '=999999')
        wb.create_sheet('転記')
        prod = wb.create_sheet('生産性指標給与支給総額計算')
        prod['B40'] = '=B38*B39'
        prod['C40'] = '=C38*C39'
        if round7:
            wb.properties.subject = ROUND7_SUBJECT
        wb.save(path)
        wb.close()
        return mapping

    def extraction(self):
        return ExtractionResult(financial=actual_financial())

    def test_confirmed_forecast_output_and_prior_hours(self):
        with tempfile.TemporaryDirectory() as td:
            source, out = Path(td) / 'template.xlsx', Path(td) / 'out.xlsx'
            mapping = self.template(source)
            fill_template(source, out, mapping, {}, self.extraction(),
                          wage_plan={'wage_total_base': 1, 'total_annual_hours': 1},
                          application_context=context(), productivity_total_annual_hours=8765.5)
            wb = openpyxl.load_workbook(out, data_only=False)
            ws = wb['申請内容']
            self.assertEqual(ws.cell(mapping.shinsei['wage_total_base'], 3).value, 10_000_000)
            self.assertEqual(ws.cell(mapping.shinsei['past_subsidies'], 3).value, 'なし')
            self.assertEqual(ws.cell(mapping.shinsei['wage_raise_method'], 3).value, '書面')
            self.assertEqual(wb['生産性指標給与支給総額計算']['B40'].value, 8765.5)
            self.assertEqual(wb['生産性指標給与支給総額計算']['C40'].value, 1000)
            self.assertIsNone(wb['生産性指標給与支給総額計算']['C38'].value)
            self.assertIsNone(wb['生産性指標給与支給総額計算']['C39'].value)
            self.assertEqual(wb['生産性指標給与支給総額計算']['C34'].value, 2_000_000)
            self.assertEqual(wb['生産性指標給与支給総額計算']['E35'].value, 17_000_000)
            self.assertIn('検算済み', wb['生産性指標給与支給総額計算']['B42'].value)
            self.assertGreater(wb['生産性指標給与支給総額計算']['C42'].value, .03)
            self.assertIn('申請前提・給与見込', wb.sheetnames)
            text = '\n'.join(str(c.value or '') for row in wb['申請前提・給与見込'] for c in row)
            self.assertIn('効果報告', text)
            self.assertIn('未確定', text)
            self.assertIn('担当者確認の合成見込表', text)
            self.assertIn('担当者確認の合成生産性予算', text)
            self.assertIn('1年度目\n2027-04-01〜2028-03-31', text)
            wb.close()

    def test_salary_and_productivity_plans_share_next_fiscal_year_periods(self):
        cases = [
            ('3月決算', '2026-04-01', '2027-03-31', '2025-04-01', '2026-03-31',
             ['2027-04-01〜2028-03-31', '2028-04-01〜2029-03-31', '2029-04-01〜2030-03-31']),
            ('12月決算', '2026-01-01', '2026-12-31', '2025-01-01', '2025-12-31',
             ['2027-01-01〜2027-12-31', '2028-01-01〜2028-12-31', '2029-01-01〜2029-12-31']),
        ]
        with tempfile.TemporaryDirectory() as td:
            source = Path(td) / 'template.xlsx'
            mapping = self.template(source)
            for name, start, end, prior_start, prior_end, expected in cases:
                with self.subTest(name=name):
                    out = Path(td) / f'{name}.xlsx'
                    c = context(application_date='2026-10-15', declaration_date='2026-10-01',
                                fiscal_year_start=start, fiscal_year_end=end)
                    fi = actual_financial(fiscal_year_start=prior_start, fiscal_year_end=prior_end)
                    fill_template(source, out, mapping, {}, ExtractionResult(financial=fi),
                                  application_context=c, productivity_total_annual_hours=1000)
                    wb = openpyxl.load_workbook(out, data_only=False)
                    rows = list(wb['申請前提・給与見込'].values)
                    salary_periods = {row[0]: row[1] for row in rows if type(row[0]) is int}
                    productivity_periods = {
                        int(row[0].split('年度目')[0]): row[0].split('\n')[1]
                        for row in rows if isinstance(row[0], str)
                        and row[0].startswith(('1年度目\n', '2年度目\n', '3年度目\n'))}
                    for idx, period in enumerate(expected, 1):
                        self.assertEqual(salary_periods[idx], period)
                        self.assertEqual(productivity_periods[idx], period)
                    self.assertEqual(salary_periods[0], f'{start}〜{end}')
                    self.assertIn(('生産性基準年度（前事業年度）', f'{prior_start}〜{prior_end}', None, None, None, None), rows)
                    self.assertNotEqual(productivity_periods[1], salary_periods[0])
                    wb.close()

    def test_missing_context_or_mismatched_template_does_not_create_file(self):
        with tempfile.TemporaryDirectory() as td:
            source, out = Path(td) / 'template.xlsx', Path(td) / 'out.xlsx'
            mapping = self.template(source, round7=False)
            with self.assertRaises(ApplicationValidationError):
                fill_template(source, out, mapping, {}, self.extraction())
            self.assertFalse(out.exists())
            with self.assertRaises(ValueError):
                fill_template(source, out, mapping, {}, self.extraction(), application_context=context(),
                              productivity_total_annual_hours=1234)
            self.assertFalse(out.exists())

    def test_missing_productivity_hours_cannot_fallback_to_forecast(self):
        with tempfile.TemporaryDirectory() as td:
            source, out = Path(td) / 'template.xlsx', Path(td) / 'out.xlsx'
            mapping = self.template(source)
            with self.assertRaises(ApplicationValidationError):
                fill_template(source, out, mapping, {}, self.extraction(),
                              wage_plan={'total_annual_hours': 1234}, application_context=context())
            self.assertFalse(out.exists())

    def test_direct_fill_cannot_bypass_productivity_checks_with_invoice_context(self):
        with tempfile.TemporaryDirectory() as td:
            source, out = Path(td) / 'template.xlsx', Path(td) / 'out.xlsx'
            mapping = self.template(source)
            with self.assertRaisesRegex(ApplicationValidationError, '申請枠'):
                fill_template(source, out, mapping, {}, self.extraction(),
                              application_context=context(frame='インボイス', productivity_forecast_years=[]),
                              productivity_total_annual_hours=1000)
            self.assertFalse(out.exists())

    def test_unselected_growth_still_outputs_confirmed_plan_and_clears_declaration(self):
        with tempfile.TemporaryDirectory() as td:
            source, out = Path(td) / 'template.xlsx', Path(td) / 'out.xlsx'
            mapping = self.template(source)
            c = context(subsidy_amount=1_000_000, declaration_confirmed=None,
                        declaration_method='', declaration_date=None)
            fill_template(source, out, mapping, {}, self.extraction(), application_context=c,
                          productivity_total_annual_hours=1000)
            wb = openpyxl.load_workbook(out)
            self.assertEqual(wb['申請内容'].cell(mapping.shinsei['wage_total_base'], 3).value, 10_000_000)
            self.assertIsNone(wb['申請内容'].cell(mapping.shinsei['wage_raise_declaration'], 3).value)
            wb.close()

    def test_direct_low_level_fill_does_not_keep_false_declarations(self):
        wb = openpyxl.Workbook()
        mapping = get_mapping('通常枠_2026')
        for field in ('past_subsidies', 'wage_raise_declaration', 'wage_raise_method'):
            wb.active.cell(mapping.shinsei[field], 3, '確認していない既定値')
        fill_shinsei_sheet(wb.active, mapping, ExtractionResult())
        for field in ('past_subsidies', 'wage_raise_declaration', 'wage_raise_method'):
            self.assertIsNone(wb.active.cell(mapping.shinsei[field], 3).value)

    def test_real_four_round7_templates(self):
        root = Path(__file__).resolve().parent.parent
        with tempfile.TemporaryDirectory() as td:
            for kind, name in APPLICATION_FILENAMES.items():
                with self.subTest(kind=kind):
                    mapping = get_mapping(kind)
                    settings = values(frame='通常' if kind.startswith('通常') else 'インボイス')
                    if '個人' in kind:
                        settings.update(fiscal_year_start='2026-01-01', fiscal_year_end='2026-12-31')
                        fi = actual_financial(fiscal_year_start='2025-01-01', fiscal_year_end='2025-12-31')
                    else:
                        fi = self.extraction().financial
                    out = Path(td) / (kind + '.xlsx')
                    fill_template(root / 'ツール' / '第7次以降' / name, out, mapping, {},
                                  ExtractionResult(financial=fi),
                                  application_context=ApplicationContext.from_dict(settings),
                                  productivity_total_annual_hours=8000)
                    wb = openpyxl.load_workbook(out)
                    ws = wb['申請内容']
                    self.assertEqual(ws.cell(mapping.shinsei['wage_total_base'], 3).value, 10_000_000)
                    self.assertEqual(ws.cell(mapping.shinsei['min_wage'], 3).value, '東京/1280円')
                    self.assertEqual(ws.cell(mapping.shinsei['past_subsidies'], 3).value, 'なし')
                    choices = [c.value for c in wb['プルダウン用']['B'] if c.value is not None]
                    self.assertIn(ws.cell(mapping.shinsei['min_wage'], 3).value, choices)
                    wb.close()

    def test_attached_historical_sheets_keep_formula_targets_and_original(self):
        from hojokin.pipeline import _attach_per_employee_wage_sheet
        with tempfile.TemporaryDirectory() as td:
            dest = Path(td) / 'application.xlsx'
            wb = openpyxl.Workbook()
            wb.active.title = '申請内容'
            wb.create_sheet('給与支給総額計算')['C1'] = 1234
            wb.save(dest)
            wb.close()

            def fake_calculation(**kwargs):
                source = openpyxl.Workbook()
                source.active.title = '給与支給総額計算'
                source.active['C1'] = "='従業員別明細'!Q7"
                source.create_sheet('従業員別明細')['Q7'] = 77
                source.save(kwargs['output_path'])
                source.close()
                return ProcessingStatus(status='完了')

            with patch('hojokin.pipeline.run_wage_calculation', side_effect=fake_calculation):
                _attach_per_employee_wage_sheet(dest, Path(td), '合成', None, None, [], None,
                                                3, None, historical_reference=True)
            wb = openpyxl.load_workbook(dest)
            self.assertEqual(wb['給与支給総額計算']['C1'].value, 1234)
            self.assertEqual(wb['前事業年度給与実績（参考）']['C1'].value,
                             "='前事業年度従業員明細（参考）'!Q7")
            self.assertEqual(wb['前事業年度従業員明細（参考）']['Q7'].value, 77)
            wb.close()


if __name__ == '__main__':
    unittest.main()
