"""第7次の画面前提・原本選択・未確認停止を、API不使用で検証する。"""
from __future__ import annotations

import ast
import hashlib
import sys
import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

from hojokin.application_context import ApplicationContext, validate_application_context
from hojokin.config import get_mapping, detect_prefecture
from hojokin.template_filler import fill_template, verify_template_labels
from hojokin.models import ExtractionResult, FinancialData
from hojokin.template_versions import (APPLICATION_FILENAMES, select_application_template,
                                      validate_application_template_version)
from hojokin.result_identity import processing_identity


class Round7IntegrationTests(unittest.TestCase):
    def test_changed_application_conditions_invalidate_old_result(self):
        original = {'company': '合成', 'call_round': 7, 'application_ym': (2026, 11),
                    'wage_total': 10000000, 'files': ['content-hash-1'],
                    'uploaded_template': ('form.xlsx', 'template-hash-1')}
        first = processing_identity(**original)
        self.assertEqual(first, processing_identity(**original))
        for key, value in [('company', '別合成'), ('call_round', 8),
                           ('application_ym', (2026, 12)), ('wage_total', 10350000),
                           ('files', ['content-hash-2']),
                           ('uploaded_template', ('form.xlsx', 'template-hash-2'))]:
            changed = dict(original, **{key: value})
            self.assertNotEqual(first, processing_identity(**changed), key)

        # The actual screen must include the uploaded template, not just the helper API.
        tree = ast.parse((ROOT/'app.py').read_text(encoding='utf-8'))
        assignment = next(node for node in tree.body if isinstance(node, ast.Assign)
                          and any(isinstance(target, ast.Name)
                                  and target.id == '_current_result_identity' for target in node.targets))
        self.assertIn('uploaded_template', {kw.arg for kw in assignment.value.keywords})

    def test_all_four_templates_are_selected_by_rule_and_keep_objects(self):
        for kind, name in APPLICATION_FILENAMES.items():
            with self.subTest(kind=kind):
                old = select_application_template(ROOT, kind, 6)
                new = select_application_template(ROOT, kind, 7)
                self.assertIsNotNone(old)
                self.assertIsNotNone(new)
                self.assertNotEqual(old, new)
                verify_template_labels(old, get_mapping(kind), kind, call_round=6)
                verify_template_labels(new, get_mapping(kind), kind, call_round=7)
                with self.assertRaises(ValueError):
                    validate_application_template_version(old, 7)
                with self.assertRaises(ValueError):
                    validate_application_template_version(new, 6)
                with zipfile.ZipFile(old) as a, zipfile.ZipFile(new) as b:
                    self.assertEqual(a.namelist(), b.namelist())
                    for part in a.namelist():
                        if part.startswith(('xl/media/', 'xl/drawings/', 'xl/charts/')):
                            self.assertEqual(a.read(part), b.read(part), part)
                        if part.endswith('.rels'):
                            self.assertNotIn(b'it2026_chingin_houkoku', b.read(part), part)

    def test_official_forms_are_unmodified_downloaded_bytes(self):
        expected = {
            '補助率引き上げ・加点措置①用.xlsx': 'd8cbf155ee2f152ad85aa12633fb630b1a7997a646ab45fdfdaf81e96278fd5e',
            '加点措置①用.xlsx': '9ee638cc843f4a78c3921a00930420ffc06949b42754540c425560c1b8483adf',
            '加点措置②用.xlsx': '707ef3970615d2150a3bd66c80ab4ba95b00515e5aa931dc4d8c219acaa01244',
        }
        for name, digest in expected.items():
            self.assertEqual(hashlib.sha256((ROOT/'補助金加点/第7次以降'/name).read_bytes()).hexdigest(), digest)

    def test_real_normal_template_records_official_next_fiscal_year_plan_periods(self):
        source = select_application_template(ROOT, '通常枠_2026', 7)
        cases = [
            ('2026-04-01', '2027-03-31', '2025-04-01', '2026-03-31',
             ['2027-04-01〜2028-03-31', '2028-04-01〜2029-03-31', '2029-04-01〜2030-03-31']),
            ('2026-01-01', '2026-12-31', '2025-01-01', '2025-12-31',
             ['2027-01-01〜2027-12-31', '2028-01-01〜2028-12-31', '2029-01-01〜2029-12-31']),
        ]
        with tempfile.TemporaryDirectory() as td:
            for start, end, prior_start, prior_end, expected in cases:
                with self.subTest(fiscal_start=start):
                    c = ApplicationContext.from_dict({
                        'call_round': 7, 'application_date': '2026-10-15',
                        'fiscal_year_start': start, 'fiscal_year_end': end,
                        'frame': '通常', 'subsidy_amount': 1_000_000, 'past_grants': [],
                        'optional_wage_bonus': False, 'wage_exempt': False,
                        'payroll_scope_confirmed': True,
                        'workplace_prefecture': '東京都', 'workplace_confirmed': True,
                        'regional_min_wage_yen': 1280,
                        # 率義務なしでも確認済みの0〜3計画をすべて入力する。
                        'forecast_years': [dict(total_salary=10_000_000, employee_count_fte=10,
                                                source='確認済み合成給与計画', confirmed=True) for _ in range(4)],
                        'productivity_forecast_years': [dict(operating_profit=2_000_000, personnel_cost=4_000_000,
                            depreciation=0, total_annual_hours=1000, source='確認済み合成事業計画', confirmed=True)
                            for _ in range(3)],
                    })
                    fi = FinancialData(fiscal_year_start=prior_start, fiscal_year_end=prior_end,
                                       operating_profit=1_000_000, salary=4_000_000)
                    output = Path(td) / f'{start}.xlsx'
                    fill_template(source, output, get_mapping('通常枠_2026'), {},
                                  ExtractionResult(financial=fi), application_context=c,
                                  productivity_total_annual_hours=1000)
                    wb = openpyxl.load_workbook(output)
                    rows = list(wb['申請前提・給与見込'].values)
                    salary = {row[0]: row[1] for row in rows if type(row[0]) is int}
                    productivity = {int(row[0].split('年度目')[0]): row[0].split('\n')[1]
                                    for row in rows if isinstance(row[0], str)
                                    and row[0].startswith(('1年度目\n', '2年度目\n', '3年度目\n'))}
                    for year, period in enumerate(expected, 1):
                        self.assertEqual(salary[year], period)
                        self.assertEqual(productivity[year], period)
                    self.assertEqual(salary[0], f'{start}〜{end}')
                    self.assertIn(('生産性基準年度（前事業年度）', f'{prior_start}〜{prior_end}', None, None, None, None), rows)
                    ws = wb['申請内容']
                    mapping = get_mapping('通常枠_2026')
                    self.assertEqual(ws.cell(mapping.shinsei['wage_total_base'], 3).value, 10_000_000)
                    self.assertIsNone(ws.cell(mapping.shinsei['wage_raise_declaration'], 3).value)
                    wb.close()

    def test_application_preflight_stops_before_extractor(self):
        tree = ast.parse((ROOT/'app.py').read_text(encoding='utf-8'))
        function = next(x for x in tree.body if isinstance(x, ast.FunctionDef) and x.name == 'run_processing')
        api = Mock(side_effect=AssertionError('未確認の申請でExtractorを作成してはいけない'))
        scope = {'Path': Path, 'validate_application_context': validate_application_context,
                 'create_extractor': api}
        exec(compile(ast.Module(body=[function], type_ignores=[]), 'app.py', 'exec'), scope)
        result = scope['run_processing']('合成テスト', '通常枠_2026', 'application', ROOT, ROOT,
                                         call_round=7, application_context=None)
        self.assertEqual(result['application']['status'], '要確認')
        api.assert_not_called()

    def test_form_selection_does_not_fall_back_to_legacy(self):
        import tempfile
        import shutil
        with tempfile.TemporaryDirectory() as d:
            work = Path(d)
            name = APPLICATION_FILENAMES['通常枠_2026']
            shutil.copyfile(ROOT/'ツール'/name, work/name)
            self.assertIsNone(select_application_template(work, '通常枠_2026', 7))

    def test_explicit_workplace_conflict_stops_before_bonus_judgment(self):
        tree = ast.parse((ROOT/'app.py').read_text(encoding='utf-8'))
        function = next(x for x in tree.body if isinstance(x, ast.FunctionDef)
                        and x.name == '_run_bonus_judgment')
        judge = Mock(side_effect=AssertionError('勤務地矛盾のまま判定してはいけない'))
        scope = {'Path': Path, '_find_bonus_ledger': lambda *args: ROOT/'synthetic.xlsx',
                 'read_bonus_wage_ledger': lambda *args, **kwargs: SimpleNamespace(prefecture='東京都'),
                 'detect_prefecture': detect_prefecture, 'judge_bonus_points': judge}
        exec(compile(ast.Module(body=[function], type_ignores=[]), 'app.py', 'exec'), scope)
        result = scope['_run_bonus_judgment'](ROOT, '合成テスト', '大阪府', ROOT,
                                               application_ym=(2026, 12), call_round=7)
        self.assertEqual(result['status'], '要確認')
        judge.assert_not_called()

    def test_reimport_refuses_original_or_existing_candidate(self):
        import tempfile
        script = ROOT/'scripts/patch_genpon_tsujo_kojin_from_drive.py'
        with tempfile.TemporaryDirectory() as d:
            original = Path(d)/'original.xlsx'
            candidate = Path(d)/'candidate.xlsx'
            original.write_bytes(b'original untouched')
            candidate.write_bytes(b'candidate untouched')
            for target in (original, candidate):
                result = subprocess.run([sys.executable, '-B', str(script), '--input',
                                         str(original), '--output', str(target)],
                                        capture_output=True, text=True, encoding='utf-8')
                self.assertNotEqual(result.returncode, 0)
                self.assertEqual(original.read_bytes(), b'original untouched')
                self.assertEqual(candidate.read_bytes(), b'candidate untouched')


if __name__ == '__main__':
    unittest.main(verbosity=2)
