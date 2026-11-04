"""第7次の画面前提・原本選択・未確認停止を、API不使用で検証する。"""
from __future__ import annotations

import ast
import hashlib
import sys
import subprocess
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

from hojokin.application_context import validate_application_context
from hojokin.config import get_mapping, detect_prefecture
from hojokin.template_filler import verify_template_labels
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
