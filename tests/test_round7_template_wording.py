"""第7次以降の申請書原本・加点台帳テンプレートの文言を固定する（API不使用）。

根拠: 第7回〜公募要領（通常枠 印刷p.9〜10・23〜25、インボイス枠 印刷p.8〜9・22）、
交付申請マニュアル11/2版 p.86〜88、公式「賃金状況報告シート」第7次公募回以降用。
2026-11-25〜: 申請書原本は原本管理担当の Drive 版（2026-11-09）が土台で、ヒアリングシートは Drive 版を
そのまま写す（文言は原本管理担当が決める）。ヒアリングの文言はここで固定せず、行の並びだけを
tests/test_round7_drive_layout.py で確かめる。
"""
from __future__ import annotations

import re
import sys
import unittest
import zipfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding='utf-8')

from hojokin.config import get_mapping
from hojokin.template_versions import APPLICATION_FILENAMES, ROUND7_FIRST_CALL as ROUND7, ROUND7_SUBJECT

ROUND7_DIR = ROOT / 'ツール' / '第7次以降'
LEGACY_DIR = ROOT / 'ツール'
NORMAL_KINDS = ('通常枠_2026', '通常枠_個人_2026')
INVOICE_KINDS = ('インボイス枠_2026', 'インボイス枠_個人_2026')
LEDGER = '加点判定用賃金台帳テンプレート.xlsx'
PRODUCTIVITY = '生産性指標給与支給総額計算'
PLAN_COLUMNS = ('C', 'D', 'E')
PLAN_INPUT_ROWS = (34, 35, 36, 38, 39)
PLAN_FORMULA_ROWS = range(34, 43)
STALE_PATTERNS = (r'R[7７][.．][7７]', r'R[7７]年度改定', r'(?<![0-9０-９,，])[6６][3３]円',
                  r'令和[6６]年\s*[1１][0０]月', r'令和[7７]年度改定', r'適用補助率確認シート', r'以内で雇用',
                  r'任意の賃上げ計画加点）で雇用', r'登記から抜粋', r'リンクされているので触らない',
                  r'特に指定がなければ', r'シート作成日', r'支給なしは空欄',
                  r'前事業年度分の賃金台帳をご提出', r'3\.0％以上であることを確認')


def round7_files() -> list[Path]:
    """こちらで作る第7次の原本（申請書4種類と加点台帳）。ヒアリングは Drive 版の写しなので含めない。"""
    return [ROUND7_DIR / name for name in APPLICATION_FILENAMES.values()] + [ROUND7_DIR / LEDGER]


def load(path: Path) -> openpyxl.Workbook:
    return openpyxl.load_workbook(path, data_only=False)


def formulas(ws) -> dict[str, str]:
    return {c.coordinate: c.value for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.startswith('=')}


class Round7TemplateWordingTests(unittest.TestCase):
    def test_no_legacy_round_or_legacy_operation_wording_remains(self):
        for path in round7_files():
            wb = load(path)
            hits = [f'{ws.title}!{c.coordinate}' for ws in wb for row in ws.iter_rows() for c in row
                    if isinstance(c.value, str) and not c.value.startswith('=')
                    and any(re.search(p, c.value) for p in STALE_PATTERNS)]
            wb.close()
            self.assertEqual(hits, [], path.name)

    def test_bonus_two_and_bonus_one_questions_use_round7_months(self):
        for kind, name in APPLICATION_FILENAMES.items():
            wb = load(ROUND7_DIR / name)
            labels = [c.value for c in wb['申請内容']['B'] if isinstance(c.value, str)]
            wb.close()
            self.assertTrue(any('「R8.6事業場内最低賃金」+55円以上' in t for t in labels), kind)
            if kind in NORMAL_KINDS:
                self.assertTrue(any('R8年度改定の地域別最低賃金未満' in t and '補助率の引き上げ' in t
                                    for t in labels), kind)

    def test_declaration_has_no_default_and_asks_confirmed_facts(self):
        for kind, name in APPLICATION_FILENAMES.items():
            rows = get_mapping(kind, ROUND7).shinsei
            wb = load(ROUND7_DIR / name)
            ws = wb['申請内容']
            for field in ('wage_raise_declaration', 'wage_raise_method'):
                self.assertNotIn('■', ws.cell(rows[field], 3).value, kind)
            no_employee = next(c.row for c in ws['B'] if isinstance(c.value, str)
                               and c.value.startswith('賃金引上げ計画期間内において、従業員を雇用する場合'))
            self.assertNotIn('■', ws.cell(no_employee, 3).value, kind)
            self.assertIn('顧客に確認した実施方法', ws.cell(rows['wage_raise_method'], 4).value, kind)
            self.assertIn('顧客に確認した表明の実施日', ws.cell(rows['wage_raise_date'], 4).value, kind)
            wb.close()

    def test_wage_growth_guide_covers_one_year_and_annual_by_category(self):
        for kind, name in APPLICATION_FILENAMES.items():
            rows = get_mapping(kind, ROUND7).shinsei
            wb = load(ROUND7_DIR / name)
            ws = wb['申請内容']
            guide = ws.cell(rows['wage_total_y3'] + 1, 2).value
            base_note = ws.cell(rows['wage_total_base'], 4).value
            wb.close()
            for phrase in ('クリック', '「1年後」と「年平均」の両方', '3.5%以上', '必須ではない',
                           'IT導入補助金2022〜2025', '年平均3%以上'):
                self.assertIn(phrase, guide, kind)
            if kind in NORMAL_KINDS:
                self.assertIn('150万円以上：3%以上', guide)
            else:
                self.assertNotIn('150万円', guide)
            self.assertIn('確認済みの0年度目見込みを入力（自動リンクなし）', base_note, kind)
            self.assertIn('交付申請日の属する事業年度', base_note, kind)

    def test_it_support_center_is_written_next_to_it_navi_for_normal_frame_only(self):
        phrases = ('IT経営サポートセンター', '2026年8月1日以降', '2026年11月30日', '加点は1つ')
        for kind in NORMAL_KINDS + INVOICE_KINDS:
            wb = load(ROUND7_DIR / APPLICATION_FILENAMES[kind])
            text = '\n'.join(str(c.value) for row in wb['申請内容'].iter_rows(min_col=2, max_col=4)
                             for c in row if isinstance(c.value, str))
            wb.close()
            for phrase in phrases:
                (self.assertIn if kind in NORMAL_KINDS else self.assertNotIn)(phrase, text, kind)

    def test_productivity_plan_is_confirmed_input_not_auto_forecast(self):
        for kind in NORMAL_KINDS:
            wb = load(ROUND7_DIR / APPLICATION_FILENAMES[kind])
            ws = wb[PRODUCTIVITY]
            for col in PLAN_COLUMNS:
                for row in PLAN_INPUT_ROWS:
                    self.assertIsNone(ws[f'{col}{row}'].value, f'{kind} {col}{row}')
            plan = {ref: f for ref, f in formulas(ws).items()
                    if ref[0] in PLAN_COLUMNS and int(ref[1:]) in PLAN_FORMULA_ROWS}
            self.assertFalse([f for f in plan.values() if re.search(r'F3[49]|1\.035|1\.05', f)], kind)
            self.assertIn('C41/$B$41', ws['C42'].value)
            self.assertIn('(D41/$B$41)^(1/2)', ws['D42'].value)
            self.assertIn('(E41/$B$41)^(1/3)', ws['E42'].value)
            self.assertIn('3%以上', ws['B42'].value)
            self.assertIn('4%以上', ws['B42'].value)
            self.assertTrue(ws['B41'].value.startswith('='), kind)
            wb.close()

    def test_bonus_ledger_rule_marks_absent_months_and_blank_as_unconfirmed(self):
        wb = load(ROUND7_DIR / '加点判定用賃金台帳テンプレート.xlsx')
        rule = wb['記入ルール']['B7'].value
        self.assertIn('「非在籍」と記入', rule)
        self.assertIn('未確認', rule)
        self.assertIn('令和7年10月〜令和8年9月', wb['記入ルール']['A7'].value)
        self.assertIn('空欄・0円は未確認', wb['加点判定用明細']['B4'].value)
        wb.close()

    def test_min_wage_note_and_visible_version_label(self):
        for kind, name in APPLICATION_FILENAMES.items():
            wb = load(ROUND7_DIR / name)
            ws = wb['申請内容']
            self.assertIn('主たる事業場（確認済み）', ws.cell(get_mapping(kind, ROUND7).shinsei['min_wage'], 4).value)
            self.assertIn('第7次以降用', ws['D2'].value)
            self.assertIn('第7次以降用', wb['転記']['A1'].value)
            self.assertNotIn('登記', ws.cell(get_mapping(kind, ROUND7).shinsei['min_wage'], 4).value)
            wb.close()

    def test_round7_marker_and_package_members_are_kept(self):
        for path in round7_files():
            with zipfile.ZipFile(path) as b:
                names = b.namelist()
                self.assertIn('docProps/core.xml', names, path.name)
                self.assertIn(b'/docProps/core.xml', b.read('[Content_Types].xml'), path.name)
                self.assertIn(b'metadata/core-properties', b.read('_rels/.rels'), path.name)
            wb = openpyxl.load_workbook(path, read_only=True)
            self.assertEqual(wb.properties.subject, ROUND7_SUBJECT, path.name)
            wb.close()
        # 加点台帳だけは旧回の台帳が土台。図形・書式・ブック定義は旧回と同じバイトのまま
        with zipfile.ZipFile(LEGACY_DIR / LEDGER) as a, zipfile.ZipFile(ROUND7_DIR / LEDGER) as b:
            self.assertEqual(a.namelist(), b.namelist())
            for part in a.namelist():
                if part.startswith(('xl/media/', 'xl/drawings/', 'xl/styles', 'xl/workbook.xml')):
                    self.assertEqual(a.read(part), b.read(part), part)


if __name__ == '__main__':
    unittest.main(verbosity=2)
