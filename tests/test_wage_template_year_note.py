# -*- coding: utf-8 -*-
"""標準の賃金台帳テンプレートが、第7次以降の通常枠で月見出しを年付きにすることを案内しているかの回帰テスト。

テンプレートの見出しは「1月」〜「12月」のまま。案内がないと、記入ルールどおりに手で埋めた台帳でも
第7次の通常枠でツールが止まり、原因がテンプレートの中から分からない。
ツール側とスキル側のテンプレートは scripts/patch_wage_template_absent_marker.py が同じバイトに保つ。
"""
import sys
import unittest
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.stdout.reconfigure(encoding='utf-8')
warnings.filterwarnings('ignore', category=UserWarning)

import openpyxl

TOOL_TEMPLATE = ROOT / 'ツール' / '賃金台帳テンプレート.xlsx'
SKILL_TEMPLATE = ROOT / '.claude' / 'skills' / 'wagebook-convert' / 'templates' / '賃金台帳テンプレート.xlsx'
NOTE_CELLS = (('従業員別明細', 'B4'), ('記入ルール', 'B8'))
YEAR_NOTE_WORDS = ('第7次以降の通常枠', '年付き', '令和7年4月', '注記', 'ツールが止ま')


class WageTemplateYearNoteTests(unittest.TestCase):
    def test_both_copies_are_byte_identical(self):
        self.assertEqual(TOOL_TEMPLATE.read_bytes(), SKILL_TEMPLATE.read_bytes())

    def test_notes_ask_for_year_stamped_month_headers(self):
        for path in (TOOL_TEMPLATE, SKILL_TEMPLATE):
            wb = openpyxl.load_workbook(path)
            try:
                for sheet, cell in NOTE_CELLS:
                    text = wb[sheet][cell].value or ''
                    for word in YEAR_NOTE_WORDS:
                        with self.subTest(file=path.parent.name, cell=f'{sheet}!{cell}', word=word):
                            self.assertIn(word, text)
            finally:
                wb.close()


if __name__ == '__main__':
    unittest.main()
