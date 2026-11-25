# -*- coding: utf-8 -*-
"""第7次の Drive 版（2026-11-09）の原本・ヒアリングの行の並びと、旧版ヒアリングとの読み替えを固定する。

Drive 版はヒアリングの「◼️表明を行った従業員名」の上に昇給時期・今期中の退職予定者・今期賞与見込みの欄が増え、
そこから下の行がずれた。ツールは第7次の対応表（get_mapping(種類, 7)）で転記し、顧客が旧版に書いた
ヒアリングでも見出しの位置で読み替える。必要書類チェック（check_docs.py）も同じ見分け方をする。
合成データだけで API は呼ばない。
"""
import importlib.util
import os
import shutil
import sys
import tempfile
import unicodedata
import unittest
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding='utf-8')
os.environ['CLAUDE_API_KEY'] = ''
warnings.filterwarnings('ignore', category=UserWarning)

import openpyxl

from hojokin.config import get_mapping
from hojokin.hearing_reader import (HEARING_DECLARER_LABELS, align_hearing_rows, read_hearing_sheet,
                                    transfer_hearing_to_tenki)
from hojokin.template_versions import APPLICATION_FILENAMES

ROUND7_DIR = ROOT / 'ツール' / '第7次以降'
LEGACY_DIR = ROOT / 'ツール'
CHECK_DOCS = ROOT / '.claude' / 'skills' / 'case-docs-check' / 'review' / 'check_docs.py'
HEARINGS = {
    '通常枠_2026': 'ヒアリングシート2026_通常枠法人.xlsx',
    '通常枠_個人_2026': 'ヒアリングシート2026_通常枠個人.xlsx',
    'インボイス枠_2026': 'ヒアリングシート2026_インボイス法人.xlsx',
    'インボイス枠_個人_2026': 'ヒアリングシート2026_インボイス個人.xlsx',
}
CHECK_FORMS = {'通常枠_2026': '通常枠法人', '通常枠_個人_2026': '通常枠個人',
               'インボイス枠_2026': 'インボイス法人', 'インボイス枠_個人_2026': 'インボイス個人'}
ADDED_LABELS = ('昇給時期', '今期中の退職予定者の有無', 'ありの場合：該当者名', '今期賞与見込み', '増減ありの場合')
DECLARER_NAMES = ('代表 花子', '給与 太郎', '最低 次郎')
LABEL_PREFIX_LEN = 4
ROUND7, LEGACY = 7, 6


def find(directory: Path, name: str) -> Path:
    return next(p for p in directory.iterdir()
                if unicodedata.normalize('NFC', p.name) == unicodedata.normalize('NFC', name))


def norm(value) -> str:
    return unicodedata.normalize('NFKC', str(value or '')).replace(' ', '').replace('\n', '').lstrip('◼️■')


def labels(path: Path, sheet: str, col: int) -> dict[int, str]:
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    out = {c.row: c.value for row in ws.iter_rows(min_col=col, max_col=col) for c in row if c.value}
    wb.close()
    return out


def load_check_docs():
    spec = importlib.util.spec_from_file_location('check_docs_under_test', CHECK_DOCS)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def filled_hearing(source: Path, dest: Path) -> Path:
    """ヒアリングの写しに、表明者3人と増えた欄の回答を書き込む（顧客記入の代わり）。"""
    shutil.copyfile(source, dest)
    wb = openpyxl.load_workbook(dest)
    ws = wb['基本情報']
    for row in range(1, ws.max_row + 1):
        label = norm(ws.cell(row, 2).value)
        for want, name in zip(HEARING_DECLARER_LABELS, DECLARER_NAMES):
            if label == want:
                ws.cell(row, 3).value = name
        if label in {norm(a) for a in ADDED_LABELS}:
            ws.cell(row, 3).value = f'回答:{label}'
    wb.save(dest)
    wb.close()
    return dest


class Round7MappingMatchesDriveLayout(unittest.TestCase):
    def test_hearing_and_tenki_labels_line_up_in_both_rounds(self):
        for call_round, directory in ((ROUND7, ROUND7_DIR), (LEGACY, LEGACY_DIR)):
            for kind, hearing in HEARINGS.items():
                with self.subTest(round=call_round, kind=kind):
                    mapping = get_mapping(kind, call_round)
                    h = labels(find(directory, hearing), '基本情報', 2)
                    t = labels(find(directory, APPLICATION_FILENAMES[kind]), '転記', 1)
                    for h_row, t_row, _ in mapping.hearing_to_tenki:
                        self.assertEqual(norm(h.get(h_row))[:LABEL_PREFIX_LEN], norm(t.get(t_row))[:LABEL_PREFIX_LEN],
                                         f'ヒアリング{h_row} → 転記{t_row}')
                    self.assertIn('表明を行った従業員名', norm(h.get(mapping.hearing_anchor_row)))
                    self.assertIn('賃上げ幅', norm(h.get(mapping.hearing_wage_raise_row)))

    def test_added_rows_reach_application_sheet_through_tenki(self):
        for kind, name in APPLICATION_FILENAMES.items():
            with self.subTest(kind=kind):
                mapping = get_mapping(kind, ROUND7)
                h = labels(find(ROUND7_DIR, HEARINGS[kind]), '基本情報', 2)
                wb = openpyxl.load_workbook(ROUND7_DIR / name)
                refs = {str(c.value) for row in wb['申請内容'].iter_rows(min_col=3, max_col=3) for c in row
                        if isinstance(c.value, str) and c.value.startswith("='転記'!B")}
                wb.close()
                wanted = {label: next(t for hr, t, _ in mapping.hearing_to_tenki if norm(h.get(hr)) == norm(label))
                          for label in ADDED_LABELS + HEARING_DECLARER_LABELS}
                for label, t_row in wanted.items():
                    self.assertIn(f"='転記'!B{t_row}", refs, label)

    def test_round7_mapping_rows_hit_the_same_labels_as_legacy(self):
        for kind, name in APPLICATION_FILENAMES.items():
            with self.subTest(kind=kind):
                old = labels(LEGACY_DIR / name, '申請内容', 2)
                new = labels(ROUND7_DIR / name, '申請内容', 2)
                old_rows, new_rows = get_mapping(kind, LEGACY).shinsei, get_mapping(kind, ROUND7).shinsei
                relabeled = {'employee_count_fte', 'wage_total_base', 'wage_total_y1', 'wage_total_y2', 'wage_total_y3'}
                for field, row in old_rows.items():
                    if field in relabeled or not old.get(row):
                        continue
                    self.assertEqual(norm(old[row])[:8], norm(new.get(new_rows[field]))[:8], field)


class AlignHearingRowsTests(unittest.TestCase):
    def record(self, label, value=None):
        return {'label': label, 'value': value}

    def sheet(self, anchor: int, added: bool) -> dict:
        data = {10: self.record('企業名', 'A社')}
        if added:
            data[anchor - 8] = self.record('昇給時期', '4月')
        data[anchor] = self.record('◼️表明を行った従業員名')
        for offset, (label, name) in enumerate(zip(HEARING_DECLARER_LABELS, DECLARER_NAMES), start=1):
            data[anchor + offset] = self.record(label, name)
        data[anchor + 6] = self.record('インボイス登録', '済み')
        return data

    def test_old_layout_is_moved_down_to_round7_rows(self):
        aligned, note = align_hearing_rows(self.sheet(83, added=False), 92)
        self.assertIsNone(note)
        self.assertEqual(aligned[93]['value'], DECLARER_NAMES[0])
        self.assertEqual(aligned[98]['value'], '済み')
        self.assertEqual(aligned[10]['value'], 'A社')

    def test_round7_layout_is_moved_up_for_legacy_mapping_and_added_rows_are_dropped(self):
        aligned, note = align_hearing_rows(self.sheet(92, added=True), 83)
        self.assertIsNone(note)
        self.assertEqual(aligned[84]['value'], DECLARER_NAMES[0])
        self.assertNotIn('4月', [rec['value'] for rec in aligned.values()])

    def test_same_layout_is_unchanged(self):
        data = self.sheet(92, added=True)
        aligned, note = align_hearing_rows(data, 92)
        self.assertIsNone(note)
        self.assertEqual(aligned, data)

    def test_missing_or_broken_anchor_drops_rows_below_instead_of_misreading(self):
        data = self.sheet(83, added=False)
        del data[83]
        aligned, note = align_hearing_rows(data, 92)
        self.assertIn('見つかりません', note)
        self.assertEqual(sorted(aligned), [10, 84, 85, 86, 89])
        broken = self.sheet(83, added=False)
        broken[84] = self.record('昇給時期', '4月')
        aligned, note = align_hearing_rows(broken, 92)
        self.assertIn('並びになっていない', note)
        self.assertTrue(all(row < 92 for row in aligned))


class HearingTransferEndToEnd(unittest.TestCase):
    def test_round7_and_legacy_hearings_fill_declarers_into_round7_tenki(self):
        with tempfile.TemporaryDirectory() as td:
            for kind, hearing in HEARINGS.items():
                mapping = get_mapping(kind, ROUND7)
                for label, directory in (('第7次の版', ROUND7_DIR), ('旧版', LEGACY_DIR)):
                    with self.subTest(kind=kind, hearing=label):
                        src = filled_hearing(find(directory, hearing), Path(td) / f'{kind}_{label}.xlsx')
                        data, note = align_hearing_rows(read_hearing_sheet(src), mapping.hearing_anchor_row)
                        self.assertIsNone(note)
                        wb = openpyxl.load_workbook(ROUND7_DIR / APPLICATION_FILENAMES[kind])
                        ws = wb['転記']
                        transfer_hearing_to_tenki(data, ws, mapping.hearing_to_tenki)
                        got = {norm(ws.cell(t, 1).value): ws.cell(t, 2).value for _, t, _ in mapping.hearing_to_tenki}
                        wb.close()
                        for want, name in zip(HEARING_DECLARER_LABELS, DECLARER_NAMES):
                            self.assertEqual(got[want], name)
                        if label == '第7次の版':
                            self.assertEqual(got['昇給時期'], '回答:昇給時期')
                            self.assertEqual(got['今期賞与見込み'], '回答:今期賞与見込み')
                        else:
                            self.assertIsNone(got['昇給時期'])


class CheckDocsReadsBothLayouts(unittest.TestCase):
    def setUp(self):
        self.cd = load_check_docs()

    def run_check(self, path: Path, kind: str):
        rep = self.cd.Report()
        frame = '通常枠' if kind.startswith('通常') else 'インボイス枠'
        entity = '個人' if '個人' in kind else '法人'
        found = self.cd._check_hearing(rep, [{'name': path.name, 'local_path': path}], frame, entity)
        return found, rep

    def test_both_layouts_are_identified_and_old_layout_is_flagged(self):
        with tempfile.TemporaryDirectory() as td:
            for kind, hearing in HEARINGS.items():
                for label, directory in (('第7次の版', ROUND7_DIR), ('旧版', LEGACY_DIR)):
                    with self.subTest(kind=kind, hearing=label):
                        path = filled_hearing(find(directory, hearing), Path(td) / f'{kind}_{label}.xlsx')
                        found, rep = self.run_check(path, kind)
                        codes = [f['code'] for f in rep.findings]
                        self.assertEqual(found, CHECK_FORMS[kind])
                        self.assertNotIn('HEARING-FORM-MISMATCH', codes)
                        self.assertNotIn('HEARING-LAYOUT-UNKNOWN', codes)
                        self.assertEqual('HEARING-OLD-LAYOUT' in codes, label == '旧版')
                        empty = ' '.join(f['message'] for f in rep.findings if f['code'] == 'HEARING-EMPTY-REQUIRED')
                        for want in HEARING_DECLARER_LABELS:
                            self.assertNotIn(want, empty)
                        if label == '第7次の版':
                            self.assertNotIn('昇給時期', empty)


if __name__ == '__main__':
    unittest.main(verbosity=2)
