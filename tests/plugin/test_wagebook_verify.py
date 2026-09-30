"""wagebook-convert（配布版）の検算スクリプト2本を、合成の台帳ブックで動かして確かめる。

verify_transcript.py（原本転記シートの受入検査）:
  正常 → 0 / 通勤手当が課税支給額に含まれているのに S列が空（控除漏れ）→ 1 / 転記シートなし → 2
verify_wagebook.py（従業員別明細の検算）:
  通勤手当欄は常に「要確認」になる設計なので正常でも 1。FAIL（役員ラベル未正規化）が出るかで見分ける。
  シートが無ければ 2。
"""
import shutil
import tempfile
import warnings
from pathlib import Path

import openpyxl

from _common import Checker, run_py, skills_dir

warnings.filterwarnings('ignore', category=UserWarning)

SKILL = skills_dir() / 'wagebook-convert'
TEMPLATE = SKILL / 'templates' / '賃金台帳テンプレート.xlsx'
VERIFY_BOOK = SKILL / 'review' / 'verify_wagebook.py'
VERIFY_TRANSCRIPT = SKILL / 'review' / 'verify_transcript.py'
LEDGER_SHEET = '従業員別明細'
DATA_START_ROW, COL_NO, COL_NAME, COL_TYPE, COL_HOURS, COL_HOURLY = 6, 2, 3, 4, 5, 6
COL_MONTH_FIRST, COL_S = 7, 19
MONTHLY_COMMUTE = 10000
# (氏名, 雇用形態, 月平均時間, 時給, 月額, S列=年間通勤手当)
EMPLOYEES = [
    ('甲野花子', '正社員', 173.33, None, 210000, MONTHLY_COMMUTE * 12),
    ('乙野太郎', 'パート', 80, 1100, 88000, None),
]
# 原本転記（月次ページ型: 行=項目 × 列=従業員）。甲野は通勤手当が課税支給額に含まれている
TRANSCRIPT_ROWS = [
    ('基本給', 200000, 88000),
    ('通勤手当', MONTHLY_COMMUTE, None),
    ('課税支給額', 210000, 88000),
    ('出勤日数', 20, 16),
    ('労働時間', 160, 80),
]


def make_book(path: Path, *, s_col=True, transcript=True, officer_label=None) -> None:
    shutil.copy(TEMPLATE, path)
    wb = openpyxl.load_workbook(path)
    ws = wb[LEDGER_SHEET]
    for i, (name, kind, hours, hourly, monthly, s) in enumerate(EMPLOYEES):
        r = DATA_START_ROW + i
        ws.cell(r, COL_NO, i + 1)
        ws.cell(r, COL_NAME, name)
        ws.cell(r, COL_TYPE, officer_label if (officer_label and i == 0) else kind)
        ws.cell(r, COL_HOURS, hours)
        if hourly:
            ws.cell(r, COL_HOURLY, hourly)
        for m in range(12):
            ws.cell(r, COL_MONTH_FIRST + m, monthly)
        if s_col and s:
            ws.cell(r, COL_S, s)
    if transcript:
        t = wb.create_sheet('原本転記_p1')
        for c, head in enumerate(['項目', EMPLOYEES[0][0], EMPLOYEES[1][0], '合計'], 1):
            t.cell(1, c, head)
        for r, (label, a, b) in enumerate(TRANSCRIPT_ROWS, 2):
            t.cell(r, 1, label)
            if a is not None:
                t.cell(r, 2, a)
            if b is not None:
                t.cell(r, 3, b)
            t.cell(r, 4, (a or 0) + (b or 0))
    wb.save(path)


def main() -> int:
    c = Checker('wagebook-convert（配布版）の検算スクリプト')
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        good, missing_s, no_tr, officer = (work / f'{n}.xlsx' for n in ('good', 'missing_s', 'no_tr', 'officer'))
        make_book(good)
        make_book(missing_s, s_col=False)
        make_book(no_tr, transcript=False)
        make_book(officer, officer_label='取締役')

        r = run_py(VERIFY_TRANSCRIPT, good, cwd=work)
        c.check('転記シート: 正常なら終了コード0', r.returncode == 0, r.stdout[-900:])
        r = run_py(VERIFY_TRANSCRIPT, missing_s, cwd=work)
        c.check('転記シート: 通勤手当の控除漏れ（S列が空）で終了コード1', r.returncode == 1 and '控除漏れ' in r.stdout,
                r.stdout[-900:])
        r = run_py(VERIFY_TRANSCRIPT, no_tr, cwd=work)
        c.check('転記シート: 転記シートが無ければ終了コード2', r.returncode == 2, r.stderr)

        r = run_py(VERIFY_BOOK, good, cwd=work)
        c.check('明細: 正常なら FAIL なし（要確認だけで終了コード1）',
                r.returncode == 1 and '修正必須 FAIL' not in r.stdout, r.stdout[-900:])
        r = run_py(VERIFY_BOOK, officer, cwd=work)
        c.check('明細: 役員ラベル未正規化は FAIL', r.returncode == 1 and '役員ラベル未正規化' in r.stdout,
                r.stdout[-900:])
        wb = openpyxl.Workbook()
        wb.save(work / 'empty.xlsx')
        r = run_py(VERIFY_BOOK, work / 'empty.xlsx', cwd=work)
        c.check('明細: 従業員別明細シートが無ければ終了コード2', r.returncode == 2, r.stderr)
    return c.finish()


if __name__ == '__main__':
    raise SystemExit(main())
