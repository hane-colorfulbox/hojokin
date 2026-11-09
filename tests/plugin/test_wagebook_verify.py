"""wagebook-convert（配布版）の検算スクリプト2本を、合成の台帳ブックで動かして確かめる。

verify_transcript.py（原本転記シートの受入検査）:
  正常 → 0 / 通勤手当が課税支給額に含まれているのに S列が空（控除漏れ）→ 1 / 転記シートなし → 2
verify_wagebook.py（従業員別明細の検算）:
  通勤手当欄は常に「要確認」になる設計なので正常でも 1。FAIL（役員ラベル未正規化）が出るかで見分ける。
  シートが無ければ 2。月の欄の「非在籍」は正常、数値でも「非在籍」でもない文字列は FAIL。
  非在籍の月の一覧は、年付きの見出しがあればその年月で示す（期首月から並ぶツールの台帳でも列の位置で月名を付けない）。
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
HEADER_ROW, DATA_START_ROW, COL_NO, COL_NAME, COL_TYPE, COL_HOURS, COL_HOURLY = 5, 6, 2, 3, 4, 5, 6
COL_MONTH_FIRST, COL_S = 7, 19
MONTHLY_COMMUTE = 10000
# 3月決算・前事業年度（令和7年4月〜令和8年3月）の年付き見出し
FISCAL_HEADERS = [f'令和7年{m}月' for m in range(4, 13)] + [f'令和8年{m}月' for m in range(1, 4)]
# ツールの台帳作成は G列から期首月の順。wagebook-convert は列の位置を暦月のまま（G=1月）にする
TOOL_HEADERS = FISCAL_HEADERS
SKILL_HEADERS = FISCAL_HEADERS[9:] + FISCAL_HEADERS[:9]
HIRED_IN_JUNE = '令和7年4月,令和7年5月'
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
# 日給制の個人台帳型: 支給欄の先頭が「平日出勤」で基本給の行が無い。勤怠欄にも同じ名前の行がある
DAILY_ROWS = [
    ('【支給】', None, None),
    ('平日出勤', 180000, 72000),
    ('休日出勤', 30000, 16000),
    ('課税支給額', 210000, 88000),
    ('【勤怠】', None, None),
    ('平日出勤(日数)', 20, 9),
    ('休日出勤(日数)', 2, 2),
]
# 原本の行名・列名をそのまま転記した形: 括弧付きの課税分合計、正の数の勤怠控除、非課税通勤費、(小計)と《合計》
BRACKET_ROWS = [
    ('基本給', 200000, 90000),
    ('勤怠控除', None, 2000),
    ('非課税通勤費', MONTHLY_COMMUTE, None),
    ('(課税分合計)', 200000, 88000),
    ('(出勤日数)', 20, 16),
]
BRACKET_HEADS = ['(小計)', '《合計》']
OTHER_PAGE_TOTAL = 123456
# 課税支給額の行が無い（行名を読めない）転記
NO_TAXABLE_ROWS = [('基本給', 200000, 88000), ('支給計', 200000, 88000), ('出勤日数', 20, 16)]
# 読み取り誤り（内訳と課税支給額が合わない）
MISREAD_ROWS = [('基本給', 200000, 88000), ('課税支給額', 210000, 88000), ('出勤日数', 20, 16)]
ORANGE = openpyxl.styles.PatternFill('solid', fgColor='FFC000')   # 原本側の不整合の印
YELLOW = openpyxl.styles.PatternFill('solid', fgColor='FFFF00')   # 読み取りに自信なしの印（免除されない）
PART_MONTHS_JOINED_JULY = [88000] * 6 + [None] * 3 + [88000] * 3
# 【支給】の見出しの次が総支給額だけ（基本給・手当を落とした集計行のみの転記）
SUMMARY_ONLY_ROWS = [('【支給】', None, None), ('総支給額', 210000, 88000), ('課税支給額', 210000, 88000),
                     ('出勤日数', 20, 16)]
# 課税支給額の行名が「課税給与」。甲野の通勤手当は課税支給額に含まれている
TAX_SALARY_ROWS = [('基本給', 200000, 88000), ('通勤手当', MONTHLY_COMMUTE, None),
                   ('課税給与', 210000, 88000), ('出勤日数', 20, 16)]
# 通勤手当が課税支給額の後ろ（非課税の欄）に印字される台帳。課税支給額には入っていない
COMMUTE_AFTER_ROWS = [('基本給', 200000, 88000), ('課税支給額', 200000, 88000),
                      ('通勤手当', MONTHLY_COMMUTE, None), ('総支給額', 210000, 88000), ('出勤日数', 20, 16)]
# 月の途中入社の日割減額を正の数で印字した台帳
PRORATE_ROWS = [('基本給', 200000, 88000), ('日割減額', 50000, None), ('課税支給額', 150000, 88000),
                ('出勤日数', 15, 16)]
# 賞与だけのページ（勤怠は無いのが普通）
BONUS_ROWS = [('賞与額', 300000, 50000), ('課税支給額', 300000, 50000)]
# 単価の行（時間給）が勤怠欄より上にある台帳。支給項目は勤怠より後ろの基本給から数える
RATE_FIRST_ROWS = [('時間給', 1100, 1100), ('出勤日数', 20, 16), ('労働時間', 160, 80), ('基本給', 200000, 88000),
                   ('通勤手当', MONTHLY_COMMUTE, None), ('課税支給額', 210000, 88000)]
# 勤怠欄が支給欄より上にあり、勤怠に「通勤日数」がある台帳。甲野の通勤手当は課税支給額に含まれている
DAYS_FIRST_ROWS = [('通勤日数', 20, 16), ('出勤日数', 20, 16), ('基本給', 200000, 88000),
                   ('通勤手当', MONTHLY_COMMUTE, None), ('課税支給額', 210000, 88000)]


def rename_first_in_ledger(path: Path) -> None:
    """従業員別明細の1人目を別名にする（転記シートの甲野が明細に載っていない状態を作る）。"""
    wb = openpyxl.load_workbook(path)
    wb[LEDGER_SHEET].cell(DATA_START_ROW, COL_NAME, '丙野三郎')
    wb.save(path)


def make_book(path: Path, *, s_col=True, transcript=True, officer_label=None, second_months=None,
              headers=None, part_hours=True, transcript_rows=None, transcript_heads=None,
              mark=None, mark_fill=None, first_type=None) -> None:
    shutil.copy(TEMPLATE, path)
    wb = openpyxl.load_workbook(path)
    ws = wb[LEDGER_SHEET]
    for m, text in enumerate(headers or []):
        ws.cell(HEADER_ROW, COL_MONTH_FIRST + m, text)
    for i, (name, kind, hours, hourly, monthly, s) in enumerate(EMPLOYEES):
        r = DATA_START_ROW + i
        ws.cell(r, COL_NO, i + 1)
        ws.cell(r, COL_NAME, name)
        label = first_type if (i == 1 and first_type) else kind
        ws.cell(r, COL_TYPE, officer_label if (officer_label and i == 0) else label)
        if part_hours or kind != 'パート':
            ws.cell(r, COL_HOURS, hours)
        if hourly:
            ws.cell(r, COL_HOURLY, hourly)
        for m in range(12):
            value = second_months[m] if (second_months and i == 1) else monthly
            ws.cell(r, COL_MONTH_FIRST + m, value)
        if s_col and s:
            ws.cell(r, COL_S, s)
    if transcript:
        t = wb.create_sheet('原本転記_p1')
        extra = transcript_heads or ['合計']
        for c, head in enumerate(['項目', EMPLOYEES[0][0], EMPLOYEES[1][0]] + extra, 1):
            t.cell(1, c, head)
        for r, (label, a, b) in enumerate(transcript_rows or TRANSCRIPT_ROWS, 2):
            t.cell(r, 1, label)
            if a is not None:
                t.cell(r, 2, a)
            if b is not None:
                t.cell(r, 3, b)
            if a is not None or b is not None:
                t.cell(r, 4, (a or 0) + (b or 0))
                if len(extra) > 1:          # 月全体の合計（他ページ分を含むので、このページの和とは合わない）
                    t.cell(r, 5, (a or 0) + (b or 0) + OTHER_PAGE_TOTAL)
            if mark and label == mark:
                t.cell(r, 2).fill = mark_fill or ORANGE
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

        daily = work / 'daily.xlsx'
        make_book(daily, transcript_rows=DAILY_ROWS)
        r = run_py(VERIFY_TRANSCRIPT, daily, cwd=work)
        c.check('転記シート: 日給制（基本給の行なし・【支給】の見出しあり・勤怠と同名の支給行）でも T1・T4 が通る',
                r.returncode == 0, r.stdout[-900:])
        bracket = work / 'bracket.xlsx'
        make_book(bracket, s_col=False, transcript_rows=BRACKET_ROWS, transcript_heads=BRACKET_HEADS)
        r = run_py(VERIFY_TRANSCRIPT, bracket, cwd=work)
        c.check('転記シート: 原本どおりの行名（(課税分合計)・正の数の勤怠控除・(小計)と《合計》）でも通る',
                r.returncode == 0 and '（2列）' in r.stdout, r.stdout[-900:])
        no_taxable = work / 'no_taxable.xlsx'
        make_book(no_taxable, transcript_rows=NO_TAXABLE_ROWS)
        r = run_py(VERIFY_TRANSCRIPT, no_taxable, cwd=work)
        c.check('転記シート: 課税支給額の行が読めず1列も検算できないときは OK にしない',
                r.returncode == 1 and '1列も検算できなかった' in r.stdout, r.stdout[-900:])
        misread = work / 'misread.xlsx'
        make_book(misread, transcript_rows=MISREAD_ROWS)
        r = run_py(VERIFY_TRANSCRIPT, misread, cwd=work)
        c.check('転記シート: 内訳と課税支給額が合わない読み取り誤りは T1 で止まる',
                r.returncode == 1 and 'NG   [T1]' in r.stdout, r.stdout[-900:])
        marked = work / 'marked.xlsx'
        make_book(marked, transcript_rows=MISREAD_ROWS, mark='課税支給額')
        r = run_py(VERIFY_TRANSCRIPT, marked, cwd=work)
        c.check('転記シート: 原本側の不整合として課税支給額をオレンジにしたら別枠の報告になる',
                r.returncode == 0 and 'オレンジの印あり: 1件' in r.stdout, r.stdout[-900:])
        yellow = work / 'yellow.xlsx'
        make_book(yellow, transcript_rows=MISREAD_ROWS, mark='課税支給額', mark_fill=YELLOW)
        r = run_py(VERIFY_TRANSCRIPT, yellow, cwd=work)
        c.check('転記シート: 黄色（読み取りに自信なし）の課税支給額は免除せず T1 で止める',
                r.returncode == 1 and 'NG   [T1]' in r.stdout, r.stdout[-900:])
        summary_only = work / 'summary_only.xlsx'
        make_book(summary_only, transcript_rows=SUMMARY_ONLY_ROWS)
        r = run_py(VERIFY_TRANSCRIPT, summary_only, cwd=work)
        c.check('転記シート: 【支給】の次が総支給額だけの集計行のみの転記は T1 で止める',
                r.returncode == 1 and '集計行のみの転記の疑い' in r.stdout, r.stdout[-900:])
        tax_salary = work / 'tax_salary.xlsx'
        make_book(tax_salary, s_col=False, transcript_rows=TAX_SALARY_ROWS)
        r = run_py(VERIFY_TRANSCRIPT, tax_salary, cwd=work)
        c.check('転記シート: 行名「課税給与」でも課税支給額として読み、通勤手当の控除漏れを止める',
                r.returncode == 1 and '控除漏れ' in r.stdout and 'OK   [T1]' in r.stdout, r.stdout[-900:])
        commute_after = work / 'commute_after.xlsx'
        make_book(commute_after, s_col=False, transcript_rows=COMMUTE_AFTER_ROWS)
        r = run_py(VERIFY_TRANSCRIPT, commute_after, cwd=work)
        c.check('転記シート: 通勤手当が課税支給額の後ろにある台帳は非課税処理済みと判定し、S列空で通る',
                r.returncode == 0 and '非課税処理済み 1名' in r.stdout, r.stdout[-900:])
        prorate = work / 'prorate.xlsx'
        make_book(prorate, transcript_rows=PRORATE_ROWS)
        r = run_py(VERIFY_TRANSCRIPT, prorate, cwd=work)
        c.check('転記シート: 正の数で印字した日割減額は差し引いて成立し、ℹ で列を挙げる',
                r.returncode == 0 and '減額項目（勤怠控除など）を差し引いて成立: 1列' in r.stdout, r.stdout[-900:])
        bonus = work / 'bonus.xlsx'
        make_book(bonus, transcript_rows=BONUS_ROWS)
        r = run_py(VERIFY_TRANSCRIPT, bonus, cwd=work)
        c.check('転記シート: 賞与だけのページは勤怠が無くても T4 で止めない',
                r.returncode == 0 and 'OK   [T4]' in r.stdout, r.stdout[-900:])
        not_listed = work / 'not_listed.xlsx'
        make_book(not_listed, s_col=False)
        rename_first_in_ledger(not_listed)
        r = run_py(VERIFY_TRANSCRIPT, not_listed, cwd=work)
        c.check('転記シート: 通勤手当が課税支給額に含まれるのに明細と突合できない名前は止める（控除漏れを見逃さない）',
                r.returncode == 1 and '明細と突合できない' in r.stdout, r.stdout[-900:])
        not_listed_ex = work / 'not_listed_ex.xlsx'
        make_book(not_listed_ex, s_col=False, transcript_rows=BRACKET_ROWS, transcript_heads=BRACKET_HEADS)
        rename_first_in_ledger(not_listed_ex)
        r = run_py(VERIFY_TRANSCRIPT, not_listed_ex, cwd=work)
        c.check('転記シート: 通勤手当が課税支給額の外の人は、明細に無い名前でも ℹ に留める',
                r.returncode == 0 and '従業員別明細に無い名前' in r.stdout, r.stdout[-900:])
        days_first = work / 'days_first.xlsx'
        make_book(days_first, transcript_rows=DAYS_FIRST_ROWS)
        r = run_py(VERIFY_TRANSCRIPT, days_first, cwd=work)
        c.check('転記シート: 勤怠欄の「通勤日数」を通勤手当と取り違えず、S列が正しければ通る',
                r.returncode == 0 and '課税支給額に含まれる 1名' in r.stdout, r.stdout[-900:])
        rate_first = work / 'rate_first.xlsx'
        make_book(rate_first, transcript_rows=RATE_FIRST_ROWS)
        r = run_py(VERIFY_TRANSCRIPT, rate_first, cwd=work)
        c.check('転記シート: 勤怠欄より上の単価の行（時間給）を支給項目の始まりにしない',
                r.returncode == 0 and '課税支給額に含まれる 1名' in r.stdout, r.stdout[-900:])
        rate_first_missing = work / 'rate_first_missing.xlsx'
        make_book(rate_first_missing, s_col=False, transcript_rows=RATE_FIRST_ROWS)
        r = run_py(VERIFY_TRANSCRIPT, rate_first_missing, cwd=work)
        c.check('転記シート: 同じ台帳で S列が空なら控除漏れで止める',
                r.returncode == 1 and '控除漏れ' in r.stdout, r.stdout[-900:])

        r = run_py(VERIFY_BOOK, good, cwd=work)
        c.check('明細: 正常なら FAIL なし（要確認だけで終了コード1）',
                r.returncode == 1 and '修正必須 FAIL' not in r.stdout, r.stdout[-900:])
        r = run_py(VERIFY_BOOK, officer, cwd=work)
        c.check('明細: 役員ラベル未正規化は FAIL', r.returncode == 1 and '役員ラベル未正規化' in r.stdout,
                r.stdout[-900:])
        part_full = work / 'part_full.xlsx'
        make_book(part_full, part_hours=False)
        r = run_py(VERIFY_BOOK, part_full, cwd=work)
        c.check('明細: 12か月支給のあるパートでE列空は FAIL（R215が過大になる）',
                r.returncode == 1 and '12か月支給のあるパートでE列空 1件' in r.stdout, r.stdout[-900:])
        part_partial = work / 'part_partial.xlsx'
        make_book(part_partial, part_hours=False, second_months=PART_MONTHS_JOINED_JULY)
        r = run_py(VERIFY_BOOK, part_partial, cwd=work)
        c.check('明細: 12か月そろわないパートでE列空は要確認（R215・R216の対象外）',
                r.returncode == 1 and '修正必須 FAIL' not in r.stdout
                and '12か月そろわないパートでE列空 1件' in r.stdout, r.stdout[-900:])
        part_officer = work / 'part_officer.xlsx'
        make_book(part_officer, part_hours=False, first_type='非常勤役員')
        r = run_py(VERIFY_BOOK, part_officer, cwd=work)
        c.check('明細: 「非常勤役員」のE列空は FAIL にしない（ツールは役員として算定から外す）',
                'パートでE列空' not in r.stdout and '修正必須 FAIL' not in r.stdout, r.stdout[-900:])
        absent = work / 'absent.xlsx'
        make_book(absent, second_months=['非在籍'] * 3 + [88000] * 9)
        r = run_py(VERIFY_BOOK, absent, cwd=work)
        c.check('明細: 入社前の月の「非在籍」は FAIL にせず一覧に出す',
                r.returncode == 1 and '修正必須 FAIL' not in r.stdout and '[非在籍の月]（1名' in r.stdout,
                r.stdout[-900:])
        tool_style = work / 'tool_style.xlsx'
        make_book(tool_style, headers=TOOL_HEADERS, second_months=['非在籍'] * 2 + [88000] * 10)
        r = run_py(VERIFY_BOOK, tool_style, cwd=work)
        c.check('明細: 期首月から並ぶ台帳でも、非在籍の月を見出しの年月で示す',
                f'{EMPLOYEES[1][0]}: {HIRED_IN_JUNE}\n' in r.stdout and ': 1月,2月' not in r.stdout,
                r.stdout[-900:])
        skill_style = work / 'skill_style.xlsx'
        make_book(skill_style, headers=SKILL_HEADERS,
                  second_months=[88000] * 3 + ['非在籍'] * 2 + [88000] * 7)
        r = run_py(VERIFY_BOOK, skill_style, cwd=work)
        c.check('明細: 列の位置が暦月の台帳でも、非在籍の月を見出しの年月で示す',
                f'{EMPLOYEES[1][0]}: {HIRED_IN_JUNE}\n' in r.stdout and ': 4月,5月' not in r.stdout,
                r.stdout[-900:])
        invalid = work / 'invalid.xlsx'
        make_book(invalid, second_months=['退職'] * 3 + [88000] * 9)
        r = run_py(VERIFY_BOOK, invalid, cwd=work)
        c.check('明細: 月の欄の数値でも「非在籍」でもない文字列は FAIL',
                r.returncode == 1 and '月の欄の不正な値 1名' in r.stdout, r.stdout[-900:])
        wb = openpyxl.Workbook()
        wb.save(work / 'empty.xlsx')
        r = run_py(VERIFY_BOOK, work / 'empty.xlsx', cwd=work)
        c.check('明細: 従業員別明細シートが無ければ終了コード2', r.returncode == 2, r.stderr)
    return c.finish()


if __name__ == '__main__':
    raise SystemExit(main())
