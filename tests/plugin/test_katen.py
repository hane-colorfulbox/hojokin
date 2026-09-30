"""katen-judge（配布版）の動作確認: 合成台帳 → 判定 → 公式様式①② → 原本無改変の確認。

スキルのフォルダだけで完結すること（チャットではスキルのフォルダしか実行環境に来ない）を、
リポジトリの外に組み立てた配布版で確かめる。エンジンを抜いたコピーでは終了コード2になること、
openpyxl で再保存した様式は check_lossless.py が不合格にすることも確かめる。
"""
import datetime
import shutil
import tempfile
from pathlib import Path

import openpyxl

from _common import Checker, run_py, skills_dir

SKILL = skills_dir() / 'katen-judge'
DRIVER = SKILL / 'review' / 'run_katen_judge.py'
LOSSLESS = SKILL / 'review' / 'check_lossless.py'
TEMPLATE = SKILL / 'ツール' / '加点判定用賃金台帳テンプレート.xlsx'
FORMS = SKILL / '補助金加点'
SHEET = '加点判定用明細'
FIRST_ROW = 7
COL_NO, COL_NAME, COL_TYPE, COL_HOURS, COL_FIRST_MONTH, COL_LATEST = 2, 3, 4, 5, 6, 18
# ②が対象になる構成: 基準月の最低 1150円（115,000÷100h）→ 直近月 1220円（差 +70円）
EMPLOYEES = [
    ('甲野花子', 'パート', 100, [115000] * 12, 122000),
    ('乙野太郎', '正社員', 173.33, [320000] * 12, 330000),
]


def make_ledger(path: Path) -> None:
    shutil.copy(TEMPLATE, path)
    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET]
    ws.cell(2, 3, '東京都')
    ws.cell(3, 3, datetime.datetime(2026, 8, 1))
    for i, (name, kind, hours, months, latest) in enumerate(EMPLOYEES):
        r = FIRST_ROW + i
        ws.cell(r, COL_NO, i + 1)
        ws.cell(r, COL_NAME, name)
        ws.cell(r, COL_TYPE, kind)
        ws.cell(r, COL_HOURS, hours)
        for m, v in enumerate(months):
            ws.cell(r, COL_FIRST_MONTH + m, v)
        ws.cell(r, COL_LATEST, latest)
    wb.save(path)


def main() -> int:
    c = Checker('katen-judge（配布版）')
    c.check('スキルに判定エンジンと様式が同梱されている',
            (SKILL / 'hojokin' / 'wage_reader.py').is_file() and len(list(FORMS.glob('*.xlsx'))) == 3)
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        ledger = work / '加点テスト台帳.xlsx'
        make_ledger(ledger)

        cp = run_py(DRIVER, ledger, '--fill', '通常', '--company', 'テスト社', '--outdir', work / 'out',
                    cwd=work)
        c.check('判定と様式生成が終了コード0で終わる', cp.returncode == 0, cp.stderr)
        c.check('加点②が対象（+70円）と判定される', '差: +70円' in cp.stdout and '判定: ◎ 対象' in cp.stdout,
                cp.stdout[-800:])
        outs = {p.name: p for p in (work / 'out').glob('*.xlsx')}
        c.check('様式①②の2つが出力される', len(outs) == 2, str(list(outs)))
        for name, out in outs.items():
            original = FORMS / name.removeprefix('テスト社_')
            r = run_py(LOSSLESS, original, out, cwd=work)
            c.check(f'{name} は原本どおり（check_lossless 合格）', r.returncode == 0, r.stdout + r.stderr)

        resaved = work / 'resaved.xlsx'
        wb = openpyxl.load_workbook(FORMS / '加点措置②用.xlsx')
        wb.save(resaved)
        r = run_py(LOSSLESS, FORMS / '加点措置②用.xlsx', resaved, cwd=work)
        c.check('openpyxl で再保存した様式は不合格（終了コード1）', r.returncode == 1, r.stdout)
        r = run_py(LOSSLESS, FORMS / '加点措置②用.xlsx', cwd=work)
        c.check('引数が足りないと終了コード2', r.returncode == 2)

        broken = work / 'broken' / 'katen-judge'
        shutil.copytree(SKILL, broken)
        shutil.rmtree(broken / 'hojokin')
        r = run_py(broken / 'review' / 'run_katen_judge.py', ledger, cwd=work)
        c.check('判定エンジンが欠けたスキルでは終了コード2で止まる', r.returncode == 2 and 'hojokin/' in r.stderr,
                f'rc={r.returncode} {r.stderr}')
    return c.finish()


if __name__ == '__main__':
    raise SystemExit(main())
