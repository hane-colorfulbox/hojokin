"""katen-judge（配布版）の動作確認: 合成台帳 → 判定 → 公式様式①② → 原本無改変の確認。

スキルのフォルダだけで完結すること（チャットではスキルのフォルダしか実行環境に来ない）を、
リポジトリの外に組み立てた配布版で確かめる。エンジンを抜いたコピーでは終了コード2になること、
openpyxl で再保存した様式は check_lossless.py が不合格にすることも確かめる。
"""
import datetime
import re
import shutil
import tempfile
from pathlib import Path

import openpyxl

from _common import Checker, run_py, skills_dir

import distribution_manifest as dm

SKILL = skills_dir() / 'katen-judge'
DRIVER = SKILL / 'review' / 'run_katen_judge.py'
LOSSLESS = SKILL / 'review' / 'check_lossless.py'
TEMPLATE = SKILL / 'ツール' / '加点判定用賃金台帳テンプレート.xlsx'
FORMS = SKILL / '補助金加点'
SHEET = '加点判定用明細'
FIRST_ROW = 7
COL_NO, COL_NAME, COL_TYPE, COL_HOURS, COL_FIRST_MONTH, COL_LATEST = 2, 3, 4, 5, 6, 18
# 第7次以降の台帳だけにある時間換算の対象月額（AH〜AS列＝対象期間12か月、AT列＝直近月）
COL_BASIS_FIRST, COL_BASIS_LATEST = 34, 46
# 旧版の適法な合成例: 基準月1200円→直近月1270円（差+70円）。
EMPLOYEES = [
    ('甲野花子', 'パート', 100, [120000] * 12, 127000),
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

        cp = run_py(DRIVER, ledger, '--call-round', 6, '--fill', '通常', '--company', 'テスト社', '--outdir', work / 'out',
                    cwd=work)
        c.check('判定と様式生成が終了コード0で終わる', cp.returncode == 0, cp.stderr)
        c.check('加点②が対象（+70円）と判定される', '差: +70円' in cp.stdout and '判定: 対象' in cp.stdout,
                cp.stdout[-800:])
        outs = {p.name: p for p in (work / 'out').glob('*.xlsx')}
        c.check('様式①②の2つが公式の様式名で出力される',
                sorted(outs) == ['テスト社_加点措置②用.xlsx', 'テスト社_補助率引き上げ・加点措置①用.xlsx'], str(list(outs)))
        originals = dict(re.findall(r'様式[①②]生成: (.+?)（原本: (.+?)）', cp.stdout))
        c.check('ドライバが原本のパスを表示する', len(originals) == 2, cp.stdout[-600:])
        for out, original in originals.items():
            r = run_py(LOSSLESS, original, out, cwd=work)
            c.check(f'{Path(out).name} は原本どおり（check_lossless 合格）', r.returncode == 0, r.stdout + r.stderr)

        cp = run_py(DRIVER, ledger, '--call-round', 6, '--fill', 'インボイス', '--company', 'テスト社', '--outdir', work / 'inv',
                    cwd=work)
        c.check('インボイス枠の様式①も見つかって生成される',
                cp.returncode == 0 and (work / 'inv' / 'テスト社_加点措置①用.xlsx').exists(), cp.stdout[-400:] + cp.stderr)
        c.check('配布版の様式ファイル名に記号が無い', all(dm.path_chars_ok(p.name) for p in FORMS.glob('*.xlsx')),
                str([p.name for p in FORMS.glob('*.xlsx')]))

        form2 = FORMS / dm.dist_path('加点措置②用.xlsx')
        resaved = work / 'resaved.xlsx'
        wb = openpyxl.load_workbook(form2)
        wb.save(resaved)
        r = run_py(LOSSLESS, form2, resaved, cwd=work)
        c.check('openpyxl で再保存した様式は不合格（終了コード1）', r.returncode == 1, r.stdout)
        r = run_py(LOSSLESS, form2, cwd=work)
        c.check('引数が足りないと終了コード2', r.returncode == 2)

        broken = work / 'broken' / 'katen-judge'
        shutil.copytree(SKILL, broken)
        shutil.rmtree(broken / 'hojokin')
        r = run_py(broken / 'review' / 'run_katen_judge.py', ledger, cwd=work)
        c.check('判定エンジンが欠けたスキルでは終了コード2で止まる', r.returncode == 2 and 'hojokin/' in r.stderr,
                f'rc={r.returncode} {r.stderr}')

        # 新版も配布されたスキルのフォルダだけで動作し、旧新取り違えを止める。
        new_template = SKILL / 'ツール' / '第7次以降' / '加点判定用賃金台帳テンプレート.xlsx'
        c.check('第7次の台帳と版別ルールが同梱されている',
                new_template.exists() and (SKILL / 'hojokin' / 'subsidy_rules.py').exists())
        if new_template.exists():
            new_ledger = work / 'new.xlsx'
            shutil.copy(new_template, new_ledger)
            wb = openpyxl.load_workbook(new_ledger)
            ws = wb[SHEET]
            ws['C2'] = '東京都'
            ws['C3'] = datetime.datetime(2026, 12, 1)
            ws['C7'], ws['D7'], ws['E7'] = '合成従業員', '正社員', 100
            # 基本給（F〜R列）は様式の【基本給】欄、対象月額（AH〜AT列）は時間換算給与の分子。
            for col in range(6, 18):
                ws.cell(7, col, 120000)
                ws.cell(7, COL_BASIS_FIRST + col - 6, 122600)
            ws['R7'] = 125000
            ws.cell(7, COL_BASIS_LATEST, 128100)  # +55円ちょうど
            wb.save(new_ledger)
            wb.close()
            cp = run_py(DRIVER, new_ledger, '--call-round', 7, '--fill', '通常', '--company', '新版テスト',
                        '--outdir', work / 'new_out', cwd=work)
            c.check('第7次で55円ちょうどを対象として公式様式を生成',
                    cp.returncode == 0 and '差: +55円' in cp.stdout,
                    cp.stdout[-500:] + cp.stderr)
            form2 = next((p for p in sorted((work / 'new_out').glob('*.xlsx'))
                          if '②' in str(openpyxl.load_workbook(p).worksheets[0]['B2'].value or '')), None)
            if form2 is not None:
                sheet = openpyxl.load_workbook(form2).worksheets[0]
                c.check('様式②の基本給欄と時間換算給与欄を別の値から作る',
                        (sheet['F17'].value, sheet['G17'].value, sheet['N17'].value, sheet['O17'].value)
                        == (120000, 1226, 125000, 1281),
                        str((sheet['F17'].value, sheet['G17'].value, sheet['N17'].value, sheet['O17'].value)))
            else:
                c.check('様式②の基本給欄と時間換算給与欄を別の値から作る', False, '様式②が生成されていない')
            blank_basis = work / 'blank_basis.xlsx'
            shutil.copy(new_ledger, blank_basis)
            wb = openpyxl.load_workbook(blank_basis)
            for col in range(COL_BASIS_FIRST, COL_BASIS_LATEST + 1):
                wb[SHEET].cell(7, col).value = None
            wb.save(blank_basis)
            wb.close()
            blank = run_py(DRIVER, blank_basis, '--call-round', 7, cwd=work)
            c.check('時間換算の対象月額が空欄なら基本給で代用せず未確定（exit3）',
                    blank.returncode == 3 and '時間換算の対象月額' in blank.stdout,
                    blank.stdout[-500:] + blank.stderr)
            wrong = run_py(DRIVER, new_ledger, '--call-round', 6, cwd=work)
            c.check('新版台帳を旧版の回として読まない', wrong.returncode == 2, wrong.stderr)
            unstamped = run_py(DRIVER, new_ledger, '--fill', '通常', cwd=work)
            c.check('公募回なしでは提出用様式を生成しない', unstamped.returncode == 2)
            wb = openpyxl.load_workbook(new_ledger)
            wb[SHEET]['E7'] = None
            wb.save(new_ledger)
            wb.close()
            missing = run_py(DRIVER, new_ledger, '--call-round', 7, '--fill', '通常',
                             '--outdir', work / 'missing_out', cwd=work)
            c.check('時間欠損の未確定判定はexit3で提出物を作らない',
                    missing.returncode == 3 and not list((work / 'missing_out').glob('*.xlsx')),
                    missing.stdout[-500:] + missing.stderr)
    return c.finish()


if __name__ == '__main__':
    raise SystemExit(main())
