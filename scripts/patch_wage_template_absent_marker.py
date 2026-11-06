# -*- coding: utf-8 -*-
"""標準の賃金台帳テンプレートに「非在籍」の書き方を足す（ツール側とスキル側を同じバイトにする）。

背景:
    第7次の通常枠は、労働生産性の基準（前事業年度の総労働時間）を賃金台帳から確かめる。
    前事業年度の途中に入社・退職した人がいると、在籍していない月を書く方法がなかったため、
    ツールは必ず停止していた。月の欄（G〜R列）に「非在籍」と書けば、その月を総労働時間から除く
    （hojokin.wage_reader.ABSENT_MARKER。空欄・0円は非在籍と推測しない）。

修正内容（表示の文言だけ。行・列・書式・他の部品は変えない）:
    - 従業員別明細!B4: 「非在籍」の書き方と、第7次通常枠でE列が全員に要ることを追記。
      あわせて、第7次以降の通常枠は5行目G〜R列の月見出しを前事業年度12か月の年付き（例：令和7年4月。
      括弧などの注記なし）に書き換えないとツールが止まることを追記（見出しのセル自体は「1月」〜「12月」のまま）
    - 記入ルール!B6（月間平均時間）・B8（1月〜12月）: 同上

実行方法:
    python -B scripts/patch_wage_template_absent_marker.py
    ツール/賃金台帳テンプレート.xlsx を書き換え、同じバイト列を
    .claude/skills/wagebook-convert/templates/賃金台帳テンプレート.xlsx へコピーする
    （配布ゲート scripts/plugin_checks.gate_sync がバイト一致を検査する）。

冪等性:
    文言を置き換えるだけなので、再実行しても同じ結果になる。
"""
from __future__ import annotations

import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from hojokin.xlsx_surgical import _patch_sheet_xml, _resolve_sheet_part  # noqa: E402

sys.stdout.reconfigure(encoding='utf-8')

TOOL_TEMPLATE = ROOT / 'ツール' / '賃金台帳テンプレート.xlsx'
SKILL_TEMPLATE = ROOT / '.claude' / 'skills' / 'wagebook-convert' / 'templates' / '賃金台帳テンプレート.xlsx'

ABSENT_GUIDE = ('在籍していない月（入社前・退職後）は、名簿・入退社日で確認できた月だけ「非在籍」と入力してください'
                '（空欄・0円を非在籍とは扱いません）。')
YEAR_HEADER_GUIDE = ('第7次以降の通常枠では、5行目G〜R列の月見出しを前事業年度12か月の年付き（例：令和7年4月）に'
                     '書き換えてください（括弧などの注記は付けません。年がないとツールが止まります）。')
B4 = ('※金額は課税支給合計（税込）を入力してください。支給がない月は空欄のままにしてください（0は入れません）。'
      + ABSENT_GUIDE + YEAR_HEADER_GUIDE +
      '　／ S列「年間通勤手当」には、毎月の欄に入れた金額に含まれている通勤手当の年間合計を入力'
      '（課税・非課税を問いません。ツールが在籍月数で均等割して各月から自動減算します）。'
      '毎月の欄が「課税支給合計」で通勤手当がそこに含まれていない場合は空欄。\n'
      'E列「月間平均時間」・F列「時給」は ★パート・アルバイトでは必須★（IT導入補助金 R215 FTE換算用、公募要領 p.10）。'
      '正社員・役員は空欄可（第7次以降の通常枠は、生産性の総労働時間に使うため役員以外の全員にE列が必要）。'
      '／ T列「年間賞与」には課税賞与の年間合計を入力（月次セルには混ぜない。ツールが給与支給総額R216に算入。'
      '賞与の支給月は問わない）。')
RULE_B6 = ('年間の月平均労働時間（例: 160）。★パート・アルバイトでは必須★（IT導入補助金 R215 FTE換算用）。'
           '正社員は空欄可（第7次以降の通常枠は、役員以外の全員に必要）')
RULE_B8 = ('各月の課税支給合計。支給がない月は空欄（0は入れない）。在籍していない月（入社前・退職後）は「非在籍」と入力'
           '（名簿・入退社日で確認できた月だけ。空欄・0円を非在籍とは扱わない。休職など在籍していて給与のない月は空欄のまま）。'
           '第7次以降の通常枠は、5行目G〜R列の月見出しを前事業年度12か月の年付き（例：令和7年4月）に書き換える'
           '（括弧などの注記は付けない。年がないとツールが止まる）')
EDITS = {0: {'B4': B4}, 1: {'B6': RULE_B6, 'B8': RULE_B8}}


def patch(path: Path) -> None:
    tmp = path.with_name(path.stem + '.tmp.xlsx')
    with zipfile.ZipFile(path) as zin:
        replacements = {}
        for index, cells in EDITS.items():
            part = _resolve_sheet_part(zin, index)
            replacements[part] = _patch_sheet_xml(zin.read(part).decode('utf-8'), cells).encode('utf-8')
        with zipfile.ZipFile(tmp, 'w') as zout:
            for info in zin.infolist():
                zout.writestr(info, replacements.get(info.filename, zin.read(info.filename)))
    tmp.replace(path)


def main() -> None:
    patch(TOOL_TEMPLATE)
    shutil.copyfile(TOOL_TEMPLATE, SKILL_TEMPLATE)
    wb = openpyxl.load_workbook(TOOL_TEMPLATE)
    checks = {
        'B4': wb['従業員別明細']['B4'].value == B4,
        '記入ルールB6': wb['記入ルール']['B6'].value == RULE_B6,
        '記入ルールB8': wb['記入ルール']['B8'].value == RULE_B8,
        'スキル側と同じバイト': TOOL_TEMPLATE.read_bytes() == SKILL_TEMPLATE.read_bytes(),
    }
    wb.close()
    for label, ok in checks.items():
        print(f'  {"OK" if ok else "NG"}: {label}')
    if not all(checks.values()):
        sys.exit(1)


if __name__ == '__main__':
    main()
