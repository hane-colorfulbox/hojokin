"""旧原本を保存し、第7次候補原本を再現可能に作る。Drive原本への書込みはしない。

公式加点様式は取得済みのバイト列をそのまま複製する。社内原本はZIP内の
表示セル・プルダウンの額・版メタデータ・通常枠の生産性計画欄だけを変更し、
図形・その他の数式・入力規則・保護を保持する。行・列の挿入削除はしない。
通常枠個人は既存のDrive由来パッチ済み原本を入力にする。
文言の根拠: 第7回〜公募要領（通常枠 印刷p.9〜10・23〜25、インボイス枠 印刷p.8〜9・22）、
交付申請マニュアル11/2版 p.86〜88、公式「賃金状況報告シート」第7次公募回以降用。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from hojokin.config import get_mapping
from hojokin.template_versions import APPLICATION_FILENAMES, ROUND7_SUBJECT
from hojokin.xlsx_surgical import _patch_sheet_xml, _resolve_sheet_part

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

FORMS = {
    'it2026_kaitei_chingin_houkoku1_hojyoritsu_katen.xlsx': '補助率引き上げ・加点措置①用.xlsx',
    'it2026_kaitei_chingin_houkoku1_katen.xlsx': '加点措置①用.xlsx',
    'it2026_kaitei_chingin_houkoku2_katen.xlsx': '加点措置②用.xlsx',
}
BASE_ROWS = {'通常枠_2026': 216, '通常枠_個人_2026': 192,
             'インボイス枠_2026': 201, 'インボイス枠_個人_2026': 173}
SHINSEI = '申請内容'
TENKI = '転記'
PRODUCTIVITY = '生産性指標給与支給総額計算'
VERSION_CELLS = ((SHINSEI, 'D2'), (TENKI, 'A1'))
GROWTH_GUIDE_PREFIX = '⇨年平均成長率'
NO_EMPLOYEE_PREFIX = '賃金引上げ計画期間内において、従業員を雇用する場合'
NAVI_PREFIX = 'IT戦略ナビwithの実施有無\n（加点）'
SELECTED, UNSELECTED = '■', '□'
# 通常枠の生産性計画（C〜E列＝1〜3年度目）。行は原本の既存レイアウト。
PLAN_COLUMNS = {'C': 1, 'D': 2, 'E': 3}
ROW_PROFIT, ROW_PERSONNEL, ROW_DEPRECIATION, ROW_VALUE_ADDED = 34, 35, 36, 37
ROW_EMPLOYEES, ROW_HOURS, ROW_TOTAL_HOURS, ROW_PRODUCTIVITY, ROW_GROWTH = 38, 39, 40, 41, 42
PLAN_INPUT_ROWS = (ROW_PROFIT, ROW_PERSONNEL, ROW_DEPRECIATION, ROW_EMPLOYEES, ROW_HOURS)
OLD_FORECAST_PARAMETERS = ('F34', 'G34', 'H34', 'F39', 'G39', 'H39')

ROUND7_LABEL = '【第7次以降用の原本】第7回公募以降の申請に使います（第1〜6次の案件には使いません）'
MIN_WAGE_NOTE = '主たる事業場（確認済み）の都道府県名のみ。住所から推測しない\n令和8年度地域別最低賃金額'
WAGE_BASE_NOTE = ('確認済みの0年度目見込みを入力（自動リンクなし）\n'
                  '0年度目＝交付申請日の属する事業年度。前事業年度の実績をそのまま入れない')
METHOD_NOTE = ('顧客に確認した実施方法を選択（初期値なし）。'
               '表明していない場合、賃上げが必須の区分は申請できず、賃上げ加点も選べない')
DATE_NOTE = '顧客に確認した表明の実施日を入力（交付申請時点までに実際に表明した日。作業日などで代用しない）'
EXEMPT_LINE = '・保険医療機関・介護サービス事業者・社会福祉事業者・学校など公募要領で適用外とされる事業者：必須ではない'
GROWTH_GUIDE = {
    '通常': ('⇨年平均成長率：上記4点を入力後「年平均成長率計算」をクリックし、申請区分の必要率を「1年後」と「年平均」の両方で確認する。\n'
             '・150万円以上：3%以上（IT導入補助金2022〜2025で交付決定ありは3.5%以上）\n'
             '・150万円未満でIT導入補助金2022〜2025の交付決定あり：3.5%以上\n'
             '・150万円未満で交付決定なし：必須ではない（賃上げ加点を希望する場合は年平均3%以上）\n' + EXEMPT_LINE),
    'インボイス': ('⇨年平均成長率：上記4点を入力後「年平均成長率計算」をクリックし、必要率を「1年後」と「年平均」の両方で確認する。\n'
              '・IT導入補助金2022〜2025で交付決定あり：3.5%以上（必須）\n'
              '・交付決定なし：必須ではない（賃上げ加点を希望する場合は年平均3%以上）\n' + EXEMPT_LINE),
}
NAVI_LABEL = '「IT戦略ナビwith」又は「IT経営サポートセンター」の利用有無\n（加点）'
IT_SUPPORT_TERMS = ('2026年8月1日以降に利用を申し込み、交付申請日以前かつ締切日の前月末日'
                    '（第7次は2026年11月30日）までに利用したものが対象。')
NAVI_NOTE = {
    '法人': ('通常枠の加点9）。どちらか一方の利用で加点（両方利用しても加点は1つ）。\nIT経営サポートセンター：'
             + IT_SUPPORT_TERMS + '申込みの法人名・住所・法人番号が申請と一致していること'),
    '個人': ('通常枠の加点9）。どちらか一方の利用で加点（両方利用しても加点は1つ）。\nIT経営サポートセンター：'
             + IT_SUPPORT_TERMS + '申込みの「企業名」欄の代表者氏名と住所が申請と一致していること'),
}
HEARING_NAVI_ITEM = 'デジwithにおける「IT戦略ナビwith」の実施\n（または中小機構「IT経営サポートセンター」の利用）'
HEARING_NAVI_NOTE = {
    kind: ('必ず実施ください。\n※実施にはgBizIDプライムが必要になります。\n'
           '※中小機構「IT経営サポートセンター」の利用も同じ加点の対象です（両方利用しても加点は1つです）。'
           '2026年8月1日以降に利用を申し込み、交付申請日までかつ2026年11月30日まで（第7次の場合）に利用したものに限ります。'
           + tail) for kind, tail in (('法人', '申込みの法人名・住所・法人番号は、申請と同じ内容にしてください。'),
                                      ('個人', '申込みの「企業名」欄には代表者氏名を入力し、住所は申請と同じ内容にしてください。'))
}
SUBSIDY_RATE_NOTE = ('◼️補助率の引上げ（1/2以内→2/3以内）を希望する場合\n'
                     'その旨お知らせいただき「賃金状況報告シート（補助率引上げ・加点措置①用）」を提出してください。\n'
                     '（対象）\n令和7年10月から令和8年9月までの間で、「当該期間における地域別最低賃金以上～'
                     '令和8年度改定の地域別最低賃金未満」で雇用している従業員が全従業員の30％以上である月が'
                     '3か月以上ある事業者（3か月は連続していなくてもかまいません）')
INVOICE_WAGE_ITEM = '給与0年度目（交付申請日の属する事業年度）の見込み総額・対象人数と、その算定根拠'
INVOICE_WAGE_NOTE = ('申請では、0年度目（交付申請日の属する事業年度）から3年度目までの給与支給総額と従業員数を入力します。\n'
                     '0年度目は見込みの値です。次の資料をご提出ください。\n'
                     '・今期のうち支給済みの月の賃金台帳と、残りの月の支給見込み（賞与・昇給の予定を含む）\n'
                     '・加点を希望する場合は、令和7年10月〜令和8年9月の各月と、申請月の前月の賃金台帳\n'
                     '前事業年度の賃金台帳だけでは、今期の見込みは作れません。')
REQUIREMENTS_HEADERS = {'C2': '申請額：150万未満', 'D2': '申請額：150万以上',
                       'B3': '2022から2025で交付決定を受けていない', 'B4': '2022から2025で交付決定を受けた'}
_DECLARE = '・賃金引き上げ計画を従業員へ表明すること。'
_GROWTH = '・1年後に、一人当たり給与支給総額を{r}%以上向上させること。\n・事業計画期間において、一人当たり給与支給総額の年平均成長率を{r}%以上とすること。'
_MIN30 = '・事業計画期間において、事業場内最低賃金（事業場内で最も低い賃金）を地域別最低賃金＋30円以上の水準にすること。'
_PLUS50 = '【加点の要件】上記に加え、事業場内最低賃金を地域別最低賃金＋50円以上の水準にすること。'
REQUIREMENTS_TABLE = {
    'C3': ('【賃上げは必須ではありません。加点を希望する場合の要件】\n' + _MIN30
           + '\n・事業計画期間において、一人当たり給与支給総額の年平均成長率を3.0%以上とすること。\n' + _DECLARE
           + '\n※ なお、上記に加え、事業計画期間において、事業場内最低賃金を地域別最低賃金＋50円以上の水準にした場合、更なる加点を行う。'),
    'D3': '\n'.join(('【必須要件】', _GROWTH.format(r='3.0'), _MIN30, _DECLARE, _PLUS50)),
    'C4': '\n'.join(('【必須要件】', _GROWTH.format(r='3.5'), _DECLARE,
                     '【加点の要件】上記に加え、' + _MIN30[1:] + '（＋50円以上の水準にした場合は更なる加点）')),
    'D4': '\n'.join(('【必須要件】', _GROWTH.format(r='3.5'), _MIN30, _DECLARE, _PLUS50)),
    'C5': '※ 基準年度（0年度目）は交付申請日の属する事業年度です。計画は交付申請時点の翌事業年度から3年間です。',
    'C6': '※ 保険医療機関・介護サービス事業者・社会福祉事業者・学校など公募要領で適用外とされる事業者は、必須要件の対象外です。',
}
LEDGER_RULES = {
    'B5': ('連番／氏名／「正社員」「パート・アルバイト」「役員」等。判定の母数は各月の常時使用する従業員'
           '（労働基準法第20条の解雇予告が必要な人）。役員・事業主本人と、その月に日雇い・2か月以内の期間雇用・'
           '季節的業務で4か月以内・試用期間14日以内の人は母数に入れない（所定の期間を超えて引き続き使用している人は入れる）。'),
    'A7': '令和7年10月〜令和8年9月 基本給（F〜Q列）',
    'B7': ('各暦月の、時間換算の元になる月額（基本給と、職務手当など毎月決まって払う通常の手当）。'
           '厚生労働省「最低賃金額以上かどうかを確認する方法」で除く6種類（臨時の賃金・賞与など1か月を超える期間ごとの賃金・'
           '時間外/休日/深夜の割増賃金・精皆勤手当・通勤手当・家族手当）は含めない。通常の手当を含めた人は、'
           '提出様式の【基本給】欄を原本の基本給に直す。'
           '在籍していない月は「非在籍」と記入。空欄・0円の月は未確認として扱い、判定が確定しません。'
           '公募回・申請月と暦月の対応を確認し、金額・所定労働時間が分からない月は確定させない。'),
    'B8': '加点②用。交付申請月の前月の月額（F〜Q列と同じ考え方：基本給＋職務手当などの通常の手当。厚労省方式で除く6種類は含めない）。',
    'B10': '本台帳には記入不要（ツールが「台帳の月額÷その月の労働時間」で自動算出します。月別労働時間が空欄の月はE列の所定労働時間を使います）。',
}
# 再生成後に残っていてはいけない旧回・旧運用の表記（数値の一部の「…63円」は対象外）。
STALE_PATTERNS = (r'R[7７][.．][7７]', r'R[7７]年度改定', r'(?<![0-9０-９,，])[6６][3３]円',
                  r'令和[6６]年\s*[1１][0０]月', r'令和[7７]年度改定', r'適用補助率確認シート',
                  r'以内で雇用', r'任意の賃上げ計画加点）で雇用', r'登記から抜粋', r'リンクされているので触らない',
                  r'特に指定がなければ', r'シート作成日', r'支給なしは空欄', r'前事業年度分の賃金台帳をご提出',
                  r'3\.0％以上であることを確認', r'it2026_chingin_houkoku',
                  r'年平均成長率を3\.5％以上とすることが申請要件')


def _hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _text(value: str) -> str:
    pairs = [
        (r'令和[6６]年\s*[1１][0０]月', '令和7年10月'),
        (r'令和[7７]年\s*[9９]月', '令和8年9月'),
        (r'令和[7７]年\s*[7７]月', '令和8年6月'),
        (r'令和[7７]年度改定', '令和8年度改定'),
        (r'令和[7７]年度地域別最低賃金', '令和8年度地域別最低賃金'),
        (r'R[7７]年度改定', 'R8年度改定'),
        (r'R[7７]改定後', 'R8改定後'),
        (r'R[7７][.．][7７](?=事業場内最低賃金)', 'R8.6'),
        (r'(?<![0-9０-９,，])[6６][3３]円', '55円'),
        (r'it2026_chingin_houkoku', 'it2026_kaitei_chingin_houkoku'),
        (r'直近[2２]期分の賃金台帳', '用途別の指定年月をすべて覆う賃金台帳（冊数は固定しない）'),
    ]
    for pattern, replacement in pairs:
        value = re.sub(pattern, replacement, value)
    return value


def _patch_formula_xml(sheet_xml: str, cells: dict[str, str | None]) -> str:
    """数式セルを差し替える。None は書式だけ残した空欄（入力欄）にする。"""
    for ref, formula in cells.items():
        match = re.search(rf'<c r="{ref}"(?P<attrs>[^>]*?)(?:/>|>.*?</c>)', sheet_xml, re.S)
        if match is None:
            if formula is None:
                continue
            raise ValueError(f'数式を書き込むセル {ref} が原本にありません。')
        style = re.search(r'\ss="(\d+)"', match.group('attrs'))
        cell = f'<c r="{ref}"' + (f' s="{style.group(1)}"' if style else '')
        cell += '/>' if formula is None else f'><f>{escape(formula[1:])}</f></c>'
        sheet_xml = sheet_xml[:match.start()] + cell + sheet_xml[match.end():]
    return sheet_xml


def _assert_no_stale_text(path: Path) -> None:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    hits = [f'{ws.title}!{cell.coordinate}' for ws in wb for row in ws for cell in row
            if isinstance(cell.value, str) and cell.data_type != 'f'
            and any(re.search(p, cell.value) for p in STALE_PATTERNS)]
    wb.close()
    if hits:
        raise AssertionError(f'{path.name} に旧表記が残っています: {hits}')


def _patch_book(source: Path, output: Path, edits: dict[str, dict[str, object]],
                formulas: dict[str, dict[str, str | None]] | None = None) -> dict:
    formulas = formulas or {}
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as zin:
        wb = openpyxl.load_workbook(source, read_only=True, data_only=False)
        replacements = {}
        for sheet_name in dict.fromkeys([*edits, *formulas]):
            part = _resolve_sheet_part(zin, wb.sheetnames.index(sheet_name))
            xml = _patch_sheet_xml(zin.read(part).decode('utf-8'), edits.get(sheet_name, {}))
            replacements[part] = _patch_formula_xml(xml, formulas.get(sheet_name, {})).encode('utf-8')
        core = zin.read('docProps/core.xml').decode('utf-8')
        subject = '<dc:subject xmlns:dc="http://purl.org/dc/elements/1.1/">' + escape(ROUND7_SUBJECT) + '</dc:subject>'
        if re.search(r'<dc:subject(?:\s[^>]*)?>.*?</dc:subject>', core, re.S):
            core = re.sub(r'<dc:subject(?:\s[^>]*)?>.*?</dc:subject>',
                          subject, core, flags=re.S)
        elif '<dc:subject/>' in core:
            core = core.replace('<dc:subject/>', subject)
        else:
            # Keep all namespace declarations, including compatibility prefixes.
            if 'xmlns:dc=' not in core:
                core = core.replace('<cp:coreProperties ', '<cp:coreProperties xmlns:dc="http://purl.org/dc/elements/1.1/" ', 1)
            core = core.replace('</cp:coreProperties>', subject + '</cp:coreProperties>')
        replacements['docProps/core.xml'] = core.encode('utf-8')
        for name in zin.namelist():
            if name.endswith('.rels'):
                raw = zin.read(name)
                linked = raw.replace(b'it2026_chingin_houkoku', b'it2026_kaitei_chingin_houkoku')
                if linked != raw:
                    replacements[name] = linked
        with zipfile.ZipFile(output, 'w') as zout:
            for info in zin.infolist():
                zout.writestr(info, replacements.get(info.filename, zin.read(info.filename)))
        wb.close()
    # Reopen and confirm no row/sheet/formula/object changes outside explicit cells.
    with zipfile.ZipFile(source) as a, zipfile.ZipFile(output) as b:
        assert a.namelist() == b.namelist()
        for name in a.namelist():
            if name not in replacements:
                assert a.read(name) == b.read(name), name
    checked = openpyxl.load_workbook(output, read_only=True, data_only=False)
    assert checked.properties.subject == ROUND7_SUBJECT
    for cells in (edits, formulas):
        for sheet, values in cells.items():
            for ref, value in values.items():
                assert checked[sheet][ref].value == value, (source.name, sheet, ref)
    checked.close()
    _assert_no_stale_text(output)
    return {'input': source.relative_to(ROOT).as_posix(), 'input_sha256': _hash(source),
            'output': output.relative_to(ROOT).as_posix(), 'output_sha256': _hash(output),
            'cells': {sheet: list(values) for sheet, values in edits.items()},
            'formulas': {sheet: list(values) for sheet, values in formulas.items()}}


def _b_labels(ws) -> dict[int, str]:
    return {cell.row: cell.value for row in ws.iter_rows(min_col=2, max_col=2)
            for cell in row if isinstance(cell.value, str)}


def _row_starting(labels: dict[int, str], prefix: str) -> int:
    rows = [row for row, text in labels.items() if text.startswith(prefix)]
    if len(rows) != 1:
        raise ValueError(f'原本のB列に「{prefix[:20]}」の行が1つだけ見つかりません: {rows}')
    return rows[0]


def _application_edits(wb, template_type: str, edits: dict) -> dict:
    """申請内容の確認文・初期値・注記と版表示。セル番地は config の対応表から引く。"""
    rows = get_mapping(template_type).shinsei
    ws = wb[SHINSEI]
    labels = _b_labels(ws)
    frame = '通常' if template_type.startswith('通常枠') else 'インボイス'
    values = edits.setdefault(SHINSEI, {})
    guide = rows['wage_total_y3'] + 1
    if not labels.get(guide, '').startswith(GROWTH_GUIDE_PREFIX):
        raise ValueError(f'{template_type} の{guide}行が賃上げ率の確認行ではありません。')
    values[f'B{guide}'] = GROWTH_GUIDE[frame]
    values[f'D{rows["wage_total_base"]}'] = WAGE_BASE_NOTE
    values[f'D{rows["min_wage"]}'] = MIN_WAGE_NOTE
    values[f'D{rows["wage_raise_method"]}'] = METHOD_NOTE
    values[f'D{rows["wage_raise_date"]}'] = DATE_NOTE
    for row in (rows['wage_raise_declaration'], rows['wage_raise_method'],
                _row_starting(labels, NO_EMPLOYEE_PREFIX)):
        choices = ws[f'C{row}'].value
        if not isinstance(choices, str) or UNSELECTED not in choices:
            raise ValueError(f'{template_type} のC{row}が表明の選択肢ではありません。')
        values[f'C{row}'] = choices.replace(SELECTED, UNSELECTED)
    for sheet, ref in VERSION_CELLS:
        edits.setdefault(sheet, {})[ref] = ROUND7_LABEL
    if frame != '通常':
        return {}
    navi = _row_starting(labels, NAVI_PREFIX)
    values[f'B{navi}'] = NAVI_LABEL
    values[f'D{navi}'] = NAVI_NOTE['個人' if '個人' in template_type else '法人']
    return _productivity_edits(edits)


def _productivity_edits(edits: dict) -> dict:
    """通常枠の旧自動予測式（+3%/年・×1.035・−10h等）を確認済み入力欄にし、1年後と年平均を表示する。"""
    labels = edits.setdefault(PRODUCTIVITY, {})
    labels.update({ref: None for ref in OLD_FORECAST_PARAMETERS})
    labels[f'A{ROW_GROWTH}'] = '1年後の伸び率（C列）／年平均成長率（D列＝2年、E列＝3年）'
    labels[f'B{ROW_GROWTH}'] = ('要件：1年後・年平均とも3%以上。IT導入補助金2023の通常枠（A・B類型）・'
                                'デジタル化基盤導入枠（複数社連携）、2024・2025の通常枠・複数社連携の交付決定ありは4%以上')
    cells: dict[str, str | None] = {f'{col}{row}': None for col in PLAN_COLUMNS for row in PLAN_INPUT_ROWS}
    for col, years in PLAN_COLUMNS.items():
        inputs = f'{col}{ROW_PROFIT}:{col}{ROW_DEPRECIATION}'
        value_added, total, productivity = (f'{col}{ROW_VALUE_ADDED}', f'{col}{ROW_TOTAL_HOURS}',
                                            f'{col}{ROW_PRODUCTIVITY}')
        ratio = f'{productivity}/$B${ROW_PRODUCTIVITY}'
        growth = ratio if years == 1 else f'({ratio})^(1/{years})'
        cells[value_added] = f'=IF(COUNT({inputs})<3,"",SUM({inputs}))'
        cells[total] = f'=IF(COUNT({col}{ROW_EMPLOYEES}:{col}{ROW_HOURS})<2,"",{col}{ROW_EMPLOYEES}*{col}{ROW_HOURS})'
        cells[productivity] = f'=IF(OR({value_added}="",N({total})<=0),"",{value_added}/{total})'
        cells[f'{col}{ROW_GROWTH}'] = (f'=IF({productivity}="","",IFERROR('
                                       f'IF(N($B${ROW_PRODUCTIVITY})<=0,"要確認",{growth}-1),"要確認"))')
    return {PRODUCTIVITY: cells}


# 基本情報の賃上げ欄にある過去交付ありの説明（第1〜6次は年平均だけ）。第7次は1年後要件が加わる（インボイス要領（ツ）一・二・三）
OLD_PAST_GRANT_NOTE = '※2022から2025までの間にIT導入補助金の交付決定を受けた事業者は年平均成長率を3.5％以上とすることが申請要件となっております'
PAST_GRANT_NOTE = ('※IT導入補助金2022〜2025の交付決定を受けた事業者は、1人当たり給与支給総額を1年後と年平均の両方で3.5％以上'
                   '伸ばす計画を立て、従業員に表明することが申請要件です')


def _hearing_edits(wb, name: str, edits: dict) -> None:
    """ヒアリングの確認事項・要件表。行の挿入はせず既存セルの文言だけを変える。"""
    kind = '個人' if '個人' in name else '法人'
    normal = '通常枠' in name
    ws = wb['確認事項']
    values = edits.setdefault(ws.title, {})
    for row in ws.iter_rows(min_col=3, max_col=3):
        for cell in row:
            text = cell.value if isinstance(cell.value, str) else ''
            if '前事業年度1年分の給与支給総額' in text:
                values[cell.coordinate] = INVOICE_WAGE_ITEM
                values[f'D{cell.row}'] = INVOICE_WAGE_NOTE
            elif normal and text.startswith('◼️補助率の引上げを希望する場合'):
                values[cell.coordinate] = SUBSIDY_RATE_NOTE
            elif normal and text.startswith('デジwithにおける「IT戦略ナビwith」の実施'):
                values[cell.coordinate] = HEARING_NAVI_ITEM
                values[f'D{cell.row}'] = HEARING_NAVI_NOTE[kind]
    if normal and not {SUBSIDY_RATE_NOTE, HEARING_NAVI_ITEM} <= set(values.values()):
        raise ValueError(f'{name} の確認事項に補助率引上げ・IT戦略ナビの行が見つかりません。')
    if '基本情報' in wb.sheetnames:
        basic = wb['基本情報']
        basic_values = edits.setdefault('基本情報', {})
        for row in basic.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                current = basic_values.get(cell.coordinate, cell.value)   # 前段の置換結果を優先して重ねる
                if OLD_PAST_GRANT_NOTE in current:
                    basic_values[cell.coordinate] = current.replace(OLD_PAST_GRANT_NOTE, PAST_GRANT_NOTE)
    if 'シート1' in wb.sheetnames:
        table = wb['シート1']
        for ref, header in REQUIREMENTS_HEADERS.items():
            if table[ref].value != header:
                raise ValueError(f'{name} のシート1 {ref} が想定の見出しではありません。')
        edits.setdefault('シート1', {}).update(REQUIREMENTS_TABLE)


def build(official_source_dir: Path) -> list[dict]:
    forms_dir = ROOT / '補助金加点' / '第7次以降'
    forms_dir.mkdir(parents=True, exist_ok=True)
    log = []
    for official, local in FORMS.items():
        source, output = official_source_dir / official, forms_dir / local
        if not source.is_file():
            raise FileNotFoundError(f'取得済みの公式原本が必要です: {source.name}')
        shutil.copyfile(source, output)
        assert _hash(source) == _hash(output)
        log.append({'official_url': 'https://it-shien.smrj.go.jp/pdf/' + official,
                    'output': output.relative_to(ROOT).as_posix(), 'sha256': _hash(output)})
    master = openpyxl.load_workbook(forms_dir / '加点措置①用.xlsx', read_only=True, data_only=True)
    amounts = {str(row[0]): int(row[2]) for row in list(master['マスタ'].values)[1:48]}
    master.close()
    assert len(amounts) == 47
    tools_dir = ROOT / 'ツール'
    source_names = list(APPLICATION_FILENAMES.values())
    source_names += [p.name for p in tools_dir.iterdir() if p.suffix == '.xlsx' and p.name.startswith('ヒアリングシート2026_')]
    for name in source_names:
        source = tools_dir / name
        wb = openpyxl.load_workbook(source, read_only=True, data_only=False)
        edits = {}
        for ws in wb:
            values = {}
            for row in ws:
                for cell in row:
                    if cell.data_type == 'f' or not isinstance(cell.value, str):
                        continue
                    changed = _text(cell.value)
                    for pref, amount in amounts.items():
                        short = pref if pref == '北海道' else pref[:-1]
                        if re.fullmatch(re.escape(short) + r'/\d+円', changed):
                            changed = f'{short}/{amount}円'
                    if changed != cell.value:
                        values[cell.coordinate] = changed
            if values:
                edits[ws.title] = values
        template_type = next((key for key, filename in APPLICATION_FILENAMES.items() if filename == name), None)
        formulas = {}
        if template_type:
            base = BASE_ROWS[template_type]
            fields = edits.setdefault(SHINSEI, {})
            fields[f'B{base}'] = '事業計画期間における給与支給総額\n（0年度目・申請年度の見込み）'
            for year in range(1, 4):
                fields[f'B{base + year}'] = f'給与支給総額\n（{year}年度目の確認済み計画数値）'
            fields[f'B{base - 1}'] = '0年度目の対象従業員数（確認済みの正社員換算人数）'
            for sheet_name in wb.sheetnames:
                if '給与支給総額計算' in sheet_name:
                    original = str(wb[sheet_name]['A2'].value or '')
                    edits.setdefault(sheet_name, {})['A2'] = original + '\n第7次：生産性は前事業年度実績、給与0年度目は申請年度の見込み。給与の提出値は「申請前提・給与見込」を確認。旧計算式を申請値へ流用しない。'
            if PRODUCTIVITY in wb:
                edits.setdefault(PRODUCTIVITY, {}).update({
                    'B27': '前事業年度（生産性の基準実績）',
                    'C27': '1年度目（申請時点の翌事業年度）確認済み計画',
                    'D27': '2年度目（申請時点の翌々事業年度）確認済み計画',
                    'E27': '3年度目（申請時点の3事業年度後）確認済み計画'})
            formulas = _application_edits(wb, template_type, edits)
        elif '確認事項' in wb:
            _hearing_edits(wb, unicodedata.normalize('NFC', name), edits)
        wb.close()
        log.append(_patch_book(source, tools_dir / '第7次以降' / name, edits, formulas))
    source = tools_dir / '加点判定用賃金台帳テンプレート.xlsx'
    wb = openpyxl.load_workbook(source, read_only=True, data_only=False)
    sheet = '加点判定用明細'
    values = {'B4': '第7次以降用：2025年10月〜2026年9月＋実際の申請月の前月。非在籍は「非在籍」と明記。空欄・0円は未確認として扱います。',
              'AG2': '2026_rounds_7_plus', 'AG3': 7}
    for idx in range(12):
        year, month = (2025, idx + 10) if idx < 3 else (2026, idx - 2)
        label = f'令和{year - 2018}年{month}月'
        values[f'{openpyxl.utils.get_column_letter(6 + idx)}6'] = label + '\n基本給'
        values[f'{openpyxl.utils.get_column_letter(19 + idx)}6'] = label + '\n労働時間'
    edits = {sheet: values}
    if '記入ルール' in wb:
        edits['記入ルール'] = dict(LEDGER_RULES)
    wb.close()
    log.append(_patch_book(source, tools_dir / '第7次以降' / source.name, edits))
    return log


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--official-source-dir', type=Path, required=True)
    parser.add_argument('--log', type=Path, required=True)
    args = parser.parse_args()
    result = build(args.official_source_dir)
    args.log.parent.mkdir(parents=True, exist_ok=True)
    args.log.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'第7次候補原本を作成・再読取確認：公式3本、社内原本8本、専用台帳1本。ログ {args.log}')


if __name__ == '__main__':
    main()
