"""第7次以降の原本を、原本管理担当が更新した Drive 版の書き出しから再現可能に作る。Drive原本への書込みはしない。

2026-11-25〜: 申請書の原本は Drive 版（2026-11-09 更新。ヒアリングの昇給時期などの受け皿と加点①の設問が
増えた版）の書き出しを入力にし、第7次の直し（版表示・最賃の額・0年度目の文言・通常枠の生産性計画欄など）を当てる。
行の位置は config の第7次の対応表（get_mapping(種類, 7)）から引く。ヒアリングシートは Drive 版をそのまま写す。
入力フォルダには正式名（APPLICATION_FILENAMES とヒアリングシート2026_*.xlsx）で置き、通常枠個人だけは
scripts/patch_genpon_tsujo_kojin_from_drive.py の機械的修正を当てたものを置く。
社内原本はZIP内の表示セル・プルダウンの額・版メタデータ・通常枠の生産性計画欄だけを変更し、
図形・その他の数式・入力規則・保護を保持する。行・列の挿入削除はしない。Drive の書き出しには
docProps/core.xml がないため、版マーカーを付けるときにその部品と登録だけを足す。
公式加点様式は取得済みのバイト列をそのまま複製する（--official-source-dir を渡したときだけ）。
加点判定用の専用台帳だけは、右端に時間換算の対象月額の列（AH〜AT）を足す。既存の行・列は動かさず、
書式は原本にある番号を使う（styles.xml・workbook.xml は旧回の台帳と同じバイトのまま）。
専用台帳だけを作り直すときは --ledger-only を使う（Drive の書き出しは不要）。
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
from hojokin.template_versions import APPLICATION_FILENAMES, ROUND7_FIRST_CALL, ROUND7_SUBJECT
from hojokin.wage_reader import (BWL_BASIS_HEADER_KEYWORD, BWL_COL_BASIS_LATEST, BWL_COL_BASIS_WINDOW_START,
                                 BWL_DATA_START_ROW, BWL_HEADER_ROW)
from hojokin.xlsx_surgical import _patch_sheet_xml, _resolve_sheet_part

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

FORMS = {
    'it2026_kaitei_chingin_houkoku1_hojyoritsu_katen.xlsx': '補助率引き上げ・加点措置①用.xlsx',
    'it2026_kaitei_chingin_houkoku1_katen.xlsx': '加点措置①用.xlsx',
    'it2026_kaitei_chingin_houkoku2_katen.xlsx': '加点措置②用.xlsx',
}
HEARING_PREFIX = 'ヒアリングシート2026_'
CORE_PART = 'docProps/core.xml'
CONTENT_TYPES_PART = '[Content_Types].xml'
ROOT_RELS_PART = '_rels/.rels'
CORE_CONTENT_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml'
CORE_REL_TYPE = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
CORE_REL_ID = 'rIdHojokinCore'
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
LEDGER_RULES = {
    'B5': ('連番／氏名／「正社員」「パート・アルバイト」「役員」等。判定の母数は各月の常時使用する従業員'
           '（労働基準法第20条の解雇予告が必要な人）。役員・事業主本人と、その月に日雇い・2か月以内の期間雇用・'
           '季節的業務で4か月以内・試用期間14日以内の人は母数に入れない（所定の期間を超えて引き続き使用している人は入れる）。'),
    'A7': '令和7年10月〜令和8年9月 基本給（F〜Q列）・交付申請直近月 基本給（R列）',
    'B7': ('各暦月の基本給（本給）。職務手当などの手当は含めない（提出様式の【基本給】欄にそのまま書く）。'
           'R列は加点②用の交付申請月の前月。'
           '在籍していない月は「非在籍」と記入。空欄・0円の月は未確認として扱い、判定が確定しません。'
           '公募回・申請月と暦月の対応を確認し、金額・所定労働時間が分からない月は確定させない。'),
    'A8': '時間換算の対象月額（AH〜AS列・交付申請直近月はAT列）',
    'B8': ('時間換算給与の分子。各暦月に支払われた賃金から、'
           '厚生労働省「最低賃金額以上かどうかを確認する方法」で除く6種類（臨時に支払われる賃金・'
           '賞与など1か月を超える期間ごとに支払われる賃金・時間外/休日/深夜の割増賃金（固定残業代を含む）・'
           '精皆勤手当・通勤手当・家族手当）を除いた額。基本給と、職務手当・役職手当・資格手当などの手当を含む。'
           '基本給に固定残業代（みなし残業代）が含まれる場合はその分を除く（そのため基本給より少なくなることがある。'
           '基本給に合わせて引き上げない。少ない月は判定を止めずに注意を表示するので、理由を原本で確かめて報告に残す）。'
           '歩合給・出来高給がある人は「歩合・出来高以外の対象額＋歩合・出来高の額×（その月の労働時間（S〜AE列、'
           '空欄の月はE列）÷その月の総労働時間）」を入れる。総労働時間を原本で確かめられない月は空欄にする（判定は未確定）。'
           '手当・歩合給がなく、基本給に固定残業代も含まれない人は基本給と同じ額を入れる（空欄・0円の月は未確認）。'
           '基本給の欄に「非在籍」と書いた月は空欄でよい。'),
    'B10': ('本台帳には記入不要（ツールが「時間換算の対象月額（AH〜AT列）÷その月の労働時間」で自動算出し、'
            '提出様式の【基本給】欄には F〜R列、【時間換算給与】欄にはこの値を書きます。'
            '月別労働時間が空欄の月はE列の所定労働時間を使います）。'),
}
LEDGER_B4 = ('第7次以降用：2025年10月〜2026年9月＋実際の申請月の前月。F〜R列に基本給、AH〜AT列に時間換算の対象月額'
             '（その月の賃金から厚労省方式で除く6種類を除いた額。手当を含み、基本給に含まれる固定残業代は除く。'
             '歩合給・出来高給は記入ルールのとおり換算）を入力。非在籍は「非在籍」と明記。空欄・0円は未確認として扱います。')
LEDGER_BASIS_BANNER = ('時間換算の対象月額（必須）— 賃金から厚労省方式の6種類（割増賃金・通勤手当など）と'
                       '基本給に含む固定残業代を除いた額。歩合給は記入ルール参照')
LEDGER_LAST_ROW = 26   # 罫線済みデータ行の末尾（原本の既存仕様）
LEDGER_BASIS_WIDTH = 11
LEDGER_STYLE = {'banner': '9', 'header': '6', 'data': '7'}   # 原本の S5 / F6 / F7 の書式番号
# 再生成後に残っていてはいけない旧回・旧運用の表記（数値の一部の「…63円」は対象外）。
STALE_PATTERNS = (r'R[7７][.．][7７]', r'R[7７]年度改定', r'(?<![0-9０-９,，])[6６][3３]円',
                  r'令和[6６]年\s*[1１][0０]月', r'令和[7７]年度改定', r'適用補助率確認シート',
                  r'以内で雇用', r'任意の賃上げ計画加点）で雇用', r'登記から抜粋', r'リンクされているので触らない',
                  r'特に指定がなければ', r'シート作成日', r'支給なしは空欄', r'前事業年度分の賃金台帳をご提出',
                  r'3\.0％以上であることを確認', r'it2026_chingin_houkoku',
                  r'年平均成長率を3\.5％以上とすることが申請要件', r'毎月決まって払う')


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


def _add_hourly_basis_block(sheet_xml: str) -> str:
    """専用台帳の明細シートに AH〜AT列（時間換算の対象月額）の書式付き空セル・結合・列幅を足す。

    既存の行・列・セルは変えない。書式は原本の番号（S5/F6/F7 と同じ）を使い、styles.xml は触らない。
    値（見出し・バナー）は後段の _patch_sheet_xml が既存セルの書式を保って書く。
    """
    first, last = BWL_COL_BASIS_WINDOW_START, BWL_COL_BASIS_LATEST
    letters = [openpyxl.utils.get_column_letter(c) for c in range(first, last + 1)]
    if f'<c r="{letters[0]}{BWL_HEADER_ROW}"' in sheet_xml:
        raise ValueError('台帳の原本にすでに AH〜AT列があります。旧回の台帳を入力にしてください。')
    for name, style in LEDGER_STYLE.items():
        ref = {'banner': 'S5', 'header': 'F6', 'data': 'F7'}[name]
        if not re.search(rf'<c r="{ref}" s="{style}"', sheet_xml):
            raise ValueError(f'原本の {ref} の書式番号が想定（{style}）と違います。')

    def append_cells(xml: str, row: int, cells: str) -> str:
        match = re.search(rf'<row r="{row}"[^>]*>.*?</row>', xml, re.S)
        if match is None:
            raise ValueError(f'台帳の原本に {row} 行目がありません。')
        block = match.group(0)
        block = block[:-len('</row>')] + cells + '</row>'
        return xml[:match.start()] + block + xml[match.end():]

    sheet_xml = append_cells(sheet_xml, BWL_HEADER_ROW - 1,
                             f'<c r="{letters[0]}{BWL_HEADER_ROW - 1}" s="{LEDGER_STYLE["banner"]}"/>')
    sheet_xml = append_cells(sheet_xml, BWL_HEADER_ROW, ''.join(
        f'<c r="{col}{BWL_HEADER_ROW}" s="{LEDGER_STYLE["header"]}"/>' for col in letters))
    for row in range(BWL_DATA_START_ROW, LEDGER_LAST_ROW + 1):
        sheet_xml = append_cells(sheet_xml, row, ''.join(
            f'<c r="{col}{row}" s="{LEDGER_STYLE["data"]}" t="n"></c>' for col in letters))
    merge = f'<mergeCell ref="{letters[0]}{BWL_HEADER_ROW - 1}:{letters[-1]}{BWL_HEADER_ROW - 1}"/>'
    count = re.search(r'<mergeCells count="(\d+)">', sheet_xml)
    sheet_xml = sheet_xml.replace(count.group(0), f'<mergeCells count="{int(count.group(1)) + 1}">', 1)
    sheet_xml = sheet_xml.replace('</mergeCells>', merge + '</mergeCells>', 1)
    cols = ''.join(f'<col width="{LEDGER_BASIS_WIDTH}" customWidth="1" min="{c}" max="{c}"/>'
                   for c in range(first, last + 1))
    sheet_xml = sheet_xml.replace('</cols>', cols + '</cols>', 1)
    sheet_xml = re.sub(r'<dimension ref="([A-Z]+\d+):[A-Z]+(\d+)"/>',
                       lambda m: f'<dimension ref="{m.group(1)}:{letters[-1]}{m.group(2)}"/>', sheet_xml, count=1)
    return sheet_xml


def _assert_no_stale_text(path: Path) -> None:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    hits = [f'{ws.title}!{cell.coordinate}' for ws in wb for row in ws for cell in row
            if isinstance(cell.value, str) and cell.data_type != 'f'
            and any(re.search(p, cell.value) for p in STALE_PATTERNS)]
    wb.close()
    if hits:
        raise AssertionError(f'{path.name} に旧表記が残っています: {hits}')


def _new_core_xml() -> str:
    """Drive の書き出しにない docProps/core.xml を、版マーカー（件名）だけで作る。"""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
            ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
            ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:subject>{escape(ROUND7_SUBJECT)}</dc:subject></cp:coreProperties>')


def _register_core_part(content_types: str, root_rels: str) -> tuple[str, str]:
    """core.xml を足すときの [Content_Types].xml と _rels/.rels の登録。既にあれば変えない。"""
    if f'PartName="/{CORE_PART}"' not in content_types:
        content_types = content_types.replace(
            '</Types>', f'<Override PartName="/{CORE_PART}" ContentType="{CORE_CONTENT_TYPE}"/></Types>')
    if CORE_REL_TYPE not in root_rels:
        if f'Id="{CORE_REL_ID}"' in root_rels:
            raise ValueError(f'_rels/.rels に {CORE_REL_ID} が既にあります。')
        root_rels = root_rels.replace(
            '</Relationships>',
            f'<Relationship Id="{CORE_REL_ID}" Type="{CORE_REL_TYPE}" Target="{CORE_PART}"/></Relationships>')
    return content_types, root_rels


def _patch_book(source: Path, output: Path, edits: dict[str, dict[str, object]],
                formulas: dict[str, dict[str, str | None]] | None = None,
                structure: dict | None = None) -> dict:
    formulas = formulas or {}
    structure = structure or {}
    output.parent.mkdir(parents=True, exist_ok=True)
    added: dict[str, bytes] = {}
    with zipfile.ZipFile(source) as zin:
        wb = openpyxl.load_workbook(source, read_only=True, data_only=False)
        replacements = {}
        for sheet_name in dict.fromkeys([*structure, *edits, *formulas]):
            part = _resolve_sheet_part(zin, wb.sheetnames.index(sheet_name))
            xml = zin.read(part).decode('utf-8')
            if sheet_name in structure:
                xml = structure[sheet_name](xml)
            xml = _patch_sheet_xml(xml, edits.get(sheet_name, {}))
            replacements[part] = _patch_formula_xml(xml, formulas.get(sheet_name, {})).encode('utf-8')
        if CORE_PART not in zin.namelist():
            # Drive の書き出しは文書情報の部品を持たない。版マーカーの部品と登録だけを足す
            added[CORE_PART] = _new_core_xml().encode('utf-8')
            content_types, root_rels = _register_core_part(
                zin.read(CONTENT_TYPES_PART).decode('utf-8'), zin.read(ROOT_RELS_PART).decode('utf-8'))
            replacements[CONTENT_TYPES_PART] = content_types.encode('utf-8')
            replacements[ROOT_RELS_PART] = root_rels.encode('utf-8')
        else:
            core = zin.read(CORE_PART).decode('utf-8')
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
            replacements[CORE_PART] = core.encode('utf-8')
        for name in zin.namelist():
            if name.endswith('.rels'):
                raw = replacements.get(name, zin.read(name))
                linked = raw.replace(b'it2026_chingin_houkoku', b'it2026_kaitei_chingin_houkoku')
                if linked != raw:
                    replacements[name] = linked
        with zipfile.ZipFile(output, 'w') as zout:
            for info in zin.infolist():
                zout.writestr(info, replacements.get(info.filename, zin.read(info.filename)))
            for name, data in added.items():
                zout.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
        wb.close()
    # Reopen and confirm no row/sheet/formula/object changes outside explicit cells.
    with zipfile.ZipFile(source) as a, zipfile.ZipFile(output) as b:
        assert a.namelist() + list(added) == b.namelist()
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
    """申請内容の確認文・初期値・注記と版表示。セル番地は config の第7次の対応表から引く。"""
    rows = get_mapping(template_type, ROUND7_FIRST_CALL).shinsei
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


def _copy_official_forms(official_source_dir: Path | None, forms_dir: Path) -> list[dict]:
    """公式3様式。取得し直したときだけ複製し、それ以外は手元の版を記録する。"""
    forms_dir.mkdir(parents=True, exist_ok=True)
    log = []
    for official, local in FORMS.items():
        output = forms_dir / local
        if official_source_dir is not None:
            source = official_source_dir / official
            if not source.is_file():
                raise FileNotFoundError(f'取得済みの公式原本が必要です: {source.name}')
            shutil.copyfile(source, output)
            assert _hash(source) == _hash(output)
        elif not output.is_file():
            raise FileNotFoundError(f'公式様式がありません（--official-source-dir で取得済みの原本を渡す）: {output}')
        log.append({'official_url': 'https://it-shien.smrj.go.jp/pdf/' + official,
                    'output': output.relative_to(ROOT).as_posix(), 'sha256': _hash(output)})
    return log


def _text_edits(wb, amounts: dict[str, int]) -> dict:
    """旧回の表記と最賃の額（公式マスタの令和8年度改定額）の置換。数式は触らない。"""
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
    return edits


def _application_template_edits(wb, template_type: str, amounts: dict[str, int]) -> tuple[dict, dict]:
    edits = _text_edits(wb, amounts)
    base = get_mapping(template_type, ROUND7_FIRST_CALL).shinsei['wage_total_base']
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
    return edits, _application_edits(wb, template_type, edits)


def _find_input(drive_dir: Path, name: str) -> Path:
    """日本語名の濁点の合成・分解の違い（Windows/Mac）を吸収して入力を探す。"""
    want = unicodedata.normalize('NFC', name)
    hits = [p for p in drive_dir.iterdir() if unicodedata.normalize('NFC', p.name) == want]
    if len(hits) != 1:
        raise FileNotFoundError(f'Drive の書き出しフォルダに {name} がありません（または重複）: {drive_dir}')
    return hits[0]


def build(drive_dir: Path, official_source_dir: Path | None = None) -> list[dict]:
    drive_dir = drive_dir.resolve()
    forms_dir = ROOT / '補助金加点' / '第7次以降'
    log = _copy_official_forms(official_source_dir, forms_dir)
    master = openpyxl.load_workbook(forms_dir / '加点措置①用.xlsx', read_only=True, data_only=True)
    amounts = {str(row[0]): int(row[2]) for row in list(master['マスタ'].values)[1:48]}
    master.close()
    assert len(amounts) == 47
    tools_dir = ROOT / 'ツール'
    out_dir = tools_dir / '第7次以降'
    for template_type, name in APPLICATION_FILENAMES.items():
        source = _find_input(drive_dir, name)
        wb = openpyxl.load_workbook(source, read_only=True, data_only=False)
        edits, formulas = _application_template_edits(wb, template_type, amounts)
        wb.close()
        log.append(_patch_book(source, out_dir / name, edits, formulas))
    hearings = sorted(p for p in drive_dir.iterdir()
                      if p.suffix == '.xlsx' and unicodedata.normalize('NFC', p.name).startswith(HEARING_PREFIX))
    if len(hearings) != len(APPLICATION_FILENAMES):
        raise FileNotFoundError(f'ヒアリングシートは4様式そろえて置いてください: {[p.name for p in hearings]}')
    for source in hearings:
        # ヒアリングは原本管理担当の Drive 版が正。文言は変えずにそのまま写す
        output = out_dir / unicodedata.normalize('NFC', source.name)
        for stale in out_dir.iterdir():
            if stale != output and unicodedata.normalize('NFC', stale.name) == output.name:
                stale.unlink()
        shutil.copyfile(source, output)
        assert _hash(source) == _hash(output)
        log.append({'input': source.relative_to(ROOT).as_posix(), 'output': output.relative_to(ROOT).as_posix(),
                    'sha256': _hash(output), 'copied_as_is': True})
    log.append(build_ledger_template(tools_dir))
    return log


def build_ledger_template(tools_dir: Path) -> dict:
    """旧回の専用台帳から第7次以降の専用台帳を作る（見出し・記入ルール・AH〜AT列）。"""
    source = tools_dir / '加点判定用賃金台帳テンプレート.xlsx'
    wb = openpyxl.load_workbook(source, read_only=True, data_only=False)
    sheet = '加点判定用明細'
    values = {'B4': LEDGER_B4, 'AG2': '2026_rounds_7_plus', 'AG3': 7}
    basis_letter = openpyxl.utils.get_column_letter
    values[f'{basis_letter(BWL_COL_BASIS_WINDOW_START)}{BWL_HEADER_ROW - 1}'] = LEDGER_BASIS_BANNER
    for idx in range(12):
        year, month = (2025, idx + 10) if idx < 3 else (2026, idx - 2)
        label = f'令和{year - 2018}年{month}月'
        values[f'{basis_letter(6 + idx)}6'] = label + '\n基本給'
        values[f'{basis_letter(19 + idx)}6'] = label + '\n労働時間'
        values[f'{basis_letter(BWL_COL_BASIS_WINDOW_START + idx)}{BWL_HEADER_ROW}'] = (
            f'{label}\n{BWL_BASIS_HEADER_KEYWORD}')
    values[f'{basis_letter(BWL_COL_BASIS_LATEST)}{BWL_HEADER_ROW}'] = f'交付申請直近月\n{BWL_BASIS_HEADER_KEYWORD}'
    edits = {sheet: values}
    if '記入ルール' in wb:
        edits['記入ルール'] = dict(LEDGER_RULES)
    wb.close()
    return _patch_book(source, tools_dir / '第7次以降' / source.name, edits,
                       structure={sheet: _add_hourly_basis_block})


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--drive-export-dir', type=Path,
                        help='Drive 版の書き出しを正式名で置いたフォルダ（申請書4種類・ヒアリング4種類）')
    parser.add_argument('--official-source-dir', type=Path,
                        help='公式3様式を取得し直したときだけ渡す（省略時は 補助金加点/第7次以降/ の版を使う）')
    parser.add_argument('--ledger-only', action='store_true',
                        help='加点判定用の専用台帳だけを作り直す（Drive の書き出しは不要）')
    parser.add_argument('--log', type=Path, required=True)
    args = parser.parse_args()
    if args.ledger_only:
        result = [build_ledger_template(ROOT / 'ツール')]
        args.log.parent.mkdir(parents=True, exist_ok=True)
        args.log.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        print(f'第7次以降の専用台帳を作成・再読取確認。ログ {args.log}')
        return
    if args.drive_export_dir is None:
        parser.error('--drive-export-dir が必要です（専用台帳だけなら --ledger-only）')
    result = build(args.drive_export_dir, args.official_source_dir)
    args.log.parent.mkdir(parents=True, exist_ok=True)
    args.log.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'第7次以降の原本を作成・再読取確認：公式3本、申請書4本、ヒアリング4本、専用台帳1本。ログ {args.log}')


if __name__ == '__main__':
    main()
