# -*- coding: utf-8 -*-
"""公式様式 Excel を「原本コピー＋対象セルの値のみ差し替え」で埋める外科的パッチ。

openpyxl の load→save はブック全体を再構築するため、公式様式が内包する
図形・記入例画像（xl/drawings, xl/media）・秘密度ラベル（docMetadata/LabelInfo.xml）・
クエリテーブル・印刷設定などが消えたり書式が変わったりする。
本モジュールは原本の ZIP 構成を丸ごと温存し、対象ワークシート XML 内の
指定セルの値だけを書き換える。成果物は人間が原本ファイルに直接入力した
ものと同等になる（差分は入力セルと再計算フラグのみ）。

制約:
- 値の書き込みのみ（数式・書式・行列構成は変更しない）
- 書き込み先の行が原本 XML に存在しない場合はエラー
  （公式様式の集計式が及ばない行への書き込みは無意味なため、黙って書かない）
"""
from __future__ import annotations

import datetime
import logging
import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import to_excel

logger = logging.getLogger(__name__)

_CELL_REF_RE = re.compile(r'^([A-Z]+)(\d+)$')

# 数式セルの計算結果キャッシュ <v>。数式(<f>)の直後の <v> のみ対象
# （値のみセル <c><v>..</v></c> や inlineStr は <f> を持たないので非対象）。
_FORMULA_CACHE_RE = re.compile(r'(<f\b[^>]*>[^<]*</f>|<f\b[^>]*/>)\s*<v>[^<]*</v>')


def _split_ref(ref: str) -> tuple[str, int]:
    m = _CELL_REF_RE.match(ref)
    if not m:
        raise ValueError(f'セル参照が不正です: {ref}')
    return m.group(1), int(m.group(2))


def _resolve_sheet_part(zf: zipfile.ZipFile, sheet_index: int = 0) -> str:
    """workbook.xml のシート定義順から sheet_index 番目のシートのパート名を解決する。"""
    wb_xml = zf.read('xl/workbook.xml').decode('utf-8')
    sheet_tags = re.findall(r'<sheet [^>]*/>', wb_xml)
    if sheet_index >= len(sheet_tags):
        raise ValueError(f'シート {sheet_index} が存在しません（{len(sheet_tags)}シート）')
    rid_m = re.search(r'r:id="([^"]+)"', sheet_tags[sheet_index])
    if not rid_m:
        raise ValueError('workbook.xml のシート定義に r:id がありません')
    rid = rid_m.group(1)

    rels_xml = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rel_m = re.search(rf'<Relationship [^>]*Id="{re.escape(rid)}"[^>]*/>', rels_xml)
    if not rel_m:
        raise ValueError(f'workbook.xml.rels に {rid} がありません')
    target_m = re.search(r'Target="([^"]+)"', rel_m.group(0))
    target = target_m.group(1)
    if target.startswith('/'):
        return target.lstrip('/')
    return 'xl/' + target


def _merge_anchor(sheet_xml: str, ref: str) -> str:
    """ref が結合セル範囲の内側なら範囲左上（アンカー）の参照を返す。"""
    col_s, row = _split_ref(ref)
    col = column_index_from_string(col_s)
    for m in re.finditer(r'<mergeCell ref="([A-Z]+\d+):([A-Z]+\d+)"/>', sheet_xml):
        c1, r1 = _split_ref(m.group(1))
        c2, r2 = _split_ref(m.group(2))
        min_c, max_c = column_index_from_string(c1), column_index_from_string(c2)
        if r1 <= row <= r2 and min_c <= col <= max_c:
            return m.group(1)
    return ref


def _serialize_value(value) -> tuple[str, str]:
    """値を (t属性文字列, セル内側XML) に変換する。"""
    if isinstance(value, bool):
        return (' t="b"', f'<v>{1 if value else 0}</v>')
    if isinstance(value, (datetime.datetime, datetime.date)):
        serial = to_excel(value)
        if isinstance(serial, float) and serial.is_integer():
            serial = int(serial)
        return ('', f'<v>{serial}</v>')
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if isinstance(value, (int, float)):
        return ('', f'<v>{value}</v>')
    text = str(value)
    space = ' xml:space="preserve"' if text != text.strip() else ''
    return (' t="inlineStr"', f'<is><t{space}>{escape(text)}</t></is>')


def _build_cell_xml(ref: str, style: str | None, value) -> str:
    t_attr, inner = _serialize_value(value)
    s_attr = f' s="{style}"' if style else ''
    return f'<c r="{ref}"{s_attr}{t_attr}>{inner}</c>'


def _patch_row(row_xml: str, row_cells: dict[str, object]) -> str:
    """1つの <row>...</row> ブロック内の対象セルを差し替える。"""
    for ref, value in sorted(row_cells.items(),
                             key=lambda kv: column_index_from_string(_split_ref(kv[0])[0])):
        # [^>]*? は非貪欲必須（貪欲だと <c .../> の / を属性として飲み込み、
        # 後続セルまで巻き込んで置換してしまう）
        cell_re = re.compile(rf'<c r="{ref}"[^>]*?(?:/>|>.*?</c>)', re.S)
        m = cell_re.search(row_xml)
        if m:
            s_m = re.search(r'\ss="(\d+)"', m.group(0))
            new_cell = _build_cell_xml(ref, s_m.group(1) if s_m else None, value)
            row_xml = row_xml[:m.start()] + new_cell + row_xml[m.end():]
        else:
            # 行はあるがセル要素が無い（書式なし空セル）→ 列順を保って挿入
            new_cell = _build_cell_xml(ref, None, value)
            target_col = column_index_from_string(_split_ref(ref)[0])
            insert_at = None
            for cm in re.finditer(r'<c r="([A-Z]+)(\d+)"', row_xml):
                if column_index_from_string(cm.group(1)) > target_col:
                    insert_at = cm.start()
                    break
            if insert_at is None:
                end_m = re.search(r'</row>$', row_xml)
                if end_m:
                    insert_at = end_m.start()
                else:  # <row .../> 自己終結
                    open_m = re.match(r'<row [^>]*?/>', row_xml)
                    row_xml = open_m.group(0)[:-2].rstrip() + '>' + new_cell + '</row>'
                    continue
            row_xml = row_xml[:insert_at] + new_cell + row_xml[insert_at:]
    return row_xml


def _patch_sheet_xml(sheet_xml: str, cell_values: dict[str, object]) -> str:
    """シート XML の対象セルの値を差し替える（他は一切変更しない）。"""
    by_row: dict[int, dict[str, object]] = {}
    for ref, value in cell_values.items():
        if value is None:
            continue
        anchored = _merge_anchor(sheet_xml, ref.upper())
        _, row = _split_ref(anchored)
        by_row.setdefault(row, {})[anchored] = value

    for row_num in sorted(by_row):
        open_re = re.compile(rf'<row r="{row_num}"[^>]*?(/>|>)')
        m = open_re.search(sheet_xml)
        if not m:
            raise ValueError(
                f'テンプレートの {row_num} 行目が原本に存在しません。'
                f'公式様式の入力可能行数を超えています（集計式の範囲外のため書き込みできません）。'
            )
        if m.group(1) == '/>':
            start, end = m.span()
        else:
            start = m.start()
            end = sheet_xml.index('</row>', m.end()) + len('</row>')
        patched = _patch_row(sheet_xml[start:end], by_row[row_num])
        sheet_xml = sheet_xml[:start] + patched + sheet_xml[end:]
    return sheet_xml


def _strip_formula_cache(sheet_xml: str) -> str:
    """数式セルの計算結果キャッシュ <v> を除去し、開いた時に必ず再計算させる。

    公式様式は空データ時の計算結果（例: 判定=対象外、事業場所在地=未選択、MIN=0）を
    キャッシュとして内包する。値だけ差し替えても数式セルのキャッシュは古いまま残り、
    fullCalcOnLoad を立てても一部ビューア（LibreOffice 等）は再計算せず古い表示を出す。
    数式・書式・様式・入力値は一切変えず、数式の計算結果キャッシュのみ捨てることで、
    どのソフトでも開いた瞬間に再計算させる。値のみのセル（<f>なし）は対象外。
    """
    return _FORMULA_CACHE_RE.sub(r'\1', sheet_xml)


def _patch_workbook_xml(wb_xml: str) -> str:
    """開いた時に全数式を再計算させる（入力値に依存する判定式・MIN式の陳腐値対策）。"""
    m = re.search(r'<calcPr [^>]*/>', wb_xml)
    if not m:
        logger.warning('workbook.xml に calcPr が無いため fullCalcOnLoad を設定できません')
        return wb_xml
    if 'fullCalcOnLoad' in m.group(0):
        return wb_xml
    new_tag = m.group(0)[:-2].rstrip() + ' fullCalcOnLoad="1"/>'
    return wb_xml[:m.start()] + new_tag + wb_xml[m.end():]


def patch_xlsx(
    template_path: Path,
    output_path: Path,
    cell_values: dict[str, object],
    sheet_index: int = 0,
) -> Path:
    """原本 xlsx をコピーし、対象シートの指定セルの値だけを差し替えて保存する。

    cell_values: {'D5': datetime(...), 'C17': '氏名', 'F17': 180000, ...}
    値が None の項目はスキップ（セルは原本のまま）。
    """
    with zipfile.ZipFile(str(template_path), 'r') as zin:
        sheet_part = _resolve_sheet_part(zin, sheet_index)
        patched_sheet = _patch_sheet_xml(
            zin.read(sheet_part).decode('utf-8'), cell_values
        )
        # 公式様式が内包する古い計算結果キャッシュ（空データ時の判定=対象外等）を捨て、
        # 開いた時に必ず再計算させる（数式・書式・様式・入力値は不変）。
        patched_sheet = _strip_formula_cache(patched_sheet)
        replacements = {
            sheet_part: patched_sheet.encode('utf-8'),
            'xl/workbook.xml': _patch_workbook_xml(
                zin.read('xl/workbook.xml').decode('utf-8')
            ).encode('utf-8'),
        }
        with zipfile.ZipFile(str(output_path), 'w') as zout:
            for item in zin.infolist():
                data = replacements.get(item.filename, zin.read(item.filename))
                zout.writestr(item, data)
    return output_path
