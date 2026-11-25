# -*- coding: utf-8 -*-
"""ヒアリングシートExcel読み取り"""
from __future__ import annotations

import logging
from pathlib import Path
import openpyxl

from .choice_masters import decode_choice_field

logger = logging.getLogger(__name__)

# 行をそろえる目印の見出しと、その下に続く表明者3人の見出し（どの版・枠でも同じ並び）
HEARING_ANCHOR_LABEL = '表明を行った従業員名'
HEARING_DECLARER_LABELS = ('従業員代表者', '給与担当者', '事業所内最低賃金者')

# 全角→半角変換テーブル
_ZEN2HAN = str.maketrans(
    '０１２３４５６７８９'
    'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ'
    'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ'
    '－（）　＠．／：，',
    '0123456789'
    'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    'abcdefghijklmnopqrstuvwxyz'
    '-() @./：,',
)


def normalize_value(value):
    """全角数字・英字を半角に変換。数値化できるものは数値にする。"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    s = str(value).translate(_ZEN2HAN).strip()
    # 先頭0の数字列は電話番号等なので文字列のまま返す
    s_clean = s.replace(',', '').replace('円', '').replace('人', '').replace('時間', '')
    if s_clean.startswith('0') and s_clean.isdigit() and len(s_clean) >= 2:
        return s_clean
    try:
        if '.' in s_clean:
            return float(s_clean)
        return int(s_clean)
    except ValueError:
        return s


def normalize_phone(value) -> str | None:
    """数値の電話番号を先頭0補完した文字列に変換"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        s = str(int(value))
        if len(s) == 9:
            return '0' + s
        if len(s) == 10 and not s.startswith('0'):
            return '0' + s
        return s
    return str(value)


def read_hearing_sheet(path: Path) -> dict[int, any]:
    """
    ヒアリングシートを読み取り、{行番号: 値} の辞書を返す。
    シート名は「基本情報」を使用。
    """
    wb = openpyxl.load_workbook(path, data_only=True)

    # シート名を探す（「基本情報」優先、なければ最初のシート）
    sheet_name = None
    for name in wb.sheetnames:
        if '基本情報' in name:
            sheet_name = name
            break
    if sheet_name is None:
        sheet_name = wb.sheetnames[0]
        logger.warning(f'「基本情報」シートが見つからず。{sheet_name} を使用')

    ws = wb[sheet_name]
    data = {}

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        row_num = row[0].row
        label = row[1].value if len(row) > 1 else None  # B列
        value = row[2].value if len(row) > 2 else None   # C列

        if label is not None:
            # 通常枠の選択肢フィールド（強み/弱み/IT投資プロセス）は番号入力を
            # 「N 本文」へ復元する。対象外・本文入りは None が返り従来処理に委ねる。
            decoded = decode_choice_field(label, value)
            data[row_num] = {
                'label': str(label).strip(),
                'value': decoded if decoded is not None else normalize_value(value),
            }

    wb.close()
    logger.info(f'ヒアリングシート読込: {path.name}, {len(data)}行')
    return data


def align_hearing_rows(hearing_data: dict, anchor_row: int) -> tuple[dict, str | None]:
    """ヒアリングの行番号を、対応表が前提とする版の行番号にそろえる。

    第7次の版（2026-11-09）は「◼️表明を行った従業員名」の上に昇給時期などの欄が増え、
    それより下の行がずれた。顧客が旧版・新版のどちらに書いても正しい欄を読むため、
    見出しの実際の行と anchor_row の差だけ、見出しから下の行を動かす。
    - 実際の見出しが下にある（対応表が知らない欄がある）: その欄は外し、下の行を上げる
    - 実際の見出しが上にある（対応表が前提とする欄がない）: 下の行を下げ、欠けた欄は空のまま
    戻り値は (そろえたヒアリング, 注意の文。問題なければ None)。
    見出しが見つからない・表明者3人の見出しが続かないときは、ずれて読むより転記しない方を選び、
    見出しから下の行を外して注意を返す。
    """
    if not hearing_data or not anchor_row:
        return hearing_data, None
    found = [row for row, rec in hearing_data.items() if HEARING_ANCHOR_LABEL in str(rec.get('label', ''))]
    if len(found) != 1:
        kept = {row: rec for row, rec in hearing_data.items() if row < anchor_row}
        note = (f'ヒアリングの「◼️表明を行った従業員名」の見出しが{"見つかりません" if not found else "複数あります"}。'
                f'行の位置を確かめられないため、{anchor_row}行目から下（表明を行った従業員名など）は転記していません。')
        logger.warning(note)
        return kept, note
    actual = found[0]
    delta = anchor_row - actual
    aligned = {}
    for row, rec in hearing_data.items():
        if row < min(actual, anchor_row):
            aligned[row] = rec
        elif row >= actual:
            aligned[row + delta] = rec
    names = [str(aligned.get(anchor_row + offset, {}).get('label', '')) for offset in (1, 2, 3)]
    if not all(label in name for label, name in zip(HEARING_DECLARER_LABELS, names)):
        kept = {row: rec for row, rec in aligned.items() if row < anchor_row}
        note = ('ヒアリングの「◼️表明を行った従業員名」の下が「従業員代表者・給与担当者・事業所内最低賃金者」'
                f'の並びになっていないため、{anchor_row}行目から下は転記していません。様式が改変されていないか確認してください。')
        logger.warning(note)
        return kept, note
    if delta:
        logger.info(f'ヒアリングの版に合わせて行をそろえました（{actual}行目の見出しを{anchor_row}行目へ）')
    return aligned, None


def transfer_hearing_to_tenki(hearing_data: dict, ws_tenki, mapping: list[tuple[int, int, bool]]) -> int:
    """
    ヒアリングデータを転記シートに転記。
    mapping: [(ヒアリング行, 転記行, 電話番号変換フラグ), ...]
    転記した件数を返す。
    """
    from .template_filler import _safe_write_cell

    count = 0
    for h_row, t_row, is_phone in mapping:
        if h_row not in hearing_data:
            continue
        value = hearing_data[h_row]['value']
        if value is None:
            continue

        if is_phone:
            value = normalize_phone(value)

        _safe_write_cell(ws_tenki, t_row, 2, value)
        count += 1
        logger.debug(f'転記: ヒアリング行{h_row} → 転記行{t_row}: {value!r}')

    return count
