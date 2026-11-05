# -*- coding: utf-8 -*-
"""加点判定用賃金台帳テンプレートへの書き出し。

read_bonus_wage_ledger（wage_reader.py）と対になる writer。レイアウト定数は
wage_reader.BWL_* を単一の真実として共有する（列がズレると判定が壊れるため）。
入力は BonusWageEmployee（(年,月)→基本給・時間換算の対象月額 を保持）。月列は暦月固定。
テンプレートが AH〜AT列（時間換算の対象月額）を持つ場合（第7次以降の台帳）は、そこへ対象月額を書く。
判定期間・基準月・金額は公募回ごとに subsidy_rules.rules_for_round から取る
（第7次以降：令和7年10月〜令和8年9月・令和8年6月比55円。
第1〜6次（旧回）の訂正用：令和6年10月〜令和7年9月・令和7年7月比63円）。
"""
from __future__ import annotations

import datetime as _dt
import logging
from pathlib import Path

import openpyxl

from .wage_reader import (
    ABSENT_MARKER,
    BonusWageEmployee,
    BWL_APPYM_CELL,
    BWL_COL_BASIS_LATEST,
    BWL_COL_BASIS_WINDOW_START,
    BWL_COL_EMPTYPE,
    BWL_COL_HOURS,
    BWL_COL_HOURS_LATEST,
    BWL_COL_HOURS_WINDOW_START,
    BWL_COL_LATEST,
    BWL_COL_NAME,
    BWL_COL_NO,
    BWL_COL_WINDOW_START,
    BWL_DATA_START_ROW,
    BWL_PREF_CELL,
    BWL_SHEET_NAME,
    BWL_HEADER_ROW,
    BWL_RULE_CELL,
    BWL_ROUND_CELL,
    parse_ym_header,
    prev_month,
    _has_hourly_basis_block,
)
from .subsidy_rules import rules_for_round

logger = logging.getLogger(__name__)


def template_has_hourly_basis(template_path: Path, call_round: int) -> bool:
    """台帳テンプレートが AH〜AT列（時間換算の対象月額）を持つか（第7次以降のテンプレートは持つ）。"""
    wb = openpyxl.load_workbook(str(template_path))
    try:
        if BWL_SHEET_NAME not in wb.sheetnames:
            return False
        return _has_hourly_basis_block(wb[BWL_SHEET_NAME], rules_for_round(call_round).bonus1_window)
    finally:
        wb.close()


def write_bonus_wage_ledger(
    employees: list[BonusWageEmployee],
    template_path: Path,
    output_path: Path,
    *,
    prefecture: str = '',
    application_ym: tuple[int, int] | None = None,
    call_round: int = 6,
) -> Path:
    """BonusWageEmployee リストを加点判定用賃金台帳テンプレに書き出す。

    C2=都道府県、C3=交付申請月（日付）。各従業員行に No/氏名/雇用形態/月間所定労働時間/
    公募回の加点①判定期間12か月の基本給（F〜Q）/交付申請直近月の基本給（R）を書く。
    判定期間は rules_for_round(call_round).bonus1_window（第7次以降は令和7年10月〜令和8年9月、
    第1〜6次（旧回）の訂正用は令和6年10月〜令和7年9月）。テンプレの見出しと合わなければ止める。
    call_round の既定値6は旧回の訂正用の既存呼び出しのため。第7次以降は必ず実際の公募回を渡す。
    monthly_hours_override（時給制等の月別労働時間）がある月は S〜AD/AE にも書き出し、
    読み戻し時に E列固定値で時間換算給与が歪むのを防ぐ。判定期間外かつ直近月以外の
    override は判定に使われないため書き出さない（往復で落ちるのは仕様）。
    テンプレートに AH〜AT列があれば、hourly_basis_separate の従業員の時間換算の対象月額を書く。
    対象月額を持たない従業員（旧形式）の AH〜AT列は空欄のまま（基本給で埋めない＝読み戻すと未確認）。
    AH〜AT列のない旧回用テンプレート（月額1つ。旧回の記入ルールは「所定内賃金で記入」）には、
    hourly_basis_separate の従業員は時間換算の対象月額を F〜R列に書く（時間換算給与を正しく保つため。
    様式の【基本給】欄は原本の基本給に手で直す運用のまま）。
    """
    if Path(template_path).resolve() == Path(output_path).resolve():
        raise ValueError('台帳原本と出力先は別のファイルにしてください。')
    rules = rules_for_round(call_round)
    wb = openpyxl.load_workbook(str(template_path))
    if BWL_SHEET_NAME not in wb.sheetnames:
        wb.close()
        raise ValueError(f'専用シート「{BWL_SHEET_NAME}」がありません。')
    ws = wb[BWL_SHEET_NAME]
    for j, ym in enumerate(rules.bonus1_window):
        if parse_ym_header(ws.cell(BWL_HEADER_ROW, BWL_COL_WINDOW_START + j).value) != ym:
            wb.close()
            raise ValueError('台帳テンプレートの年月見出しが指定公募回と一致しません。')
    has_basis_block = _has_hourly_basis_block(ws, rules.bonus1_window)
    ws.cell(*BWL_RULE_CELL, value=rules.rule_id)
    ws.cell(*BWL_ROUND_CELL, value=call_round)

    if prefecture:
        ws.cell(*BWL_PREF_CELL, value=prefecture)
    if application_ym:
        cell = ws.cell(BWL_APPYM_CELL[0], BWL_APPYM_CELL[1],
                       _dt.datetime(application_ym[0], application_ym[1], 1))
        cell.number_format = 'yyyy/mm'
    latest_ym = prev_month(application_ym) if application_ym else None

    for i, emp in enumerate(employees):
        row = BWL_DATA_START_ROW + i
        ws.cell(row, BWL_COL_NO, emp.no if emp.no else i + 1)
        ws.cell(row, BWL_COL_NAME, emp.name)
        ws.cell(row, BWL_COL_EMPTYPE, emp.employment_type or '')
        ws.cell(row, BWL_COL_HOURS).value = None
        if emp.scheduled_hours is not None:
            ws.cell(row, BWL_COL_HOURS, round(emp.scheduled_hours, 4))
        amount_columns = [(ym, BWL_COL_WINDOW_START + j, BWL_COL_BASIS_WINDOW_START + j)
                          for j, ym in enumerate(rules.bonus1_window)]
        if latest_ym is not None:
            amount_columns.append((latest_ym, BWL_COL_LATEST, BWL_COL_BASIS_LATEST))
        for ym, base_col, basis_col in amount_columns:
            if has_basis_block:
                targets = [(base_col, emp.monthly_base)]
                if emp.hourly_basis_separate:
                    targets.append((basis_col, emp.monthly_hourly_basis))
            else:
                single = emp.monthly_hourly_basis if emp.hourly_basis_separate else emp.monthly_base
                targets = [(base_col, single)]
            for col, store in targets:
                value = store.get(ym)
                if emp.monthly_presence.get(ym) is False:
                    ws.cell(row, col, ABSENT_MARKER)
                elif value is not None:
                    ws.cell(row, col, round(value))
        # 時間は小数第4位まで保持する。第1位丸めだと 57:45→57.8h のように歪み、時間換算給与が
        # 1〜2円ズレて加点②の境界判定（第7次以降55円、第1〜6次（旧回）63円）を覆し得る。
        for j, ym in enumerate(rules.bonus1_window):
            h = emp.monthly_hours_override.get(ym)
            if h is not None and h > 0:
                ws.cell(row, BWL_COL_HOURS_WINDOW_START + j, round(h, 4))
        if latest_ym is not None:
            h = emp.monthly_hours_override.get(latest_ym)
            if h is not None and h > 0:
                ws.cell(row, BWL_COL_HOURS_LATEST, round(h, 4))

    wb.save(str(output_path))
    wb.close()
    logger.info(f'加点判定用賃金台帳保存: {output_path}（{len(employees)}名）')
    return output_path
