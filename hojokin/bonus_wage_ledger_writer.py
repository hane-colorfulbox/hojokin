# -*- coding: utf-8 -*-
"""加点判定用賃金台帳テンプレートへの書き出し。

read_bonus_wage_ledger（wage_reader.py）と対になる writer。レイアウト定数は
wage_reader.BWL_* を単一の真実として共有する（列がズレると判定が壊れるため）。
入力は BonusWageEmployee（(年,月)→基本給 を保持）。月列は暦月固定。
"""
from __future__ import annotations

import datetime as _dt
import logging
from pathlib import Path

import openpyxl

from .wage_reader import (
    BonusWageEmployee,
    BONUS1_WINDOW,
    BWL_APPYM_CELL,
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
)
from .subsidy_rules import rules_for_round

logger = logging.getLogger(__name__)


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
    令和6年10月〜令和7年9月の基本給（F〜Q）/交付申請直近月の基本給（R）を書く。
    monthly_hours_override（時給制等の月別労働時間）がある月は S〜AD/AE にも書き出し、
    読み戻し時に E列固定値で時間換算給与が歪むのを防ぐ。BONUS1_WINDOW 外かつ直近月以外の
    override は判定に使われないため書き出さない（往復で落ちるのは仕様）。
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
        for j, ym in enumerate(rules.bonus1_window):
            base = emp.monthly_base.get(ym)
            if emp.monthly_presence.get(ym) is False:
                ws.cell(row, BWL_COL_WINDOW_START + j, '非在籍')
            elif base is not None:
                ws.cell(row, BWL_COL_WINDOW_START + j, round(base))
        if latest_ym is not None:
            base = emp.monthly_base.get(latest_ym)
            if emp.monthly_presence.get(latest_ym) is False:
                ws.cell(row, BWL_COL_LATEST, '非在籍')
            elif base is not None:
                ws.cell(row, BWL_COL_LATEST, round(base))
        # 時間は小数第4位まで保持する。第1位丸めだと 57:45→57.8h のように歪み、
        # 時間換算給与が1〜2円ズレて加点②の63円境界判定を覆し得る。
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
