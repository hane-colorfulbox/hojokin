"""必要書類チェックの新旧台帳・実申請月・3期またぎを合成データで確認。API不使用。"""
import importlib.util
import json
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
ROOT = Path(__file__).resolve().parents[1]
CHECK = ROOT / '.claude/skills/case-docs-check/review/check_docs.py'
spec = importlib.util.spec_from_file_location('case_docs_rounds_under_test', CHECK)
mod = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = mod
spec.loader.exec_module(mod)


def make_ledger(path, call_round, app_month='2026/11', mixed_hours=False):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = mod.BWL_SHEET
    ws['C2'] = '東京都'
    ws['C3'] = app_month
    if call_round >= 7:
        ws['AG2'] = '2026_rounds_7_plus'
        ws['AG3'] = call_round
    for col, text in enumerate(mod.bonus_headers(call_round), 2):
        ws.cell(6, col, text)
    hours_round = 6 if mixed_hours else call_round
    for col, text in enumerate(mod.bonus_headers(hours_round)[4:-1], 19):
        ws.cell(6, col, text.replace('基本給', '労働時間'))
    ws['AE6'] = '交付申請直近月\n労働時間'
    wb.save(path)
    wb.close()


def check_ledger(path, call_round, application_ym=(2026, 11)):
    report = mod.Report()
    detected = {'wage_ledger': [{'name': path.name, 'local_path': str(path), 'ext': '.xlsx'}]}
    mod.check_excel_structures(report, detected, 'katen', '不明', '不明',
                               call_round, application_ym)
    return {f['code'] for f in report.findings if f['level'] == 'FAIL'}


def main():
    assert mod.bonus_headers(6)[4] == '令和6年10月\n基本給'
    assert mod.bonus_headers(7)[4] == '令和7年10月\n基本給'
    assert mod.bonus_headers(8)[-2] == '令和8年9月\n基本給'
    assert mod.parse_application_month('2026/11') == (2026, 11)
    assert mod.parse_application_month('2026/13') is None
    assert mod.parse_application_month('11月') is None
    months = mod.required_wage_months(7, (2026, 12))
    assert (2025, 10) in months and (2026, 6) in months and (2026, 11) in months
    assert len(months) == 13
    # 10月末決算：2025/10・2026/6・2026/11の期末年が3種類になる。
    fiscal_ends = {y if m <= 10 else y + 1 for y, m in months}
    assert fiscal_ends == {2025, 2026, 2027}
    assert (2026, 12) in mod.required_wage_months(8, (2027, 1))
    try:
        mod.bonus_headers(0)
        raise AssertionError('invalid call round was accepted')
    except ValueError:
        pass
    with tempfile.TemporaryDirectory(prefix='docs-rounds-') as tmp:
        folder = Path(tmp)
        path = folder / '合成_加点判定用賃金台帳.xlsx'
        make_ledger(path, 6)
        assert not check_ledger(path, 6)
        assert 'TPL-BONUS-HEADER' in check_ledger(path, 7)
        make_ledger(path, 7)
        assert not check_ledger(path, 7)
        assert 'BONUS-ROUND-MISMATCH' in check_ledger(path, 8)
        assert 'TPL-BONUS-HEADER' in check_ledger(path, 6)
        make_ledger(path, 7, mixed_hours=True)
        assert 'TPL-BONUS-HOURS-HEADER' in check_ledger(path, 7)
        make_ledger(path, 7, app_month='2026/12')
        assert 'BONUS-APPYM-MISMATCH' in check_ledger(path, 7)
        make_ledger(path, 7, app_month='2026/13')
        assert 'BONUS-APPYM-INVALID' in check_ledger(path, 7)
        make_ledger(path, 7)
        out = folder / 'check.json'
        cp = subprocess.run([sys.executable, '-B', str(CHECK), '--local', str(folder),
                             '--task', 'katen', '--call-round', '7', '--app-month', '2026/11',
                             '--json', str(out)], capture_output=True, text=True, encoding='utf-8')
        assert cp.returncode == 0, cp.stdout + cp.stderr
        result = json.loads(out.read_text(encoding='utf-8'))
        assert result['context']['call_round'] == 7
        assert result['context']['application_ym'] == [2026, 11]
        assert [2026, 10] in result['context']['bonus_required_months']
        cp = subprocess.run([sys.executable, '-B', str(CHECK), '--local', str(folder),
                             '--task', 'katen', '--json', str(out)],
                            capture_output=True, text=True, encoding='utf-8')
        result = json.loads(out.read_text(encoding='utf-8'))
        assert any(f['code'] == 'CALL-ROUND-UNCONFIRMED' for f in result['findings'])
    print('PASS: 新旧ヘッダー、時間列、実申請月、必要月、3期またぎ、年またぎ、未確認回の報告')
    return 0


def test_document_rounds():
    assert main() == 0


if __name__ == '__main__':
    raise SystemExit(main())
