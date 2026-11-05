"""case-docs-check の賃金状況報告シート版チェックを、リポジトリの新旧公式様式（原本コピー）で確かめる。
API不使用。顧客データなし（公式様式の白紙原本だけを使う）。"""
import importlib.util
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
ROOT = Path(__file__).resolve().parents[1]
CHECK = ROOT / '.claude/skills/case-docs-check/review/check_docs.py'
FORMS = ROOT / '補助金加点'
spec = importlib.util.spec_from_file_location('cdc_under_test', CHECK)
mod = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = mod
spec.loader.exec_module(mod)

OLD = {'s1h': FORMS / '補助率引き上げ・加点措置①用.xlsx', 's1': FORMS / '加点措置①用.xlsx',
       's2': FORMS / '加点措置②用.xlsx'}
NEW = {k: FORMS / '第7次以降' / v.name for k, v in OLD.items()}


def run(args, out):
    cp = subprocess.run([sys.executable, '-B', str(CHECK), *map(str, args), '--json', str(out)],
                        capture_output=True, text=True, encoding='utf-8')
    assert cp.returncode in (0, 1), cp.stdout + cp.stderr
    return cp, json.loads(out.read_text(encoding='utf-8'))


def run_local(folder, *args):
    return run(['--local', folder, '--task', 'katen', *args], folder.parent / f'{folder.name}.json')


def codes(data, level=None):
    return [(f['code'], tuple(f['files'])) for f in data['findings'] if level is None or f['level'] == level]


def wage_fails(data):
    return [c for c in codes(data, 'FAIL') if c[0].startswith('WAGE-REPORT')]


def test_wage_report_version():
    # 1) 単体: 版と種類の判定（公式様式の原本で実測）
    for key, path in OLD.items():
        sheet = mod._read_wage_report(str(path))
        assert mod.wage_report_version(sheet[1], sheet[2]) == mod.WAGE_REPORT_ROUNDS_1_6, (key, sheet)
    for key, path in NEW.items():
        sheet = mod._read_wage_report(str(path))
        assert mod.wage_report_version(sheet[1], sheet[2]) == mod.WAGE_REPORT_ROUND7, (key, sheet)
    kinds = {k: mod.wage_report_kind(mod._read_wage_report(str(p))[1]) for k, p in NEW.items()}
    assert kinds == {'s1h': ('補助率引上げ・加点措置①用', '通常枠'), 's1': ('加点措置①用', 'インボイス枠'),
                     's2': ('加点措置②用', None)}, kinds
    # 新旧の目印が混ざる・無いときは推測しない
    assert mod.wage_report_version('【第７次公募回以降用】賃金状況報告シート',
                                   '令和7年7月の事業場内最低賃金＋６３円') is None
    assert mod.wage_report_version('賃金状況報告シート', '') is None
    # 全角・改行の揺れを吸収
    assert mod.wage_report_version('賃金状況報告シート',
                                   '令和８年６月の事業場内最低賃金\n＋５５円以上') == mod.WAGE_REPORT_ROUND7

    with tempfile.TemporaryDirectory(prefix='wage-report-') as tmp:
        tmp = Path(tmp)
        # 2) 旧版を「賃金状況報告」名で置き、第7次で FAIL
        a = tmp / 'a'
        a.mkdir()
        shutil.copy(OLD['s2'], a / '賃金状況報告シート_加点措置②用.xlsx')
        cp, data = run_local(a, '--call-round', '7', '--app-month', '2026/11')
        assert cp.returncode == 1, cp.stdout
        assert ('WAGE-REPORT-OLD-VERSION', ('賃金状況報告シート_加点措置②用.xlsx',)) in codes(data, 'FAIL'), cp.stdout
        assert data['context']['wage_reports'][0]['version'] == mod.WAGE_REPORT_ROUNDS_1_6
        # 同じ旧版を第6次（旧回の訂正）として見ると版の FAIL は出ない
        cp, data = run_local(a, '--call-round', '6', '--app-month', '2026/11')
        assert not wage_fails(data), codes(data)
        # 3) katen-judge の出力名（「賃金状況報告」なし）の新版。第7次は版 OK、シート名で見つけた INFO
        b = tmp / 'b'
        b.mkdir()
        shutil.copy(NEW['s2'], b / '顧客企業A_加点措置②用.xlsx')
        shutil.copy(NEW['s1h'], b / '顧客企業A_補助率引き上げ・加点措置①用.xlsx')
        cp, data = run_local(b, '--call-round', '7', '--frame', '通常枠')
        assert not wage_fails(data), codes(data)
        assert sum(1 for c in codes(data, 'INFO') if c[0] == 'WAGE-REPORT-BY-SHEET') == 2, codes(data)
        assert {t['name'] for t in data['content_check_targets'] if t['category'] == 'wage_report'} == {
            '顧客企業A_加点措置②用.xlsx', '顧客企業A_補助率引き上げ・加点措置①用.xlsx'}
        # 新版を第6次として指定すると FAIL
        cp, data = run_local(b, '--call-round', '6')
        assert any(c[0] == 'WAGE-REPORT-ROUND-MISMATCH' for c in codes(data, 'FAIL')), codes(data)
        # 4) 枠の取り違え: インボイス枠の案件に補助率引上げ用
        cp, data = run_local(b, '--call-round', '7', '--frame', 'インボイス枠')
        assert ('WAGE-REPORT-FRAME-MISMATCH', ('顧客企業A_補助率引き上げ・加点措置①用.xlsx',)) in codes(data, 'FAIL'), codes(data)
        # 通常枠の案件に通常枠以外用の①（公式のダウンロード名のまま）
        c = tmp / 'c'
        c.mkdir()
        shutil.copy(NEW['s1'], c / 'it2026_kaitei_chingin_houkoku1_katen.xlsx')
        cp, data = run_local(c, '--call-round', '7', '--frame', '通常枠')
        assert any(x[0] == 'WAGE-REPORT-FRAME-MISMATCH' for x in codes(data, 'FAIL')), codes(data)
        # 5) 公募回未指定: 版は INFO で報告、FAIL にしない
        cp, data = run_local(a)
        assert any(x[0] == 'WAGE-REPORT-VERSION' for x in codes(data, 'INFO')), codes(data)
        assert not wage_fails(data)
        # 6) 賃金状況報告シートでない未分類の Excel は黙って無視
        d = tmp / 'd'
        d.mkdir()
        wb = openpyxl.Workbook()
        wb.active.title = 'メモ'
        wb.save(d / '加点措置メモ.xlsx')
        cp, data = run_local(d, '--call-round', '7')
        assert not [x for x in codes(data) if x[0].startswith('WAGE-REPORT')], codes(data)
        # 7) Drive 経路（未DL）: 名前の手がかりで中身チェック（DL）対象に載る
        m = tmp / 'manifest.json'
        m.write_text(json.dumps({'files': [{'name': '顧客企業A_加点措置②用.xlsx', 'mimeType': 'application/x',
                                            'parent_path': ''}]}, ensure_ascii=False), encoding='utf-8')
        out = tmp / 'm.json'
        cp, data = run(['--manifest', m, '--task', 'katen', '--call-round', '7'], out)
        assert any(t['name'] == '顧客企業A_加点措置②用.xlsx' for t in data['content_check_targets']), \
            data['content_check_targets']
        # 2回目（DL済み）で版を判定
        files = tmp / 'files'
        files.mkdir()
        shutil.copy(OLD['s2'], files / '顧客企業A_加点措置②用.xlsx')
        cp, data = run(['--manifest', m, '--files-dir', files, '--task', 'katen', '--call-round', '7'], out)
        assert cp.returncode == 1 and any(x[0] == 'WAGE-REPORT-OLD-VERSION' for x in codes(data, 'FAIL')), cp.stdout
        # 8) 入力フォルダに何も書いていない
        assert sorted(p.name for p in b.iterdir()) == ['顧客企業A_加点措置②用.xlsx',
                                                      '顧客企業A_補助率引き上げ・加点措置①用.xlsx']
    print('PASS: 賃金状況報告シートの版（新旧）・種類・枠・公募回未指定・名前なし検出・Drive 2段階')


if __name__ == '__main__':
    test_wage_report_version()
