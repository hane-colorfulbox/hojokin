"""旧原本を保存し、第7次候補原本を再現可能に作る。Drive原本への書込みはしない。

公式加点様式は取得済みのバイト列をそのまま複製する。社内原本はZIP内の
表示セル・プルダウンの額・版メタデータだけを変更し、図形・数式・保護を保持する。
通常枠個人は既存のDrive由来パッチ済み原本を入力にする。
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


def _hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _text(value: str) -> str:
    pairs = [
        (r'令和[6６]年\s*[1１][0０]月', '令和7年10月'),
        (r'令和[7７]年\s*[9９]月', '令和8年9月'),
        (r'令和[7７]年\s*[7７]月', '令和8年6月'),
        (r'令和[7７]年度改定', '令和8年度改定'),
        (r'令和[7７]年度地域別最低賃金', '令和8年度地域別最低賃金'),
        (r'R7年度改定|R7改定後', 'R8改定後'),
        (r'[6６][3３]円', '55円'),
        (r'it2026_chingin_houkoku', 'it2026_kaitei_chingin_houkoku'),
        (r'直近[2２]期分の賃金台帳', '用途別の指定年月をすべて覆う賃金台帳（冊数は固定しない）'),
        (r'地域別最低賃金[＋+]\s*[5５][0０]円以内', '地域別最低賃金＋50円以上（任意の賃上げ計画加点）'),
    ]
    for pattern, replacement in pairs:
        value = re.sub(pattern, replacement, value)
    return value


def _patch_book(source: Path, output: Path, edits: dict[str, dict[str, object]]) -> dict:
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as zin:
        wb = openpyxl.load_workbook(source, read_only=True, data_only=False)
        replacements = {}
        for sheet_name, values in edits.items():
            part = _resolve_sheet_part(zin, wb.sheetnames.index(sheet_name))
            replacements[part] = _patch_sheet_xml(zin.read(part).decode('utf-8'), values).encode('utf-8')
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
    # Reopen and confirm no row/sheet/formula/object changes outside explicit label inputs.
    with zipfile.ZipFile(source) as a, zipfile.ZipFile(output) as b:
        assert a.namelist() == b.namelist()
        for name in a.namelist():
            if name not in replacements:
                assert a.read(name) == b.read(name), name
    checked = openpyxl.load_workbook(output, read_only=True, data_only=False)
    assert checked.properties.subject == ROUND7_SUBJECT
    for sheet, values in edits.items():
        for ref, value in values.items():
            assert checked[sheet][ref].value == value, (source.name, sheet, ref)
    checked.close()
    return {'input': source.relative_to(ROOT).as_posix(), 'input_sha256': _hash(source),
            'output': output.relative_to(ROOT).as_posix(), 'output_sha256': _hash(output),
            'cells': {sheet: list(values) for sheet, values in edits.items()}}


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
        if template_type:
            base = BASE_ROWS[template_type]
            fields = edits.setdefault('申請内容', {})
            fields[f'B{base}'] = '事業計画期間における給与支給総額\n（0年度目・申請年度の見込み）'
            for year in range(1, 4):
                fields[f'B{base + year}'] = f'給与支給総額\n（{year}年度目の確認済み計画数値）'
            fields[f'B{base - 1}'] = '0年度目の対象従業員数（確認済みの正社員換算人数）'
            for sheet_name in wb.sheetnames:
                if '給与支給総額計算' in sheet_name:
                    original = str(wb[sheet_name]['A2'].value or '')
                    edits.setdefault(sheet_name, {})['A2'] = original + '\n第7次：生産性は前事業年度実績、給与0年度目は申請年度の見込み。給与の提出値は「申請前提・給与見込」を確認。旧計算式を申請値へ流用しない。'
            if '生産性指標給与支給総額計算' in wb:
                labels = edits.setdefault('生産性指標給与支給総額計算', {})
                labels.update({'B27': '前事業年度（生産性の実績）',
                               'C27': '1年度目（申請当年度）確認済み計画',
                               'D27': '2年度目（翌事業年度）確認済み計画',
                               'E27': '3年度目（翌々事業年度）確認済み計画'})
                labels['B42'] = '生産性計画は未確認（確認済み値と3%／4%の要件を別途検算）'
        elif '確認事項' in wb:
            # Existing comments contain fixed collection assumptions. No row insertion.
            ws = wb['確認事項']
            for row in ws:
                for cell in row:
                    if isinstance(cell.value, str) and '前事業年度1年分の給与支給総額' in cell.value:
                        edits.setdefault(ws.title, {})[cell.coordinate] = '給与0年度目の見込み総額・対象人数・算定根拠（確認済み値）。生産性用の前事業年度実績は別資料。'
        wb.close()
        log.append(_patch_book(source, tools_dir / '第7次以降' / name, edits))
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
        edits['記入ルール'] = {'A7': '公募回・申請月と暦月を確認。2025年10月〜2026年9月＋直近月。基本給・所定時間が不明な月は確定しない。非在籍は「非在籍」と記入。'}
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
