# -*- coding: utf-8 -*-
"""生成した公式様式①②が原本を壊していないかを確かめる（katen-judge スキル同梱・標準ライブラリのみ）。

run_katen_judge.py --fill の出力は「原本 xlsx を丸ごとコピーし、入力セルの値だけを差し替える」
方式で作る。このスクリプトは原本と出力を突き合わせ、次のすべてを満たせば合格にする。
  1. xlsx 内のパーツ（ファイル）の集合が同じ
  2. 中身が違うパーツは、先頭シートの XML と xl/workbook.xml だけ
  3. 図形・記入例画像・印刷設定・秘密度ラベルのパーツがバイト単位で同じ
  4. 原本にある拡張データ検証（x14:dataValidation）が先頭シートに残っている
  5. docProps/app.xml に openpyxl の刻印が無い（再保存されていない）

チャットからダウンロードした様式も、Excel で開いて保存する前ならこれで確かめられる
（Excel で保存すると全パーツが書き換わるので、保存後のファイルは対象外）。

使い方:
  python check_lossless.py <原本.xlsx> <出力.xlsx>
終了コード: 0=合格 / 1=不合格 / 2=ファイル・形式エラー
"""
import re
import sys
import zipfile
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, 'reconfigure'):
        _stream.reconfigure(encoding='utf-8')

WORKBOOK_PART = 'xl/workbook.xml'
WORKBOOK_RELS = 'xl/_rels/workbook.xml.rels'
IMPORTANT_PREFIXES = ('xl/drawings', 'xl/media', 'xl/printerSettings')
IMPORTANT_SUBSTR = ('LabelInfo', 'docMetadata')
X14_DV_TAG = 'x14:dataValidation'
OPENPYXL_MARK = 'openpyxl'


def read_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as zf:
        return {i.filename: zf.read(i.filename) for i in zf.infolist()}


def first_sheet_part(parts: dict[str, bytes]) -> str:
    """workbook.xml の先頭シートの r:id を rels で引いて、シート XML のパスを返す。"""
    wb = parts[WORKBOOK_PART].decode('utf-8', 'replace')
    m = re.search(r'<(?:\w+:)?sheet\b[^>]*\br:id="([^"]+)"', wb)
    if not m:
        raise ValueError('workbook.xml にシートがありません')
    rels = parts[WORKBOOK_RELS].decode('utf-8', 'replace')
    for rel in re.finditer(r'<Relationship\b[^>]*>', rels):
        tag = rel.group(0)
        if f'Id="{m.group(1)}"' in tag:
            target = re.search(r'Target="([^"]+)"', tag).group(1)
            return target.lstrip('/') if target.startswith('/') else f'xl/{target}'
    raise ValueError(f'先頭シートの関係 {m.group(1)} が見つかりません')


def compare(orig: dict[str, bytes], new: dict[str, bytes]) -> list[str]:
    problems: list[str] = []
    lost, added = sorted(set(orig) - set(new)), sorted(set(new) - set(orig))
    if lost or added:
        problems.append(f'パーツの集合が違う（消えた {lost} / 増えた {added}）')

    main = first_sheet_part(orig)
    allowed = {main, WORKBOOK_PART}
    changed = sorted(n for n in orig if n in new and orig[n] != new[n])
    unexpected = [n for n in changed if n not in allowed]
    if unexpected:
        problems.append(f'値を入れるシート以外のパーツが変わっている: {unexpected}')

    for n in orig:
        if n.startswith(IMPORTANT_PREFIXES) or any(s in n for s in IMPORTANT_SUBSTR):
            if orig[n] != new.get(n):
                problems.append(f'図形・画像・印刷設定・ラベルのパーツが変わった: {n}')

    if X14_DV_TAG in orig[main].decode('utf-8', 'ignore') and \
            X14_DV_TAG not in new.get(main, b'').decode('utf-8', 'ignore'):
        problems.append('拡張データ検証（プルダウン等）が消えている')

    app = new.get('docProps/app.xml', b'').decode('utf-8', 'ignore').lower()
    if OPENPYXL_MARK in app:
        problems.append('openpyxl で再保存された形跡がある（原本コピー方式を通っていない）')
    return problems


def main(argv=None) -> int:
    args = sys.argv[1:] if argv is None else argv
    if len(args) != 2:
        print(__doc__.strip().splitlines()[0])
        print('使い方: python check_lossless.py <原本.xlsx> <出力.xlsx>', file=sys.stderr)
        return 2
    paths = [Path(a) for a in args]
    for p in paths:
        if not p.exists():
            print(f'❌ ファイルが見つかりません: {p}', file=sys.stderr)
            return 2
    try:
        orig, new = (read_parts(p) for p in paths)
        problems = compare(orig, new)
    except (zipfile.BadZipFile, KeyError, ValueError) as e:
        print(f'❌ xlsx として読めません: {e}', file=sys.stderr)
        return 2
    if problems:
        print(f'❌ 原本と違う箇所があります（{paths[1].name}）:')
        for msg in problems:
            print(f'  - {msg}')
        return 1
    print(f'✅ 原本どおり（値を入れたシート以外は無改変）: {paths[1].name}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
