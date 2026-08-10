# -*- coding: utf-8 -*-
"""通常枠×個人事業主の TemplateMapping を機械生成する。

考え方:
    新テンプレは build_genpon_tsujo_kojin.py が
      E = 通常枠法人 v2（通常枠固有ブロック）
      G = インボイス枠個人（個人事業主ブロック）
    の行セグメントを連結して作っている。よって既存マッピングの行番号は、
      E 由来の項目 → SHINSEI_MAP_E で平行移動
      G 由来の項目 → インボイス個人マッピングの行を SHINSEI_MAP_G で平行移動
    で機械的に求まる。

    hearing_to_tenki は行番号ではなくラベル一致で組み直す
    （通常枠個人のヒアリングシートは行構成が既存2様式のどちらとも違うため）。

検証:
    生成した行の B列ラベルが、参照元（E または G）の同項目ラベルと一致するかを全件照合する。
    1件でも合わなければ生成を中止する。

出力:
    標準出力に検証サマリ、--emit 指定時に Python リテラルをファイルへ書き出す。
"""
import argparse
import sys
import unicodedata
from pathlib import Path

import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from hojokin import config  # noqa: E402
from build_genpon_tsujo_kojin import (  # noqa: E402
    SHINSEI_MAP_E, SHINSEI_MAP_G, TENKI_SPEC,
)

TOOL = ROOT / 'ツール'


def find(name: str) -> Path:
    want = unicodedata.normalize('NFC', name)
    for p in TOOL.iterdir():
        if unicodedata.normalize('NFC', p.name) == want:
            return p
    raise FileNotFoundError(name)


def norm(s) -> str:
    return unicodedata.normalize('NFKC', str(s or '')).replace(' ', '').replace('　', '').strip()


TSUJO = config.MAPPING_2026_TSUJO
KOJIN = config.MAPPING_2026_INVOICE_KOJIN

# 個人事業主テンプレには存在しない項目（インボイス個人マッピングにも無い）。
# 役員セクションは丸ごと削除、設立年月日は生年月日/事業開始年月日に置き換わっている。
EXPECTED_ABSENT = {'established_date', 'rep_title', 'officer_count'} | {
    f'officer_{i}_{k}' for i in range(1, 11) for k in ('title', 'name', 'kana')
}

# 転記シートの見出し行（データ項目ではない）
SECTION_PREFIX = ('◼', '【', '⇨', '↓', '※')
SECTION_EXACT = {'項目', '記入欄', '補足'}


def build_shinsei(new_ws, e_ws, g_ws):
    """フィールド名 → 新テンプレの行。E 由来を優先し、無ければ G 由来へ落とす。"""
    out, report = {}, []
    for field, e_row in TSUJO.shinsei.items():
        if e_row in SHINSEI_MAP_E:
            new_row, src, src_row, src_ws = SHINSEI_MAP_E[e_row], 'E', e_row, e_ws
        elif field in KOJIN.shinsei and KOJIN.shinsei[field] in SHINSEI_MAP_G:
            g_row = KOJIN.shinsei[field]
            new_row, src, src_row, src_ws = SHINSEI_MAP_G[g_row], 'G', g_row, g_ws
        elif field in EXPECTED_ABSENT:
            report.append((field, e_row, None, None, None, 'OK(意図的に除外: 個人事業主に無い項目)'))
            continue
        else:
            report.append((field, e_row, None, None, None, 'NG: 新レイアウトに対応行が見つからない'))
            continue
        want = norm(src_ws.cell(src_row, 2).value)
        got = norm(new_ws.cell(new_row, 2).value)
        ok = (want == got) or (want and got and (want.startswith(got) or got.startswith(want)))
        out[field] = new_row
        report.append((field, e_row, src, src_row, new_row, 'OK' if ok else f'NG: {want[:22]!r} != {got[:22]!r}'))
    return out, report


def build_hearing_to_tenki(new_tenki_ws, hear_ws):
    """転記シートA列ラベル ⇔ ヒアリング基本情報B列ラベル をラベル一致で対応づける。"""
    hear = {}
    for r in range(1, 110):
        v = hear_ws.cell(r, 2).value
        if v and not str(v).strip().startswith(('◼', '【', '※', '⇨', '↓')):
            hear.setdefault(norm(v), r)
    # 電話番号変換フラグは通常枠法人の設定を項目名で引き継ぐ
    tel_by_label = {}
    e_hear = openpyxl.load_workbook(find('ヒアリングシート2026_通常枠法人.xlsx'), data_only=True)['基本情報']
    for hr, _tr, tel in TSUJO.hearing_to_tenki:
        tel_by_label[norm(e_hear.cell(hr, 2).value)] = tel

    rows, unmatched = [], []
    for dst, _src, _srow in TENKI_SPEC:
        raw = str(new_tenki_ws.cell(dst, 1).value or '').strip()
        label = norm(raw)
        if not label or raw.startswith(SECTION_PREFIX) or raw in SECTION_EXACT:
            continue
        hr = hear.get(label)
        if hr is None:
            cands = [k for k in hear if k.startswith(label[:12]) or label.startswith(k[:12])]
            if len(cands) == 1:
                hr = hear[cands[0]]
            else:
                unmatched.append((dst, label[:30]))
                continue
        rows.append((hr, dst, bool(tel_by_label.get(label, False))))
    rows.sort()
    return rows, unmatched


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--emit', help='生成した Python リテラルの書き出し先')
    args = ap.parse_args()

    new_wb = openpyxl.load_workbook(find('【原本_個人】企業名_通常枠_個人2026.xlsx'))
    e_wb = openpyxl.load_workbook(find('【原本_法人】企業名_通常枠_法人2026_v2.xlsx'))
    g_wb = openpyxl.load_workbook(find('【原本_個人】企業名_インボイス枠_個人2026.xlsx'))
    hear = openpyxl.load_workbook(find('ヒアリングシート2026_通常枠個人.xlsx'), data_only=True)['基本情報']

    global new_wb_ref, new_tenki_ws_ref
    new_wb_ref, new_tenki_ws_ref = new_wb, new_wb['転記']
    shinsei, report = build_shinsei(new_wb['申請内容'], e_wb['申請内容'], g_wb['申請内容'])
    h2t, unmatched = build_hearing_to_tenki(new_wb['転記'], hear)

    ng = [r for r in report if not r[5].startswith('OK')]
    print(f'shinsei: {len(shinsei)}/{len(TSUJO.shinsei)} 項目を解決')
    for r in ng:
        print(f'  {r[5]}  field={r[0]} (E{r[1]} -> {r[4]})')
    print(f'hearing_to_tenki: {len(h2t)} 件')
    for d, lab in unmatched:
        print(f'  未対応 転記r{d}: {lab}')

    if ng or unmatched:
        print('\n生成中止: 未解決あり')
        return 1

    lines = ['MAPPING_2026_TSUJO_KOJIN = TemplateMapping(', '    hearing_to_tenki=[']
    for hr, tr, tel in h2t:
        label = str(hear.cell(hr, 2).value or '').split('\n')[0][:26]
        lines.append(f'        ({hr}, {tr}, {tel}),   # {label}')
    lines.append('    ],')
    lines.append('    shinsei={')
    for field, row in shinsei.items():
        label = str(new_wb['申請内容'].cell(row, 2).value or '').split('\n')[0][:26]
        lines.append(f"        {field!r}: {row},   # {label}")
    lines.append('    },')
    # 給与計算シートは E 由来の「生産性指標給与支給総額計算」をそのまま継承しているので通常枠と同一
    lines.append(f'    kyuyo_sheet_name={TSUJO.kyuyo_sheet_name!r},')
    lines.append(f'    kyuyo={TSUJO.kyuyo!r},')
    new_shinsei_ws = new_wb_ref['申請内容']
    last = max((c.row for row in new_shinsei_ws.iter_rows() for c in row if c.value is not None), default=240)
    lines.append(f'    shinsei_clear_range=(5, {last + 6}),')
    # 自由記述ブロック（主な事業内容〜取引先属性）。end は range() の排他側
    tcol = {norm(new_tenki_ws_ref.cell(r, 1).value): r for r in range(1, 111)
            if new_tenki_ws_ref.cell(r, 1).value}
    t_start = tcol.get(norm('主な事業内容'))
    t_end = tcol.get(norm('どのような属性の取引先が多いですか？'))
    if not (t_start and t_end):
        print('生成中止: 転記の自由記述ブロックを特定できない')
        return 1
    lines.append(f'    tenki_text_range=({t_start}, {t_end + 1}),')
    lines.append('    is_kojin=True,')
    preserve = sorted(SHINSEI_MAP_E[r] for r in TSUJO.preserve_rows if r in SHINSEI_MAP_E)
    lines.append(f'    preserve_rows={preserve!r},')
    wage_row = None
    e_hear = openpyxl.load_workbook(find('ヒアリングシート2026_通常枠法人.xlsx'), data_only=True)['基本情報']
    if TSUJO.hearing_wage_raise_row:
        want = norm(e_hear.cell(TSUJO.hearing_wage_raise_row, 2).value)
        for r in range(1, 110):
            if norm(hear.cell(r, 2).value) == want:
                wage_row = r
                break
    lines.append(f'    hearing_wage_raise_row={wage_row or 0},')
    lines.append(')')
    text = '\n'.join(lines)

    print(f'\npreserve_rows(通常枠 {TSUJO.preserve_rows} -> 個人 {preserve})')
    print(f'hearing_wage_raise_row(通常枠 {TSUJO.hearing_wage_raise_row} -> 個人 {wage_row})')
    if args.emit:
        Path(args.emit).write_text(text + '\n', encoding='utf-8')
        print(f'\nemit -> {args.emit}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
