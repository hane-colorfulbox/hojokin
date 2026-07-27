# -*- coding: utf-8 -*-
"""原本転記シート群（SKILL.md §4.1.2）の受け入れ検査。openpyxl のみで動く（API不使用）。

使い方:
    python .claude/skills/wagebook-convert/review/verify_transcript.py "{会社名}_賃金台帳一覧.xlsx"

verify_wagebook.py が「従業員別明細」を検算するのに対し、こちらは「原本転記*」シート群が
**原本の代わりに使える状態か**を検証する。転記シートの用途（確認者が Excel の関数で
「課税支給額 − 通勤手当」等を計算する／加点判定に数値を転用する）が成り立つ条件を機械化した。

  T1 内訳の完全性 : 課税支給額 ＝ 支給項目（基本給〜課税支給額直前）の和 が各列で成立。
                    成立すれば、通勤手当に限らずどの内訳項目でも関数で加減算できる。
                    非課税項目が内訳に混在する台帳では、部分集合一致で課税対象の項目群を特定して報告。
  T2 数値セル     : 課税支給額と内訳が数値型（文字列だと Excel 関数が計算できない）。
  T3 合計の一致   : 原本の「合計」列 ＝ 転記値の SUM が全行一致（読み取り誤りの検出）。
                    原本側の不整合で一致しない行は、合計セルに色を付けておけば別枠で報告される。
  T4 加点転用     : 各月次ページに 基本給 と 労働時間の手掛かり（日額単価／出勤日数／労働時間）がある。

1ページ内に複数ブロック（月次／賞与／総合計）が横並びの形式に対応する（項目名列で区切る）。
終了コード: 0=全項目OK / 1=要確認あり / 2=転記シートなし・構造解析不可。
"""
import re
import sys
import unicodedata
from itertools import combinations
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
from openpyxl import load_workbook  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

if len(sys.argv) < 2:
    print('使い方: python verify_transcript.py <出力xlsx>', file=sys.stderr)
    raise SystemExit(2)
BOOK = Path(sys.argv[1])
if not BOOK.exists():
    print(f'❌ ファイルが見つかりません: {BOOK}', file=sys.stderr)
    raise SystemExit(2)
LABEL_HEADS = ('項目', '氏名', 'No.', 'No')
SKIP_HEADS = ('合計', '検算', '項目', '氏名')
TAXABLE = ('課税支給額', '総支給額(課税)', '課税分給与額')
BASE = ('基本給', '賞与額', '基本給区分')
COMMUTE = '通勤手当'
NON_PAY = ('項目', '氏名', '扶養家族', '出勤日数', '休日出勤', '有給休暇', '普通残業', '休日残業',
           '欠勤日数', '遅刻早退', '期間', '所属', '日額単価', 'No', '支給日', '出勤時間数',
           '欠勤', '不就労', '平日普通', '平日深夜', '休日普通', '休日深夜', '法定休日普通',
           '法定休日深夜', '内60時間超過', '内45時間超過', '有休日数', '有休残日数', '性別',
           '生年月日', '入社日', '退社日', '月額表区分', '役職')
TIME_HINTS = ('日額単価', '出勤日数', '出勤時間数', '労働時間', '時間数', '平日普通')

results, notes = [], []


def check(rid, label, cond, detail=''):
    results.append((rid, cond))
    print(f"  {'OK  ' if cond else 'NG  '} [{rid}] {label} {detail}")


def norm(v):
    return unicodedata.normalize('NFKC', str(v or '')).replace(' ', '').replace('　', '')


def find_header_row(ws):
    """右に2つ以上の見出しが並ぶ行を、この転記シートのヘッダー行とみなす。"""
    best, best_n = None, 0
    for r in range(1, min(ws.max_row, 14) + 1):
        if norm(ws.cell(r, 1).value) not in LABEL_HEADS:
            continue
        n = sum(1 for c in range(2, ws.max_column + 1) if ws.cell(r, c).value not in (None, ''))
        if n > best_n:
            best, best_n = r, n
    return best


def blocks_of(ws, hdr):
    """項目名列（'項目'/'氏名'）で区切ってブロックに分ける。"""
    label_cols = [c for c in range(1, ws.max_column + 1)
                  if norm(ws.cell(hdr, c).value) in LABEL_HEADS]
    out = []
    for i, lc in enumerate(label_cols):
        end = label_cols[i + 1] - 1 if i + 1 < len(label_cols) else ws.max_column
        data, total = [], None
        for c in range(lc + 1, end + 1):
            h = norm(ws.cell(hdr, c).value)
            if not h:
                continue
            if h.startswith('合計') and total is None:
                total = c
            elif not h.startswith(SKIP_HEADS):
                data.append(c)
        rows = {}
        for r in range(hdr, ws.max_row + 1):
            lab = norm(ws.cell(r, lc).value)
            if lab and lab not in LABEL_HEADS:
                rows.setdefault(lab, r)
        if data:
            out.append({'label_col': lc, 'data': data, 'total': total, 'rows': rows,
                        'heads': [norm(ws.cell(hdr, c).value) for c in data]})
    return out


def pay_items(ws, blk):
    base_r = next((v for k, v in blk['rows'].items() if k.startswith(BASE)), None)
    tax_r = next((v for k, v in blk['rows'].items() if k.startswith(TAXABLE)), None)
    if base_r is None or tax_r is None or tax_r <= base_r:
        return None, tax_r
    items = []
    for r in range(base_r, tax_r):
        lab = norm(ws.cell(r, blk['label_col']).value)
        if lab and lab not in NON_PAY and not lab.startswith('日額単価'):
            items.append((r, lab))
    return items, tax_r


wb = load_workbook(BOOK, data_only=False)
# 「原本転記」で始まる全シートを候補にする。表紙（ヘッダー行を持たないシート）は
# find_header_row が None を返すので自動的に対象外になる（旧仕様の1枚構成にも対応）。
sheets = [s for s in wb.sheetnames if s.startswith('原本転記')]
print(f'ブック: {BOOK.name}\n転記データシート: {len(sheets)}枚\n')

parsed = {}
for sn in sheets:
    ws = wb[sn]
    hdr = find_header_row(ws)
    if hdr:
        parsed[sn] = {'ws': ws, 'hdr': hdr, 'blocks': blocks_of(ws, hdr)}
    else:
        notes.append(sn)   # 表紙シート等（データ行を持たない）
if not sheets:
    print('❌ 「原本転記*」シートが無い。§4.1.2 の転記シート群を作ってから実行する'
          '（既存JSON流用でPDF原本が無い場合は §9 に省略理由を明記）', file=sys.stderr)
    raise SystemExit(2)
if not parsed:
    print('❌ 転記シートの構造を解析できない（ヘッダー行に「項目」or「氏名」が必要）。'
          '§4.1.2 のレイアウトを確認する', file=sys.stderr)
    raise SystemExit(2)
nblocks = sum(len(v['blocks']) for v in parsed.values())
print(f'解析: {len(parsed)}シート / {nblocks}ブロック'
      + (f' / 未解析 {notes}' if notes else '') + '\n')

print('=== T1 内訳の完全性（課税支給額 ＝ 支給項目の和） ===')
ok1, ng1, partial, noitem = 0, [], [], []
for sn, p in parsed.items():
    ws = p['ws']
    for bi, blk in enumerate(p['blocks']):
        items, tax_r = pay_items(ws, blk)
        if items is None:
            is_summary_block = blk['heads'] and all(
                h.startswith(('賞与', '総合計')) for h in blk['heads'])
            if tax_r is not None and not is_summary_block:
                # 課税支給額はあるのに内訳行が1つもない＝集計行だけの転記（§4.1.2 🔴 違反）
                ng1.append(f'{sn}: 課税支給額はあるが支給項目（基本給・各手当）の行が無い'
                           '＝集計行のみの転記の疑い')
            else:
                noitem.append(f'{sn}#b{bi}')
            continue
        for c in blk['data']:
            tax = ws.cell(tax_r, c).value
            if not isinstance(tax, (int, float)):
                continue
            vals = [(lab, ws.cell(r, c).value) for r, lab in items
                    if isinstance(ws.cell(r, c).value, (int, float)) and ws.cell(r, c).value]
            s = sum(v for _, v in vals)
            head = norm(ws.cell(p['hdr'], c).value)
            if abs(s - tax) < 0.5:
                ok1 += 1
                continue
            sub = None
            if len(vals) <= 14:
                for k in range(len(vals), 0, -1):
                    for comb in combinations(range(len(vals)), k):
                        if abs(sum(vals[i][1] for i in comb) - tax) < 0.5:
                            sub = comb
                            break
                    if sub:
                        break
            if sub is not None:
                partial.append(f'{sn}/{head}: 課税対象外＝{[vals[i][0] for i in range(len(vals)) if i not in sub]}')
                ok1 += 1
            else:
                ng1.append(f'{sn}/{head}: 内訳和{s:,.0f} vs 課税{tax:,.0f}')
if partial:
    uniq = sorted({x.split(': ', 1)[1] for x in partial})
    print(f'    非課税項目が内訳に混在: {len(partial)}件（部分集合一致で課税対象を特定）→ {uniq[:3]}')
if noitem:
    print(f'    支給項目行を持たないブロック（総合計欄など）: {len(noitem)}件 {noitem[:6]}')
check('T1', f'課税支給額＝内訳の和が全列で成立（{ok1}列）', not ng1, '; '.join(ng1[:4]))

print('=== T2 関数計算の可能性（数値セルであること） ===')
bad2 = []
for sn, p in parsed.items():
    ws = p['ws']
    for blk in p['blocks']:
        items, tax_r = pay_items(ws, blk)
        if items is None:
            continue
        for c in blk['data']:
            for r, lab in [(tax_r, '課税支給額')] + items:
                v = ws.cell(r, c).value
                if v is not None and not isinstance(v, (int, float)):
                    bad2.append(f'{sn}!{get_column_letter(c)}{r}({lab})={v!r}')
check('T2', '課税支給額と内訳がすべて数値セル', not bad2, '; '.join(bad2[:4]))

print('=== T3 原本の網羅性（合計列 ＝ 転記値のSUM） ===')
ok3, ng3, ng3_marked, nototal = 0, [], [], []
for sn, p in parsed.items():
    ws = p['ws']
    for bi, blk in enumerate(p['blocks']):
        if not blk['total']:
            nototal.append(f'{sn}#b{bi}')
            continue
        for r in range(p['hdr'], ws.max_row + 1):
            printed = ws.cell(r, blk['total']).value
            if not isinstance(printed, (int, float)):
                continue
            s = sum(ws.cell(r, c).value for c in blk['data']
                    if isinstance(ws.cell(r, c).value, (int, float)))
            if abs(s - printed) < 0.5:
                ok3 += 1
            else:
                lab = norm(ws.cell(r, blk['label_col']).value)
                cellfill = ws.cell(r, blk['total']).fill
                marked = bool(cellfill and cellfill.fill_type and
                              (cellfill.fgColor.rgb or '') not in ('00000000', None))
                msg = f'{sn} r{r}({lab}): SUM{s:,.0f} vs 印字{printed:,.0f}'
                (ng3_marked if marked else ng3).append(msg)
if nototal:
    print(f'    合計列を持たないブロック: {len(nototal)}件 {nototal[:6]}')
if ng3_marked:
    print(f'    原本側の不整合として黄色マーク済み: {len(ng3_marked)}件 {ng3_marked[:3]}')
check('T3', f'原本の合計とSUMが全行一致（{ok3}行・マーク済みの原本不整合は除く）',
      not ng3, '; '.join(ng3[:4]))

print('=== T4 加点判定への転用（基本給＋労働時間の手掛かり） ===')
ng4, rate_pages = [], []
for sn, p in parsed.items():
    monthly = [blk for blk in p['blocks']
               if blk['heads'] and not all(h.startswith(('賞与', '総合計')) for h in blk['heads'])]
    if not monthly:
        continue          # 賞与/総合計だけのページは月次の勤怠を持たない
    for blk in monthly:
        has_base = any(k.startswith(BASE) for k in blk['rows'])
        has_time = any(k.startswith(TIME_HINTS) for k in blk['rows'])
        if not (has_base and has_time):
            ng4.append(f'{sn}: 基本給={has_base} 時間手掛かり={has_time}')
    rr = next((v for blk in monthly for k, v in blk['rows'].items() if k.startswith('日額単価')), None)
    if rr:
        n = sum(1 for blk in monthly for c in blk['data']
                if isinstance(p['ws'].cell(rr, c).value, (int, float)))
        rate_pages.append((sn, n))
check('T4', f'月次ページすべてで基本給＋労働時間の手掛かりあり', not ng4, '; '.join(ng4[:4]))
print(f'    日額単価が転記されたページ: {len(rate_pages)}/{len(parsed)}'
      + (f' {rate_pages}' if rate_pages else '（この台帳形式には日額単価の印字なし）'))

nfail = sum(1 for _, ok in results if not ok)
print()
if nfail:
    print(f'=== 要確認 {nfail}件 ===（NGはPDF原本で確認し、§9 報告に判断を明記してから完了報告へ。'
          '原本側の不整合なら該当セルを黄色にして注記する）')
else:
    print('=== PASS ===（T1〜T4 すべて成立。ただし転記の「真」はPDF原本なので '
          '§5.5-3 の目視突合は別途必要）')
sys.exit(1 if nfail else 0)
