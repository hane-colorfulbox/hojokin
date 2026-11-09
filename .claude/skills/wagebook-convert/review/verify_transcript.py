# -*- coding: utf-8 -*-
"""原本転記シート群（SKILL.md §4.1.2）の受け入れ検査。openpyxl のみで動く（API不使用）。

使い方:
    python "<スキルのフォルダ>/review/verify_transcript.py" "{会社名}_賃金台帳一覧.xlsx"

verify_wagebook.py が「従業員別明細」を検算するのに対し、こちらは「原本転記*」シート群が
**原本の代わりに使える状態か**を検証する。転記シートの用途（確認者が Excel の関数で
「課税支給額 − 通勤手当」等を計算する／加点判定に数値を転用する）が成り立つ条件を機械化した。

  T1 内訳の完全性 : 課税支給額 ＝ 支給項目（基本給〜課税支給額直前）の和 が各列で成立。
                    成立すれば、通勤手当に限らずどの内訳項目でも関数で加減算できる。
                    非課税項目が内訳に混在する台帳では、部分集合一致で課税対象の項目群を特定して報告。
                    基本給の行が無いページは支給欄の見出し行（【支給】等）の次から数える。勤怠控除などの
                    減額項目は正の数で転記されていても差し引く向きで計算する（ℹ で列を挙げる）。
                    支給項目が集計行だけなら NG。1列も検算できなければ NG。
                    課税支給額のセルがオレンジ（原本側の不整合の印）なら別枠で報告する（T3 と同じ扱い）。
  T2 数値セル     : 課税支給額が数値か空欄で、内訳に「数字の入った文字列」が無い（Excel 関数が計算できない）。
  T3 合計の一致   : 原本の合計列（「小計」「合計」で始まる列）＝ 転記値の SUM が全行一致（読み取り誤りの検出）。
                    原本側の不整合で一致しない行は、合計セルをオレンジにしておけば別枠で報告される。
  T4 加点転用     : 各月次ページに 基本給 と 労働時間の手掛かり（日額単価／時給／出勤日数／労働時間）がある。
                    賞与だけ・役員報酬だけのページは対象外。
  T5 通勤手当     : 通勤手当が課税支給額に含まれているか（included / excluded）を判定し、
                    従業員別明細の S列と突合する。控除漏れ（included なのに S列が空）と
                    二重控除（excluded なのに S列に値）を検知する（SKILL.md §3.1 の全額控除方式）。
                    どちらの式も成立しない場合は「判定不能」＝控除せず報告に明記させる。

1ページ内に複数ブロック（月次／賞与／総合計）が横並びの形式に対応する（項目名列で区切る）。
T5 は行=項目×列=従業員（月次ページ型）と行=従業員×列=項目（一覧型）の両レイアウトを扱う
（entity_maps で正規化。後者を無言で素通りさせない）。
終了コード: 0=全項目OK / 1=要確認あり / 2=転記シートなし・構造解析不可。
"""
import re
import sys
import unicodedata
from itertools import combinations
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
from openpyxl import load_workbook  # noqa: E402

if len(sys.argv) < 2:
    print('使い方: python verify_transcript.py <出力xlsx>', file=sys.stderr)
    raise SystemExit(2)
BOOK = Path(sys.argv[1])
if not BOOK.exists():
    print(f'❌ ファイルが見つかりません: {BOOK}', file=sys.stderr)
    raise SystemExit(2)
LABEL_HEADS = ('項目', '氏名', 'No.', 'No', '社員名', '従業員名', '社員氏名', '従業員氏名')
SKIP_HEADS = ('合計', '小計', '検算', '項目', '氏名')
# 原本の合計列。「(小計)」（そのページ分）と「《合計》」（月全体）が並ぶ台帳では小計を使う
TOTAL_HEADS = ('小計', '合計')
# 課税支給額に当たる行名（前方一致。「累計」は除く）。「課税対象額」は社会保険料を引いた後の額のことがあるので入れない
TAXABLE = ('課税支給額', '総支給額(課税)', '課税分給与額', '課税分合計', '課税支給合計', '課税合計',
           '課税給与', '課税支給計', '課税計')
# 支給項目の始まりになる行名（前方一致）。賞与の行も始まりになる（is_bonus_label）
BASE_WORDS = ('基本給', '本給', '月給', '給料', '時間給', '日給', '基本賃金', '基本時給')
BASE = BASE_WORDS + ('基本給区分', '賞与額', '役員報酬')
BONUS_WORDS = ('賞与', '期末手当', '一時金')
BONUS_NOT = ('引当', '累計', '算定', '支給日', '支給月')
# 支給欄の見出し行（原本の縦帯・区分名を転記した行。値は空）。日給制で先頭が「平日出勤」のように
# 基本給の行が無いページは、この行の次から課税支給額の直前までを支給項目とみなす。
# 先頭の記号と末尾の(…)は付いていてよい。「▼ {PDF名} p.N」の出典行（§4.1.2）と取り違えないよう、
# 記号の後ろが見出し語そのものの行だけを見出しとみなす
PAY_SECTION_RE = re.compile(r'[▼■◆●]?(支給|支給額|支給項目|支給欄)(\(.*\))?')
# 集計の行名（完全一致）。支給項目がこれだけなら「集計行だけの転記」とみなす（§4.1.2 🔴）。
# 「基本給合計」「手当合計」のような項目ごとの合計は集計行に含めない
AGGREGATE_RE = re.compile(r'総支給(額|金額)?|(支給額?|給与|支払)?(合計|総額|総計|計|金額)|小計')
# 支給欄の中の減額項目。原本で正の数のまま印字されていても課税支給額では差し引かれている
DEDUCT_WORDS = ('控除', '欠勤', '不就労', '遅刻', '早退', '遅早', '減額', '減給')
# 原本側の不整合の印（§4.1.2）。黄色は「読み取りに自信がない」印に使うので、別の色に分けている
MISMATCH_RGB = 'FFC000'              # オレンジ（標準の色。テーマの色のオレンジは認識しない）
# 行名全体を囲む括弧は外して読む（「(課税分合計)」「《合計》」「【支給】」「(出勤日数)」）
BRACKETS = {'(': ')', '[': ']', '<': '>', '【': '】', '《': '》', '〈': '〉', '「': '」',
            '『': '』', '〔': '〕'}
COMMUTE = ('通勤手当', '通勤費', '交通費', '通勤')
NONTAX_COMMUTE = ('非課税通勤',)      # 「非課税通勤費」等＝通勤の非課税分が単独で印字される形式
TAXED_COMMUTE = ('課税通勤',)         # 「課税通勤費」等
# 通勤の行でも金額ではないもの（勤怠欄の通勤日数・通勤距離など）
COMMUTE_NOT_AMOUNT = ('日数', '回数', '距離', '区間', '経路', '方法', '時間', '区分', '手段', 'km')
# 従業員別明細（同じブック内）の S列＝年間通勤手当。T5 でこの値と転記シートの判定を突合する。
LEDGER_SHEET, LEDGER_HDR_ROW = '従業員別明細', 5
LEDGER_COL_NAME, LEDGER_COL_S = 3, 19
# 属性・単価の行。金額ではないので、どの判定でも支給項目の和に入れない（完全一致）
HARD_NON_PAY = ('項目', '氏名', '扶養家族', '期間', '所属', '日額単価', 'No', '支給日', '支給年月日',
                '性別', '生年月日', '入社日', '退社日', '月額表区分', '役職', '時給単価', '時間単価', '単価')
# 勤怠の行（完全一致。前方一致にすると「休日出勤」が「休日出勤手当」を巻き添えにする）。
# 台帳によっては同じ名前の支給額の行がある（日給制の「休日出勤」など）ので、
# これを外した和が合わないときは入れて計算し直す（sum_verdict）
AMOUNT_OR_ATTENDANCE = ('休日出勤', '欠勤', '不就労', '遅刻早退')   # 台帳によっては支給額・減額の行
NON_PAY = HARD_NON_PAY + AMOUNT_OR_ATTENDANCE + (
    '出勤日数', '有給休暇', '普通残業', '休日残業', '欠勤日数', '出勤時間数', '平日普通', '平日深夜',
    '休日普通', '休日深夜', '法定休日普通', '法定休日深夜', '内60時間超過', '内45時間超過', '有休日数',
    '有休残日数')
# 再計算（broad）で足し直すのは、支給額にもなる勤怠語の行だけ（時間・日数専用の行は足さない）
NON_PAY_BROAD = tuple(k for k in NON_PAY if k not in AMOUNT_OR_ATTENDANCE)
# 単価の行（加点判定の時間換算給与の分子）。時給者・日給者でラベルが違う
RATE_HINTS = ('日額単価', '時給', '時間単価', '単価')
# 加点判定に転用できるか（T4）＝分子（単価/基本給）と分母（時間）の手掛かりがあるか
TIME_HINTS = RATE_HINTS + ('出勤日数', '出勤時間数', '労働時間', '総労働時間', '実労働時間',
                           '所定労働時間', '就業時間', '稼働時間', '時間数', '平日普通')
# 「平日出勤(日数)」「勤務日数」「実働時間」なども時間の手掛かりにする
TIME_SUFFIXES = ('(日数)', '(勤怠)', '(時間)', '日数', '時間', '時間数')
TIME_STEMS = ('勤務', '実働', '就労', '出勤', '労働', '所定')
# 月全体の合計（他ページ分を含む）を表す見出しの語。原本の《合計》もこの扱い（§4.1.2）
MONTH_WIDE_WORDS = ('月全体', '全ページ', '全体', '当月計', '他ページ')
MONTH_NAME_RE = re.compile(r'((令和|平成|R|H)?\d+年)?\d{1,2}月(分)?.*')

results, notes = [], []


def check(rid, label, cond, detail=''):
    results.append((rid, cond))
    print(f"  {'OK  ' if cond else 'NG  '} [{rid}] {label} {detail}")


def norm(v):
    s = unicodedata.normalize('NFKC', str(v or '')).replace(' ', '').replace('　', '')
    if len(s) > 2 and BRACKETS.get(s[0]) == s[-1]:
        s = s[1:-1]
    return s


def numof(v):
    return v if isinstance(v, (int, float)) else None


def is_marked(cell):
    """原本側の不整合として転記者がオレンジの印を付けたセルか（黄色・灰色・白は印とみなさない）。"""
    f = cell.fill
    rgb = f.fgColor.rgb if (f and f.fill_type and f.fgColor is not None) else None
    return isinstance(rgb, str) and rgb.upper().endswith(MISMATCH_RGB)


def all_marked(cells):
    """2つ以上のセルがすべてオレンジ＝行・列ごとの装飾。個々の不整合の印とはみなさない。"""
    cells = list(cells)
    return len(cells) >= 2 and all(is_marked(c) for c in cells)


def is_pay_section(label, cells=()):
    """支給欄の見出し行か。値の入った行（「支給額」という名前の合計行など）は見出しとみなさない。"""
    if any(numof(c.value) is not None for c in cells):
        return False
    return bool(PAY_SECTION_RE.fullmatch(label))


def is_taxable(label):
    return label.startswith(TAXABLE) and '累計' not in label


def is_bonus_label(label):
    return any(w in label for w in BONUS_WORDS) and not any(w in label for w in BONUS_NOT)


def is_aggregate(label):
    return bool(AGGREGATE_RE.fullmatch(label) or PAY_SECTION_RE.fullmatch(label))


def is_base(label):
    return (label.startswith(BASE) or is_bonus_label(label)) and not AGGREGATE_RE.fullmatch(label)


def is_bonus_head(h):
    return is_bonus_label(h) or h.startswith('総合計')


def is_time_label(h):
    if h in ('支給日', '支給年月日'):
        return False
    return (h.startswith(TIME_HINTS) or h.endswith(TIME_SUFFIXES)
            or (any(w in h for w in TIME_STEMS) and h.endswith('日')))


def is_commute_label(label):
    return label.startswith(COMMUTE) and not any(w in label for w in COMMUTE_NOT_AMOUNT)


def has_digit_text(v):
    """数字の入った文字列か（'200,000'・'215,000*'・'(5,000)' 等）。式（=…）は除く。"""
    if not isinstance(v, str):
        return False
    s = unicodedata.normalize('NFKC', v).strip()
    return not s.startswith('=') and bool(re.search(r'\d', s))


def name_key(name):
    """氏名の突合用キー。原本どおりの社員番号つき（「0001山田太郎」）でも明細の氏名と合わせる。"""
    return re.sub(r'^[\d.\-:_]+', '', name)


def is_check_name(name):
    """=SUM の検算の行・列（§4.1.2 で原本の合計の隣に置くもの。見出しは「検算」で始める）。
    「SUM」を含むかでは判定しない（KUSUMA・SUMAN などのローマ字の氏名を巻き込む）。"""
    return name.startswith('検算')


def is_deduct(label):
    """支給欄の中の減額項目か（「減額補填手当」のような加算の手当は除く）。"""
    return any(w in label for w in DEDUCT_WORDS) and not label.endswith(('手当', '補填', '補償', '加算'))


def non_person(name):
    """氏名でない列・行（合計行・部門計・検算・番号だけ）。T5 の突合から外す。"""
    return (not name or name in LABEL_HEADS or bool(AGGREGATE_RE.fullmatch(name))
            or name.endswith('計') or is_check_name(name))


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
    """項目名列（'項目'/'氏名'）で区切ってブロックに分ける。

    形式②（行=従業員×列=項目。列見出しに課税支給額の名前がある）は、合計・小計の列もデータ列として扱い、
    列方向の合計照合（T3）はしない（その形式の合計は行=合計行のほう）。
    """
    label_cols = [c for c in range(1, ws.max_column + 1)
                  if norm(ws.cell(hdr, c).value) in LABEL_HEADS]
    out = []
    for i, lc in enumerate(label_cols):
        end = label_cols[i + 1] - 1 if i + 1 < len(label_cols) else ws.max_column
        heads_all = [(norm(ws.cell(hdr, c).value), c) for c in range(lc + 1, end + 1)]
        heads_all = [(h, c) for h, c in heads_all if h]
        transposed = any(is_taxable(h) for h, _c in heads_all)
        data, totals = [], []
        for h, c in heads_all:
            raw = unicodedata.normalize('NFKC', str(ws.cell(hdr, c).value)).strip()
            if is_check_name(h):
                continue                       # =SUM の検算列（合計列とも従業員の列とも扱わない）
            if transposed:
                if h not in LABEL_HEADS:
                    data.append(c)
            elif h.startswith(TOTAL_HEADS):
                # 月全体の《合計》（他ページ分を含む）と自分で計算した列は、このページの照合に使わない
                if not ('《' in raw or '検算' in h or any(w in h for w in MONTH_WIDE_WORDS)):
                    totals.append((h, c))
            elif not h.startswith(SKIP_HEADS):
                data.append(c)
        cells_of = [[ws.cell(r, c).value for r in range(hdr + 1, ws.max_row + 1)] for _h, c in totals]
        had_formula = any(str(v or '').startswith('=') for col in cells_of for v in col)
        # 数値の印字がある列だけを合計列の候補にする（=SUM の検算式の列を合計列と取り違えない）
        totals = [(h, c) for (h, c), col in zip(totals, cells_of) if any(numof(v) is not None for v in col)]
        total = next((c for h, c in totals if h.startswith('小計')), totals[0][1] if totals else None)
        rows = {}
        for r in range(hdr, ws.max_row + 1):
            lab = norm(ws.cell(r, lc).value)
            if lab and lab not in LABEL_HEADS:
                rows.setdefault(lab, r)
        if data:
            out.append({'label_col': lc, 'data': data, 'total': total, 'totals': totals,
                        'rows': rows, 'transposed': transposed,
                        'formula_only_total': had_formula and total is None,
                        'heads': [norm(ws.cell(hdr, c).value) for c in data]})
    return out


def is_transposed(blk):
    """行=従業員 × 列=項目 の形式か（一覧型ページ）。列見出しに課税支給額の名前がある。"""
    return blk['transposed']


def entity_maps(ws, p, blk):
    """ブロックを [(氏名, {項目ラベル: Cell})] に正規化する。

    形式①（行=項目・列=従業員）と形式②（行=従業員・列=項目）を同じ形にして、
    どちらでも同一ロジックで判定できるようにする（②を素通りさせない）。
    dict は挿入順＝原本の印字順を保つので、支給項目の範囲を順序で決められる。
    Cell を返すので、値だけでなくセル位置（coordinate）も報告に使える。
    """
    if is_transposed(blk):
        cols = [(norm(ws.cell(p['hdr'], c).value), c) for c in blk['data']]
        return [(name, {lab: ws.cell(r, c) for lab, c in cols})
                for name, r in blk['rows'].items()]
    out = []
    for c in blk['data']:
        name = norm(ws.cell(p['name_row'], c).value) or norm(ws.cell(p['hdr'], c).value)
        out.append((name, {lab: ws.cell(r, c) for lab, r in blk['rows'].items()}))
    return out


def label_cells(ws, blk, label):
    """ブロックの中で、その項目名の値が入るセル全部（形式①は行、形式②は列）。"""
    if is_transposed(blk):
        c = blk['data'][blk['heads'].index(label)]
        return [ws.cell(r, c) for r in blk['rows'].values()]
    return [ws.cell(blk['rows'][label], c) for c in blk['data']]


def block_labels(blk):
    return blk['heads'] if is_transposed(blk) else list(blk['rows'])


def valued_labels(ws, blk):
    return [h for h in block_labels(blk) if any(numof(c.value) is not None for c in label_cells(ws, blk, h))]


def valued_pay_labels(ws, p, blk):
    """ブロックの支給項目のうち、値の入った金額の行（集計行・減額項目を除く）。"""
    labs = set()
    for _name, cells in entity_maps(ws, p, blk):
        items, _tax, _cell = item_sum(cells)
        labs.update(k for k, c in (items or []) if numof(c.value) and not is_aggregate(k) and not is_deduct(k))
    return labs


def is_bonus_block(ws, p, blk):
    """賞与だけのブロック（列見出しが賞与・総合計、または値の入った支給項目が賞与の行だけ）。勤怠は無いのが普通。"""
    if blk['heads'] and all(is_bonus_head(h) for h in blk['heads']):
        return True
    pay = valued_pay_labels(ws, p, blk)
    return bool(pay) and all(is_bonus_label(k) for k in pay)


def is_officer_block(ws, p, blk):
    """値の入った支給項目が役員報酬だけのブロック。役員は加点判定の対象外で、勤怠も無いのが普通。"""
    pay = valued_pay_labels(ws, p, blk)
    return bool(pay) and all(k.startswith('役員報酬') for k in pay)


def item_sum(cells, exclude=NON_PAY):
    """支給項目（基本給〜課税支給額の直前）と課税支給額を返す。

    returns ([(ラベル, Cell)] または None, 課税支給額の値 または None, 課税支給額の Cell または None)
    順序は原本の印字順（entity_maps が挿入順を保つ）。支給項目の始まりは基本給・役員報酬・賞与などの行。
    それが無いページ（日給制で先頭が「平日出勤」など）は、支給欄の見出し行（【支給】等）の次から。
    """
    keys = list(cells)
    i_tax = next((i for i, k in enumerate(keys) if is_taxable(k)), None)
    tax_cell = cells[keys[i_tax]] if i_tax is not None else None
    tax = numof(tax_cell.value) if tax_cell is not None else None
    if i_tax is None:
        return None, tax, tax_cell
    pre = keys[:i_tax]
    sections = [i for i, k in enumerate(pre) if is_pay_section(k, [cells[k]])]
    if sections:
        start = sections[-1] + 1             # 支給欄の見出し行があれば、いちばん下の見出しの次から
    else:
        # 見出しが無ければ基本給などの行から。ただし単価・区分の行（「時間給 1,100」「給料区分 月給」）が
        # 勤怠欄より上に並ぶ台帳があるので、勤怠の行より後ろにある最初の基本給の行を優先する
        bases = [i for i, k in enumerate(pre) if is_base(k) and not isinstance(cells[k].value, str)
                 and not k.endswith(('区分', '形態', '単価'))]
        if not bases:
            return None, tax, tax_cell
        last_time = max((i for i, k in enumerate(pre) if is_time_label(k) or k in NON_PAY_BROAD), default=-1)
        start = next((i for i in bases if i > last_time), bases[0])
    items = [(k, cells[k]) for k in keys[start:i_tax]
             if k not in exclude and not k.startswith('日額単価') and not is_check_name(k)
             and not is_pay_section(k, [cells[k]])]
    return items, tax, tax_cell


def _signed(vals):
    """支給欄の減額項目（勤怠控除など）は、正の数で転記されていても差し引く向きにする。"""
    return [(k, -v if v > 0 and is_deduct(k) else v) for k, v in vals]


def _amounts(items):
    """支給項目の値。集計行（総支給額・支給合計など）は和に入れない（二重に数えない・誤読を隠さない）。"""
    return [(k, c.value) for k, c in items if numof(c.value) and not (is_aggregate(k) and not is_base(k))]


def sum_verdict(cells):
    """T1 の判定。returns (verdict, detail, 課税支給額の Cell)。

    ok      : 課税支給額 ＝ 支給項目の和
    signed  : 減額項目（勤怠控除など）を差し引く向きにして成立（原本の印字が正の数かを確認する）
    partial : 非課税項目が内訳に混在し、部分集合が課税支給額に一致（課税対象を特定できた）
    ng      : どの部分集合とも一致しない＝転記漏れ・読み取り誤りの疑い
    noitem  : 課税支給額はあるが支給項目が無い／集計行だけ（集計行のみの転記の疑い）
    skip    : 課税支給額が数値でない（この列/行は対象外）
    """
    items, tax, tax_cell = item_sum(cells)
    if tax is None:
        return 'skip', '', tax_cell
    if items is None:
        return 'noitem', '', tax_cell
    valued = [k for k, c in items if numof(c.value)]
    if valued and all(is_aggregate(k) and not is_base(k) for k in valued):
        return 'noitem', '', tax_cell        # 値のある支給項目が集計行だけ（【支給】→総支給額→課税支給額 等）
    vals = _amounts(items)
    broad_items, _tax, _cell = item_sum(cells, exclude=NON_PAY_BROAD)
    broad = _amounts(broad_items)
    # 勤怠語と同名の支給額の行（日給制の「休日出勤」など）を入れた和も試す
    for cand, verdict in ((vals, 'ok'), (broad, 'ok'), (_signed(vals), 'signed'), (_signed(broad), 'signed')):
        if abs(sum(v for _, v in cand) - tax) < 0.5:
            return verdict, '', tax_cell
    if len(vals) <= 14:
        for cand in (vals, _signed(vals)):
            for k in range(len(cand), 0, -1):
                for comb in combinations(range(len(cand)), k):
                    if abs(sum(cand[i][1] for i in comb) - tax) < 0.5:
                        excl = [cand[i][0] for i in range(len(cand)) if i not in comb]
                        return 'partial', f'課税対象外＝{excl}', tax_cell
    s = sum(v for _, v in vals)
    return 'ng', f'内訳和{s:,.0f} vs 課税{tax:,.0f}', tax_cell


def commute_state(cells):
    """通勤手当が課税支給額に含まれているかを判定する（SKILL.md §3.1 の判定表）。

    included : 課税支給額 ＝ Σ支給項目（通勤手当を含む）→ S列に通勤手当の年額を入れて控除する
    excluded : 課税支給額 ＝ Σ支給項目 − 通勤手当、または通勤手当の行が課税支給額より後ろ（非課税の欄）で
               Σ支給項目 ＝ 課税支給額 → 既に除外済み。S列は空が正
    unknown  : どれも成立しない（内訳が読めない）→ 控除せず §9 に「未控除」を明記
    none     : 通勤手当の印字がない（＝支給なし）

    通勤手当の行は金額の行だけを見る（勤怠欄の通勤日数・通勤距離は除く）。支給項目の範囲の中にある行を優先する。
    T1 の部分集合探索は複数解があり得るのでここでは使わない。
    """
    keys = list(cells)

    def amount(k):
        v = numof(cells[k].value)
        return v if v and v > 0 else None

    def amount_row(prefix):
        return next((k for k in keys if k.startswith(prefix)
                     and not any(w in k for w in COMMUTE_NOT_AMOUNT) and amount(k)), None)

    tk, nk = amount_row(TAXED_COMMUTE), amount_row(NONTAX_COMMUTE)
    if tk or nk:
        # 課税通勤費／非課税通勤費が別行で印字される形式は、その値でそのまま決まる
        return ('included', amount(tk)) if tk else ('excluded', amount(nk))

    cands = [k for k in keys if is_commute_label(k) and not k.startswith(NONTAX_COMMUTE + TAXED_COMMUTE)
             and amount(k)]
    if not cands:
        return 'none', 0.0
    items, tax, tax_cell = item_sum(cells)
    if items is None or tax is None:
        return 'unknown', amount(cands[0])
    in_items = [k for k in cands if any(k == ik for ik, _c in items)]
    i_tax = next(i for i, k in enumerate(keys) if cells[k] is tax_cell)
    after_tax = [k for k in cands if keys.index(k) > i_tax]
    if in_items:
        com_k, in_range = in_items[0], True
    elif after_tax:
        com_k, in_range = after_tax[0], False
    else:
        return 'unknown', amount(cands[0])      # 支給欄より上の通勤の行は金額か判断できない
    com = amount(com_k)
    vals = _amounts(items)
    for cand in (vals, _signed(vals)):
        s = sum(v for _, v in cand)
        if abs(s - tax) < 0.5:
            return ('included' if in_range else 'excluded'), com
        if in_range and abs((s - com) - tax) < 0.5:
            return 'excluded', com
    return 'unknown', com


def ledger_transport(wb):
    """従業員別明細の S列（年間通勤手当）を氏名キーで読む。シートが無ければ None。"""
    if LEDGER_SHEET not in wb.sheetnames:
        return None
    ws = wb[LEDGER_SHEET]
    out = {}
    for r in range(LEDGER_HDR_ROW + 1, ws.max_row + 1):
        nm = name_key(norm(ws.cell(r, LEDGER_COL_NAME).value))
        if nm:
            v = ws.cell(r, LEDGER_COL_S).value
            out[nm] = v if isinstance(v, (int, float)) else None
    return out


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
        name_row = next((r for r in range(1, min(ws.max_row, 14) + 1)
                         if norm(ws.cell(r, 1).value).startswith('氏名')), None)
        parsed[sn] = {'ws': ws, 'hdr': hdr, 'name_row': name_row or hdr,
                      'blocks': blocks_of(ws, hdr)}
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
ok1, ng1, ng1_marked, partial, noitem, signed1, styled1 = 0, [], [], [], [], [], []
for sn, p in parsed.items():
    for bi, blk in enumerate(p['blocks']):
        # 総合計欄・課税支給額だけの賞与欄は支給項目の内訳を持たないのが正常（集計のみでも違反にしない）。
        # 賞与欄でも総支給額などの集計行に値があるなら、賞与額の行を落とした転記として止める
        is_summary_block = blk['heads'] and all(is_bonus_head(h) for h in blk['heads']) and (
            all(h.startswith('総合計') for h in blk['heads'])
            or all(is_taxable(h) for h in valued_labels(p['ws'], blk)))
        bkey, seen = f'{sn}#b{bi}', 0
        judged = [(name, *sum_verdict(cells)) for name, cells in entity_maps(p['ws'], p, blk)]
        # 課税支給額の行（列）がまるごとオレンジ＝装飾。個々の不整合の印とはみなさない
        row_styled = all_marked(tc for _n, v, _d, tc in judged if v != 'skip' and tc is not None)
        if row_styled:
            styled1.append(bkey)
        for name, verdict, detail, tax_cell in judged:
            if verdict == 'skip':
                continue
            seen += 1
            if verdict == 'ok':
                ok1 += 1
            elif verdict == 'signed':
                signed1.append(f'{sn}/{name}')
                ok1 += 1
            elif verdict == 'partial':
                partial.append(f'{sn}/{name}: {detail}')
                ok1 += 1
            elif verdict == 'noitem':
                # 課税支給額はあるのに内訳が1つもない＝集計行だけの転記（§4.1.2 🔴 違反）
                if is_summary_block:
                    if bkey not in noitem:
                        noitem.append(bkey)
                else:
                    ng1.append(f'{sn}/{name}: 課税支給額はあるが支給項目（基本給・各手当）が'
                               '無い＝集計行のみの転記の疑い（基本給の行が無いページは【支給】の見出し行を残す）')
            elif is_marked(tax_cell) and not row_styled:
                # 原本の課税支給額そのものが内訳と合わない（原本側の不整合）。転記者がオレンジの印を付けたもの
                ng1_marked.append(f'{sn}!{tax_cell.coordinate}({name}): {detail}')
            else:
                ng1.append(f'{sn}/{name}: {detail}')
        if not seen and bkey not in noitem:
            noitem.append(bkey)
if partial:
    uniq = sorted({x.split(': ', 1)[1] for x in partial})
    print(f'    非課税項目が内訳に混在: {len(partial)}件（部分集合一致で課税対象を特定）→ {uniq[:3]}')
if signed1:
    print(f'    ℹ 減額項目（勤怠控除など）を差し引いて成立: {len(signed1)}列 {signed1[:4]}'
          '（原本で正の数の印字かを確認。原本が△・マイナスなら符号ごと転記する）')
if noitem:
    print(f'    支給項目行を持たないブロック（総合計欄など）: {len(noitem)}件 {noitem[:6]}')
if ng1_marked:
    print(f'    原本側の不整合としてオレンジの印あり: {len(ng1_marked)}件 {ng1_marked[:3]}'
          '（セル番地と差額を表紙と §9 に書く）')
if styled1:
    print(f'    ⚠ 課税支給額の行全体がオレンジのブロック: {styled1[:4]} — 行ごとの色は印とみなさない。'
          '食い違うセルだけをオレンジにする')
if not (ok1 or ng1 or ng1_marked):
    # 1列も検算できていないのに成立扱いにしない（課税支給額の行名を読めていない）
    ng1.append('課税支給額の行を読めず、1列も検算できなかった。まず集計行だけの転記になっていないかを確かめる。'
               f'課税支給額の行名が {list(TAXABLE)} のどれでも始まらないなら、その行名を §9 に書いて'
               'スクリプトへの追加を依頼する。課税支給額に当たる行がそもそも無い台帳なら §9 にその旨を書く')
check('T1', f'課税支給額＝内訳の和が全列で成立（{ok1}列）', not ng1, '; '.join(ng1[:4]))

print('=== T2 関数計算の可能性（数値セルであること） ===')
bad2 = []
for sn, p in parsed.items():
    for blk in p['blocks']:
        for name, cells in entity_maps(p['ws'], p, blk):
            if is_check_name(name):
                continue                       # 検算の行（=SUM の式）は対象外
            items, _tax, tax_cell = item_sum(cells)
            if tax_cell is not None and tax_cell.value not in (None, '') and numof(tax_cell.value) is None:
                # 課税支給額は検算の基準なので、式や記号つきの文字列も認めない（数値か空欄）
                tax_k = next(k for k in cells if cells[k] is tax_cell)
                bad2.append(f'{sn}!{tax_cell.coordinate}({tax_k})={tax_cell.value!r}')
            for lab, c in (items or []):
                # 内訳は「数字の入った文字列」だけを見る（=SUM の検算や「月給」等の文字の行は対象外）
                if has_digit_text(c.value):
                    bad2.append(f'{sn}!{c.coordinate}({lab})={c.value!r}')
check('T2', '課税支給額は数値か空欄、内訳に「数字の入った文字列」が無い', not bad2, '; '.join(bad2[:4]))

print('=== T3 原本の網羅性（合計列 ＝ 転記値のSUM） ===')
ok3, ng3, ng3_marked, nototal, formula_only_totals, styled3 = 0, [], [], [], set(), []


def row_sum(ws, blk, r):
    return sum(ws.cell(r, c).value for c in blk['data'] if isinstance(ws.cell(r, c).value, (int, float)))


def page_sum_cols(ws, p, blk):
    """そのページの和を印字した合計列（数値のある行の半分以上で SUM と一致する列）。
    自分で計算した「小計」と原本の「合計」が並んでいても、原本の列を照合から外さないために使う。
    月全体の《合計》や区切りごとの小計はほとんどの行で合わないので、ここには入らない。"""
    out = []
    for h, c in blk['totals']:
        rows = [r for r in range(p['hdr'], ws.max_row + 1) if numof(ws.cell(r, c).value) is not None]
        hits = sum(1 for r in rows if abs(ws.cell(r, c).value - row_sum(ws, blk, r)) < 0.5)
        if rows and hits * 2 >= len(rows):
            out.append(c)
    return out


def t3_row_ok(ws, blk, r, s, page_cols):
    """ページの和の列があれば、そのすべてが行の SUM と一致すること。無ければ、合計の候補のどれかが一致するか、
    「小計」が2本以上のときは区切りごとの小計がすべて一致し、どの区切りにも入らない列が無いこと。"""
    if page_cols:
        return all(numof(ws.cell(r, c).value) is None or abs(ws.cell(r, c).value - s) < 0.5 for c in page_cols)
    cand = [(h, c, ws.cell(r, c).value) for h, c in blk['totals'] if numof(ws.cell(r, c).value) is not None]
    if any(abs(v - s) < 0.5 for _h, _c, v in cand):
        return True
    subs = [x for x in cand if x[0].startswith('小計')]
    if len(subs) < 2:
        return False
    prev = blk['label_col']
    for h, c, v in sorted(cand, key=lambda x: x[1]):
        seg = sum(ws.cell(r, d).value for d in blk['data']
                  if prev < d < c and isinstance(ws.cell(r, d).value, (int, float)))
        if h.startswith('小計') and abs(seg - v) >= 0.5:
            return False
        prev = c
    return not any(d > prev for d in blk['data'])   # 最後の小計より右の列は、どの小計でも照合されていない


for sn, p in parsed.items():
    ws = p['ws']
    for bi, blk in enumerate(p['blocks']):
        if not blk['total'] or is_transposed(blk):
            nototal.append(f'{sn}#b{bi}')
            if blk['formula_only_total']:
                formula_only_totals.add(f'{sn}#b{bi}')
            continue
        col_cells = [ws.cell(r, blk['total']) for r in range(p['hdr'], ws.max_row + 1)
                     if numof(ws.cell(r, blk['total']).value) is not None]
        col_styled = all_marked(col_cells)
        if col_styled:
            styled3.append(f'{sn}#b{bi}')
        page_cols = page_sum_cols(ws, p, blk)
        for r in range(p['hdr'], ws.max_row + 1):
            printed = ws.cell(r, blk['total']).value
            if not isinstance(printed, (int, float)):
                continue
            s = row_sum(ws, blk, r)
            if t3_row_ok(ws, blk, r, s, page_cols):
                ok3 += 1
            else:
                lab = norm(ws.cell(r, blk['label_col']).value)
                msg = f'{sn} r{r}({lab}): SUM{s:,.0f} vs 印字{printed:,.0f}'
                marked = not col_styled and any(is_marked(ws.cell(r, c)) for _h, c in blk['totals'])
                (ng3_marked if marked else ng3).append(msg)
if nototal:
    print(f'    合計列を持たないブロック（一覧型を含む）: {len(nototal)}件 {nototal[:6]}')
if ng3_marked:
    print(f'    原本側の不整合としてオレンジの印あり: {len(ng3_marked)}件 {ng3_marked[:3]}')
if styled3:
    print(f'    ⚠ 合計列全体がオレンジのブロック: {styled3[:4]} — 列ごとの色は印とみなさない')
formula_total = [k for k in nototal if k in formula_only_totals]
if formula_total:
    print(f'    ℹ 合計の見出しはあるが数値が無いブロック: {len(formula_total)}件 {formula_total[:6]}'
          '（=SUM の式だけの列は検算に使えない。原本に合計の印字があれば値で入れ、式の列は見出しを「検算」で始める）')
check('T3', f'原本の合計とSUMが全行一致（{ok3}行・オレンジの印の原本不整合は除く）',
      not ng3, '; '.join(ng3[:4]))

print('=== T4 加点判定への転用（基本給＋労働時間の手掛かり） ===')
ng4, rate_pages, empty_time, skipped4 = [], [], [], []


for sn, p in parsed.items():
    monthly = []
    for blk in p['blocks']:
        if not blk['heads']:
            continue
        if is_bonus_block(p['ws'], p, blk) or is_officer_block(p['ws'], p, blk):
            skipped4.append(sn)
            continue
        monthly.append(blk)
    for blk in monthly:
        labels = block_labels(blk)
        # 日給制で基本給の行が無いページは、支給欄の見出し行（値の無い【支給】等）か単価の行があれば分子が取れる
        has_base = any(is_base(h) or h.startswith(RATE_HINTS)
                       or is_pay_section(h, label_cells(p['ws'], blk, h)) for h in labels)
        time_labels = [h for h in labels if is_time_label(h)]
        if not (has_base and time_labels):
            ng4.append(f'{sn}: 基本給={has_base} 時間手掛かり={bool(time_labels)}')
        elif not any(numof(c.value) for h in time_labels for c in label_cells(p['ws'], blk, h)):
            empty_time.append(sn)
    rr = next((v for blk in monthly if not is_transposed(blk)
               for k, v in blk['rows'].items() if k.startswith(RATE_HINTS)), None)
    if rr:
        n = sum(1 for blk in monthly for c in blk['data']
                if isinstance(p['ws'].cell(rr, c).value, (int, float)))
        rate_pages.append((sn, n))
check('T4', '月次ページすべてで基本給＋労働時間の手掛かりあり', not ng4, '; '.join(ng4[:4]))
if ng4:
    print('    → 原本に時間・日数・単価の印字が無い台帳なら NG のままでよい。§9 に「原本に印字なし」と書く'
          '（加点判定には勤怠の資料が別に要る）。印字があるのに転記していなければ転記する')
if skipped4:
    print(f'    ℹ 賞与だけ・役員報酬だけのページ（勤怠が無いのが普通なので対象外）: {len(set(skipped4))}件')
if empty_time:
    print(f'    ℹ 時間・日数の行はあるが値が1つも無いページ: {len(empty_time)}件 {empty_time[:6]}'
          '（原本が空欄ならそのままでよい。加点判定に転用するときは別の勤怠資料が要る）')
print(f'    単価（日額/時給）が転記されたページ: {len(rate_pages)}/{len(parsed)}'
      + (f' {rate_pages}' if rate_pages else '（この台帳形式には単価の印字なし）'))

print('=== T5 通勤手当のS列整合（控除漏れ・二重控除の検知。§3.1） ===')
states, amounts = {}, {}
anon = {'included': 0, 'excluded': 0, 'unknown': 0}
for sn, p in parsed.items():
    for blk in p['blocks']:
        for name, m in entity_maps(p['ws'], p, blk):
            st, amt = commute_state(m)
            if st == 'none':
                continue
            if non_person(name):
                anon[st] += 1          # 氏名でない列・行（合計行・部門計・検算）
                continue
            # 月の見出し・番号だけの列（氏名の行が無い個人台帳など）は、誰の S列とも突き合わせられない
            key = name if (MONTH_NAME_RE.fullmatch(name) or name.isdigit()) else name_key(name)
            states.setdefault(key, set()).add(st)
            if amt:
                amounts.setdefault(key, []).append(amt)

s_col = ledger_transport(wb)
missing, doubled, unknown_names, mixed, unlisted, unmatched, ok5 = [], [], [], [], [], [], 0
for name, sts in sorted(states.items()):
    if s_col is not None and name not in s_col:
        # 従業員別明細に無い名前。通勤手当が課税支給額に含まれているなら控除漏れを見逃さないよう止める。
        # 月の見出し・番号だけの列は誰の S列とも突き合わせられないので、判定がつけば止める（氏名の行を置かせる）
        unattributable = bool(MONTH_NAME_RE.fullmatch(name) or name.isdigit())
        stop = 'included' in sts or (unattributable and 'excluded' in sts)
        (unmatched if stop else unlisted).append(name)
        continue
    if 'included' in sts and 'excluded' in sts:
        mixed.append(name)
        continue
    st = 'included' if 'included' in sts else ('excluded' if 'excluded' in sts else 'unknown')
    if st == 'unknown':
        unknown_names.append(name)
        continue
    if s_col is None:
        ok5 += 1
        continue
    sv = s_col.get(name)
    if st == 'included' and not sv:
        missing.append((name, sum(amounts.get(name, []))))
    elif st == 'excluded' and sv:
        doubled.append((name, sv))
    else:
        ok5 += 1

n_inc = sum(1 for s in states.values() if s == {'included'})
n_exc = sum(1 for s in states.values() if s == {'excluded'})
if not states and not any(anon.values()):
    print('    通勤手当の印字なし＝支給なし（控除不要）。§9 には「対象なし（支給なし）」と書く')
else:
    print(f'    判定: 課税支給額に含まれる {n_inc}名 / 非課税処理済み {n_exc}名 / '
          f'判定不能 {len(unknown_names)}名 / 混在 {len(mixed)}名'
          + (f' / 氏名が取れない列・行 {anon}' if any(anon.values()) else ''))
if s_col is None and states:
    print(f'    ⚠ 同じブックに「{LEDGER_SHEET}」シートが無いため S列との突合はスキップした')
if missing:
    print(f'    🔴 [控除漏れ] {len(missing)}名 — 課税支給額に通勤手当が含まれているのに S列が空:')
    for nm, tot in missing[:10]:
        print(f'      {nm}: 転記シート上の通勤手当合計 {tot:,.0f} 円')
    print('    → 2026-07-27 運用では課税扱いの通勤手当も控除する（§3.1）。転記シートは対象年度外の'
          'ページも含むので、S列には**対象12ヶ月分だけ**を集計した年額を入れる')
if doubled:
    print(f'    🔴 [二重控除] {len(doubled)}名 — 非課税処理済みなのに S列に値がある:')
    for nm, sv in doubled[:10]:
        print(f'      {nm}: S列={sv:,.0f} 円 → 空欄にする')
    print('    → 課税支給合計にはそもそも通勤手当が入っていない。S列を埋めると引かれるのは'
          '基本給・残業手当で、R216 が過小になる')
if unmatched:
    print(f'    🔴 [明細と突合できない] {len(unmatched)}件 {unmatched[:8]} — 課税支給額に通勤手当が含まれているが、'
          '従業員別明細に同じ氏名が無い。対象年度の従業員なら氏名の表記を明細とそろえる（1人1ページで列が月の'
          '台帳は、各列の上に「氏名」の行を置く）。対象年度外の退職者なら §9 にその旨を書く')
if mixed:
    print(f'    ⚠ [判定が混在] {mixed[:10]} — ページによって課税/非課税処理が違う。'
          '対象事業年度のページの処理に合わせ、根拠を §9 に書く')
if unknown_names:
    print(f'    ℹ [判定不能] {len(unknown_names)}名 {unknown_names[:8]} — 内訳から課税支給額との関係が'
          '決まらない（諸手当に合算・項目の取り違え等）。**控除せず** §9 に「未控除（内訳不明）」と'
          '理由・人数を書く（顧客に内訳を催促しない。§3.1.1）')
if unlisted:
    print(f'    ℹ 従業員別明細に無い名前（通勤手当は課税支給額の外か判定不能）: {len(unlisted)}件 {unlisted[:8]}'
          ' — 対象年度外の退職者・月の見出しなどなら問題なし')
check('T5', f'S列が §3.1 の判定表と整合（整合 {ok5}名）',
      not (missing or doubled or mixed or unmatched),
      f'控除漏れ{len(missing)}名 / 二重控除{len(doubled)}名 / 混在{len(mixed)}名 / 突合できない{len(unmatched)}件')

nfail = sum(1 for _, ok in results if not ok)
print()
if nfail:
    print(f'=== 要確認 {nfail}件 ===（NGはPDF原本で確認し、§9 報告に判断を明記してから完了報告へ。'
          '原本側の不整合なら食い違うセルだけをオレンジ（標準の色 FFC000）にして表紙に注記する。'
          '黄色は「読み取りに自信なし」の印で、免除されない）')
else:
    print('=== PASS ===（T1〜T5 すべて成立。ただし転記の「真」はPDF原本なので '
          '§5.5-3 の目視突合は別途必要）')
sys.exit(1 if nfail else 0)
