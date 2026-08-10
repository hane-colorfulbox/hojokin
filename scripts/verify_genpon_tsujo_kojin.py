# 生成した通常枠×個人 原本下書きの構造検証
import sys, re, unicodedata
from pathlib import Path
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
BASE = Path(r'C:\Users\user\projects\カラフルボックス\補助金')
SCRATCH = Path(__file__).parent
NG = []
OK = []


def resolve(rel):
    p = BASE / rel if not Path(rel).is_absolute() else Path(rel)
    if p.exists():
        return p
    want = unicodedata.normalize('NFC', p.name)
    for cand in p.parent.iterdir():
        if unicodedata.normalize('NFC', cand.name) == want:
            return cand
    raise FileNotFoundError(rel)


def check(cond, ok_msg, ng_msg):
    (OK.append(ok_msg) if cond else NG.append(ng_msg))


def norm(s):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(s or '')))


wb = openpyxl.load_workbook(resolve('ツール/【原本_個人】企業名_通常枠_個人2026.xlsx'))
tenki, shinsei, seisan = wb['転記'], wb['申請内容'], wb['生産性指標給与支給総額計算']

# 1. 全シートの数式スキャン: #REF / インボイス専用シート参照 / 未解決転記参照
bad_tokens = []
tenki_labels = {r: tenki.cell(r, 1).value for r in range(1, 111)}
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not (isinstance(v, str) and v.startswith('=')):
                continue
            if '#REF' in v:
                bad_tokens.append(f'{ws.title}!{cell.coordinate}: #REF')
            if "'給与支給総額計算'" in v or "'給与支給総額計算 (旧)'" in v:
                bad_tokens.append(f'{ws.title}!{cell.coordinate}: インボイス版シート参照残存')
            if "'ツールマスタ'" in v:
                bad_tokens.append(f'{ws.title}!{cell.coordinate}: ツールマスタ参照(通常枠はシート9のはず)')
            for m in re.finditer(r"'転記'!\$?B\$?(\d+)", v):
                r = int(m.group(1))
                if not (tenki_labels.get(r) and str(tenki_labels[r]).strip()):
                    bad_tokens.append(f'{ws.title}!{cell.coordinate}: 転記!B{r} のラベルが空')
check(not bad_tokens, '数式スキャン: #REF/残存参照/空ラベル参照 なし', f'数式スキャンNG: {bad_tokens}')

# 2. 申請内容⇔転記のラベル整合
# 申請内容の表示ラベルとヒアリング側の設問文が意図的に違うペア。
# 下4件は素材テンプレ（通常枠法人 v2 / インボイス個人）にも同じ形で存在する継承分で、
# この組み立てが作った不整合ではない（2026-08-10 に両テンプレで実測確認）。
ALLOW = {('事業所開始年月日', '事業開始年月日'), ('現在住所の郵便番号', '現住所：郵便番号'),
         ('現在住所', '現住所'), ('事業所所在地の郵便番号', '事業所所在地：郵便番号'),
         ('SECURITYACTION自己宣言ID', 'SECURITYACTION自己宣言ID'),
         ('屋号・商号フリガナ', '屋号・商号（フリガナ）'),
         ('強み（転記）', '自社の強み（複数選択可）'),
         ('弱み（転記）', '自社の弱み（複数選択可）'),
         ('どのようなプロセスに対してＩＴ投資を行ったか（転記）',
          'どのようなプロセスに対してIT投資を行いましたか。（複数選択可）')}
mism = []
for r in range(1, 240):
    v = shinsei.cell(r, 3).value
    if isinstance(v, str):
        m = re.fullmatch(r"='転記'!B(\d+)", v)
        if m:
            tr = int(m.group(1))
            a, b = norm(shinsei.cell(r, 2).value), norm(tenki_labels.get(tr))
            if a and b and a != b and not a.startswith(b) and not b.startswith(a) \
               and (shinsei.cell(r, 2).value, tenki_labels.get(tr)) not in ALLOW \
               and (a, b) not in {(norm(x), norm(y)) for x, y in ALLOW}:
                mism.append(f'申請内容r{r}[{a[:20]}] vs 転記r{tr}[{b[:20]}]')
check(not mism, f'転記参照のラベル整合: 全一致（許容差分除く）', f'ラベル不整合: {mism}')

# 3. 転記 と 通常枠個人ヒアリングシート基本情報のフィールド対応（順序保存の包含）
# 参照先はリポジトリ管理下の ツール/ に置く（_debug/ は gitignore で他環境に無いため）
hear = openpyxl.load_workbook(resolve('ツール/ヒアリングシート2026_通常枠個人.xlsx'))['基本情報']
hear_fields = []
for r in range(5, 104):
    v = hear.cell(r, 2).value
    if v and str(v).strip() and not str(v).startswith(('※', '⇨', '◼', '【')):
        hear_fields.append((r, norm(v)))
tenki_fields = [(r, norm(l)) for r, l in tenki_labels.items()
                if l and not str(l).startswith(('項目', '◼', '【', '⇨', '↓'))]
tf_norms = [t for _, t in tenki_fields]
missing = [f'ヒアr{r}:{v[:22]}' for r, v in hear_fields if v not in tf_norms]
# 転記だけにあるフィールド（ヒアに無い）も列挙
hf_norms = [t for _, t in hear_fields]
extra = [f'転記r{r}:{v[:22]}' for r, v in tenki_fields if v not in hf_norms]
print('— ヒアにあり転記に無い:', missing if missing else 'なし')
print('— 転記にありヒアに無い:', extra if extra else 'なし')

# 4. 主要固定値
check(norm(shinsei.cell(77, 3).value) == '12月', '決算月=12月固定', f'決算月NG: {shinsei.cell(77,3).value!r}')
check(shinsei.cell(62, 3).value == 0, '資本金(基本情報)=0', f'資本金NG: {shinsei.cell(62,3).value!r}')
check(shinsei.cell(134, 3).value == 0, '資本金(財務)=0', f'財務資本金NG: {shinsei.cell(134,3).value!r}')
check('個人事業主は基本「1」' in str(shinsei.cell(126, 4).value), '役員数note=個人は1', f'役員数noteNG: {shinsei.cell(126,4).value!r}')

# 5. シート9連動とツール名
check("VLOOKUP($C$74,'シート9'" in str(shinsei.cell(151, 3).value), 'シート9連動アンカー=$C$74', f'シート9アンカーNG: {shinsei.cell(151,3).value!r}')
check(norm(shinsei.cell(74, 2).value) == 'ツール名', 'r74=ツール名', f'r74ラベルNG: {shinsei.cell(74,2).value!r}')
dv_map = {}
for dv in shinsei.data_validations.dataValidation:
    dv_map[str(dv.sqref)] = dv.formula1
check(dv_map.get('C74') == "'シート9'!$A$5:$A$12", 'C74のDV=シート9ツールリスト', f'C74 DV NG: {dv_map.get("C74")!r}')

# 6. 生産性指標の参照先ラベル
expect = {123: '従業員数：正規雇用', 124: '従業員数：契約社員', 125: '従業員数：パートアルバイト',
          126: '代表者・役員数', 127: '年間の平均労働時間'}
for r, lab in expect.items():
    check(norm(shinsei.cell(r, 2).value) == norm(lab), f'r{r}={lab}', f'r{r}ラベルNG: {shinsei.cell(r,2).value!r}')

# 7. 郵便番号IMPORTXMLの参照先
z1, z2 = str(shinsei.cell(54, 3).value), str(shinsei.cell(59, 3).value)
check('ENCODEURL(C55)' in z1, '郵便番号①→C55(現在住所)', f'郵便番号①NG: {z1[:80]}')
check('ENCODEURL(C60)' in z2, '郵便番号②→C60(事業所所在地)', f'郵便番号②NG: {z2[:80]}')

# 8. プロンプトの参照と枠名
pv = str(shinsei.cell(75, 3).value)
for tok in ['C65', 'C66', 'C67', 'C68', 'C69', 'C70', 'C71', 'C72', 'C73', 'C74']:
    check(re.search(rf'(?<![A-Z0-9]){tok}(?![0-9])', pv), f'プロンプト参照{tok}あり', f'プロンプト参照{tok}なし')
check('（通常枠）' in pv and 'インボイス枠' not in pv, 'プロンプト枠名=通常枠', 'プロンプト枠名NG')
check("DUMMYFUNCTION(\"AI(C75)" in str(shinsei.cell(76, 3).value), 'AI()→C75', f'AI参照NG: {str(shinsei.cell(76,3).value)[:60]}')
check(str(shinsei.cell(76, 5).value) == '=LEN(C76)', 'LEN→C76', f'LEN NG: {shinsei.cell(76,5).value!r}')

# 9. 添付欄・必要書類欄
for r, lab in [(9, '身分証明書'), (10, '令和7年所得税の納税証明書'), (11, '確定申告書'),
               (13, '賃金状況報告シート'), (163, '身分証明書'), (165, '所得税納税証明書'),
               (167, '確定申告書'), (169, '所得税の青色申告決算書又は収支内訳書'), (171, 'その他資料')]:
    cellv = shinsei.cell(r, 3).value if r < 20 else shinsei.cell(r, 2).value
    check(norm(lab)[:10] in norm(cellv), f'r{r}≈{lab}', f'r{r}NG: {cellv!r}')

# 10. 計画年ラベル（暦年）
y = [str(shinsei.cell(r, 2).value) for r in (195, 196, 197)]
check('2027/1～2027/12' in y[0] and '2028/1' in y[1] and '2029/1' in y[2], '計画年=暦年2027-2029', f'計画年NG: {y}')

# 11. 通常枠固有ブロックの存置
for r, lab in [(37, '申請類型選択'), (43, '指定する一定期間において3ヶ月以上'),
               (140, '経営意欲'), (151, '補助金を利用してもっとも改善したい業務プロセス'),
               (212, '事業実施年度内における賃金状況について'), (220, '労働生産性指標')]:
    check(norm(lab)[:12] in norm(shinsei.cell(r, 2).value), f'通常枠ブロックr{r}={lab[:12]}', f'通常枠ブロックr{r}NG: {shinsei.cell(r,2).value!r}')

# 12. DV・結合・CFの総数と範囲
dvs = shinsei.data_validations.dataValidation
check(len(dvs) == 14, f'DV数=14', f'DV数NG: {len(dvs)}')
oob = [str(dv.sqref) for dv in dvs if any(int(m.group(1)) > 240 for m in re.finditer(r'[A-Z](\d+)', str(dv.sqref)))]
check(not oob, 'DV範囲すべて240行以内', f'DV範囲外: {oob}')
merges = [str(m) for m in shinsei.merged_cells.ranges]
check(all(int(re.search(r'(\d+)$', m).group(1)) <= 240 for m in merges), f'結合{len(merges)}件すべて240行以内', '結合範囲外あり')
cfs = list(shinsei.conditional_formatting)
check(len(cfs) == 5, 'CF=5系統', f'CF数: {len(cfs)}')

# 13. 法人前提の継承文言が残っていないか（個人事業主に存在しない書類を案内してしまう）
inherited = []
for r in range(1, 240):
    for c in range(2, 6):
        v = shinsei.cell(r, c).value
        if isinstance(v, str):
            for kw in ['履歴事項', '法人税', '貸借対照表', '損益計算書', '登記', 'インボイス']:
                if kw in v:
                    inherited.append(f'r{r}{chr(64+c)}:{kw}')
check(not inherited, '法人前提の継承文言なし', f'法人前提の文言が残存: {inherited}')

# 14. 最終行（法人版の孤児『提出完了しました』が残っていないか）
last_row = max((c.row for row in shinsei.iter_rows() for c in row if c.value is not None), default=0)
check(last_row == 239, '申請内容の最終非空行=239（孤児なし）', f'最終非空行NG: {last_row}（法人版r261の消し残りの疑い）')
check('提出完了' in str(shinsei.cell(239, 2).value), 'r239=提出完了しました', f'r239NG: {shinsei.cell(239,2).value!r}')

# 15. シート9 の平文ガイドが新レイアウトの行番号を指しているか（数式ではないので式スキャンに出ない）
sheet9 = wb['シート9']
guide = {'A2': ['C74', 'C151', 'C153'], 'B4': ['C151'], 'C4': ['C152'], 'D4': ['C153'],
         'E4': ['C75'], 'A17': ['C151'], 'B17': ['C152'], 'C17': ['C153']}
bad_guide = []
for coord, needs in guide.items():
    v = str(sheet9[coord].value or '')
    if not all(n in v for n in needs):
        bad_guide.append(f'{coord}={v[:40]!r}')
    for old in ('C71', 'C72', 'C179', 'C180', 'C181'):
        if re.search(rf'(?<![0-9]){old}(?![0-9])', v):
            bad_guide.append(f'{coord} に法人版の {old} が残存')
check(not bad_guide, 'シート9の案内文が新レイアウト（C74/C75/C151-153）', f'シート9案内文NG: {bad_guide}')

print()
print(f'OK {len(OK)} 件')
for m in OK:
    print('  ✓', m)
print(f'NG {len(NG)} 件')
for m in NG:
    print('  ✗', m)
sys.exit(1 if NG else 0)
