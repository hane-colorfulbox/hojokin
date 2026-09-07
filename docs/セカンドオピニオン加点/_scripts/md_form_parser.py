"""会社別 面談ヒアリングシート(05) の Markdown → 設問データ(FORM 形式) 変換。

使い方（単体では検証のみ。生成は make_* / build_company_sheets.py 側）:
    python md_form_parser.py <05_面談ヒアリングシート_XXX.md> [...]

なぜパースするのか:
  汎用版は `hearing_form.FORM`（Python データ）が単一の真実で、そこから
  入力用PDF / チェック可能xlsx / md を生成している。会社別版は6社ぶんあり、
  同じ内容を Python データと Markdown で二重管理すると必ずズレる。
  そこで**会社別は Markdown を単一の真実**とし、生成物はここでパースして作る。

出力ブロックは hearing_form.FORM と同じ形（+ 'subsection'）:
    ('title', text) / ('note', text) / ('section', title) / ('subsection', title)
    ('row', [segments]) / ('space', dy)
segment も同じ:
    ('t', text) 静的文字 / ('c', label) チェック肢 / ('f', width) 記入欄（PDFのミリ幅相当）

Markdown 側の記法（会社別6本で実測した前提。崩すとパース結果が変わる）:
    # …                 タイトル
    ## N. …             セクション
    ### …               サブセクション
    ---                 区切り（読み飛ばす）
    □ラベル             チェック肢。行に □ が2つ以上／`：□`／`- □` で「設問行」と判定
    （　　　）           全角スペースだけの括弧＝記入欄。〔　〕() も同様
    　　（2連以上）      括弧の内側などに現れる裸の記入欄（例:「衛生士　　名」「　　時　分」）
    **…**               強調は外す（Word/PDF では見出し扱いしない）
"""
import re
import sys
import pathlib

sys.stdout.reconfigure(encoding="utf-8")

# ---- 記入欄の判定・幅 ----
BLANK_CHARS = "　 ／/"                     # 記入欄の内側に許す文字（全角/半角スペース・スラッシュ）
PAREN_BLANK = re.compile(r"[（(〔\[][　 ／/]*[）)〕\]]")   # 例: （　　） （　／　） 〔　　〕
BARE_BLANK = re.compile(r"　{2,}")           # 例: 衛生士　　名 ／ 　　時　分
CHECK = "□"
FIELD_W_PER_CHAR = 15                        # 空白1文字あたりの記入欄幅（pt）
FIELD_W_MIN, FIELD_W_MAX = 38, 300

# ---- 行の種類 ----
RE_TITLE = re.compile(r"^#\s+(.*)$")
RE_SECTION = re.compile(r"^##\s+(.*)$")
RE_SUBSECTION = re.compile(r"^###\s+(.*)$")
RE_HR = re.compile(r"^-{3,}$")
RE_BULLET = re.compile(r"^[-*]\s+")
RE_BOLD = re.compile(r"\*\*(.+?)\*\*")
# チェック肢ラベルの終わり: 次の □ ／ 空白+スラッシュ ／ 空白だけの括弧
LABEL_END = re.compile(r"□|(?<=\S)[\s　]+[／/]|" + PAREN_BLANK.pattern)


def _field_width(blank_text):
    n = sum(1 for ch in blank_text if ch in BLANK_CHARS)
    return min(FIELD_W_MAX, max(FIELD_W_MIN, n * FIELD_W_PER_CHAR))


def _first_blank(text):
    """text 中の最初の記入欄を (start, end, matched) で返す。無ければ None。"""
    cands = [m for m in (PAREN_BLANK.search(text), BARE_BLANK.search(text)) if m]
    if not cands:
        return None
    m = min(cands, key=lambda m: m.start())
    return m.start(), m.end(), m.group(0)


def _text_segments(text):
    """静的文字列を ('t', …) と ('f', 幅) の並びに分解する。"""
    segs = []
    while text:
        found = _first_blank(text)
        if not found:
            if text:
                segs.append(("t", text))
            break
        s, e, matched = found
        if s:
            segs.append(("t", text[:s]))
        segs.append(("f", _field_width(matched)))
        text = text[e:]
    return segs


def _check_segments(raw_label):
    """□ に続く生ラベルを ('c', ラベル) + 後続の静的文字に分解する。

    「複数（　　か所）」のようにラベル内に記入欄があるものは、括弧の手前で
    チェック肢を切り、残りを静的文字（記入欄を含む）として返す。
    """
    found = _first_blank(raw_label)
    if not found:
        return [("c", raw_label.strip())] if raw_label.strip() else []
    s, _, _ = found
    head, rest = raw_label[:s], raw_label[s:]
    for opener in "（(〔[":                      # 「複数（」→ チェック肢は「複数」、括弧は静的文字へ
        if head.rstrip().endswith(opener):
            head = head.rstrip()[:-1]
            rest = opener + rest
            break
    segs = []
    if head.strip():
        segs.append(("c", head.strip()))
    segs.extend(_text_segments(rest))
    return segs


def parse_line(line):
    """1行を segments に分解する。"""
    segs, pos = [], 0
    while pos < len(line):
        idx = line.find(CHECK, pos)
        if idx < 0:
            segs.extend(_text_segments(line[pos:]))
            break
        if idx > pos:
            segs.extend(_text_segments(line[pos:idx]))
        after = line[idx + 1:].lstrip("　 ")
        skipped = len(line[idx + 1:]) - len(after)
        m = LABEL_END.search(after)
        raw_label = after[: m.start()] if m else after
        segs.extend(_check_segments(raw_label))
        pos = idx + 1 + skipped + len(raw_label)
    return segs


def is_question_line(line):
    """設問行（チェック肢または記入欄を持つ行）かどうか。

    地の文にも「あてはまる□にチェック」のように □ が出るため、
    チェック肢としては「□が2つ以上」「：□」「行頭の - □」のいずれかを要求する。
    """
    if _first_blank(line):
        return True
    if line.count(CHECK) >= 2:
        return True
    if re.search(r"[：:]\s*" + CHECK, line):
        return True
    return bool(re.match(r"^" + CHECK, line))


def parse_md(path):
    """Markdown ファイルを {'title', 'lead', 'blocks'} に変換する。"""
    raw = pathlib.Path(path).read_text(encoding="utf-8").replace("\r\n", "\n")
    title, lead, blocks = "", [], []
    for line in raw.split("\n"):
        line = RE_BOLD.sub(r"\1", line.strip())
        if not line or RE_HR.match(line):
            continue
        m = RE_TITLE.match(line)
        if m:
            title = m.group(1).strip()
            blocks.append(("title", title))
            continue
        m = RE_SUBSECTION.match(line)          # ### は ## より先に判定する
        if m:
            blocks.append(("subsection", m.group(1).strip()))
            continue
        m = RE_SECTION.match(line)
        if m:
            blocks.append(("section", m.group(1).strip()))
            continue
        body = RE_BULLET.sub("", line)
        if is_question_line(body):
            segs = parse_line(body)
            if segs:
                blocks.append(("row", segs))
        else:
            blocks.append(("note", body))
            if not any(b[0] in ("section", "row") for b in blocks):
                lead.append(body)
    return {"title": title, "lead": lead, "blocks": blocks}


# ---- 検証（取りこぼし検知）----

def render_plain(blocks):
    """パース結果を素のテキストへ戻す（元の md と突き合わせる検証用）。"""
    out = []
    for block in blocks:
        kind = block[0]
        if kind in ("title", "note", "section", "subsection"):
            out.append(block[1])
        elif kind == "row":
            out.append("".join(
                seg[1] if seg[0] == "t" else (CHECK + seg[1] if seg[0] == "c" else "_")
                for seg in block[1]
            ))
    return "\n".join(out)


def _normalize(text):
    """比較用の正規化（記号・空白・記入欄の中身の差を無視する）。"""
    text = RE_BOLD.sub(r"\1", text)
    text = re.sub(r"^[-*#]+\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"^-{3,}$", "", text, flags=re.MULTILINE)
    text = PAREN_BLANK.sub("_", text)
    text = BARE_BLANK.sub("_", text)
    text = re.sub(r"[（(〔\[）)〕\]_\s　／/]+", "", text)
    return text


def verify(path):
    """元 md と パース→再描画 の文字列が（記号・空白を除いて）一致するか検証する。"""
    src = pathlib.Path(path).read_text(encoding="utf-8")
    parsed = parse_md(path)
    a, b = _normalize(src), _normalize(render_plain(parsed["blocks"]))
    if a == b:
        n_row = sum(1 for x in parsed["blocks"] if x[0] == "row")
        n_check = sum(1 for x in parsed["blocks"] if x[0] == "row"
                      for s in x[1] if s[0] == "c")
        n_field = sum(1 for x in parsed["blocks"] if x[0] == "row"
                      for s in x[1] if s[0] == "f")
        return True, f"設問行 {n_row} / チェック肢 {n_check} / 記入欄 {n_field}"
    # 差分の位置を1か所だけ示す
    i = next((i for i, (x, y) in enumerate(zip(a, b)) if x != y), min(len(a), len(b)))
    return False, f"不一致 @{i}: 元「{a[max(0,i-25):i+25]}」→ 生成「{b[max(0,i-25):i+25]}」"


def main():
    if len(sys.argv) < 2:
        print("input .md を指定してください")
        sys.exit(1)
    ng = 0
    for arg in sys.argv[1:]:
        ok, msg = verify(arg)
        print(f"{'OK  ' if ok else 'NG  '}{pathlib.Path(arg).name}: {msg}")
        ng += 0 if ok else 1
    sys.exit(1 if ng else 0)


if __name__ == "__main__":
    main()
