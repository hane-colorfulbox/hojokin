"""第7次の担当者向け配布物（PDF・案内HTML・README.txt・ZIP）を正本Markdownから生成する。

既定ではGitにコミット済みの正本（--source-ref、既定HEAD）から読む。未コミットの正本で見た目を
確かめるだけなら --from-worktree を付ける（配布用には使わない）。既存の出力は上書きしない。通信しない。

    python -B scripts/build_official_guide.py --edition-date 2026-11-05 --out-dir <未使用の出力先>
    python -B scripts/build_official_guide.py --from-worktree --edition-date 2026-11-05 --out-dir <確認用>
"""
from datetime import date
from pathlib import Path
from xml.sax.saxutils import escape
import argparse
import hashlib
import html
import json
import posixpath
import re
import subprocess
import sys
import zipfile

import markdown
from bs4 import BeautifulSoup, Comment, NavigableString
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

sys.path.insert(0, str(Path(__file__).resolve().parent))   # テストがパス指定で読み込む場合も同じフォルダの補正を使う
import pdf_kinsoku  # noqa: E402

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = 'docs'
YOTEN = '第7次申請_担当者向け要点.md'
SCOPE = '第7次申請_公式資料で判断できることと確認の範囲.md'
STAFF_GUIDES = (YOTEN, '第7次申請_変更点と使い方.md', '第7次申請_確認待ち事項.md', SCOPE,
                '第7次申請_ツール入力手順.md')
ADMIN_GUIDES = ('第7次原本_改訂内容と再取込手順.md',)
PDF_GUIDES = (YOTEN, SCOPE)
HTML_DIR = '案内HTML'
ADMIN_DIR = HTML_DIR + '/管理者・開発者向け'
PDF_TITLE = '第7次申請 担当者向け要点と公式資料の確認範囲'
PDF_NAME = '第7次申請_担当者向け要点と公式資料の確認範囲_{date}.pdf'
ZIP_NAME = '第7次申請_配布資料_{date}.zip'
README_NAME = 'README.txt'
SUMMARY_NAME = 'guide_artifact_summary.json'
DEFAULT_LABEL = '第7次対応版'
DEFAULT_TOOL_MIN = 'v0.2.99'
DEFAULT_PLUGIN_MIN = '2026.11.05'
PDF_PAGE_RANGE = (4, 12)
FIRST_PAGE_MARK = '第7次申請：担当者向け要点'
PDF_REQUIRED_TERMS = ('担当者向け要点', '17:00', '55円', '令和8年度改定', '常時使用する従業員',
                      '翌事業年度', '見込み値', '保留', '30円', '50円')
TEXT_NEWLINE = '\n'
README_NEWLINE = '\r\n'
BLOCK_TAGS = frozenset({'h1', 'h2', 'h3', 'p', 'ul', 'ol', 'table'})
INLINE_TAGS = frozenset({'strong', 'em', 'code', 'a', 'br'})
INK, LINK, LINE, HEAD_BG, NOTE_BG, NOTE_LINE = '#183146', '#145bb4', '#d9e2ea', '#edf3f9', '#fff4d9', '#ad761a'
TABLE_RATIOS = {2: [.25, .75], 3: [.22, .44, .34], 4: [.25, .25, .25, .25]}
MARGIN, TOP_MARGIN, BOTTOM_MARGIN, CELL_PADDING = 44, 42, 49, 4
CSS = ('*{box-sizing:border-box}body{margin:0;background:#f3f6fa;color:#183146;font:15px/1.85 "Yu Gothic UI",Meiryo,sans-serif}'
       'main{max-width:1120px;margin:auto;padding:30px 26px 70px}header{background:#163c59;color:#fff;padding:26px 30px;border-radius:12px}'
       'header p{color:#dae8f4;margin-bottom:0}h1{font-size:26px;line-height:1.5;margin:0 0 12px}h2{font-size:21px;margin:30px 0 14px}h3{font-size:17px}'
       'article{background:#fff;padding:26px 30px;border:1px solid #d9e2ea;border-radius:10px;margin-top:22px}a{color:#145bb4;text-underline-offset:3px}'
       'li{margin:7px 0}.scroll{overflow:auto}table{border-collapse:collapse;width:100%;font-size:14px}th,td{text-align:left;vertical-align:top;padding:11px;border:1px solid #d9e2ea}th{background:#edf3f9}'
       'pre{padding:15px;background:#f1f5f9;overflow:auto}code{font-size:13px;overflow-wrap:anywhere}'
       '@media(max-width:650px){main{padding:16px 10px}header,article{padding:20px 15px}h1{font-size:23px}table{font-size:13px}th,td{padding:8px}}')


class GuideBuildError(Exception):
    """配布物を作れない（正本の不備・未対応の書式・検査不合格）。"""


def jp_date(d: date) -> str:
    return f'{d.year}年{d.month}月{d.day}日'


def git(*args) -> str:
    return subprocess.check_output(['git', *args], cwd=ROOT).decode('utf-8')


def load_sources(source_ref, from_worktree):
    """(正本名→本文, 生成元の記録) を返す。改行はLFにそろえる。"""
    names = STAFF_GUIDES + ADMIN_GUIDES
    head = git('rev-parse', '--verify', '--end-of-options', (source_ref or 'HEAD') + '^{commit}').strip()
    if from_worktree:
        texts = {n: (ROOT / DOCS_DIR / n).read_text(encoding='utf-8') for n in names}
        dirty = bool(git('status', '--porcelain', '--', *[f'{DOCS_DIR}/{n}' for n in names]).strip())
        meta = {'mode': 'worktree', 'base_commit': head, 'uncommitted_changes': dirty}
    else:
        try:
            texts = {n: git('show', f'{head}:{DOCS_DIR}/{n}') for n in names}
        except subprocess.CalledProcessError as e:
            raise GuideBuildError(f'コミット {head[:12]} に正本がそろっていません。正本をコミットしてから生成してください') from e
        meta = {'mode': 'git', 'commit': head, 'commit_date': git('show', '-s', '--format=%cI', head).strip()}
    return {n: t.replace('\r\n', '\n') for n, t in texts.items()}, meta


def split_title(name, text):
    first, _, rest = text.partition('\n')
    if not first.startswith('# '):
        raise GuideBuildError(f'{name}: 1行目が「# 題名」ではありません')
    return first[2:].strip(), rest


def html_path(name):
    folder = ADMIN_DIR if name in ADMIN_GUIDES else HTML_DIR
    return f'{folder}/{name[:-3]}.html'


def render_html(name, text, edition):
    title, body_md = split_title(name, text)
    body = markdown.markdown(body_md, extensions=['tables', 'fenced_code'])
    here = posixpath.dirname(html_path(name))
    for other in STAFF_GUIDES + ADMIN_GUIDES:
        target = posixpath.relpath(html_path(other), here)
        body = body.replace(f'href="{other}"', f'href="{html.escape(target)}"')
    if re.search(r'href="(?!https?://)[^"]*\.md', body):
        raise GuideBuildError(f'{name}: 配布物の外を指す相対リンクがあります')
    body = body.replace('<table>', '<div class="scroll"><table>').replace('</table>', '</table></div>')
    return ('<!doctype html><html lang="ja"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1">'
            f'<title>{html.escape(title)}</title><style>{CSS}</style></head><body><main>'
            f'<header><h1>{html.escape(title)}</h1><p>{edition["jp"]}版　{html.escape(edition["label"])}　社内配布用</p></header>'
            f'<article>{body}</article></main></body></html>\n')


def make_styles(font, bold):
    ink = colors.HexColor(INK)

    def style(name, size, leading, face=font, **kw):
        return ParagraphStyle(name, fontName=face, fontSize=size, leading=leading, textColor=ink, wordWrap='CJK', **kw)
    return {'title': style('title', 21, 30, bold, spaceAfter=10),
            'part': style('part', 16, 24, bold, spaceAfter=12, keepWithNext=True),
            'h2': style('h2', 13.5, 20, bold, spaceBefore=8, spaceAfter=10, keepWithNext=True),
            'h3': style('h3', 11, 17, bold, spaceBefore=8, spaceAfter=6, keepWithNext=True),
            'body': style('body', 9.2, 14.8, spaceAfter=6),
            'li': style('li', 9.2, 14.2, spaceAfter=2, leftIndent=10, firstLineIndent=-10),
            'small': style('small', 8, 12, spaceAfter=10),
            'cell': style('cell', 8.1, 12.1)}


def render_inline(node, where):
    """許可したインライン要素だけを ReportLab の書式に変える。それ以外は止める。"""
    out = []
    for child in node.children:
        if isinstance(child, Comment):
            continue
        if isinstance(child, NavigableString):
            out.append(escape(str(child)))
            continue
        if child.name not in INLINE_TAGS:
            raise GuideBuildError(f'{where}: PDFに未対応の要素 <{child.name}> があります')
        inner = render_inline(child, where)
        href = child.get('href', '') if child.name == 'a' else ''
        if child.name == 'strong':
            out.append(f'<b>{inner}</b>')
        elif child.name == 'em':
            out.append(f'<i>{inner}</i>')
        elif child.name == 'br':
            out.append('<br/>')
        elif re.match(r'https?://', href):
            out.append(f'<a href="{escape(href, {chr(34): "&quot;"})}" color="{LINK}"><u>{inner}</u></a>')
        else:
            out.append(inner)   # code と相対リンクは文字だけ残す（置き場所はPDF末尾に書く）
    return ''.join(out)


def render_list_item(item, where):
    blocks = [c for c in item.children if getattr(c, 'name', None) and c.name not in INLINE_TAGS]
    if any(b.name != 'p' for b in blocks):
        raise GuideBuildError(f'{where}: リストの中の入れ子のリスト・表・コードはPDFに未対応です')
    return '<br/>'.join(render_inline(b, where) for b in blocks) if blocks else render_inline(item, where)


def make_table(node, styles, where):
    rows = [[Paragraph(render_inline(c, where), styles['cell']) for c in tr.find_all(['th', 'td'], recursive=False)]
            for tr in node.find_all('tr')]
    width = A4[0] - 2 * MARGIN
    ratios = TABLE_RATIOS.get(len(rows[0]), [1 / len(rows[0])] * len(rows[0]))
    table = Table(rows, colWidths=[width * r for r in ratios], repeatRows=1, hAlign='LEFT')
    table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(HEAD_BG)), ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                               ('GRID', (0, 0), (-1, -1), .45, colors.HexColor(LINE))]
                              + [(k, (0, 0), (-1, -1), CELL_PADDING) for k in ('LEFTPADDING', 'RIGHTPADDING', 'TOPPADDING', 'BOTTOMPADDING')]))
    return [table, Spacer(1, 8)]


def markdown_to_story(name, text, styles):
    """正本1本をPDFの部品にする。未対応の要素は黙って落とさず GuideBuildError で止める。"""
    soup = BeautifulSoup(markdown.markdown(text, extensions=['tables', 'fenced_code']), 'html.parser')
    story = []
    for node in soup.children:
        if isinstance(node, Comment) or (isinstance(node, NavigableString) and not node.strip()):
            continue
        tag = getattr(node, 'name', None)
        where = f'{name} の <{tag}>'
        if tag not in BLOCK_TAGS:
            raise GuideBuildError(f'{name}: PDFに未対応の要素 <{tag or "文字列"}> があります')
        if tag in ('h1', 'h2', 'h3', 'p'):
            story.append(Paragraph(render_inline(node, where), styles[{'h1': 'part', 'p': 'body'}.get(tag, tag)]))
        elif tag in ('ul', 'ol'):
            for no, item in enumerate(node.find_all('li', recursive=False), 1):
                mark = f'{no}. ' if tag == 'ol' else '・ '
                story.append(Paragraph(mark + render_list_item(item, where), styles['li']))
        else:
            story.extend(make_table(node, styles, where))
    return story


def location_story(styles, pdf_name):
    intro = f'このPDF（{pdf_name}）は、配布ZIPの直下にあります。HTML版の資料はZIP内の「{HTML_DIR}」フォルダにあります。'
    items = [html_path(n) for n in STAFF_GUIDES]
    items += [f'{html_path(n)}（管理者・開発者向け。申請担当者は読まなくてかまいません）' for n in ADMIN_GUIDES]
    outro = f'読む順番と版の見分け方は、ZIP直下の {README_NAME} にあります。'
    return ([Paragraph('関連資料の置き場所', styles['h2']), Paragraph(escape(intro), styles['body'])]
            + [Paragraph('・ ' + escape(s), styles['li']) for s in items] + [Paragraph(escape(outro), styles['body'])])


def build_pdf(path, texts, fonts, edition, status_note):
    pdf_kinsoku.apply()   # 「）。」のように行頭禁則が続くとき「。」が行頭に来ないようにする
    pdfmetrics.registerFont(TTFont('GuideJP', str(fonts[0]), subfontIndex=0))
    bold = 'GuideJP'
    if fonts[1].is_file():
        pdfmetrics.registerFont(TTFont('GuideJP-Bold', str(fonts[1]), subfontIndex=0))
        bold = 'GuideJP-Bold'
    else:
        print(f'⚠ 太字フォントが見つからないため、見出しは通常の太さになります: {fonts[1]}', file=sys.stderr)
    pdfmetrics.registerFontFamily('GuideJP', normal='GuideJP', bold=bold, italic='GuideJP', boldItalic=bold)
    styles = make_styles('GuideJP', bold)
    story = [Paragraph(escape(PDF_TITLE).replace(' ', '<br/>', 1), styles['title']),
             Paragraph(escape(f'{edition["jp"]}版 / {edition["label"]} / 社内配布用'), styles['small'])]
    if status_note:
        box = Table([[Paragraph(escape(f'{edition["jp"]}時点の状況：{status_note}'), styles['body'])]],
                    colWidths=[A4[0] - 2 * MARGIN])
        box.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(NOTE_BG)),
                                 ('BOX', (0, 0), (-1, -1), .8, colors.HexColor(NOTE_LINE))]))
        story += [box, Spacer(1, 12)]
    for i, name in enumerate(PDF_GUIDES):
        if i:
            story.append(PageBreak())
        story += markdown_to_story(name, texts[name], styles)
    story += location_story(styles, path.name)

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor(LINE))
        canvas.line(MARGIN, 36, A4[0] - MARGIN, 36)
        canvas.setFont('GuideJP', 7.5)
        canvas.setFillColor(colors.HexColor(INK))
        canvas.drawString(MARGIN, 23, f'社内配布用 / {PDF_TITLE} / {edition["jp"]}版')
        canvas.drawRightString(A4[0] - MARGIN, 23, str(doc.page))
        canvas.restoreState()
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=MARGIN, rightMargin=MARGIN, topMargin=TOP_MARGIN,
                            bottomMargin=BOTTOM_MARGIN, title=PDF_TITLE, author='カラフルボックス', pageCompression=1,
                            invariant=1)
    doc.build(story, onFirstPage=footer, onLaterPages=footer)


def verify_pdf(path):
    reader = PdfReader(path)
    pages = len(reader.pages)
    if not PDF_PAGE_RANGE[0] <= pages <= PDF_PAGE_RANGE[1]:
        raise GuideBuildError(f'PDFのページ数が想定外です: {pages}（想定 {PDF_PAGE_RANGE[0]}〜{PDF_PAGE_RANGE[1]}）')
    texts = [re.sub(r'\s+', '', page.extract_text() or '') for page in reader.pages]
    missing = [w for w in PDF_REQUIRED_TERMS if re.sub(r'\s+', '', w) not in ''.join(texts)]
    first_page_ok = FIRST_PAGE_MARK in texts[0]
    if missing or not first_page_ok:
        raise GuideBuildError(f'PDFの本文検査に不合格: 見つからない語 {missing}、1ページ目に要点の見出しがある={first_page_ok}')
    if not all(page.get('/Resources', {}).get('/Font') for page in reader.pages):
        raise GuideBuildError('フォントが埋め込まれていないページがあります')
    return pages


def readme_text(edition, pdf_name, versions, status_note):
    staff = [html_path(n) for n in STAFF_GUIDES]
    lines = [f'第7次申請 配布資料（{edition["jp"]}版・{edition["label"]}）', '',
             '対象読者：申請担当者（申請書づくり・加点判定を担当する社員）。', '',
             '■ 読む順番',
             f'1. {pdf_name}',
             '   前半が「担当者向け要点」、後半が「公式資料で判断できることと確認の範囲」です。まず前半を読みます。',
             f'2. {staff[1]}', f'3. {staff[2]}',
             f'4. {staff[4]}（前半の「画面で作成する場合」だけ。後半のローカルCLIは管理者・開発者向け）',
             f'※ {staff[0]} と {staff[3]} は、PDFと同じ内容のHTML版です。', '',
             '■ 管理者・開発者向け（申請担当者は読まなくてかまいません）']
    lines += [f'- {html_path(n)}' for n in ADMIN_GUIDES]
    lines += ['- ツール入力手順の後半「ローカルCLIで作成する場合」', '',
              '■ ツールとプラグインの版の見分け方',
              f'- 書類作成ツール（Streamlit）：画面のいちばん下のフッターが {versions[0]} 以上なら第7次対応版です。',
              f'- 組織プラグイン hojokin：hojokin-guide に「プラグインの版を確認して」と頼み、版の日付が {versions[1]} 以降なら第7次対応版です。',
              '- それより前の版では、第7次の申請書・加点判定を作りません。管理者に連絡し、第7次対応版の反映を待ちます。']
    if status_note:
        lines += ['', f'■ {edition["jp"]}時点の状況', status_note]
    return README_NEWLINE.join(lines) + README_NEWLINE


def write_zip(zip_path, out_dir, members, edition_date):
    stamp = (edition_date.year, edition_date.month, edition_date.day, 0, 0, 0)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for rel in sorted(members):
            info = zipfile.ZipInfo(rel, stamp)
            info.compress_type, info.external_attr = zipfile.ZIP_DEFLATED, 0o644 << 16
            zf.writestr(info, (out_dir / rel).read_bytes())


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def parse_args(argv):
    p = argparse.ArgumentParser(description='第7次の担当者向け配布物（PDF・HTML・README・ZIP）を生成する（通信なし）')
    p.add_argument('--out-dir', required=True, type=Path, help='未使用（存在しないか空）の出力先')
    p.add_argument('--edition-date', required=True, type=date.fromisoformat, help='版の日付（YYYY-MM-DD）')
    p.add_argument('--edition-label', default=DEFAULT_LABEL, help=f'版の名前（既定「{DEFAULT_LABEL}」）')
    p.add_argument('--status-note', default='', help='PDF1ページ目とREADMEに載せる、その日時点の状況（任意）')
    p.add_argument('--tool-min-version', default=DEFAULT_TOOL_MIN, help='第7次対応のツールの最小版（要点と一致が必要）')
    p.add_argument('--plugin-min-version', default=DEFAULT_PLUGIN_MIN, help='第7次対応のプラグインの最小版（要点と一致が必要）')
    src = p.add_mutually_exclusive_group()
    src.add_argument('--source-ref', help='生成元のGitコミット（既定HEAD）')
    src.add_argument('--from-worktree', action='store_true', help='（確認用）作業ツリーの正本から生成する')
    p.add_argument('--font', type=Path, default=Path('C:/Windows/Fonts/meiryo.ttc'), help='日本語TrueTypeフォント')
    p.add_argument('--font-bold', type=Path, default=Path('C:/Windows/Fonts/meiryob.ttc'), help='太字の日本語フォント')
    args = p.parse_args(argv)
    if args.out_dir.exists() and (not args.out_dir.is_dir() or any(args.out_dir.iterdir())):
        p.error('出力先が空ではありません。既存の配布物を上書きせず、未使用の出力先を指定してください')
    if not args.font.is_file():
        p.error('日本語フォントが見つかりません。--font で指定してください')
    return args


def main(argv=None):
    args = parse_args(argv)
    texts, source = load_sources(args.source_ref, args.from_worktree)
    versions = (args.tool_min_version, args.plugin_min_version)
    if not all(v in texts[YOTEN] for v in versions):
        raise GuideBuildError(f'担当者向け要点に版の基準 {versions} が書かれていません（READMEと食い違うため停止）')
    edition = {'jp': jp_date(args.edition_date), 'label': args.edition_label}
    out = args.out_dir.resolve()
    out.mkdir(parents=True, exist_ok=True)
    html_files = []
    for name in STAFF_GUIDES + ADMIN_GUIDES:
        target = out / html_path(name)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(render_html(name, texts[name], edition), encoding='utf-8', newline=TEXT_NEWLINE)
        html_files.append(html_path(name))
    pdf_name = PDF_NAME.format(date=args.edition_date.isoformat())
    build_pdf(out / pdf_name, texts, (args.font, args.font_bold), edition, args.status_note)
    pages = verify_pdf(out / pdf_name)
    (out / README_NAME).write_bytes(readme_text(edition, pdf_name, versions, args.status_note).encode('utf-8'))
    members = [README_NAME, pdf_name] + html_files
    zip_name = ZIP_NAME.format(date=args.edition_date.isoformat())
    write_zip(out / zip_name, out, members, args.edition_date)
    summary = {'edition_date': args.edition_date.isoformat(), 'edition_label': args.edition_label, 'source': source,
               'pdf': {'name': pdf_name, 'pages': pages}, 'zip': zip_name,
               'files_sha256': {m: sha256(out / m) for m in members + [zip_name]},
               'source_inputs_sha256': {f'{DOCS_DIR}/{n}': hashlib.sha256(t.encode('utf-8')).hexdigest() for n, t in texts.items()},
               'text_checks_passed': True}
    (out / SUMMARY_NAME).write_text(json.dumps(summary, ensure_ascii=False, indent=2) + '\n', encoding='utf-8', newline=TEXT_NEWLINE)
    print(json.dumps({'out_dir': str(out), 'pdf_pages': pages, 'zip': zip_name, 'html': len(html_files),
                      'source': source['mode']}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except GuideBuildError as e:
        print(f'❌ {e}', file=sys.stderr)
        raise SystemExit(1)
