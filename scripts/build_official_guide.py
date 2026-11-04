"""第7次配布ガイドの再生成。正本をGitコミットから読み、既存出力を上書きしない。"""
from pathlib import Path
import argparse
import html
import json
import re
import subprocess
from datetime import datetime
import sys
import textwrap

import markdown
from bs4 import BeautifulSoup
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                               TableStyle, PageBreak, KeepTogether)

sys.stdout.reconfigure(encoding='utf-8')

def main(argv=None):
    parser=argparse.ArgumentParser(description='コミット済みの第7次配布正本からPDF・HTMLを生成する（通信なし）')
    parser.add_argument('--source-ref',default='HEAD',help='生成元のGitコミット。既定HEAD')
    parser.add_argument('--out-dir',required=True,type=Path,help='未使用の出力ディレクトリ')
    parser.add_argument('--font',type=Path,default=Path('C:/Windows/Fonts/meiryo.ttc'),help='日本語TrueTypeフォント')
    args=parser.parse_args(argv)
    if args.out_dir.exists() and any(args.out_dir.iterdir()):
        parser.error('出力先が空ではありません。既存配布物を上書きせず、未使用の出力先を指定してください')
    if not args.font.is_file():
        parser.error('日本語フォントが見つかりません。--font で指定してください')
    ROOT=Path(__file__).resolve().parents[1]
    DIST=args.out_dir.resolve()
    SHA=subprocess.check_output(['git','rev-parse','--verify','--end-of-options',args.source_ref+'^{commit}'],cwd=ROOT).decode().strip()
    commit_date=subprocess.check_output(['git','show','-s','--format=%cI',SHA],cwd=ROOT).decode().strip()
    VERSION=datetime.fromisoformat(commit_date).strftime('%Y.%m.%d')+'-'+SHA[:12]
    def read_source(name):
        return subprocess.check_output(['git','show',SHA+':docs/'+name],cwd=ROOT).decode('utf-8').replace('\r\n','\n')
    GUIDES=['第7次申請_公式資料で判断できることと確認の範囲.md',
            '第7次申請_変更点と使い方.md','第7次申請_確認待ち事項.md',
            '第7次申請_ツール入力手順.md','第7次原本_改訂内容と再取込手順.md']
    MAIN=GUIDES[0]
    HTML_DIR=DIST/'第7次_配布案内_2026-11-04'
    HTML_DIR.mkdir(parents=True,exist_ok=True)
    CSS='''
    *{box-sizing:border-box}body{margin:0;background:#f3f6fa;color:#183146;font:15px/1.85 "Yu Gothic UI",Meiryo,sans-serif}
    main{max-width:1120px;margin:auto;padding:30px 26px 70px}header{background:#163c59;color:#fff;padding:26px 30px;border-radius:12px}
    header p{color:#dae8f4;margin-bottom:0}h1{font-size:26px;line-height:1.5;margin:0 0 12px}h2{font-size:21px;margin:30px 0 14px}h3{font-size:17px}
    article{background:white;padding:26px 30px;border:1px solid #d9e2ea;border-radius:10px;margin-top:22px}a{color:#145bb4;text-underline-offset:3px}
    li{margin:7px 0}.scroll{overflow:auto}table{border-collapse:collapse;width:100%;font-size:14px}th,td{text-align:left;vertical-align:top;padding:11px;border:1px solid #d9e2ea}th{background:#edf3f9}
    pre{padding:15px;background:#f1f5f9;overflow:auto}code{font-size:13px;overflow-wrap:anywhere}.note{padding:14px 18px;background:#fff4d9;border-left:4px solid #ad761a}
    @media(max-width:650px){main{padding:16px 10px}header,article{padding:20px 15px}h1{font-size:23px}table{font-size:13px}th,td{padding:8px}}
    '''
    CSS=textwrap.dedent(CSS)
    for name in GUIDES:
        text=read_source(name)
        title=text.splitlines()[0].lstrip('# ')
        body=markdown.markdown(text,extensions=['tables','fenced_code'])
        for other in GUIDES:
            body=body.replace('href="'+other+'"','href="'+other[:-3]+'.html"')
        body=body.replace('<table>','<div class="scroll"><table>').replace('</table>','</table></div>')
        doc='<!doctype html><html lang="ja"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
        doc+=f'<title>{html.escape(title)}</title><style>{CSS}</style></head><body><main>'
        doc+=f'<header><h1>{html.escape(title)}</h1><p>公式資料再照合・訂正版　{VERSION}　社内配布用</p></header>'
        doc+='<p class="note">制度の一般的な定義の再照会は不要です。案件の不足事実、社内調整、今回は保留する事項を区別します。</p>'
        doc+='<article>'+body+'</article></main></body></html>'
        (HTML_DIR/(name[:-3]+'.html')).write_text(doc,encoding='utf-8')

    pdfmetrics.registerFont(TTFont('GuideJP',str(args.font),subfontIndex=0))
    INK=colors.HexColor('#183146');BLUE=colors.HexColor('#145bb4');LINE=colors.HexColor('#d9e2ea')
    styles={
     'body':ParagraphStyle('body',fontName='GuideJP',fontSize=9.4,leading=15.4,textColor=INK,spaceAfter=8,wordWrap='CJK'),
     'small':ParagraphStyle('small',fontName='GuideJP',fontSize=8,leading=12,textColor=INK,spaceAfter=6,wordWrap='CJK'),
     'cell':ParagraphStyle('cell',fontName='GuideJP',fontSize=8.2,leading=12.5,textColor=INK,wordWrap='CJK'),
     'h1':ParagraphStyle('h1',fontName='GuideJP',fontSize=21,leading=30,textColor=INK,spaceAfter=16,wordWrap='CJK'),
     'h2':ParagraphStyle('h2',fontName='GuideJP',fontSize=14,leading=21,textColor=INK,spaceBefore=8,spaceAfter=12,keepWithNext=True,wordWrap='CJK'),
     'h3':ParagraphStyle('h3',fontName='GuideJP',fontSize=11,leading=17,textColor=INK,spaceBefore=10,spaceAfter=8,keepWithNext=True,wordWrap='CJK'),
    }
    def inline(node):
        s=''.join(str(x) for x in node.contents)
        s=s.replace('<strong>','<b>').replace('</strong>','</b>').replace('<em>','<i>').replace('</em>','</i>')
        s=re.sub(r'<code>(.*?)</code>',r'\1',s,flags=re.S)
        s=re.sub(r'<a href="(?!https?://)[^"]*">(.*?)</a>',r'\1',s,flags=re.S)
        return s

    source=read_source(MAIN)
    soup=BeautifulSoup(markdown.markdown(source,extensions=['tables']), 'html.parser')
    story=[]
    for node in soup.find_all(recursive=False):
        tag=node.name
        if tag=='h1':
            story.append(Paragraph('第7次申請<br/>公式資料に基づく対応ガイド',styles['h1']))
            story.append(Paragraph(f'2026年11月4日 / 公式資料再照合・訂正版 / {VERSION}',styles['small']))
            continue
        if tag=='h2':
            if node.get_text().startswith(('2.','4.','5.','6.','7.')):story.append(PageBreak())
            story.append(Paragraph(inline(node),styles['h2']))
        elif tag=='h3':story.append(Paragraph(inline(node),styles['h3']))
        elif tag=='p':story.append(Paragraph(inline(node),styles['body']))
        elif tag in ('ul','ol'):
            for item in node.find_all('li',recursive=False):
                story.append(Paragraph('・ '+inline(item),styles['body']))
        elif tag=='table':
            rows=[]
            for row in node.find_all('tr'):
                rows.append([Paragraph(inline(c),styles['cell']) for c in row.find_all(['td','th'],recursive=False)])
            count=len(rows[0]);width=A4[0]-88
            ratios={2:[.25,.75],3:[.26,.38,.36],4:[.25,.25,.25,.25]}.get(count,[1/count]*count)
            table=Table(rows,colWidths=[width*v for v in ratios],repeatRows=1,hAlign='LEFT')
            table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#edf3f9')),
                                      ('VALIGN',(0,0),(-1,-1),'TOP'),('GRID',(0,0),(-1,-1),.45,LINE),
                                      ('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),
                                      ('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7)]))
            story.extend([table,Spacer(1,10)])

    output=DIST/'第7次申請_公式資料に基づく対応ガイド_2026-11-04_訂正版.pdf'
    def footer(canvas,doc):
        canvas.saveState();canvas.setStrokeColor(LINE);canvas.line(44,36,A4[0]-44,36)
        canvas.setFont('GuideJP',7.5);canvas.setFillColor(INK)
        canvas.drawString(44,23,'社内配布用 / 公式資料再照合・訂正版 / 本番・Drive・全員配信は未実施')
        canvas.drawRightString(A4[0]-44,23,str(doc.page));canvas.restoreState()
    doc=SimpleDocTemplate(str(output),pagesize=A4,rightMargin=44,leftMargin=44,topMargin=42,bottomMargin=49,
                          title='第7次申請 公式資料に基づく対応ガイド（訂正版）',author='カラフルボックス',pageCompression=1)
    doc.build(story,onFirstPage=footer,onLaterPages=footer)
    reader=PdfReader(output)
    assert 5<=len(reader.pages)<=9,len(reader.pages)
    text='\n'.join(page.extract_text() or '' for page in reader.pages)
    for word in ['再照会','翌事業年度','30円','50円','計画入力','一般定義','保留']:
        assert word in text,word
    for page in reader.pages:
        resources=page.get('/Resources',{})
        assert resources.get('/Font'), 'font resource absent'
    (DIST/'guide_artifact_summary.json').write_text(json.dumps({'pdf':str(output),'pages':len(reader.pages),
     'source_commit':SHA,'version':VERSION,'html_count':len(GUIDES),'text_checks_passed':True},ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'pdf':str(output),'pages':len(reader.pages),'html_count':len(GUIDES)},ensure_ascii=False))


if __name__=='__main__':
    main()
