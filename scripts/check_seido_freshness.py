"""公式サイトの資料更新を検知して、制度コンテキストのドリフトを知らせる。

デジタル化・AI導入補助金2026 の公式ダウンロードページから「資料名 → 更新日 → URL」を
取得し、docs/制度一次情報/sources.json に記録したスナップショットと突合する。
変わった資料と、その資料に依存している自社ドキュメントだけを出力する。

なぜ必要か:
    公募要領のような「制度の芯」は年1〜2回しか動かないが、セカンドオピニオンマニュアル・
    後年手続きマニュアル等の運用マニュアルは月1〜2回動く（2026年7月は5回）。
    芯だけ見ていると、今まさに回している業務の手順が古いまま残る。

依存:
    標準ライブラリのみ（引き継ぎパックに同梱しても追加インストール不要にするため）。
    Anthropic API は呼ばない＝課金ゼロ。

実行方法:
    python scripts/check_seido_freshness.py           変更を検知して表示（変更ありなら exit 1）
    python scripts/check_seido_freshness.py --update  検知したうえで sources.json を更新
    python scripts/check_seido_freshness.py --init    スナップショットを新規作成
    python scripts/check_seido_freshness.py --deep    PDF に HEAD を投げてサイズ・更新時刻も記録

回すタイミング:
    各申請回の頭（資料依頼を始めるとき） / 毎年10月の最低賃金改定時 / それ以外でも月1回程度。
"""
import argparse
import json
import re
import sys
import urllib.error
import urllib.request
from datetime import date
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path(__file__).resolve().parent.parent
SOURCES_PATH = ROOT / 'docs' / '制度一次情報' / 'sources.json'

BASE_URL = 'https://it-shien.smrj.go.jp'
DOWNLOAD_URL = f'{BASE_URL}/download/'
USER_AGENT = 'Mozilla/5.0 (compatible; hojokin-freshness-check/1.0)'
TIMEOUT_SEC = 30

# 公式ページ（Nuxt の SSR 出力）の資料ブロック。1件 = c-banner-pdf-item。
ITEM_SPLIT_RE = re.compile(r'class="c-banner-pdf-item"')
TITLE_RE = re.compile(r'<h4[^>]*>(.*?)</h4>', re.DOTALL)
UPDATED_RE = re.compile(r'更新日：([^<]+)')
HREF_RE = re.compile(r'href="([^"]+\.(?:pdf|xlsx|docx|xls|doc))"', re.IGNORECASE)
TAG_RE = re.compile(r'<[^>]+>')
JP_DATE_RE = re.compile(r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日')

# 出力の見出し
MARK_CHANGED = '🔴 更新あり（対応が要るもの）'
MARK_INFO = '🟡 更新あり（依存ドキュメントの登録なし）'
MARK_NEW = '🆕 公式ページに新しく現れた資料'
MARK_GONE = '⚠️ 公式ページから消えた資料'


def fetch(url: str) -> str:
    req = urllib.request.Request(url, headers={'User-Agent': USER_AGENT})
    with urllib.request.urlopen(req, timeout=TIMEOUT_SEC) as res:
        charset = res.headers.get_content_charset() or 'utf-8'
        return res.read().decode(charset, 'replace')


def normalize_date(raw: str) -> str:
    """「2026年5月15日」→「2026-05-15」。取れなければ原文を残す（「準備中」等）。"""
    m = JP_DATE_RE.search(raw)
    if not m:
        return raw.strip()
    y, mo, d = m.groups()
    return f'{int(y):04d}-{int(mo):02d}-{int(d):02d}'


def parse_download_page(html: str) -> dict[str, dict]:
    """URL をキーに {url: {title, updated}} を返す。URL が資料の安定した識別子。

    資料名（h4）は「公募要領」のように枠をまたいで重複するため、キーには使わない。
    """
    items: dict[str, dict] = {}
    for chunk in ITEM_SPLIT_RE.split(html)[1:]:
        # 次の資料ブロックまでを1件とみなす（split 済みなので chunk 全体が1件分）
        href_m = HREF_RE.search(chunk)
        if not href_m:
            continue
        url = href_m.group(1)
        if url.startswith('/'):
            url = BASE_URL + url
        title_m = TITLE_RE.search(chunk)
        title = TAG_RE.sub('', title_m.group(1)).strip() if title_m else '(名称不明)'
        updated_m = UPDATED_RE.search(chunk)
        updated = normalize_date(updated_m.group(1)) if updated_m else '(更新日なし)'
        # 同一 URL が複数箇所に出る場合は先勝ち
        items.setdefault(url, {'title': title, 'updated': updated})
    return items


def head_info(url: str) -> dict:
    """PDF に HEAD を投げて Last-Modified / Content-Length を取る（--deep 用）。

    公式ページの更新日を動かさずに中身だけ差し替えられた場合の保険。
    """
    req = urllib.request.Request(url, headers={'User-Agent': USER_AGENT}, method='HEAD')
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT_SEC) as res:
            return {
                'last_modified': res.headers.get('Last-Modified', ''),
                'content_length': res.headers.get('Content-Length', ''),
            }
    except (urllib.error.URLError, TimeoutError) as e:
        return {'last_modified': f'(取得失敗: {e})', 'content_length': ''}


def load_sources() -> dict:
    if not SOURCES_PATH.exists():
        return {}
    # utf-8-sig: Windows のエディタや PowerShell の Out-File が付ける BOM を許容する
    return json.loads(SOURCES_PATH.read_text(encoding='utf-8-sig'))


def save_sources(data: dict) -> None:
    SOURCES_PATH.parent.mkdir(parents=True, exist_ok=True)
    SOURCES_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8'
    )


def build_snapshot(live: dict[str, dict], previous: dict, deep: bool) -> dict:
    """既存の depends・label を引き継ぎつつ、公式の最新状態でスナップショットを作る。"""
    prev_docs = {d['url']: d for d in previous.get('documents', [])}
    docs = []
    for url, info in live.items():
        old = prev_docs.get(url, {})
        doc = {
            'url': url,
            'title': info['title'],
            'label': old.get('label') or info['title'],
            'official_updated': info['updated'],
            'depends': old.get('depends', []),
        }
        if deep:
            doc.update(head_info(url))
        elif 'last_modified' in old:
            doc['last_modified'] = old['last_modified']
            doc['content_length'] = old.get('content_length', '')
        docs.append(doc)
    docs.sort(key=lambda d: d['url'])
    return {
        '_readme': (
            '公式ダウンロードページの資料スナップショット。'
            'scripts/check_seido_freshness.py が突合に使う。'
            'depends には、その資料が変わったら読み直すべき自社ドキュメントを書く。'
        ),
        'download_page': DOWNLOAD_URL,
        'last_verified': date.today().isoformat(),
        'documents': docs,
    }


def report(live: dict[str, dict], previous: dict) -> tuple[list, list, list, list]:
    prev_docs = {d['url']: d for d in previous.get('documents', [])}
    changed, info_only, added, gone = [], [], [], []

    for url, cur in live.items():
        old = prev_docs.get(url)
        if old is None:
            added.append((url, cur))
        elif old.get('official_updated') != cur['updated']:
            entry = (url, old, cur)
            (changed if old.get('depends') else info_only).append(entry)

    for url, old in prev_docs.items():
        if url not in live:
            gone.append((url, old))

    return changed, info_only, added, gone


def print_report(changed, info_only, added, gone, previous) -> None:
    print(f'公式ダウンロードページ: {DOWNLOAD_URL}')
    print(f'前回の確認日: {previous.get("last_verified", "(記録なし)")}／本日: {date.today().isoformat()}')
    print()

    if changed:
        print(MARK_CHANGED)
        for url, old, cur in changed:
            print(f'  ● {old.get("label", cur["title"])}')
            print(f'      {old.get("official_updated")} → {cur["updated"]}')
            print(f'      {url}')
            print('      読み直す自社ドキュメント:')
            for dep in old['depends']:
                print(f'        - {dep}')
        print()

    if info_only:
        print(MARK_INFO)
        for url, old, cur in info_only:
            print(f'  ● {old.get("label", cur["title"])}: '
                  f'{old.get("official_updated")} → {cur["updated"]}  {url}')
        print()

    if added:
        print(MARK_NEW)
        for url, cur in added:
            print(f'  ● {cur["title"]}（更新日 {cur["updated"]}）  {url}')
        print('  → 依存する自社ドキュメントがあれば sources.json の depends に登録する。')
        print()

    if gone:
        print(MARK_GONE)
        for url, old in gone:
            print(f'  ● {old.get("label", url)}  {url}')
        print('  → 公募年度の切り替わりや資料の統廃合の可能性。URL を確認する。')
        print()

    if not (changed or info_only or added or gone):
        print('✅ 変更なし（追跡中の資料はすべて記録どおり）')


def main() -> int:
    parser = argparse.ArgumentParser(description='補助金 公式資料の更新を検知する')
    parser.add_argument('--update', action='store_true',
                        help='検知後に sources.json を最新状態で書き換える')
    parser.add_argument('--init', action='store_true',
                        help='スナップショットを新規作成する（初回のみ）')
    parser.add_argument('--deep', action='store_true',
                        help='各資料に HEAD を投げて Last-Modified / サイズも記録する')
    args = parser.parse_args()

    try:
        html = fetch(DOWNLOAD_URL)
    except (urllib.error.URLError, TimeoutError) as e:
        print(f'❌ 公式ページを取得できませんでした: {e}', file=sys.stderr)
        return 2

    live = parse_download_page(html)
    if not live:
        print('❌ 資料を1件も抽出できませんでした。公式ページの HTML 構造が変わった可能性が'
              'あります（parse_download_page の正規表現を見直してください）。', file=sys.stderr)
        return 2
    print(f'公式ページから {len(live)} 件の資料を検出しました。\n')

    previous = load_sources()

    if args.init or not previous:
        save_sources(build_snapshot(live, previous, args.deep))
        print(f'✅ スナップショットを作成しました: {SOURCES_PATH.relative_to(ROOT)}')
        print('   depends（資料が変わったら読み直す自社ドキュメント）を手で埋めてください。')
        return 0

    changed, info_only, added, gone = report(live, previous)
    print_report(changed, info_only, added, gone, previous)

    if args.update:
        save_sources(build_snapshot(live, previous, args.deep))
        print(f'📝 sources.json を更新しました: {SOURCES_PATH.relative_to(ROOT)}')
        return 0

    if changed or added or gone:
        print('（内容を反映したら --update で sources.json を更新してください）')
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
