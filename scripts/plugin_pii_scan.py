"""配布プラグインの PII スキャナ（顧客名・人名・社員メール・Drive ID・金額）。

build_plugin.py のゲートと、プラグイン用リポジトリの pre-commit から呼ぶ。
検出した語そのものは画面に出さない（先頭1文字以外を伏せ字にする）。

使い方:
    python scripts/plugin_pii_scan.py <ディレクトリまたはファイル>...
    python scripts/plugin_pii_scan.py --git-range origin/main..HEAD   # 追加行だけを走査
    python scripts/plugin_pii_scan.py --staged                        # pre-commit 用

終了コード: 0=ヒットなし / 1=ヒットあり / 2=入力・設定エラー（名簿なし等）

許可（誤検知の承認）:
    scripts/plugin_pii_allowlist.json に「ルールID＋行」の sha256 だけを持つ。名前は保存しない。
    --approve を付けると、今回のヒットを理由付きで追記する（人が中身を確認した後に使う）。
"""
import argparse
import hashlib
import html
import json
import os
import re
import subprocess
import sys
import unicodedata
import zipfile
from dataclasses import dataclass
from pathlib import Path

from plugin_pii_rules import build_rules, exempt  # noqa: F401  （build_rules は呼び出し側の窓口）

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

HERE = Path(__file__).resolve().parent
ALLOWLIST_PATH = HERE / 'plugin_pii_allowlist.json'
ROSTER_ENV = 'HOJOKIN_PII_ROSTER'
ROSTER_DEFAULT = Path.home() / '.hojokin_pii' / 'roster.txt'
ROSTER_MIN_ENTRIES = 50
TEXT_SUFFIXES = {'.md', '.py', '.json', '.txt', '.toml', '.yml', '.yaml', '.csv', '.tsv', '.html'}
XLSX_SUFFIXES = {'.xlsx', '.xlsm'}
# xlsx の中で文字を持ちうるパーツはすべて見る（リンク先は _rels、図形の文字は drawings、表の名前は tables）
XLSX_XML_PART = re.compile(r'\.(xml|rels|vml)$', re.I)
XML_TAG = re.compile(r'<[^>]+>')
XML_ATTR_VALUE = re.compile(r'\s[\w:.\-]+="([^"]*)"')
SHARED_STRING_ITEM = re.compile(r'<si>(.*?)</si>', re.S)
# 図形の座標（EMU 単位の大きな数）やスタイルの識別子が並ぶパーツでは、金額と長い ID のルールを外す
LAYOUT_PART = re.compile(r'^xl/(drawings/|styles\.xml|theme/)')
LAYOUT_SKIP_RULES = frozenset({'amount', 'long-id'})
MAX_FILE_BYTES = 20 * 1024 * 1024

@dataclass
class Hit:
    path: str
    line_no: int
    rule: str
    matched: str
    line: str

    def key(self) -> str:
        return hashlib.sha256(f'{self.rule}\x00{_norm(self.line).strip()}'.encode('utf-8')).hexdigest()

    def masked(self) -> str:
        m = self.matched
        return m[:1] + '●' * max(1, len(m) - 1)


def _norm(s: str) -> str:
    return unicodedata.normalize('NFKC', s)


def _git_common_dir(cwd: Path) -> Path | None:
    try:
        out = subprocess.run(['git', 'rev-parse', '--git-common-dir'], cwd=cwd, capture_output=True,
                             text=True, encoding='utf-8', check=True).stdout.strip()
    except (OSError, subprocess.CalledProcessError):
        return None
    p = Path(out)
    return p if p.is_absolute() else (cwd / p).resolve()


def load_roster(path: Path | None = None) -> list[str]:
    path = path or Path(os.environ.get(ROSTER_ENV, ROSTER_DEFAULT))
    if not path.exists():
        raise FileNotFoundError(f'顧客名簿が見つかりません: {path}（{ROSTER_ENV} で場所を変えられます）')
    names = [_norm(l).strip() for l in path.read_text(encoding='utf-8').splitlines()]
    names = sorted({n for n in names if n and not n.startswith('#')}, key=len, reverse=True)
    if len(names) < ROSTER_MIN_ENTRIES:
        raise ValueError(f'顧客名簿が {len(names)} 語しかありません（下限 {ROSTER_MIN_ENTRIES}）: {path}')
    return names


def load_extra_patterns(repo: Path) -> list[tuple[str, re.Pattern]]:
    """開発リポジトリの pre-commit の PATTERNS と .gitleaks.toml の独自ルール（path allowlist は使わない）。"""
    pats: list[tuple[str, re.Pattern]] = []
    common = _git_common_dir(repo)
    hook = common / 'hooks' / 'pre-commit' if common else None
    if hook and hook.exists():
        m = re.search(r"PATTERNS='(.*?)'", hook.read_text(encoding='utf-8', errors='replace'))
        if m:
            pats.append(('hook-patterns', re.compile(m.group(1))))
    gl = repo / '.gitleaks.toml'
    if gl.exists():
        body = gl.read_text(encoding='utf-8', errors='replace')
        for rid, rx in re.findall(r'id\s*=\s*"([^"]+)".*?regex\s*=\s*\'\'\'(.*?)\'\'\'', body, re.S):
            try:
                pats.append((f'gitleaks-{rid}', re.compile(rx)))
            except re.error:
                continue
    return pats


def scan_text(label: str, text: str, rules, skip_rules=frozenset()) -> list[Hit]:
    hits: list[Hit] = []
    for no, raw in enumerate(text.splitlines(), 1):
        line = _norm(raw)
        for rule, rx in rules:
            if rule in skip_rules:
                continue
            for m in rx.finditer(line):
                if not exempt(rule, m, line):
                    hits.append(Hit(label, no, rule, m.group(0), raw))
    return hits


def _xml_lines(part: str, xml: str):
    """XML パーツから、検査する文字を1行ずつ取り出す（文字参照 &#...; は復号する）。"""
    for s in XML_TAG.sub('\n', xml).splitlines():          # 要素の文字（1要素1行）
        if s.strip():
            yield html.unescape(s)
    for tag in XML_TAG.findall(xml):                       # 属性値（リンク先 Target・表示名など）
        for value in XML_ATTR_VALUE.findall(tag):
            if value.strip():
                yield html.unescape(value)
    if part.endswith('sharedStrings.xml'):                 # 書式ごとに分かれた文字を1セル分につなぐ
        for item in SHARED_STRING_ITEM.findall(xml):
            yield html.unescape(XML_TAG.sub('', item))


def _read_text(path: Path) -> str | None:
    """UTF-8 で読めなければ Shift_JIS（給与ソフトの CSV に多い）で読み直す。どちらも無理なら None。"""
    data = path.read_bytes()
    for enc in ('utf-8', 'cp932'):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return None


def _xlsx_texts(path: Path):
    with zipfile.ZipFile(path) as zf:
        for info in zf.infolist():
            if XLSX_XML_PART.search(info.filename):
                xml = zf.read(info).decode('utf-8', errors='replace')
                yield info.filename, '\n'.join(_xml_lines(info.filename, xml))


def scan_path(path: Path, rules, base: Path | None = None) -> list[Hit]:
    label = path.relative_to(base).as_posix() if base else path.as_posix()
    if path.stat().st_size > MAX_FILE_BYTES:
        return [Hit(label, 0, 'too-large', path.name, f'{path.stat().st_size} bytes')]
    suffix = path.suffix.lower()
    if suffix in XLSX_SUFFIXES:
        hits: list[Hit] = []
        for part, text in _xlsx_texts(path):
            skip = LAYOUT_SKIP_RULES if LAYOUT_PART.match(part) else frozenset()
            hits.extend(scan_text(f'{label}!{part}', text, rules, skip))
        return hits
    if suffix in TEXT_SUFFIXES or not suffix or path.name.startswith('.'):
        text = _read_text(path)
        if text is None:
            return [Hit(label, 0, 'unscannable', path.name, '文字コードを判定できない（UTF-8 でも Shift_JIS でもない）')]
        return scan_text(label, text, rules)
    # 中身を検査できない形式（PDF・画像・Word など）は、通さずに止める（配る前に人が確かめる）
    return [Hit(label, 0, 'unscannable', path.name, f'検査できない形式: {suffix}')]


def scan_tree(targets: list[Path], rules) -> list[Hit]:
    hits: list[Hit] = []
    for t in targets:
        if t.is_dir():
            for f in sorted(t.rglob('*')):
                if f.is_file() and '.git' not in f.relative_to(t).parts and '__pycache__' not in f.parts:
                    hits.extend(scan_path(f, rules, t))
        elif t.is_file():
            hits.extend(scan_path(t, rules))
    return hits


def scan_git_diff(repo: Path, diff_args: list[str], rules) -> list[Hit]:
    """git diff の追加行だけを走査する（公開前の未 push コミット・pre-commit 用）。"""
    out = subprocess.run(['git', '-c', 'core.quotepath=false', 'diff', '--no-color', '-U0', *diff_args],
                         cwd=repo, capture_output=True, text=True, encoding='utf-8', errors='replace',
                         check=True).stdout
    hits: list[Hit] = []
    cur, line_no = '', 0
    for raw in out.splitlines():
        if raw.startswith('+++ '):
            cur = raw[6:] if raw.startswith('+++ b/') else raw[4:]
        elif raw.startswith('@@'):
            m = re.search(r'\+(\d+)', raw)
            line_no = int(m.group(1)) if m else 0
        elif raw.startswith('+') and not raw.startswith('+++'):
            hits.extend(Hit(cur, line_no, h.rule, h.matched, h.line) for h in scan_text(cur, raw[1:], rules))
            line_no += 1
    return hits


def load_allowlist() -> dict:
    if not ALLOWLIST_PATH.exists():
        return {}
    return json.loads(ALLOWLIST_PATH.read_text(encoding='utf-8'))


def filter_allowed(hits: list[Hit], allow: dict) -> list[Hit]:
    return [h for h in hits if h.key() not in allow]


def approve(hits: list[Hit], reason: str) -> int:
    allow = load_allowlist()
    added = 0
    for h in hits:
        if h.key() not in allow:
            allow[h.key()] = f'{h.rule}: {reason}'
            added += 1
    ALLOWLIST_PATH.write_text(json.dumps(dict(sorted(allow.items())), ensure_ascii=False, indent=1) + '\n',
                              encoding='utf-8')
    return added


def run(targets=None, git_range=None, staged=False, repo=None, roster_path=None):
    """ヒット一覧（許可済みを除く）を返す。設定エラーは例外で上げる。"""
    repo = repo or HERE.parent
    rules = build_rules(load_roster(roster_path), load_extra_patterns(repo))
    if staged:
        hits = scan_git_diff(Path.cwd(), ['--cached'], rules)
    elif git_range:
        hits = scan_git_diff(repo, [git_range], rules)
    else:
        hits = scan_tree([Path(t) for t in targets], rules)
    return filter_allowed(hits, load_allowlist())


def print_hits(hits: list[Hit], limit: int = 200) -> None:
    for h in hits[:limit]:
        print(f'  {h.path}:{h.line_no}  [{h.rule}]  {h.masked()}')
    if len(hits) > limit:
        print(f'  … ほか {len(hits) - limit} 件')


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description='配布プラグインの PII スキャン')
    ap.add_argument('targets', nargs='*', help='走査するディレクトリ／ファイル')
    ap.add_argument('--git-range', help='git diff の範囲（追加行のみ走査）例: origin/main..HEAD')
    ap.add_argument('--staged', action='store_true', help='ステージ済みの追加行を走査（pre-commit 用）')
    ap.add_argument('--roster', type=Path, default=None, help=f'顧客名簿（既定: {ROSTER_DEFAULT}）')
    ap.add_argument('--approve', metavar='理由', help='今回のヒットを許可リストへ追記する（確認済みのときだけ）')
    ap.add_argument('--only-rule', help='このルールのヒットだけを表示・許可する（例: amount）')
    args = ap.parse_args(argv)
    if not (args.targets or args.git_range or args.staged):
        ap.error('走査対象を指定してください')
    try:
        hits = run(args.targets, args.git_range, args.staged, roster_path=args.roster)
        if args.only_rule:
            hits = [h for h in hits if h.rule == args.only_rule]
    except (FileNotFoundError, ValueError, subprocess.CalledProcessError) as e:
        print(f'❌ PII スキャン設定エラー: {e}', file=sys.stderr)
        return 2
    if args.approve and hits:
        print(f'許可リストに {approve(hits, args.approve)} 件追加しました: {ALLOWLIST_PATH.name}')
        return 0
    if hits:
        print(f'❌ PII の疑い {len(hits)} 件（語は伏せ字）:')
        print_hits(hits)
        return 1
    print('✅ PII スキャン: ヒットなし')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
