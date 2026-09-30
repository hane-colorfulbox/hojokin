"""配布プラグインのゲート（1つでも NG ならビルドを止める検査の集まり）。

どの関数も「プラグイン内の相対パス → 中身(bytes)」の辞書を受け取り、問題の説明文のリストを返す
（空なら合格）。build_plugin.py から呼ぶ。検査対象は組み立て後の配布物そのもの。
"""
import posixpath
import re
import unicodedata
import urllib.parse
import zipfile
import io

import yaml

import distribution_manifest as dm

LINK_RE = re.compile(r'(?<!!)\[[^\]]*\]\(([^)\s]+)(?:\s+"[^"]*")?\)')
FENCE_RE = re.compile(r'^\s*(```|~~~)')
BLOCK_RE = r'<!-- {name}:begin -->.*?<!-- {name}:end -->'
SKILL_PREFIX = 'skills/'
XLSX_META_PARTS = ('docProps/core.xml', 'docProps/app.xml')
XLSX_ALLOWED_AUTHORS = {'', 'openpyxl'}
XLSX_FORBIDDEN_PARTS = re.compile(r'^xl/(comments\d*\.xml|threadedComments/|persons/)')


def _text(b: bytes) -> str:
    return b.decode('utf-8')


def skill_of(path: str) -> str | None:
    if not path.startswith(SKILL_PREFIX):
        return None
    return path[len(SKILL_PREFIX):].split('/', 1)[0]


def parse_frontmatter(text: str) -> tuple[dict, str]:
    if not text.startswith('---\n'):
        raise ValueError('frontmatter（先頭の --- ブロック）がありません')
    end = text.find('\n---\n', 4)
    if end < 0:
        raise ValueError('frontmatter の終わり（---）がありません')
    data = yaml.safe_load(text[4:end]) or {}
    if not isinstance(data, dict):
        raise ValueError('frontmatter が key: value の形になっていません')
    return data, text[end + 5:]


def check_frontmatter(files: dict[str, bytes]) -> list[str]:
    problems = []
    for path in sorted(p for p in files if re.fullmatch(r'skills/[^/]+/SKILL\.md', p)):
        folder = skill_of(path)
        try:
            fm, _ = parse_frontmatter(_text(files[path]))
        except (ValueError, yaml.YAMLError) as e:
            problems.append(f'{path}: {e}')
            continue
        extra = set(fm) - dm.ALLOWED_FRONTMATTER_KEYS
        if extra:
            problems.append(f'{path}: claude.ai で使えないキー {sorted(extra)}（使えるのは6キーだけ）')
        name, desc = fm.get('name'), fm.get('description')
        if not isinstance(name, str) or not re.match(dm.SKILL_NAME_PATTERN, name) or \
                len(name) > dm.SKILL_NAME_MAX:
            problems.append(f'{path}: name が規則（小文字英数とハイフン・64字以内）に合いません: {name!r}')
        elif name != folder:
            problems.append(f'{path}: name（{name}）とフォルダ名（{folder}）が違います')
        if not isinstance(desc, str) or not desc.strip():
            problems.append(f'{path}: description がありません')
        elif len(desc) > dm.DESCRIPTION_MAX or '<' in desc or '>' in desc:
            problems.append(f'{path}: description は {dm.DESCRIPTION_MAX} 字以内で < > を含めない'
                            f'（現在 {len(desc)} 字）')
        tools = fm.get('allowed-tools')
        if tools is not None and not isinstance(tools, str):
            problems.append(f'{path}: allowed-tools は1行の文字列で書く（リスト形式は使わない）')
    skills = {skill_of(p) for p in files if skill_of(p)}
    for s in sorted(skills):
        if f'skills/{s}/SKILL.md' not in files:
            problems.append(f'skills/{s}/: SKILL.md がありません')
    return problems


def check_layout(files: dict[str, bytes], plugin_root: str) -> list[str]:
    problems = []
    rel = [p[len(plugin_root):] for p in files if p.startswith(plugin_root)]
    if any(p.split('/', 1)[0] == 'bin' for p in rel):
        problems.append('プラグイン最上位に bin/ があります（claude.ai のアップロードで拒否される）')
    for p in sorted(rel):
        if unicodedata.normalize('NFC', p) != p:
            problems.append(f'{p}: パスが NFC 正規化されていません（Linux のチャット環境で名前が合わなくなる）')
        if not dm.path_chars_ok(p):
            bad = ''.join(sorted({ch for ch in p if not dm.path_chars_ok(ch)}))
            problems.append(f'{p}: claude.ai が受け付けない文字がパスにあります（{bad}）。'
                            'distribution_manifest.PATH_CHAR_REPLACEMENTS に置き換えを足す')
    if len(rel) > dm.MAX_FILES:
        problems.append(f'ファイル数 {len(rel)} が上限 {dm.MAX_FILES} を超えています')
    total = sum(len(files[plugin_root + p]) for p in rel)
    if total > dm.MAX_TOTAL_BYTES:
        problems.append(f'合計サイズ {total} バイトが上限 {dm.MAX_TOTAL_BYTES} を超えています')
    return problems


def _md_links(text: str):
    in_fence = False
    for no, line in enumerate(text.splitlines(), 1):
        if FENCE_RE.match(line):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        line = re.sub(r'`[^`]*`', '', line)
        for m in LINK_RE.finditer(line):
            yield no, m.group(1)


def check_links(files: dict[str, bytes]) -> list[str]:
    """相対リンクが「同じスキルのフォルダ内」で解決するか。スキルをまたぐリンクも不合格。"""
    problems = []
    names = set(files)   # 正規化せずに比べる（Linux のチャット環境では NFC/NFD の違いでリンクが切れる）
    for path in sorted(p for p in files if p.endswith('.md') and skill_of(p)):
        skill_root = f'skills/{skill_of(path)}/'
        for no, target in _md_links(_text(files[path])):
            if re.match(r'^[a-z][a-z0-9+.-]*:', target, re.I) or target.startswith('#'):
                continue
            target_path = urllib.parse.unquote(target.split('#', 1)[0])
            if not target_path:
                continue
            resolved = posixpath.normpath(posixpath.join(posixpath.dirname(path), target_path))
            if not resolved.startswith(skill_root):
                problems.append(f'{path}:{no}: スキルの外を指すリンク → {target}')
            elif resolved not in names and not any(n.startswith(resolved + '/') for n in names):
                problems.append(f'{path}:{no}: リンク先がありません → {target}')
    return problems


def _strip_blocks(text: str) -> str:
    for name in ('legacy-pack-check', 'legacy-migration'):
        text = re.sub(BLOCK_RE.format(name=name), '', text, flags=re.S)
    return text


def check_stale_terms(files: dict[str, bytes]) -> list[str]:
    """ZIP 配布時代の手順語が残っていないか（旧パックの確認ブロックと開発リポジトリの注記は例外）。"""
    problems = []
    for path in sorted(p for p in files if p.endswith(('.md', '.py')) and skill_of(p)):
        text = _strip_blocks(_text(files[path]))
        for no, line in enumerate(text.splitlines(), 1):
            if '開発リポジトリ' in line:
                continue
            for term in dm.STALE_TERMS:
                if term in line:
                    problems.append(f'{path}:{no}: 古い語「{term}」が残っています')
    return problems


def extract_block(text: str, name: str) -> str | None:
    m = re.search(BLOCK_RE.format(name=name), text, re.S)
    return m.group(0) if m else None


def check_shared_blocks(files: dict[str, bytes], canon_path: str,
                        required: dict[str, tuple[str, ...]]) -> list[str]:
    """共通の文言ブロックが、正本（ガイドの SKILL.md）と一字一句同じか。"""
    problems = []
    canon = _text(files[canon_path])
    for name in dm.SHARED_BLOCKS:
        if extract_block(canon, name) is None:
            problems.append(f'{canon_path}: 正本に共通ブロック「{name}」がありません')
    for path, names in required.items():
        if path not in files:
            problems.append(f'{path}: ファイルがありません（共通ブロックの確認先）')
            continue
        text = _text(files[path])
        for name in names:
            got, want = extract_block(text, name), extract_block(canon, name)
            if got is None:
                problems.append(f'{path}: 共通ブロック「{name}」がありません')
            elif want is not None and got != want:
                problems.append(f'{path}: 共通ブロック「{name}」が正本（{canon_path}）と違います')
    return problems


def check_xlsx_metadata(files: dict[str, bytes]) -> list[str]:
    problems = []
    for path in sorted(p for p in files if p.lower().endswith('.xlsx')):
        try:
            zf = zipfile.ZipFile(io.BytesIO(files[path]))
        except zipfile.BadZipFile:
            problems.append(f'{path}: xlsx として開けません')
            continue
        with zf:
            names = zf.namelist()
            bad = [n for n in names if XLSX_FORBIDDEN_PARTS.match(n)]
            if bad:
                problems.append(f'{path}: コメント・作成者情報のパーツがあります {bad}')
            for part in XLSX_META_PARTS:
                if part not in names:
                    continue
                xml = zf.read(part).decode('utf-8', 'replace')
                for tag in ('dc:creator', 'cp:lastModifiedBy'):
                    for v in re.findall(rf'<{tag}(?:\s[^>]*)?>(.*?)</{tag}>', xml, re.S):
                        if v.strip() not in XLSX_ALLOWED_AUTHORS:
                            problems.append(f'{path}: {part} の {tag} に人名などが入っています')
                for tag in ('Company', 'Manager'):
                    if re.search(rf'<{tag}>\s*[^<\s][^<]*</{tag}>', xml):
                        problems.append(f'{path}: {part} の {tag} が空ではありません')
    return problems
