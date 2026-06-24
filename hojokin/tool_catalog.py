# -*- coding: utf-8 -*-
"""導入ツールの実機能を「事業内容255字」生成プロンプトに注入するためのカタログ層。

補助金で導入する第三者製ITツール（Scale人事評価 / クラフトバンクオフィス / ClipLine /
ピスケスアポ / irohana / ブレイン 等。社内の自動化ツールとは別物）の実機能・効果を
docs/ツール情報/*.md から実行時に読み込み、ヒアリング/見積由来の tool_name（表記ゆれあり）
から該当ツールを引き当てて、その公式登録機能をプロンプトに差し込む。

設計の要点:
- マッチ用キー（別名）の単一の真実は本ファイルの TOOL_CATALOG。
  docs/ツール情報/INDEX.md「申請ツール名（ヒアリングB53）との照合用 別名」節は人間用
  ドキュメントであり、コードはパースしない。両者は手動同期する（別名を足したら両方直す）。
- 機能本文の真実は各 md の「## …事業内容255字用サマリ」節（無ければ「公式登録情報」節）。
  本ファイルは本文を持たず、実行時に md から抽出する。
- 引き当て失敗・曖昧・サマリ未整備（ブレイン等）はすべて None を返し、呼び出し側は
  「注入なし＝従来挙動（ツール名のみで生成）」に倒す。誤注入より非注入に倒す方針。
"""
from __future__ import annotations

import logging
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

logger = logging.getLogger(__name__)

# docs/ツール情報/ は hojokin/ の親（リポジトリルート）直下。gitignore 外でデプロイ済み。
# config.BASE_DIR は CWD 依存で Streamlit Cloud の起動位置に左右されるため使わない。
_DOCS_DIR = Path(__file__).resolve().parent.parent / 'docs' / 'ツール情報'

# 見出しを部分一致で拾うキーワード（番号 "8." や 【】 のゆれに非依存にする）
_SUMMARY_HEADING_KW = '事業内容255字用サマリ'
_OFFICIAL_HEADING_KW = '公式登録情報'

# _normalize で畳む区切り記号（空白・中黒・各種ハイフン/ダッシュ・スラッシュ・括弧・アンダースコア）。
# カタカナ長音「ー」は別ツール同士の誤一致を避けるため残す。範囲解釈の曖昧さを避け re.escape で明示。
_STRIP_RE = re.compile('[' + re.escape(
    ' \t\n\r　・･/／()（）[]［］｢｣「」-‐‑–—―－_'
) + ']')


@dataclass(frozen=True)
class MatchResult:
    """tool_name の引き当て結果。summary_text はプロンプトに差し込む実機能本文。"""

    key: str
    display: str
    summary_text: str
    ai: bool | None
    subproduct: str | None = None


# 別名→ツール定義。aliases は素の表記でよい（_normalize で畳んで比較する）。
# ai: True=生成AI搭載 / False=AI非搭載 / None=製品次第。
# subproducts: {サブ製品キー: [そのサブ製品を示す語...]}（irohana のような製品ファミリー型）。
#   引き当て後にサブ製品を特定し、md の "### <display> <key>" 節（あれば公式概要の引用）を使う。
TOOL_CATALOG: dict[str, dict] = {
    'scale_hr': {
        'display': 'Scale人事評価',
        'file': 'ツール情報_Scale人事評価.md',
        'ai': True,
        'aliases': [
            'Scale人事評価', 'スケール人事評価', 'スケール', 'Scale', 'Scale HR',
            'scale-hr', 'カラフルボックス 人事評価', 'AI人事評価システム',
        ],
    },
    'craftbank': {
        'display': 'クラフトバンクオフィス',
        'file': 'ツール情報_クラフトバンクオフィス.md',
        'ai': False,
        'aliases': [
            'クラフトバンクオフィス', 'クラフトバンク office', 'クラフトバンク',
            'CraftBank Office', 'Craft Bank Office', 'CBO',
        ],
    },
    'clipline': {
        'display': 'ClipLine（ABILI Clip）',
        'file': 'ツール情報_ClipLine_ABILI.md',
        'ai': True,
        'aliases': [
            'ClipLine', 'クリップライン', 'ABILI Clip', 'アビリクリップ',
            'ABILI', 'アビリ',
        ],
    },
    'pisces': {
        'display': 'ピスケスアポ',
        'file': 'ツール情報_ピスケスアポ.md',
        'ai': False,
        'aliases': [
            'ピスケスアポ', 'ピスケス・アポ', 'ピスケス', 'Pisces Apo',
            'e-pisces', '歯科予約ピスケス',
        ],
    },
    'irohana': {
        'display': 'irohana',
        'file': 'ツール情報_irohana.md',
        'ai': True,
        'aliases': [
            'irohana', 'いろはな', 'イロハナ', '特定技能 管理システム', '外国人雇用DX',
        ],
        # 素の製品語(visa/match/study)も含める。サブ製品特定はツール確定後にのみ走るため、
        # 短い英語語でも誤爆しない（ツール特定には display+aliases しか使わない）。
        'subproducts': {
            'visa': ['irohana visa', 'visa', 'ビザ', '在留'],
            'match': ['irohana match', 'match', '人材紹介'],
            'study': ['irohana study', 'study', '2号', 'eラーニング'],
        },
    },
    'brain': {
        'display': 'ブレイン',
        'file': 'ツール情報_ブレイン.md',
        'ai': None,
        'aliases': [
            'ブレイン', '株式会社ブレインコンサルティングオフィス',
            'ブレインコンサルティング', 'PSRコンソーシアム', 'PSR',
        ],
        # 案件ごとに製品が異なりサマリ節を持たない md。サマリ未整備で None に倒れる。
        'subproducts': {},
    },
}


def _normalize(s: str) -> str:
    """表記ゆれを畳む。NFKC（全角→半角・Ａ→A）→ 小文字 → 区切り記号除去。

    入力・別名の両方に同じ正規化を掛けるため、変換が多少 lossy でも照合の一貫性は保たれる。
    カタカナ長音「ー」は区別保持のため残す（消すと別ツール同士の誤一致が増える）。
    """
    if not s:
        return ''
    s = unicodedata.normalize('NFKC', str(s)).lower()
    return _STRIP_RE.sub('', s)


def _read_md(file_name: str) -> str | None:
    path = _DOCS_DIR / file_name
    try:
        return path.read_text(encoding='utf-8')
    except OSError as e:
        logger.warning(f'[tool_catalog] md 読込失敗: {path} ({e})')
        return None


def _section_body(lines: list[str], heading_prefix: str, keyword: str) -> str | None:
    """`heading_prefix`（'## ' / '### '）で始まり keyword を含む最初の見出しの本文を返す。

    本文は「同レベルの次の見出し」直前まで。'### ' 探索時は '## '（上位見出し）でも打ち切る。
    '### ' を含む行は '## ' で始まらない（3文字目が空白でない）ため、'## ' 節は '### ' で割れない。
    """
    start = None
    for i, line in enumerate(lines):
        if line.startswith(heading_prefix) and keyword in line:
            start = i + 1
            break
    if start is None:
        return None
    body: list[str] = []
    for line in lines[start:]:
        if line.startswith(heading_prefix):
            break
        if heading_prefix == '### ' and line.startswith('## '):
            break
        body.append(line)
    text = '\n'.join(body).strip()
    return text or None


def _extract_blockquote(block: str) -> str | None:
    """節本文から引用（'> …'）行だけを連結して返す。引用が無ければ None。"""
    quoted = []
    for ln in block.splitlines():
        s = ln.strip()
        if s.startswith('>'):
            q = s[1:].strip()
            if q:
                quoted.append(q)
    text = ' '.join(quoted)
    return text or None


def extract_summary(file_name: str, subproduct: str | None = None) -> str | None:
    """md からプロンプト注入用の本文を抽出する。

    優先順:
    1. subproduct 指定があり md に "### <display> <subproduct>" 節があれば、その公式概要
       （引用 '> …' があれば引用、無ければ節本文）を返す＝製品単位の正確な概要。
    2. 「…事業内容255字用サマリ」節（概要/課題/解決/効果の4要素）。
    3. 「公式登録情報」節。
    いずれも無ければ None（ブレインのようにサマリ節を持たない md は None）。
    """
    text = _read_md(file_name)
    if not text:
        return None
    lines = text.splitlines()

    if subproduct:
        spec = next((s for s in TOOL_CATALOG.values() if s['file'] == file_name), None)
        display = spec['display'] if spec else ''
        sub_keyword = f'{display} {subproduct}'.strip()  # 例 'irohana visa'
        block = _section_body(lines, '### ', sub_keyword)
        if block:
            return _extract_blockquote(block) or block

    summary = _section_body(lines, '## ', _SUMMARY_HEADING_KW)
    if summary:
        return summary
    return _section_body(lines, '## ', _OFFICIAL_HEADING_KW)


def match_tool(name: str) -> MatchResult | None:
    """tool_name（表記ゆれ含む）から該当ツールを引き当てる。

    ツール特定は各ツールの display + aliases を部分一致で走査し最長一致を採用。
    同点で複数ツールに当たれば曖昧として None（黙って一方に倒さず手動指定へ誘導）。
    サブ製品語（visa/match 等）はツール特定には使わず（'match' 等が無関係名に誤爆するため）、
    ツール確定後のサブ製品特定にのみ使う。サマリ未整備なら None。
    """
    norm = _normalize(name)
    if not norm:
        return None

    best_key: str | None = None
    best_len = 0
    ambiguous = False
    for key, spec in TOOL_CATALOG.items():
        matched_len = 0
        for alias in [spec['display']] + spec.get('aliases', []):
            na = _normalize(alias)
            if na and na in norm:
                matched_len = max(matched_len, len(na))
        if matched_len == 0:
            continue
        if matched_len > best_len:
            best_key, best_len, ambiguous = key, matched_len, False
        elif matched_len == best_len and key != best_key:
            ambiguous = True

    if best_key is None:
        return None
    if ambiguous:
        logger.warning(f'[tool_catalog] ツール特定が曖昧: name={name!r} → 注入なし（手動指定推奨）')
        return None

    spec = TOOL_CATALOG[best_key]
    subproduct = None
    for sub_key, sub_aliases in spec.get('subproducts', {}).items():
        for alias in sub_aliases:
            na = _normalize(alias)
            if na and na in norm:
                subproduct = sub_key
                break
        if subproduct:
            break

    summary = extract_summary(spec['file'], subproduct=subproduct)
    if not summary:
        logger.warning(f'[tool_catalog] サマリ未整備: key={best_key} → 注入なし')
        return None

    return MatchResult(
        key=best_key,
        display=spec['display'],
        summary_text=summary,
        ai=spec.get('ai'),
        subproduct=subproduct,
    )


def build_tool_reference(match: MatchResult | None) -> str:
    """PROMPT_AI_JUDGMENT に差し込むツール実機能ブロックを返す。match=None なら空文字。"""
    if match is None:
        return ''
    if match.ai is True:
        ai_clause = '可（生成AIを搭載済みのため、AI機能を具体的に記述してよい）'
    elif match.ai is False:
        ai_clause = (
            '不可（公式登録上はAI非搭載。AI・生成AIには言及せず、デジタル化による効率化として記述する）'
        )
    else:
        ai_clause = '当該製品の公式登録に従う（不明ならAIに言及しない）'
    product = match.display + (f'（{match.subproduct}）' if match.subproduct else '')
    return (
        '\n【導入ツールの実機能（IT補助金コンソールの公式登録に基づく事実。'
        '事業内容の(3)解決策・(4)効果はこの範囲で記述すること）】\n'
        f'ツール: {product}\n'
        f'{match.summary_text}\n'
        '重要: 上記は当該ツールの公式登録に基づく実機能である。事業内容の要素(3)解決策・(4)効果は'
        'この範囲の機能・用途に厳密に沿って記述し、ここに記載の無い機能・効果を創作・誇張してはならない。'
        f'AI機能への言及は{ai_clause}。\n'
    )


def catalog_display_names() -> list[str]:
    """手動プルダウン用のツール表示名一覧。md に製品節を持つファミリーはサブ製品も併記。"""
    names: list[str] = []
    for spec in TOOL_CATALOG.values():
        names.append(spec['display'])
        for sub_key in spec.get('subproducts', {}):
            names.append(f"{spec['display']} {sub_key}")  # 例 'irohana visa'
    return names
