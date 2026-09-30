"""ゲート（scripts/plugin_gates.py）が、壊れた配布物を確実に止めることを合成の入力で確かめる。"""
import io
import zipfile

from _common import Checker

import plugin_gates as g

GOOD_MD = '---\nname: sample-skill\ndescription: 見本のスキル。\nallowed-tools: \'Read Glob\'\n---\n\n# 見本\n'
BLOCK = '<!-- chat-env:begin -->\n注意\n<!-- chat-env:end -->'


def skill(md: str, name: str = 'sample-skill') -> dict:
    return {f'skills/{name}/SKILL.md': md.encode('utf-8')}


def xlsx(core_xml: str, extra_parts: tuple = ()) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as zf:
        zf.writestr('docProps/core.xml', core_xml)
        for part in extra_parts:
            zf.writestr(part, '<x/>')
    return buf.getvalue()


def main() -> int:
    c = Checker('ゲート（合成の入力）')
    c.check('正しい frontmatter は通る', not g.check_frontmatter(skill(GOOD_MD)))
    c.check('claude.ai で使えないキーは止める',
            bool(g.check_frontmatter(skill(GOOD_MD.replace('allowed-tools', 'when_to_use: いつ\nallowed-tools')))))
    c.check('name とフォルダ名の不一致は止める', bool(g.check_frontmatter(skill(GOOD_MD, 'other-name'))))
    c.check('description の < > は止める',
            bool(g.check_frontmatter(skill(GOOD_MD.replace('見本のスキル。', '見本<b>スキル')))))
    c.check('allowed-tools のリスト形式は止める',
            bool(g.check_frontmatter(skill(GOOD_MD.replace("allowed-tools: 'Read Glob'",
                                                           'allowed-tools:\n  - Read')))))
    c.check('frontmatter が無ければ止める', bool(g.check_frontmatter(skill('# 見出しだけ\n'))))

    files = {
        'skills/a/SKILL.md': b'[ok](ref/x.md) [in](#top) [web](https://example.com)\n'
                             b'```\n[fence](nothing.md)\n```\n',
        'skills/a/ref/x.md': b'[up](../SKILL.md)\n',
        'skills/b/SKILL.md': b'[cross](../a/SKILL.md) [none](missing.md)\n',
    }
    problems = g.check_links(files)
    c.check('スキル内で解決するリンク・ページ内・外部 URL・コードブロックは通る',
            not [p for p in problems if p.startswith('skills/a/')], str(problems))
    c.check('スキルをまたぐリンクは止める', any('スキルの外' in p for p in problems), str(problems))
    c.check('リンク先が無ければ止める', any('リンク先がありません' in p for p in problems), str(problems))

    stale = {
        'skills/a/SKILL.md': ('hojokin-handoff.zip を展開する\n'
                              '<!-- legacy-pack-check:begin -->\n.claude/skills/ があれば\n'
                              '<!-- legacy-pack-check:end -->\n'
                              '開発リポジトリでは .claude/skills/ にある\n').encode('utf-8'),
    }
    problems = g.check_stale_terms(stale)
    c.check('古い語は止め、旧パック確認ブロックと開発リポジトリの注記は例外にする',
            len(problems) == 1 and 'hojokin-handoff.zip' in problems[0], str(problems))

    canon = {'skills/guide/SKILL.md': f'{BLOCK}\n'.encode()}
    blocks = ('chat-env',)
    same = {**canon, 'skills/x/SKILL.md': f'前置き\n{BLOCK}\n'.encode()}
    diff = {**canon, 'skills/x/SKILL.md': BLOCK.replace('注意', '注意!').encode()}
    miss = {**canon, 'skills/x/SKILL.md': b'none'}
    orig = g.dm.SHARED_BLOCKS
    g.dm.SHARED_BLOCKS = blocks
    try:
        c.check('共通ブロックが同じなら通る',
                not g.check_shared_blocks(same, 'skills/guide/SKILL.md', {'skills/x/SKILL.md': blocks}))
        c.check('共通ブロックの文言違いは止める',
                bool(g.check_shared_blocks(diff, 'skills/guide/SKILL.md', {'skills/x/SKILL.md': blocks})))
        c.check('共通ブロックの欠落は止める',
                bool(g.check_shared_blocks(miss, 'skills/guide/SKILL.md', {'skills/x/SKILL.md': blocks})))
    finally:
        g.dm.SHARED_BLOCKS = orig

    ok_core = '<cp:coreProperties><dc:creator>openpyxl</dc:creator></cp:coreProperties>'
    ng_core = '<cp:coreProperties><dc:creator>甲野花子</dc:creator></cp:coreProperties>'
    c.check('xlsx の作成者が空か openpyxl なら通る', not g.check_xlsx_metadata({'a.xlsx': xlsx(ok_core)}))
    c.check('xlsx の作成者に人名があれば止める', bool(g.check_xlsx_metadata({'a.xlsx': xlsx(ng_core)})))
    c.check('xlsx のコメント・作成者パーツは止める',
            bool(g.check_xlsx_metadata({'a.xlsx': xlsx(ok_core, ('xl/comments1.xml',))})))
    c.check('プラグイン最上位の bin/ は止める', bool(g.check_layout({'bin/tool': b'x'}, '')))
    return c.finish()


if __name__ == '__main__':
    raise SystemExit(main())
