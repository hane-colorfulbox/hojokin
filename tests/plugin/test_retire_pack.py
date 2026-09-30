"""旧パックの退役版 ZIP を、模擬の旧パックフォルダで旧経路（update_handoff.py）に通して確かめる。

旧パックの利用者が「引き継ぎを最新に更新して」を実行した場合と同じ動きを再現し、
旧スキルが廃止案内に置き換わる・古い資料とエンジンが消える・配布版 CLAUDE.md が移行案内になる・
本人が書いた CLAUDE.md は守られる、の4点を見る。
"""
import json
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

from _common import REPO, Checker, run_py

import retire_pack

MARKER = retire_pack.CLAUDE_MARKER
OLD_MANAGED_FILES = {
    '.claude/skills/wagebook-convert/SKILL.md': '---\nname: wagebook-convert\ndescription: 旧版\n---\n',
    '.claude/skills/wagebook-convert/review/verify_wagebook.py': 'print("old")\n',
    '.claude/skills/katen-judge/SKILL.md': '---\nname: katen-judge\ndescription: 旧版\n---\n',
    '.claude/skills/case-docs-check/SKILL.md': '---\nname: case-docs-check\ndescription: 旧版\n---\n',
    '引き継ぎ/00_はじめに.md': '# 旧 00\n',
    '引き継ぎ/02_制度知識ベース.md': '# 旧 02\n',
}
OLD_LISTED_FILES = {
    'docs/運用マニュアル.md': '# 旧マニュアル\n',
    'hojokin/wage_reader.py': '# 旧エンジン\n',
    'hojokin/__init__.py': '',
}


class _Src:
    def commit_date(self):
        return datetime(2026, 9, 30)


def make_old_pack(root: Path, claude_md: str, with_history: bool = True) -> None:
    for rel, text in {**OLD_MANAGED_FILES, **OLD_LISTED_FILES}.items():
        p = root / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text, encoding='utf-8')
    (root / 'CLAUDE.md').write_text(claude_md, encoding='utf-8')
    (root / 'scripts').mkdir()
    shutil.copy(REPO / 'scripts' / 'update_handoff.py', root / 'scripts' / 'update_handoff.py')
    (root / '案件メモ_自分用.md').write_text('利用者が自分で置いたファイル\n', encoding='utf-8')
    (root / '引き継ぎ' / '自分のメモ.md').write_text('利用者が引き継ぎ/ に置いたファイル\n', encoding='utf-8')
    if with_history:
        listed = sorted([*OLD_MANAGED_FILES, *OLD_LISTED_FILES, 'CLAUDE.md', 'scripts/update_handoff.py'])
        (root / '.handoff_version').write_text(json.dumps({'built_on': '2026-08-17', 'files': listed}),
                                               encoding='utf-8')


def main() -> int:
    c = Checker('旧パックの退役版 ZIP')
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        retire_zip = tmp / 'hojokin-handoff.zip'
        retire_pack.write_retire_zip(retire_zip, _Src())

        old = tmp / 'old_pack'
        make_old_pack(old, f'{MARKER}\n# 旧・配布版の索引\n')
        r = run_py(old / 'scripts' / 'update_handoff.py', retire_zip, cwd=old)
        c.check('旧経路の更新が終了コード0で終わる', r.returncode == 0, r.stdout + r.stderr)
        for name in retire_pack.OLD_SKILLS:
            skill_md = (old / '.claude/skills' / name / 'SKILL.md').read_text(encoding='utf-8')
            c.check(f'旧スキル {name} は廃止案内だけになり、自動では使われない',
                    'disable-model-invocation: true' in skill_md and '廃止' in skill_md)
        c.check('旧スキルの付属ファイルは消える', not (old / '.claude/skills/wagebook-convert/review').exists())
        c.check('引き継ぎ/ には触らない（利用者が置いたファイルも残る）',
                (old / '引き継ぎ' / '自分のメモ.md').exists() and (old / '引き継ぎ' / '02_制度知識ベース.md').exists())
        c.check('更新の記録がある利用者は、前回配った古い資料・エンジン・更新スクリプトが消える',
                not any((old / rel).exists() for rel in [*OLD_LISTED_FILES, 'scripts/update_handoff.py']))
        claude_md = (old / 'CLAUDE.md').read_text(encoding='utf-8')
        c.check('配布版 CLAUDE.md は移行案内に置き換わる', '旧パックは廃止されました' in claude_md and MARKER in claude_md)
        c.check('利用者が自分で置いたファイルは残る', (old / '案件メモ_自分用.md').exists())

        mine = tmp / 'own_claude_md'
        make_old_pack(mine, '# 自分で書いた CLAUDE.md\n')
        r = run_py(mine / 'scripts' / 'update_handoff.py', retire_zip, cwd=mine)
        c.check('本人が書いた CLAUDE.md は上書きされない',
                r.returncode == 0 and (mine / 'CLAUDE.md').read_text(encoding='utf-8') == '# 自分で書いた CLAUDE.md\n')
        c.check('その場合は移行案内が CLAUDE.handoff.md に置かれる',
                '旧パックは廃止されました' in (mine / 'CLAUDE.handoff.md').read_text(encoding='utf-8'))

        manual = tmp / 'manual_unzip'
        make_old_pack(manual, f'{MARKER}\n# 旧・配布版の索引\n', with_history=False)
        r = run_py(manual / 'scripts' / 'update_handoff.py', retire_zip, cwd=manual)
        skill_md = (manual / '.claude/skills/katen-judge/SKILL.md').read_text(encoding='utf-8')
        c.check('手で展開しただけの利用者でも、旧スキルは廃止案内に置き換わる',
                r.returncode == 0 and 'disable-model-invocation: true' in skill_md)
        c.check('手で展開しただけの利用者の古い資料は消さない（片付けは本人）',
                all((manual / rel).exists() for rel in OLD_LISTED_FILES))
        c.check('片付け手順に「このファイル」のような置き場所で意味が変わる書き方がない',
                'このファイル' not in retire_pack.CLEANUP_STEPS)
    return c.finish()


if __name__ == '__main__':
    raise SystemExit(main())
