"""プラグイン用リポジトリへの書き出し（plugin_output.sync_repo）が、他のフォルダを壊さないことを確かめる。"""
import json
import tempfile
from pathlib import Path

from _common import REPO, Checker

import plugin_output as po


def files_v(version: str, extra: dict | None = None) -> dict:
    files = {'plugins/hojokin/skills/a/SKILL.md': f'v{version}'.encode(), 'README.md': b'readme'}
    files.update(extra or {})
    files[po.BUILD_INFO_NAME] = json.dumps({'plugin': 'hojokin', 'files': {p: 'x' for p in files}}).encode()
    return files


def refuses(fn) -> bool:
    try:
        fn()
    except SystemExit:
        return True
    return False


def main() -> int:
    c = Checker('プラグイン用リポジトリへの書き出し')
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        dest = tmp / 'plugins-repo'
        po.sync_repo(dest, files_v('1', {'plugins/hojokin/skills/a/old.md': b'old'}), REPO)
        c.check('空のフォルダには書き出せる', (dest / 'plugins/hojokin/skills/a/SKILL.md').read_bytes() == b'v1')
        (dest / 'メモ_手で置いた.txt').write_text('消えてはいけない', encoding='utf-8')
        po.sync_repo(dest, files_v('2'), REPO)
        c.check('前回配ったが今回無いファイルは消える', not (dest / 'plugins/hojokin/skills/a/old.md').exists())
        c.check('前回配っていないファイル（手で置いたもの）は消さない', (dest / 'メモ_手で置いた.txt').exists())
        c.check('中身は今回の版になる', (dest / 'plugins/hojokin/skills/a/SKILL.md').read_bytes() == b'v2')

        other = tmp / 'other-work'
        other.mkdir()
        (other / '顧客メモ.txt').write_text('大事なファイル', encoding='utf-8')
        c.check('このビルドの BUILD_INFO.json が無いフォルダには書き出さない',
                refuses(lambda: po.sync_repo(other, files_v('3'), REPO)) and (other / '顧客メモ.txt').exists())
        c.check('開発リポジトリやその中には書き出さない',
                refuses(lambda: po.sync_repo(REPO, files_v('3'), REPO))
                and refuses(lambda: po.sync_repo(REPO / 'scripts', files_v('3'), REPO)))
        foreign = tmp / 'foreign'
        foreign.mkdir()
        (foreign / po.BUILD_INFO_NAME).write_text(json.dumps({'plugin': 'other'}), encoding='utf-8')
        c.check('別のプラグインのリポジトリには書き出さない', refuses(lambda: po.sync_repo(foreign, files_v('3'), REPO)))
    return c.finish()


if __name__ == '__main__':
    raise SystemExit(main())
