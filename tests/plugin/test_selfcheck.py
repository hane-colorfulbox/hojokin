"""各スキルの selfcheck.py（版と中身の確認）を配布版で確かめる。"""
import shutil
import tempfile
from pathlib import Path

from _common import Checker, run_py, skills_dir

import distribution_manifest as dm


def main() -> int:
    c = Checker('selfcheck.py（配布版）')
    skills = skills_dir()
    for spec in dm.SKILLS:
        r = run_py(skills / spec.name / dm.SELFCHECK_NAME, cwd=skills)
        c.check(f'{spec.name}: 配布時のままなら終了コード0', r.returncode == 0, r.stdout)
    r = run_py(skills / 'hojokin-guide' / dm.SELFCHECK_NAME, cwd=skills)
    c.check('隣のスキルもまとめて確かめる（4本）',
            all(f'✅ {s.name}:' in r.stdout for s in dm.SKILLS), r.stdout)
    c.check('プラグインの版と元コミットを表示する', 'プラグイン  : hojokin 版' in r.stdout, r.stdout)
    with tempfile.TemporaryDirectory() as tmp:
        copy = Path(tmp) / 'skills' / 'katen-judge'
        shutil.copytree(skills / 'katen-judge', copy)
        target = copy / 'review' / 'run_katen_judge.py'
        target.write_text(target.read_text(encoding='utf-8') + '\n# 書き換え\n', encoding='utf-8')
        r = run_py(copy / dm.SELFCHECK_NAME, cwd=tmp)
        c.check('書き換わったファイルがあれば終了コード1', r.returncode == 1 and '書き換わっている' in r.stdout, r.stdout)
        (copy / 'reference').rename(copy / 'reference_x')
        r = run_py(copy / dm.SELFCHECK_NAME, cwd=tmp)
        c.check('欠けたファイルを報告する', '欠けている' in r.stdout, r.stdout)
        (copy / dm.BUILD_INFO_NAME).unlink()
        r = run_py(copy / dm.SELFCHECK_NAME, cwd=tmp)
        c.check('_build.json が無ければ終了コード2', r.returncode == 2, r.stdout)
    return c.finish()


if __name__ == '__main__':
    raise SystemExit(main())
