"""tests/plugin の共通部品。run_all.py が組み立てた配布版プラグインの場所を環境変数で受け取る。"""
import os
import subprocess
import sys
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, 'reconfigure'):
        _stream.reconfigure(encoding='utf-8')

REPO = Path(__file__).resolve().parents[2]
SKILLS_ENV = 'HOJOKIN_PLUGIN_SKILLS'
STAGE_ENV = 'HOJOKIN_PLUGIN_STAGE'
sys.path.insert(0, str(REPO / 'scripts'))


def skills_dir() -> Path:
    value = os.environ.get(SKILLS_ENV)
    if not value:
        raise SystemExit(f'{SKILLS_ENV} が未設定です。tests/plugin/run_all.py から実行してください')
    return Path(value)


def stage_dir() -> Path:
    return Path(os.environ[STAGE_ENV])


def run_py(script: Path, *args, cwd: Path | None = None) -> subprocess.CompletedProcess:
    env = dict(os.environ, PYTHONIOENCODING='utf-8', PYTHONDONTWRITEBYTECODE='1')
    return subprocess.run([sys.executable, '-B', str(script), *map(str, args)], cwd=cwd,
                          capture_output=True, text=True, encoding='utf-8', errors='replace', env=env)


class Checker:
    def __init__(self, title: str):
        self.passed, self.failed = 0, 0
        print(f'=== {title} ===')

    def check(self, label: str, cond: bool, detail: str = '') -> bool:
        if cond:
            self.passed += 1
            print(f'  PASS: {label}')
        else:
            self.failed += 1
            print(f'  FAIL: {label}  {detail[:600]}')
        return cond

    def finish(self) -> int:
        print(f'--- {self.passed} PASS / {self.failed} FAIL')
        return 1 if self.failed else 0
