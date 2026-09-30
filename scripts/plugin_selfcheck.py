# -*- coding: utf-8 -*-
"""補助金プラグインの版と中身の確認（各スキルのフォルダに selfcheck.py として同梱）。

1回の実行で次を出す。標準ライブラリだけで動く（チャット・Cowork・Claude Code 共通）。
  - Python の版、openpyxl と PyMuPDF（fitz）の有無
  - プラグインの版と、元になった開発リポジトリのコミット
  - このスキル（と、隣に並んでいる他のスキル）のファイルが、配布時の sha256 と一致するか

使い方: python selfcheck.py
終了コード: 0=すべて一致 / 1=欠け・書き換わりあり / 2=_build.json が無い
"""
import hashlib
import importlib
import json
import platform
import sys
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, 'reconfigure'):
        _stream.reconfigure(encoding='utf-8')

BUILD_INFO = '_build.json'
OPTIONAL_MODULES = (('openpyxl', 'openpyxl'), ('fitz', 'PyMuPDF'))


def module_version(mod: str) -> str:
    try:
        m = importlib.import_module(mod)
    except Exception:  # noqa: BLE001 - 未導入・壊れた導入のどちらも「なし」と表示する
        return 'なし'
    return str(getattr(m, '__version__', None) or getattr(m, 'VersionBind', None) or 'あり')


def verify_skill(folder: Path) -> tuple[dict, list[str]]:
    info = json.loads((folder / BUILD_INFO).read_text(encoding='utf-8'))
    problems = []
    for rel, want in sorted(info.get('files', {}).items()):
        p = folder / rel
        if not p.is_file():
            problems.append(f'欠けている: {rel}')
        elif hashlib.sha256(p.read_bytes()).hexdigest() != want:
            problems.append(f'書き換わっている: {rel}')
    return info, problems


def main() -> int:
    here = Path(__file__).resolve().parent
    if not (here / BUILD_INFO).exists():
        print(f'❌ {BUILD_INFO} がありません（配布版のスキルではない可能性）: {here}')
        return 2
    print(f'Python      : {platform.python_version()}（{platform.system()}）')
    for mod, label in OPTIONAL_MODULES:
        print(f'{label:<12}: {module_version(mod)}')
    folders = [here] + sorted(p for p in here.parent.iterdir()
                              if p != here and (p / BUILD_INFO).is_file())
    failed = False
    for i, folder in enumerate(folders):
        info, problems = verify_skill(folder)
        if i == 0:
            print(f'プラグイン  : {info.get("plugin")} 版 {info.get("version")}'
                  f'（元コミット {info.get("source_commit", "")[:12]}）')
        mark = '✅' if not problems else '❌'
        print(f'{mark} {info.get("skill")}: {len(info.get("files", {}))} ファイル'
              + ('' if not problems else f'・問題 {len(problems)} 件'))
        for msg in problems[:20]:
            print(f'     - {msg}')
        failed |= bool(problems)
    if len(folders) == 1:
        print('（他のスキルのフォルダは見えない環境です。チャットでは使ったスキルの分だけが見えます）')
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
