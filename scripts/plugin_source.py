"""配布プラグインの材料の取り出し元（git の HEAD、または開発中の確認用に作業ツリー）。

HEAD から取り出すのが正規の経路。作業ツリーの改行コード（autocrlf）や未コミットの変更に
左右されず、コミット済みのものだけを配るため。
"""
import subprocess
from datetime import datetime, timedelta, timezone
from pathlib import Path

JST = timezone(timedelta(hours=9))
TEXT_SUFFIXES = {'.md', '.py', '.json', '.txt', '.toml', '.yml', '.yaml', '.csv', '.tsv'}


class SourceError(Exception):
    pass


def _git(repo: Path, *args: str, input_bytes: bytes | None = None) -> bytes:
    try:
        return subprocess.run(['git', '-c', 'core.quotepath=false', *args], cwd=repo, input=input_bytes,
                              capture_output=True, check=True).stdout
    except (OSError, subprocess.CalledProcessError) as e:
        err = getattr(e, 'stderr', b'') or b''
        raise SourceError(f'git {" ".join(args[:2])} に失敗: {err.decode("utf-8", "replace").strip()}') from e


class GitSource:
    """HEAD にコミット済みの内容だけを返す。"""
    version_suffix = ''

    def __init__(self, repo: Path):
        self.repo = repo
        self._sha = _git(repo, 'rev-parse', 'HEAD').decode().strip()
        iso = _git(repo, 'show', '-s', '--format=%cI', 'HEAD').decode().strip()
        self._date = datetime.fromisoformat(iso).astimezone(JST)
        raw = _git(repo, 'ls-tree', '-r', '-z', '--name-only', self._sha)
        self._tree = {p.decode('utf-8') for p in raw.split(b'\0') if p}
        self._cache: dict[str, bytes] = {}

    def commit_sha(self) -> str:
        return self._sha

    def commit_date(self) -> datetime:
        return self._date

    def list_dir(self, d: str) -> list[str]:
        found = sorted(p for p in self._tree if p.startswith(d.rstrip('/') + '/'))
        if not found:
            raise SourceError(f'HEAD にフォルダがありません（または空）: {d}')
        return found

    def exists(self, path: str) -> bool:
        return path in self._tree

    def prefetch(self, paths) -> None:
        todo = sorted({p for p in paths if p in self._tree and p not in self._cache})
        if not todo:
            return
        req = ''.join(f'{self._sha}:{p}\n' for p in todo).encode('utf-8')
        out = _git(self.repo, 'cat-file', '--batch', input_bytes=req)
        pos = 0
        for p in todo:
            nl = out.index(b'\n', pos)
            header = out[pos:nl].split()
            if len(header) != 3 or header[1] != b'blob':
                raise SourceError(f'HEAD から読めません: {p}')
            size = int(header[2])
            self._cache[p] = out[nl + 1:nl + 1 + size]
            pos = nl + 1 + size + 1

    def read(self, path: str) -> bytes:
        if path not in self._tree:
            raise SourceError(f'許可リストのファイルが HEAD にありません: {path}')
        if path not in self._cache:
            self.prefetch([path])
        return self._cache[path]

    def read_paths(self) -> set[str]:
        return set(self._cache)

    def dirty_paths(self, paths) -> list[str]:
        """作業ツリーやステージに、HEAD と違う変更があるパス（未追跡の新規ファイルも含む）。"""
        paths = sorted(set(paths))
        if not paths:
            return []
        raw = _git(self.repo, 'status', '--porcelain', '-z', '--untracked-files=all', '--', *paths)
        entries = raw.split(b'\0')
        dirty, i = set(), 0
        while i < len(entries):
            entry = entries[i].decode('utf-8')
            i += 1
            if len(entry) < 4:
                continue
            dirty.add(entry[3:])
            if entry[0] in 'RC':   # 改名・複製は次の要素が元のパスなので読み飛ばす
                i += 1
        return sorted(dirty)


class WorktreeSource(GitSource):
    """開発中の確認用。未コミットの作業ツリーから読む（出力には使わない）。"""
    version_suffix = '-worktree'

    def __init__(self, repo: Path):
        super().__init__(repo)
        raw = _git(repo, 'ls-files', '-z', '--cached', '--others', '--exclude-standard')
        self._tree = {p.decode('utf-8') for p in raw.split(b'\0') if p and (repo / p.decode('utf-8')).is_file()}

    def prefetch(self, paths) -> None:
        return None

    def read(self, path: str) -> bytes:
        p = self.repo / path
        if path not in self._tree or not p.is_file():
            raise SourceError(f'作業ツリーにありません: {path}')
        data = p.read_bytes()
        if p.suffix.lower() in TEXT_SUFFIXES:
            data = data.replace(b'\r\n', b'\n')
        self._cache[path] = data
        return data

    def dirty_paths(self, paths) -> list[str]:
        return []
