"""申請原本の適用回を検査し、旧版の暗黙採用を防ぐ。"""
from __future__ import annotations

import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROUND7_SUBJECT = 'hojokin:2026_7_plus'
APPLICATION_FILENAMES = {
    '通常枠_2026': '【原本_法人】企業名_通常枠_法人2026_v2.xlsx',
    '通常枠_個人_2026': '【原本_個人】企業名_通常枠_個人2026.xlsx',
    'インボイス枠_2026': '【原本_法人】企業名_インボイス枠_法人2026_v2.xlsx',
    'インボイス枠_個人_2026': '【原本_個人】企業名_インボイス枠_個人2026.xlsx',
}


def template_rule_marker(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        core = ET.fromstring(archive.read('docProps/core.xml'))
    return core.findtext('{http://purl.org/dc/elements/1.1/}subject', default='') or ''


def validate_application_template_version(path: Path, call_round: int) -> None:
    if isinstance(call_round, bool) or not isinstance(call_round, int) or call_round < 1:
        raise ValueError('原本の照合には公募回の指定が必要です。')
    marker = template_rule_marker(path)
    if call_round >= 7 and marker != ROUND7_SUBJECT:
        raise ValueError('第7次以降用と確認できない申請原本です。新版の原本を指定してください。')
    if call_round <= 6 and marker == ROUND7_SUBJECT:
        raise ValueError('第1〜6次の案件に第7次以降の申請原本は使えません。')


def select_application_template(base_dir: Path, template_type: str, call_round: int) -> Path | None:
    """版別ディレクトリを優先。アップロード原本もメタデータで検査する。"""
    if template_type not in APPLICATION_FILENAMES:
        raise ValueError(f'対応していない申請原本種別です: {template_type}')
    root = Path(base_dir)
    name = APPLICATION_FILENAMES[template_type]
    if call_round >= 7:
        dirs = [root / 'ツール' / '第7次以降', root / '第7次以降', root]
    else:
        dirs = [root / 'ツール', root]
    for directory in dirs:
        if not directory.is_dir():
            continue
        matches = [p for p in directory.iterdir()
                   if p.is_file() and unicodedata.normalize('NFC', p.name) == name]
        for path in matches:
            try:
                validate_application_template_version(path, call_round)
            except (ValueError, KeyError, zipfile.BadZipFile, ET.ParseError):
                continue
            return path
    return None
