"""申請原本の適用回を検査し、旧版の暗黙採用を防ぐ。

第7次以降の原本は、次のどちらかを満たすときだけ使う。
1. 版マーカー（docProps/core.xml の subject）が第7次以降用である。
2. 版マーカーがない原本（Drive から書き出した原本、名前を変えた原本など）は、
   申請内容シートの項目位置が転記対応表と一致し（W-TPL-001 と同じ照合）、
   第7次以降用の項目名があり、第1〜6次（旧回）の基準の表記が残っていない。
使えない原本は、担当者が直し方を判断できる理由の文章を返す。
"""
from __future__ import annotations

import re
import unicodedata
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

ROUND7_SUBJECT = 'hojokin:2026_7_plus'
APPLICATION_FILENAMES = {
    '通常枠_2026': '【原本_法人】企業名_通常枠_法人2026_v2.xlsx',
    '通常枠_個人_2026': '【原本_個人】企業名_通常枠_個人2026.xlsx',
    'インボイス枠_2026': '【原本_法人】企業名_インボイス枠_法人2026_v2.xlsx',
    'インボイス枠_個人_2026': '【原本_個人】企業名_インボイス枠_個人2026.xlsx',
}
ROUND7_FIRST_CALL = 7
# 第7次以降の原本だけにある申請内容B列の項目名（scripts/build_round7_templates.py が書き込む）
ROUND7_EMPLOYEE_COUNT_LABEL = '0年度目の対象従業員数（確認済みの正社員換算人数）'
APPLICATION_SHEET = '申請内容'
# ツールが作った申請書（{会社名}_{種別}_AI版.xlsx）は原本として使わない。
TOOL_OUTPUT_MARKER = '_AI版'
LOCK_FILE_PREFIX = '~$'
CORE_PROPERTIES_PART = 'docProps/core.xml'
WORKBOOK_PART = 'xl/workbook.xml'
_DC_SUBJECT = '{http://purl.org/dc/elements/1.1/}subject'
_SHEET_TAG = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'
# 第1〜6次の加点基準の表記（scripts/build_round7_templates.py が置き換える語句）。
# 版マーカーのない原本にこれが残っていれば旧版として扱う。「1,063円」等の金額は除く。
LEGACY_RULE_TEXT = re.compile(
    r'(?<![0-9０-９,，])[6６][3３]円|令和[6６]年\s*[1１][0０]月|令和[7７]年\s*[7７]月'
    r'|令和[7７]年度改定|R7年度改定|R7改定後')
SHOWN_LEGACY_TEXTS = 3
_WS_RE = re.compile(r'\s+')


@dataclass(frozen=True)
class TemplateSelection:
    """原本探索の結果。path が None のとき rejected に使えなかった原本と理由が入る。"""
    path: Path | None
    rejected: tuple[str, ...] = ()


def normalize_label(value) -> str:
    """B列ラベル比較用の正規化: NFKC ＋ 全空白（改行・タブ・全半角スペース）除去"""
    return _WS_RE.sub('', unicodedata.normalize('NFKC', str(value if value is not None else '')))


def check_shinsei_anchor_labels(ws, mapping, call_round: int | None = None) -> list[tuple[int, str, str]]:
    """申請内容シートのアンカー行 B列ラベルを照合し、不一致 [(行, 期待, 実際), ...] を返す（W-TPL-001）。

    比較は先頭 ANCHOR_PREFIX_LEN 文字の相互切詰め等値（startswith だと
    「強み」が隣の「強み（転記）」に一致して ±1 行ズレを見逃す）。
    配布版の判定エンジンにも入るため、このモジュールは template_filler に依存しない。
    """
    from . import config as _config
    n = _config.ANCHOR_PREFIX_LEN
    mismatches = []
    for field, expected in _config.SHINSEI_ANCHOR_LABELS.items():
        if call_round is not None and call_round >= ROUND7_FIRST_CALL and field == 'employee_count_fte':
            expected = ROUND7_EMPLOYEE_COUNT_LABEL
        row = mapping.shinsei.get(field)
        if row is None:
            continue
        candidates = expected if isinstance(expected, tuple) else (expected,)
        actual_raw = ws.cell(row=row, column=2).value
        got = normalize_label(actual_raw)[:n]
        if any(got == normalize_label(want)[:n] for want in candidates):
            continue
        shown = str(actual_raw)[:30] if actual_raw is not None else '（空欄）'
        mismatches.append((row, candidates[0], shown))
    return mismatches


def template_rule_marker(path: Path) -> str:
    """版マーカー。Drive から書き出した原本など core.xml がない場合は空文字。"""
    with zipfile.ZipFile(path) as archive:
        if CORE_PROPERTIES_PART not in archive.namelist():
            return ''
        core = ET.fromstring(archive.read(CORE_PROPERTIES_PART))
    return core.findtext(_DC_SUBJECT, default='') or ''


def _sheet_names(path: Path) -> list[str]:
    """ブック全体を読み込まずにシート名だけを取得する（顧客資料の大きい台帳も速く判定する）。"""
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read(WORKBOOK_PART))
    return [node.get('name', '') for node in root.iter(_SHEET_TAG)]


def _round7_layout_issue(path: Path, template_type: str | None) -> str | None:
    """版マーカーのない原本を、項目位置と第7次以降用の表記で確認する。問題なしなら None。"""
    import openpyxl
    from .config import get_mapping

    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    try:
        if APPLICATION_SHEET not in wb.sheetnames:
            return f'「{APPLICATION_SHEET}」シートがありません。申請書の原本ではない可能性があります。'
        ws = wb[APPLICATION_SHEET]
        kinds = [template_type] if template_type else list(APPLICATION_FILENAMES)
        if not any(not check_shinsei_anchor_labels(ws, get_mapping(kind), call_round=ROUND7_FIRST_CALL)
                   for kind in kinds):
            return ('第7次以降用の項目（「0年度目の対象従業員数」など）が、ツールの転記対応表の位置にありません。'
                    '旧版の原本、選んだ申請枠・法人／個人と違う原本、または行が増減した原本です。')
        stale = sorted({match.group(0) for row in ws.iter_rows(values_only=True) for value in row
                        if isinstance(value, str) for match in LEGACY_RULE_TEXT.finditer(value)})
        if stale:
            shown = '、'.join(stale[:SHOWN_LEGACY_TEXTS])
            return f'第1〜6次（旧回）の基準の表記（{shown}）が残っています。第7次以降用に改訂した原本を使ってください。'
        return None
    finally:
        wb.close()


def application_template_issues(path: Path, call_round: int,
                                template_type: str | None = None) -> list[str]:
    """原本を使えない理由の一覧。空なら使える。template_type を省くと4種のどれかに合えばよい。"""
    if isinstance(call_round, bool) or not isinstance(call_round, int) or call_round < 1:
        return ['原本の照合には公募回の指定が必要です。']
    if template_type is not None and template_type not in APPLICATION_FILENAMES:
        raise ValueError(f'対応していない申請原本種別です: {template_type}')
    name = unicodedata.normalize('NFC', Path(path).name)
    if TOOL_OUTPUT_MARKER in name:
        return [f'「{name}」はツールが作成した申請書です。未記入の原本を指定してください。']
    try:
        marker = template_rule_marker(path)
        if call_round >= ROUND7_FIRST_CALL:
            if marker == ROUND7_SUBJECT:
                return []
            issue = _round7_layout_issue(path, template_type)
            return [f'「{name}」は第7次以降用と確認できない申請原本です。{issue}'] if issue else []
        # 第1〜6次（旧回）の訂正用。第7次以降用の原本を旧回へ流用しない。
        if marker == ROUND7_SUBJECT or _round7_layout_issue(path, template_type) is None:
            return [f'「{name}」は第7次以降用の原本です。第1〜6次（旧回）の訂正には旧版の原本を使ってください。']
        return []
    except (OSError, KeyError, ValueError, zipfile.BadZipFile, ET.ParseError):
        return [f'「{name}」をExcel（.xlsx）の原本として読み取れません。ファイルが壊れていないか確認してください。']


def validate_application_template_version(path: Path, call_round: int,
                                          template_type: str | None = None) -> None:
    issues = application_template_issues(path, call_round, template_type)
    if issues:
        raise ValueError('\n'.join(issues))


def _search_dirs(base_dir: Path, call_round: int) -> list[Path]:
    root = Path(base_dir)
    if call_round >= ROUND7_FIRST_CALL:
        dirs = [root / 'ツール' / '第7次以降', root / '第7次以降', root]
    else:
        dirs = [root / 'ツール', root]
    return [d for d in dirs if d.is_dir()]


def _is_workbook_file(path: Path) -> bool:
    name = unicodedata.normalize('NFC', path.name)
    return path.is_file() and path.suffix.lower() == '.xlsx' and not name.startswith(LOCK_FILE_PREFIX)


def _has_application_sheet(path: Path) -> bool:
    try:
        return APPLICATION_SHEET in _sheet_names(path)
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError):
        return False


def find_application_template(base_dir: Path, template_type: str, call_round: int) -> TemplateSelection:
    """原本を探す。正式名の原本を優先し、第7次以降は名前を変えた原本も中身で確認して使う。"""
    if template_type not in APPLICATION_FILENAMES:
        raise ValueError(f'対応していない申請原本種別です: {template_type}')
    official = APPLICATION_FILENAMES[template_type]
    dirs = _search_dirs(base_dir, call_round)
    rejected: list[str] = []
    for directory in dirs:
        matches = sorted(p for p in directory.iterdir()
                         if _is_workbook_file(p) and unicodedata.normalize('NFC', p.name) == official)
        for path in matches:
            issues = application_template_issues(path, call_round, template_type)
            if not issues:
                return TemplateSelection(path, tuple(rejected))
            rejected.extend(issues)
    if call_round < ROUND7_FIRST_CALL:
        return TemplateSelection(None, tuple(rejected))
    accepted: list[Path] = []
    for directory in dirs:
        for path in sorted(directory.iterdir()):
            if (not _is_workbook_file(path) or unicodedata.normalize('NFC', path.name) == official
                    or not _has_application_sheet(path)):
                continue
            issues = application_template_issues(path, call_round, template_type)
            if issues:
                rejected.extend(issues)
            else:
                accepted.append(path)
    if len(accepted) == 1:
        return TemplateSelection(accepted[0], tuple(rejected))
    if accepted:
        names = '、'.join(f'「{unicodedata.normalize("NFC", p.name)}」' for p in accepted)
        rejected.append(f'使える原本が複数あります（{names}）。使う原本を1つだけにしてください。')
    return TemplateSelection(None, tuple(rejected))


def select_application_template(base_dir: Path, template_type: str, call_round: int) -> Path | None:
    """版別ディレクトリを優先。アップロード原本も版マーカーか中身で検査する。"""
    return find_application_template(base_dir, template_type, call_round).path
