# -*- coding: utf-8 -*-
"""処理パイプライン: ファイル検出 → 抽出 → 転記 → 出力"""
from __future__ import annotations

import logging
import re
import tempfile
import unicodedata
from pathlib import Path

from .config import get_mapping, CLAUDE_API_KEY
from .models import ExtractionResult, ProcessingStatus, CompanyInfo
from .ai_extractor import create_extractor, BaseExtractor, StubExtractor
from .hearing_reader import read_hearing_sheet
from .template_filler import fill_template
from .wage_calculator import (
    create_wage_calculation,
    PayrollEmployee,
    calculate_per_capita_wage,
)
from .wage_reader import read_wage_ledger, read_wage_ledgers, export_wage_ledger_summary
from .pdf_reader import pdf_to_images

logger = logging.getLogger(__name__)


# ───────────────────────── ファイル名年月パース ─────────────────────────
# 決算書ファイル名から「期末年月 (year, month)」を取り出すユーティリティ。
# 例:
#   令和7年3月決算書        → (2025, 3)
#   令和6年3月決算書        → (2024, 3)
#   R7.3決算書              → (2025, 3)
#   2025年3月期決算書        → (2025, 3)
#   2025.03_決算書          → (2025, 3)
#   平成31年4月決算書        → (2019, 4)
# 取れなければ None。年月の妥当性チェック付き（年 1900-2100、月 1-12）。

_REIWA_RE = re.compile(r'令和\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月')
_REIWA_GANNEN_RE = re.compile(r'令和\s*元\s*年\s*(\d{1,2})\s*月')
_HEISEI_RE = re.compile(r'平成\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月')
_HEISEI_GANNEN_RE = re.compile(r'平成\s*元\s*年\s*(\d{1,2})\s*月')
_RY_DOT_RE = re.compile(r'(?<![A-Za-z0-9])R\s*(\d{1,2})\s*[\.\-_／/]\s*(\d{1,2})(?!\d)')
_YYYY_KANJI_RE = re.compile(r'(\d{4})\s*年\s*(\d{1,2})\s*月')
_YYYY_SEP_RE = re.compile(r'(?<!\d)(\d{4})[\.\-_／/](\d{1,2})(?!\d)')


def _parse_fiscal_end_from_filename(name: str) -> tuple[int, int] | None:
    """ファイル名から期末年月 (year, month) を取り出す。失敗時 None。

    NFC 正規化してから複数のパターンを試行し、最初に成立したものを返す。
    """
    s = unicodedata.normalize('NFC', name)

    def _valid(y: int, m: int) -> tuple[int, int] | None:
        if 1900 <= y <= 2100 and 1 <= m <= 12:
            return (y, m)
        return None

    # 令和元年 (=2019)
    m = _REIWA_GANNEN_RE.search(s)
    if m:
        return _valid(2019, int(m.group(1)))
    # 令和N年M月 (令和N = 2018 + N)
    m = _REIWA_RE.search(s)
    if m:
        return _valid(2018 + int(m.group(1)), int(m.group(2)))
    # 平成元年 (=1989)
    m = _HEISEI_GANNEN_RE.search(s)
    if m:
        return _valid(1989, int(m.group(1)))
    # 平成N年M月 (平成N = 1988 + N)
    m = _HEISEI_RE.search(s)
    if m:
        return _valid(1988 + int(m.group(1)), int(m.group(2)))
    # RN.M / RN-M / RN_M (略式の令和)
    m = _RY_DOT_RE.search(s)
    if m:
        return _valid(2018 + int(m.group(1)), int(m.group(2)))
    # YYYY年M月
    m = _YYYY_KANJI_RE.search(s)
    if m:
        return _valid(int(m.group(1)), int(m.group(2)))
    # YYYY-MM / YYYY.MM / YYYY_MM
    m = _YYYY_SEP_RE.search(s)
    if m:
        return _valid(int(m.group(1)), int(m.group(2)))
    return None


_WAGE_PERIOD_REIWA_RE = re.compile(
    r'R\s*(\d{1,2})\s*[\.\-_／/]\s*(\d{1,2})\s*[-〜~～ー]\s*R\s*(\d{1,2})\s*[\.\-_／/]\s*(\d{1,2})'
)
_WAGE_PERIOD_YYYY_RE = re.compile(
    r'(\d{4})\s*[\.\-_／/年]\s*(\d{1,2})\s*月?\s*[-〜~～ー]\s*(\d{4})\s*[\.\-_／/年]\s*(\d{1,2})'
)


def _parse_wage_ledger_period(name: str) -> tuple[tuple[int, int], tuple[int, int]] | None:
    """賃金台帳のファイル名から (期首年月, 期末年月) を取り出す。失敗時 None。

    例:
      R6.4-R7.3賃金台帳    → ((2024, 4), (2025, 3))
      2024-04～2025-03賃金 → ((2024, 4), (2025, 3))
    """
    s = unicodedata.normalize('NFC', name)

    def _valid(y: int, mo: int) -> bool:
        return 1900 <= y <= 2100 and 1 <= mo <= 12

    m = _WAGE_PERIOD_REIWA_RE.search(s)
    if m:
        sy = 2018 + int(m.group(1))
        sm = int(m.group(2))
        ey = 2018 + int(m.group(3))
        em = int(m.group(4))
        if _valid(sy, sm) and _valid(ey, em):
            return ((sy, sm), (ey, em))
    m = _WAGE_PERIOD_YYYY_RE.search(s)
    if m:
        sy = int(m.group(1))
        sm = int(m.group(2))
        ey = int(m.group(3))
        em = int(m.group(4))
        if _valid(sy, sm) and _valid(ey, em):
            return ((sy, sm), (ey, em))
    return None


def _parse_year_month_from_iso(s: str) -> tuple[int, int] | None:
    """'2025-03' / '2025-03-31' / '2025/03' → (2025, 3)。失敗時 None。"""
    if not s:
        return None
    s = s.strip()
    m = re.match(r'^(\d{4})[\-\./](\d{1,2})', s)
    if not m:
        return None
    y, mo = int(m.group(1)), int(m.group(2))
    if 1900 <= y <= 2100 and 1 <= mo <= 12:
        return (y, mo)
    return None


def _check_pl_wage_period_consistency(
    detector: 'FileDetector',
    financial,  # FinancialData | None
    fiscal_month_override: int | None,
) -> tuple[object, str]:
    """賃金台帳ファイル名の期間と PL 期末年月の整合性をチェック。

    ズレを検出したら financial の数値情報を空に戻して「決算書由来の値を出力に
    乗せない」+ 強警告メッセージを返す。給与支給総額の本算出は賃金台帳ベースで
    続行されるが、テンプレ転記用の財務値はスキップされる。

    Returns:
        (financial_or_reset, warning_message)
        - financial_or_reset: 整合OK時はそのまま。NG時は revenue=0 にリセット
        - warning_message: 空 or 強警告（status.message 末尾に追加する想定）
    """
    if financial is None:
        return financial, ''

    pl_end = _parse_year_month_from_iso(getattr(financial, 'fiscal_year_end', '') or '')

    # 賃金台帳ファイル名から期末を集める
    ledger_paths = detector.get_all('wage_ledger')
    wage_ends: list[tuple[int, int]] = []
    for p in ledger_paths:
        period = _parse_wage_ledger_period(p.name)
        if period is not None:
            wage_ends.append(period[1])

    msgs: list[str] = []

    # ── 賃金台帳期末 vs PL期末 ──
    if wage_ends and pl_end is not None:
        # 賃金台帳が複数あっても、通常は同一期末年月のはず。バラついていたら最頻値
        # を採用（端末ばらつき疑い）。ここでは最も新しいものを採用して比較
        latest_wage_end = max(wage_ends)
        if latest_wage_end != pl_end:
            msgs.append(
                f' ⛔ 致命的不整合: 賃金台帳期末={latest_wage_end[0]:04d}-{latest_wage_end[1]:02d} ／ '
                f'決算書期末={pl_end[0]:04d}-{pl_end[1]:02d}。'
                f'別の期の決算書が読み込まれた可能性が高いため、決算書由来の'
                f'財務値（売上・粗利・営業利益等）の転記をスキップしました。'
                f'資料フォルダから前期の決算書を退避してから再実行してください。'
            )

    # ── 賃金台帳期末 vs ユーザー指定の決算月 ──
    if wage_ends and fiscal_month_override is not None:
        latest_wage_end = max(wage_ends)
        if latest_wage_end[1] != fiscal_month_override:
            msgs.append(
                f' ⛔ 致命的不整合: 賃金台帳期末月={latest_wage_end[1]}月 ／ '
                f'ユーザー指定決算月={fiscal_month_override}月。'
                f'賃金台帳の対象期間がユーザー指定と矛盾しています。'
                f'決算月の設定または賃金台帳の差し替えを確認してください。'
            )

    # ── PL期末 vs ユーザー指定の決算月 ──
    if pl_end is not None and fiscal_month_override is not None and not wage_ends:
        # 賃金台帳ファイル名から期間が取れなかった時のフォールバック
        if pl_end[1] != fiscal_month_override:
            msgs.append(
                f' ⚠ 警告: 決算書期末月={pl_end[1]}月 ／ '
                f'ユーザー指定決算月={fiscal_month_override}月で不一致。'
                f'別の期の決算書が読み込まれた可能性があります。'
            )

    warning = ''.join(msgs)
    if warning and '致命的不整合' in warning:
        # 致命的不整合: financial を「revenue=0」状態にリセットして、テンプレ転記の
        # 起点である「PL値の出力」を全部止める。給与支給総額は賃金台帳から再算出されるが
        # 売上・粗利・営業利益・経常利益・減価償却費・役員報酬の決算書由来値は出ない。
        from .models import FinancialData
        reset = FinancialData()
        # 賃金台帳から取れる情報（人件費周辺）は維持したいが、PL値はリセット
        logger.warning(f'PL/賃金台帳整合性チェック失敗: {warning}')
        return reset, warning

    if warning:
        logger.warning(f'PL/賃金台帳整合性チェック警告: {warning}')
    return financial, warning


def _pick_full_version(candidates: list[Path]) -> Path:
    """同一期PDF候補から「フル版」を選ぶ。

    PDFsam 等の部分抜粋を除外してサイズ最大を採用。
    全候補が抜粋マーカー付きなら、消去法でサイズ最大を採用。
    """
    if len(candidates) == 1:
        return candidates[0]
    EXCLUDE_MARKERS = ('_PDFsam_', 'PDFsam', '部分抜粋', '抜粋')
    full_versions = [
        p for p in candidates
        if not any(m in p.name for m in EXCLUDE_MARKERS)
    ]
    pool = full_versions or candidates
    return max(pool, key=lambda p: p.stat().st_size)


class FileDetector:
    """資料フォルダからファイルを自動分類"""

    PATTERNS = {
        'hearing': ['ヒアリング'],
        'registry': ['履歴事項'],
        'identity': ['運転免許証', '運転経歴証明書', '住民票', '本人確認'],
        'tax': ['納税証明'],
        'pl': ['損益計算書', '決算報告書', '決算書', '収支内訳書', '青色申告'],
        'cost_report': ['製造原価報告書', '原価報告書'],
        'estimate': ['見積', 'お見積'],
        'wage_report': ['賃金状況報告'],
        # 賃金台帳: 「給与台帳」「賃金台帳」両方を検出（給与ソフト出力で「給与台帳」表記が多い）
        'wage_ledger': ['賃金台帳', '給与台帳'],
        'wage_data': ['支給控除一覧', '給与データ'],
    }

    # 自己参照ループ防止: 過去の出力ファイル（前回ツールで生成された Excel）を入力として誤検出しないよう、
    # ファイル名にこれらのパターンを含むファイルは検出対象から除外する。
    # 出力ファイル命名規則 (app.py / pipeline.py を参照):
    #   - {会社名}_{枠}_AI版.xlsx
    #   - {会社名}_給与支給総額計算.xlsx
    #   - {会社名}_賃金台帳一覧.xlsx
    #   - {会社名}_加点①_結果.xlsx / 加点②_結果.xlsx
    OUTPUT_FILE_MARKERS = (
        '_AI版', '_給与支給総額計算', '_賃金台帳一覧',
        '_加点①', '_加点②',
    )

    # カテゴリ別の許可拡張子（小文字で比較）
    # openpyxl系のカテゴリに .csv が混入すると読み込み時に例外が出て pipeline が全滅するため、
    # 検出段階で弾く。拡張子は後段の読み取り処理に合わせて絞っている。
    ALLOWED_EXTS = {
        'hearing':     {'.xlsx', '.xlsm'},
        'registry':    {'.pdf'},
        'identity':    {'.pdf'},
        'tax':         {'.pdf'},
        'pl':          {'.pdf'},
        'cost_report': {'.pdf'},
        'estimate':    {'.xlsx', '.xlsm', '.pdf'},
        'wage_report': {'.xlsx', '.xlsm'},
        # 2026-05 方針変更: 賃金台帳の回収は Excel/CSV に集約。PDF は受け付けない
        # （PDFで届いた場合はローカルで Excel/CSV に変換してから投入する運用）
        'wage_ledger': {'.xlsx', '.xlsm', '.csv'},
        'wage_data':   {'.pdf'},
    }

    def __init__(self, folder: Path):
        self.folder = folder
        self.files: dict[str, list[Path]] = {k: [] for k in self.PATTERNS}
        self.skipped: list[tuple[str, str, str]] = []  # (category, filename, reason)
        self._scan()

    def _scan(self):
        """フォルダを再帰的にスキャンしてファイル分類"""
        for p in self._iter_files(self.folder):
            if p.name.startswith('~$'):
                continue
            # Google Drive等で作成されたファイル名はNFD分解形式(例: グ=ク+濁点)で
            # 保存されていることがあり、NFCのキーワードと素直にin比較すると一致しない。
            # 正規化してから判定する。
            name_nfc = unicodedata.normalize('NFC', p.name)
            # 自己参照ループ防止: 過去の出力ファイルは入力として扱わない
            if any(marker in name_nfc for marker in self.OUTPUT_FILE_MARKERS):
                self.skipped.append(('output_file', p.name, '過去の出力ファイル（処理対象外）'))
                logger.info(f'除外: [output_file] {p.name} (過去の出力ファイル)')
                continue
            for category, keywords in self.PATTERNS.items():
                if any(kw in name_nfc for kw in keywords):
                    allowed = self.ALLOWED_EXTS.get(category)
                    if allowed is not None and p.suffix.lower() not in allowed:
                        self.skipped.append((category, p.name, f'拡張子{p.suffix}は{category}では非対応'))
                        logger.info(f'除外: [{category}] {p.name} (許可拡張子: {sorted(allowed)})')
                        break
                    self.files[category].append(p)
                    logger.debug(f'検出: [{category}] {p.name}')
                    break

    def _iter_files(self, directory: Path):
        """日本語パス対応の再帰ファイル探索"""
        try:
            for p in directory.iterdir():
                if p.is_dir() and not p.name.startswith(('.', '_')):
                    yield from self._iter_files(p)
                elif p.is_file():
                    yield p
        except PermissionError:
            logger.warning(f'アクセス拒否: {directory}')

    def get(self, category: str) -> Path | None:
        """カテゴリの最初のファイルを返す"""
        files = self.files.get(category, [])
        return files[0] if files else None

    def get_all(self, category: str) -> list[Path]:
        """カテゴリの全ファイルを返す"""
        return self.files.get(category, [])

    def get_pl_latest(self, fiscal_month_override: int | None = None) -> Path | None:
        """損益計算書の直近期を返す。

        判定優先順（fiscal_month_override 指定時は決算月一致を最優先）:
          1. ファイル名に「第N期」を含む → N が最大のものを採用（事業年数の進んだ会社対応）
          2. ファイル名から期末年月を抽出（令和X年Y月 / RY.M / YYYY年M月 / YYYY-MM 等）
             → fiscal_month_override が指定されていれば、月が一致する候補のみで比較
             → 期末年月が最新のものを採用
          3. 同期が複数あれば、PDFsam等の部分抜粋を除外し、フル版（サイズ最大）を採用
          4. 上記いずれも該当なければ更新日時最新を採用

        旧実装の問題: ステップ1・2が無いと「令和7年3月決算書」「令和6年3月決算書」のような
        ファイル名（第N期表記なし）が並んだ際、mtime ガチャで前期決算書が選ばれて
        賃金台帳の期間と整合しない財務値が転記される誤動作があった。
        """
        pls = self.files.get('pl', [])
        if not pls:
            return None

        # ---- ステップ1: 第N期表記 ----
        period_re = re.compile(r'第(\d+)期')

        def period_num(p: Path) -> int:
            m = period_re.search(p.name)
            return int(m.group(1)) if m else -1

        nums = [(p, period_num(p)) for p in pls]
        max_num = max(n for _, n in nums)

        if max_num >= 0:
            latest = [p for p, n in nums if n == max_num]
            return _pick_full_version(latest)

        # ---- ステップ2: ファイル名から期末年月を抽出 ----
        date_pairs = [
            (p, _parse_fiscal_end_from_filename(p.name))
            for p in pls
        ]
        with_date = [(p, ym) for p, ym in date_pairs if ym is not None]

        if with_date:
            # 決算月指定があれば、月が一致する候補に絞る
            if fiscal_month_override is not None:
                month_match = [
                    (p, ym) for p, ym in with_date
                    if ym[1] == fiscal_month_override
                ]
                if month_match:
                    with_date = month_match
                else:
                    logger.warning(
                        f'決算月{fiscal_month_override}月と一致するファイル名が無く、'
                        f'年月最新で選びます: '
                        f'{[(p.name, ym) for p, ym in with_date]}'
                    )
            # 期末年月が最新のもの
            max_ym = max(ym for _, ym in with_date)
            latest = [p for p, ym in with_date if ym == max_ym]
            if len(latest) > 1:
                logger.info(
                    f'同一期末年月のPL候補が{len(latest)}件あり → フル版優先で選択'
                )
            return _pick_full_version(latest)

        # ---- ステップ3: フォールバック（mtime 最新） ----
        logger.warning(
            f'PL候補のファイル名から期番号も期末年月も抽出できないため、'
            f'更新日時最新でフォールバックします: {[p.name for p in pls]}'
        )
        return max(pls, key=lambda p: p.stat().st_mtime)

    def summary(self) -> str:
        """検出結果のサマリ"""
        lines = ['検出されたファイル:']
        for cat, files in self.files.items():
            if files:
                names = [f.name for f in files]
                lines.append(f'  {cat}: {", ".join(names)}')
            else:
                lines.append(f'  {cat}: なし')
        if self.skipped:
            lines.append('')
            lines.append('除外されたファイル（拡張子不一致）:')
            for cat, name, reason in self.skipped:
                lines.append(f'  [{cat}] {name} — {reason}')
        return '\n'.join(lines)


def run_application_transfer(
    resource_folder: Path,
    template_path: Path,
    template_type: str,
    output_path: Path,
    extractor: BaseExtractor | None = None,
    fiscal_month_override: int | None = None,
    has_cost_report_hint: bool = False,
) -> ProcessingStatus:
    """
    タスク1: 申請書転記の実行

    Args:
        resource_folder: 資料フォルダ
        template_path: テンプレートExcelパス
        template_type: '通常枠_2026' or 'インボイス枠_2026'
        output_path: 出力ファイルパス
        extractor: AI抽出器（省略時は自動選択）
        fiscal_month_override: ユーザー指定の決算月（1〜12）。指定時はAI推定より優先
        has_cost_report_hint: ユーザーが「製造原価報告書あり」とチェック済みなら True。
            自動検出されなかった場合の警告強化に使う
    """
    status = ProcessingStatus(
        company_name=resource_folder.name,
        template_type=template_type,
        status='処理中',
    )

    try:
        from .ai_extractor import APICreditExhaustedError
        mapping = get_mapping(template_type)

        if extractor is None:
            extractor = create_extractor(CLAUDE_API_KEY)

        # ファイル検出
        detector = FileDetector(resource_folder)
        logger.info(detector.summary())

        extraction = ExtractionResult()

        # API残高切れ等で Phase 2 が部分実行になった場合の理由
        api_skipped_reason: str = ''

        # ===== Phase 1: API不要な処理（確実に動く・コスト¥0）=====
        # ヒアリングシート読取（Excel直接読取）
        hearing_path = detector.get('hearing')
        hearing_data = {}
        if hearing_path:
            hearing_data = read_hearing_sheet(hearing_path)
            logger.info(f'ヒアリングシート: {len(hearing_data)}行読込')
        else:
            logger.warning('ヒアリングシートが見つかりません')

        # 見積書: Excel なら API 不要なのでここで処理（PDF は Phase 2）
        estimate_path = detector.get('estimate')
        estimate_pdf_path = None
        if estimate_path:
            if estimate_path.suffix == '.xlsx':
                # Excel の見積書は直接読取（API不要）
                import openpyxl
                wb_est = openpyxl.load_workbook(estimate_path, data_only=True)
                ws = wb_est[wb_est.sheetnames[0]]
                tool_name_keywords = ['件名', '品名', 'ツール名', '商品名', 'サービス名']
                found = False
                for row in ws.iter_rows(min_row=1, max_row=30):
                    for i, cell in enumerate(row):
                        v = cell.value
                        if v and isinstance(v, str):
                            if any(kw in v for kw in tool_name_keywords):
                                if i + 1 < len(row) and row[i + 1].value:
                                    extraction.estimate.tool_name = str(row[i + 1].value)
                                    found = True
                                    break
                    if found:
                        break
                if not found:
                    import re
                    name = estimate_path.stem
                    for remove in ['お見積り', 'お見積', '見積り', '見積', '御見積', '_', '様']:
                        name = name.replace(remove, '')
                    name = re.sub(r'\d{8}', '', name)
                    name = re.sub(r'\d{4}[-/]\d{2}[-/]\d{2}', '', name)
                    name = name.strip()
                    if len(name) > 2:
                        extraction.estimate.tool_name = name
                wb_est.close()
            else:
                # PDF は Phase 2 で処理
                estimate_pdf_path = estimate_path

        # ===== Phase 2: API使用（残高切れなら部分スキップ。Phase 1 結果は維持）=====
        # 製造原価報告書の自動検出フラグは Phase 2 内で例外発生しても参照されるため
        # try ブロック前に初期化しておく（UnboundLocalError 回避）
        cost_report_detected: bool = False
        try:
            # 履歴事項PDF → CompanyInfo
            registry_path = detector.get('registry')
            if registry_path:
                images = pdf_to_images(registry_path)
                extraction.company = extractor.extract_registry(images)
                logger.info(f'履歴事項: {extraction.company.name}')

            # 損益計算書PDF → FinancialData
            # 決算月指定があれば、ファイル名年月と突合して直近期を確定（誤読防止）
            pl_path = detector.get_pl_latest(fiscal_month_override=fiscal_month_override)
            pl_period_warning = ''
            if pl_path:
                logger.info(f'直近期決算書として採用: {pl_path.name}')
                images = pdf_to_images(pl_path)
                cost_report_path = detector.get('cost_report')
                if cost_report_path:
                    images += pdf_to_images(cost_report_path)
                    logger.info(f'製造原価報告書も読取: {cost_report_path.name}')
                    cost_report_detected = True
                extraction.financial = extractor.extract_pl(images)
                logger.info(f'損益計算書: 売上{extraction.financial.revenue:,}')

                # 賃金台帳期間 vs PL期末 / ユーザー指定決算月の整合性チェック
                # ズレを検出したら財務値転記をスキップ + 強警告
                extraction.financial, pl_period_warning = (
                    _check_pl_wage_period_consistency(
                        detector, extraction.financial, fiscal_month_override,
                    )
                )

            # 納税証明書PDF
            tax_path = detector.get('tax')
            if tax_path:
                images = pdf_to_images(tax_path)
                extraction.tax = extractor.extract_tax(images)

            # 見積書PDF (Excelは Phase 1 で処理済)
            if estimate_pdf_path:
                images = pdf_to_images(estimate_pdf_path)
                extraction.estimate = extractor.extract_estimate(images)

            # AI判断（ヒアリングデータも渡してIT投資状況等の矛盾を防ぐ）
            extraction.ai_judgment = extractor.generate_ai_judgment(
                extraction.company,
                extraction.financial,
                extraction.estimate.tool_name,
                hearing_data=hearing_data,
            )
        except APICreditExhaustedError as e:
            # API残高切れ → Phase 1 の結果（ヒアリング・賃金台帳Excel等）で申請書を出力
            # 確認キューにAI由来項目が「未取得」として一覧される
            api_skipped_reason = str(e)
            logger.warning(
                f'[Phase 2] API残高切れで以降の API 呼出をスキップ: {e}'
            )

        # 賃金台帳 → 1人当たり給与支給総額の計画値 + 一覧Excel出力
        # 賃金台帳処理: AI 経路（PDF）が残高切れになっても Phase 1 結果は維持
        wage_extraction_method = (
            'AI抽出（Claude Sonnet 4.6）' if extractor is not None
            else '決定論パーサー（USE_AI_WAGE_EXTRACTION=false）'
        )
        try:
            wage_plan, ledger_employees, wage_status = _calc_wage_plan_from_ledger(
                detector, extraction.financial, extractor=extractor,
                fiscal_month_override=fiscal_month_override,
            )
        except APICreditExhaustedError as e:
            api_skipped_reason = api_skipped_reason or str(e)
            logger.warning(f'[Phase 2] 賃金台帳AI抽出で残高切れ: {e}')
            # フォールバック: extractor=None で決定論パーサーのみで再試行
            try:
                wage_plan, ledger_employees, wage_status = _calc_wage_plan_from_ledger(
                    detector, extraction.financial, extractor=None,
                    fiscal_month_override=fiscal_month_override,
                )
                wage_extraction_method = '決定論パーサー（AI残高切れフォールバック）'
            except Exception as e2:
                logger.warning(f'決定論パーサーも失敗: {e2}')
                wage_plan, ledger_employees, wage_status = None, [], 'error'
                wage_extraction_method = '抽出失敗（AI残高切れ＋決定論パーサーも失敗）'

        # ユーザー指定の決算月 vs AI 推定の照合（警告のみ。処理は続行）
        _, fiscal_month_warning = _resolve_fiscal_period(
            extraction.financial, fiscal_month_override,
        )

        # ユーザーが「製造原価報告書あり」とチェックしたのに自動検出されなかった場合の警告
        # （ファイル名キーワード未一致や、損益計算書PDFに統合されているケースなど）
        cost_report_warning = ''
        if has_cost_report_hint and not cost_report_detected:
            cost_report_warning = (
                ' ⚠ 「製造原価報告書あり」とチェックされていますが、製造原価報告書のPDFを'
                '検出できませんでした。ファイル名に「製造原価報告書」を含めるか、'
                '損益計算書PDFに統合されている場合は原本を目視で確認し、'
                '原価部の人件費（労務費・賞与等）が抽出値に含まれているかご確認ください。'
            )

        # AI 生成の事業内容の文字数チェック（240〜255文字が望ましい）
        biz_desc_warning = ''
        biz_desc = (extraction.ai_judgment.business_description or '').strip()
        if biz_desc:
            n = len(biz_desc)
            if n > 255:
                biz_desc_warning = (
                    f' ⚠ 事業内容が文字数制限超過（{n}文字 / 上限255文字）。'
                    f'申請書セルで切り詰められるおそれがあるため、原稿を手動短縮してください。'
                )
            elif n < 240:
                biz_desc_warning = (
                    f' ⚠ 事業内容が短すぎます（{n}文字 / 推奨240〜255文字）。'
                    f'4要素（現状・課題・解決策・期待効果）が十分書き切れているか、'
                    f'ヒアリング情報を追記して厚みを出してください。'
                )

        # 賃金台帳に PDF がアップロードされた場合の警告
        # 2026-05 方針: 賃金台帳は Excel/CSV のみ受け付ける。
        # PDF は detector でスキップされるが、ユーザーには明示警告して再投入を促す
        wage_pdf_warning = ''
        wage_pdf_files = [
            name for cat, name, _ in detector.skipped
            if cat == 'wage_ledger' and name.lower().endswith('.pdf')
        ]
        if wage_pdf_files:
            wage_pdf_warning = (
                f' ⚠ 賃金台帳PDFは受け付け対象外です（{len(wage_pdf_files)}件: '
                f'{", ".join(wage_pdf_files[:3])}'
                f'{"…" if len(wage_pdf_files) > 3 else ""}）。'
                f'ローカルでExcel/CSV形式に変換してから再投入してください。'
            )

        # テンプレート転記（Phase 1 + Phase 2成功分のみ。残高切れ項目は confidence='low' で空欄）
        empty_cells = fill_template(
            template_path=template_path,
            output_path=output_path,
            mapping=mapping,
            hearing_data=hearing_data,
            extraction=extraction,
            wage_plan=wage_plan,
        )

        status.status = '完了' if not api_skipped_reason else '部分完了'
        status.output_files = [output_path.name]
        status.empty_cells = empty_cells
        # 後続タスク（給与計算/加点判定）で再利用するためAI抽出結果をstatusに保持
        status.financial = extraction.financial
        status.ledger_employees = ledger_employees or []
        # Phase 4: 低信頼項目を「確認キュー」として集約
        status.confidence_warnings = _build_confidence_warnings(extraction.financial)
        # 賃金台帳の読み取り状況に応じて完了メッセージに警告を追記（処理は続行）
        wage_warning = ''
        if wage_status == 'no_data':
            wage_warning = ' ⚠ 賃金台帳が読み取れませんでした（給与支給総額は空欄）'
        elif wage_status == 'zero_total':
            wage_warning = ' ⚠ 賃金台帳の給与支給総額が0でした'
        elif wage_status == 'error':
            wage_warning = ' ⚠ 賃金台帳処理中にエラーが発生しました'
        elif wage_status == 'fiscal_year_mismatch':
            wage_warning = (
                ' ⛔ 強警告: 賃金台帳の記録期間と直近事業年度がズレており、'
                '「直近決算期の全月在籍者」から給与支給総額を自動算出できません。'
                '申請書 R215（従業員数）・R216（給与支給総額）・R217〜R219（賃上げ計画）は '
                '空欄のままです。【確認事項】まず賃金台帳の提出期間が直近決算期12ヶ月を'
                '含んでいるかご確認ください。含んでいなければ顧客に正しい期間の賃金台帳を'
                '再提出してもらえないか相談のうえ、手動で値を入力してください'
            )
        # 整合性チェック: 賃金台帳合計と損益計算書の人件費の差が大きいと AI 抽出ミスの疑い
        consistency_warning = _check_wage_pl_consistency(wage_plan, extraction.financial)
        # 会計式整合（売上 − 原価 = 粗利）— ()書きマイナス誤読の自動検出
        pl_accounting_warning = _check_pl_accounting_consistency(extraction.financial)
        # 業種コードのフォーマット検証（旧3桁体系・自己流コード検出）
        industry_code_warning = _check_industry_code_format(extraction.ai_judgment)
        # 賃金台帳抽出結果の自動品質検証（人数妥当性・月別カバレッジ・値分布・賞与未参照）
        from .wage_validator import run_all_validations
        validation_warnings = ''.join(
            run_all_validations(hearing_data, ledger_employees, extraction.financial)
        )
        # API残高切れで Phase 2 がスキップされた場合は冒頭にお知らせを追加
        api_skip_msg = ''
        if api_skipped_reason:
            api_skip_msg = (
                f'⛔ APIエラーでAI抽出部分がスキップされました ({api_skipped_reason})。'
                f'ヒアリングシート・賃金台帳など API不要部分のみ転記しました。'
                f'残高チャージ後に再実行すると AI 部分も埋まります。 '
            )
        status.message = (
            api_skip_msg
            + f'完了。空欄{len(empty_cells)}件{wage_warning}{consistency_warning}'
            + validation_warnings
            + fiscal_month_warning
            + pl_period_warning
            + cost_report_warning
            + biz_desc_warning
            + wage_pdf_warning
            + pl_accounting_warning
            + industry_code_warning
        )
        logger.info(f'申請書作成完了: {output_path.name} (空欄{len(empty_cells)}件{wage_warning})')

        # 賃金台帳一覧Excel出力（チェック用）— AI抽出結果をそのまま再利用してAPI呼出しの2重化を防ぐ
        if ledger_employees:
            company = output_path.stem.split('_')[0]
            ledger_output = output_path.parent / f'{company}_賃金台帳一覧.xlsx'
            export_wage_ledger_summary(
                ledger_employees, ledger_output, company,
                extraction_method=wage_extraction_method,
            )
            status.output_files.append(ledger_output.name)

    except Exception as e:
        # 通常の例外（ファイル不在等）: ステータスをエラーに
        # API残高切れは Phase 2 内で個別ハンドル済みなのでここには到達しない
        status.status = 'エラー'
        status.message = str(e)
        logger.error(f'エラー: {e}', exc_info=True)

    return status


def run_wage_calculation(
    resource_folder: Path,
    company_name: str,
    output_path: Path,
    extractor: BaseExtractor | None = None,
    cached_financial: 'FinancialData | None' = None,
    cached_ledger_employees: list | None = None,
    fiscal_month_override: int | None = None,
    cached_company: 'CompanyInfo | None' = None,
) -> ProcessingStatus:
    """
    タスク2: 給与支給総額計算の実行

    cached_financial / cached_ledger_employees が渡された場合は API 呼出を省略する
    （申請書作成タスクの結果を再利用してコスト2重化を防ぐ）。

    fiscal_month_override (1〜12) が指定された場合、ユーザー指定の決算月で
    賃金台帳の対象期間（直近12ヶ月）を確定する。AI 推定とズレていれば警告。
    """
    status = ProcessingStatus(
        company_name=company_name,
        template_type='給与計算',
        status='処理中',
    )

    try:
        if extractor is None:
            extractor = create_extractor(CLAUDE_API_KEY)

        detector = FileDetector(resource_folder)
        logger.info(detector.summary())

        # 損益計算書（任意: あれば精度向上）— キャッシュがあれば再利用
        financial = cached_financial
        pl_period_warning = ''
        if financial is None:
            pl_path = detector.get_pl_latest(fiscal_month_override=fiscal_month_override)
            if pl_path:
                logger.info(f'直近期決算書として採用: {pl_path.name}')
                images = pdf_to_images(pl_path)
                financial = extractor.extract_pl(images)
                # 賃金台帳期間 vs PL期末 / ユーザー指定決算月の整合性チェック
                financial, pl_period_warning = _check_pl_wage_period_consistency(
                    detector, financial, fiscal_month_override,
                )
        else:
            logger.info('PL: 申請書作成タスクの結果を再利用（API呼出スキップ）')
            # 申請書作成タスクで既にチェック済みだが、キャッシュ経路でも再確認
            financial, pl_period_warning = _check_pl_wage_period_consistency(
                detector, financial, fiscal_month_override,
            )

        if financial is None or financial.revenue == 0:
            from .models import FinancialData
            if financial is None:
                financial = FinancialData()
            logger.info('損益計算書なし → 賃金台帳ベースで計算')

        # 賃金状況報告シートから従業員データ読取（あれば）
        employees_detail = None
        seishain_count = 0
        part_count = 0
        yakuin_hoshu_3m = 0

        # 役員数の取得（履歴事項証明書から動的に決定。固定値ハードコードを廃止）
        # 申請書作成タスクから cached_company が渡ってきていればそれを優先。
        # 無ければ resource_folder 内の履歴事項証明書を AI 抽出する（API 1回）。
        # 全く取れない場合のフォールバックは1人（代表取締役を最低限想定）。
        #
        # 注意: AI プロンプト（ai_extractor.py:287）で「代表者は officers に含めない」
        # と明示しており、CompanyInfo.officers は代表取締役を除いたリストになる。
        # よって役員総数 = 1（代表者）+ len(officers)。template_filler.py:121 と同じ規約。
        from .ai_extractor import APICreditExhaustedError
        yakuin_count = 1
        company = cached_company
        if company is None:
            registry_path = detector.get('registry')
            if registry_path:
                try:
                    images = pdf_to_images(registry_path)
                    company = extractor.extract_registry(images)
                    logger.info(
                        f'履歴事項: {company.name}（officers={len(company.officers)}名）'
                    )
                except APICreditExhaustedError:
                    # 残高切れは他の API 呼び出しも全部失敗するので即停止
                    raise
                except Exception as e:
                    logger.warning(f'履歴事項証明書の AI 抽出に失敗: {e}')
                    company = None
        if company is not None:
            # 代表者(+1) + 取締役・監査役等(officers リスト)
            yakuin_count = 1 + len(company.officers)
            logger.info(
                f'役員数: {yakuin_count}名（代表者1 + その他役員{len(company.officers)}名）'
            )
        else:
            logger.info('役員数: 1名（履歴事項証明書なし or 取得失敗のためフォールバック）')

        wage_report_path = detector.get('wage_report')
        if wage_report_path:
            employees_detail, seishain_count, part_count, yakuin_hoshu_3m = \
                _read_wage_report(wage_report_path)
            logger.info(f'賃金状況報告シート: 正社員{seishain_count}, パート{part_count}')

        # 賃金状況報告シート未取込 + PL に役員報酬がある → PL の役員報酬 ÷ 4 で機械計算
        # （正確な3ヶ月実績ではないが、年間役員報酬を均等割した推定値として埋めておく）
        if yakuin_hoshu_3m == 0 and financial.officer_compensation > 0:
            yakuin_hoshu_3m = int(financial.officer_compensation / 4)
            logger.info(
                f'役員報酬3ヶ月合計を PL から推定: {yakuin_hoshu_3m:,}円 '
                f'(年額{financial.officer_compensation:,}円 ÷ 4)'
            )

        # フォールバック: 賃金状況報告シートで人数が取れなかった場合、賃金台帳から補完
        if seishain_count + part_count == 0:
            # キャッシュがあれば再利用（API呼出スキップ）
            ledger_emps = cached_ledger_employees
            if ledger_emps:
                logger.info(f'賃金台帳: 申請書作成タスクの結果を再利用（API呼出スキップ、{len(ledger_emps)}名）')
            else:
                ledger_paths = detector.get_all('wage_ledger')
                if ledger_paths:
                    fiscal_hint, _ = _resolve_fiscal_period(financial, fiscal_month_override)
                    ledger_emps = read_wage_ledgers(
                        ledger_paths,
                        extractor=extractor,
                        fiscal_period_hint=fiscal_hint,
                    )
            if ledger_emps:
                # fiscal_hint を渡して時系列順で直近3ヶ月を抽出（非1月始まり対応）
                _fiscal_hint_for_detail, _ = _resolve_fiscal_period(financial, fiscal_month_override)
                employees_detail = _build_employees_detail_from_ledger(
                    ledger_emps, fiscal_period_hint=_fiscal_hint_for_detail,
                )
                # 「契約社員」は正規雇用相当として seishain_count にカウント
                # （wage_calculator.is_full_time_employment と整合）
                from .wage_calculator import is_full_time_employment
                seishain_count = sum(
                    1 for e in employees_detail if is_full_time_employment(e['type'])
                )
                part_count = sum(
                    1 for e in employees_detail
                    if e['type'] in ('パート・アルバイト',)
                )
                logger.info(
                    f'賃金台帳フォールバック: 正社員{seishain_count}, '
                    f'パート・契約{part_count}'
                )

        # 給与データPDFから読取（APIが必要）
        wage_pdfs = detector.get_all('wage_data')
        if wage_pdfs and not employees_detail:
            # PDFからの読取はAPI必須
            wages_list = []
            for wp in sorted(wage_pdfs):
                images = pdf_to_images(wp)
                wages = extractor.extract_wages(images, wp.stem)
                wages_list.append(wages)
            # TODO: wages_listからemployees_detailを構築

        # 表示用の期間ラベル。ユーザーが決算月を指定 + AI 推定と不一致なら
        # override 反映後の期間を表示（賃金計算と帳票表示の整合を取る）
        _resolved_period, _ = _resolve_fiscal_period(financial, fiscal_month_override)
        if fiscal_month_override is not None and _resolved_period and '〜' in _resolved_period:
            fiscal_label = _resolved_period
        else:
            fiscal_label = f'{financial.fiscal_year_start} ～ {financial.fiscal_year_end}'

        # データソース（ファイル名）— 人間チェックの突合用に各セクションに表示する
        _pl_path = detector.get_pl_latest(fiscal_month_override=fiscal_month_override)
        _ledger_paths = detector.get_all('wage_ledger')
        _wage_report = detector.get('wage_report')
        _registry = detector.get('registry')

        # PL の各値が決算書PDFのどのページに記載されているかを機械的に逆引き
        # （AI抽出結果の検証 + 人間チェック時のページ番号案内に兼用）
        # PDFテキスト層が無い画像PDFの場合は空辞書になる
        pl_value_pages: dict[str, list[int]] = {}
        if _pl_path and financial is not None:
            try:
                from .pdf_text_extractor import get_pdf_pages_text, find_value_pages
                _pages = get_pdf_pages_text(_pl_path)
                if _pages and any(p.strip() for p in _pages):
                    for key in ('salary', 'misc_wages', 'bonus', 'legal_welfare',
                                'welfare', 'officer_compensation', 'revenue',
                                'gross_profit', 'operating_profit', 'ordinary_profit',
                                'depreciation'):
                        val = getattr(financial, key, 0)
                        # 0 円の科目は決算書に記載がないことが多いので検証スキップ
                        # （AI誤読ではなく「該当する経費が無い」が正しいケースが大半）
                        if not val:
                            continue
                        pl_value_pages[key] = find_value_pages(_pages, val)
                    logger.info(
                        f'PL 値→ページ逆引き完了: '
                        f'{ {k: v for k, v in pl_value_pages.items() if v} }'
                    )
                else:
                    logger.info(
                        '決算書PDFのテキスト層が取得できなかったため、ページ番号特定はスキップ'
                    )
            except Exception as e:
                logger.warning(f'PL ページ番号逆引きに失敗: {e}')

        source_files = {
            'pl': _pl_path.name if _pl_path else '',
            'wage_ledger': (
                _ledger_paths[0].name if len(_ledger_paths) == 1
                else f'{_ledger_paths[0].name} 他 {len(_ledger_paths) - 1} 件'
                if _ledger_paths else ''
            ),
            'wage_report': _wage_report.name if _wage_report else '',
            'registry': _registry.name if _registry else '',
            'pl_value_pages': pl_value_pages,
        }

        create_wage_calculation(
            output_path=output_path,
            company_name=company_name,
            fiscal_year_label=fiscal_label,
            financial=financial,
            seishain_count=seishain_count,
            part_count=part_count,
            yakuin_count=yakuin_count,
            yakuin_hoshu_3m=yakuin_hoshu_3m,
            employees_detail=employees_detail,
            source_files=source_files,
        )

        # ユーザー指定の決算月 vs AI 推定の照合（警告のみ）
        _, fiscal_month_warning = _resolve_fiscal_period(financial, fiscal_month_override)

        status.status = '完了'
        status.output_files = [output_path.name]
        status.message = '給与支給総額計算 完了' + fiscal_month_warning + pl_period_warning
        logger.info(f'給与計算完了: {output_path.name}')

    except Exception as e:
        from .ai_extractor import APICreditExhaustedError
        if isinstance(e, APICreditExhaustedError):
            status.status = 'エラー'
            status.message = (
                f'⛔ {e}\n'
                f'処理を中断しました。チャージ後にもう一度実行してください。'
            )
            logger.error(f'API残高切れ: 給与計算中断 ({e})')
        else:
            status.status = 'エラー'
            status.message = str(e)
            logger.error(f'エラー: {e}', exc_info=True)

    return status


def _format_fiscal_period(financial: 'FinancialData') -> str | None:
    """FinancialData の fiscal_year_start/end から AI 用ヒント文字列を組み立てる。

    例: '2024-05-01' / '2025-04-30' → '2024-05〜2025-04'
    """
    start = (financial.fiscal_year_start or '').strip()
    end = (financial.fiscal_year_end or '').strip()
    if not start and not end:
        return None
    # YYYY-MM-DD → YYYY-MM
    def _ym(s: str) -> str:
        if len(s) >= 7 and s[4] == '-':
            return s[:7]
        return s
    s_ym = _ym(start)
    e_ym = _ym(end)
    if s_ym and e_ym:
        return f'{s_ym}〜{e_ym}'
    return s_ym or e_ym


def _guess_recent_fiscal_end_year(fiscal_month: int) -> int:
    """指定された決算月から、今日基準で「直近の確定済み決算期末年」を推定する。

    判定は安全側に倒す: 決算月そのもの の月内は **まだ確定していない**とみなし前年扱い。
    （実際の月末日は 28/30/31 の差があり、決算実務でも申告までは「直近期＝先期」と
    扱うのが普通）

    例:
      今日=2026-05-14, fiscal_month=3 → 2026（2026-03 はすでに過ぎている）
      今日=2026-05-14, fiscal_month=5 → 2025（2026-05 はまだ進行中）
      今日=2026-05-14, fiscal_month=6 → 2025（2026-06 はまだ来てない）
    """
    from datetime import date
    today = date.today()
    # 今日の月が決算月より大きい場合のみ「直近期末は今年」と判定。
    # 同月の場合は決算月内＝進行中なので、前年扱いに倒す。
    if today.month > fiscal_month:
        return today.year
    return today.year - 1


def _resolve_fiscal_period(
    financial: 'FinancialData',
    fiscal_month_override: int | None = None,
) -> tuple[str | None, str]:
    """fiscal_period_hint を解決する。

    fiscal_month_override（1〜12）が指定されている場合：
      - その月を期末月としてヒント文字列を再構築（ユーザー指定優先）
      - 期末年は financial.fiscal_year_end が取れていればその年、なければ今日から推定
      - AI 推定の期末月と override がズレていれば警告メッセージを返す

    fiscal_month_override が None の場合：
      - 従来通り AI 抽出の fiscal_year_start/end から組み立てる

    Returns:
        (fiscal_period_hint, warning_message)
          - fiscal_period_hint: '2024-05〜2025-04' 形式 or None
          - warning_message: 不一致警告。一致時 or 未指定時は ''
    """
    warning = ''
    if not fiscal_month_override:
        return _format_fiscal_period(financial), warning

    # 期末年の決定:
    #   - AI 推定の月が override と一致 → AI の year を採用（AI が信頼できる）
    #   - AI 推定の月が override と不一致 → AI の year も誤読の可能性が高いので
    #     今日基準で推定し直す（例: AI=2026-01 のはずが実は 2025-12 だった場合、
    #     AI year 2026 をそのまま流用すると未来の 2026-12 を生成してしまう）
    #   - AI 推定なし → 今日基準で推定
    end_str = (financial.fiscal_year_end or '').strip()
    ai_month: int | None = None
    ai_year: int | None = None
    if end_str and len(end_str) >= 7 and '-' in end_str:
        try:
            ai_year = int(end_str.split('-')[0])
            ai_month = int(end_str.split('-')[1])
        except (ValueError, IndexError):
            ai_year = None
            ai_month = None

    if ai_month is not None and ai_month == fiscal_month_override and ai_year is not None:
        # 月が一致するなら AI year を信用
        end_year = ai_year
    else:
        # 月が不一致 or AI 推定なし → 今日基準で推定
        end_year = _guess_recent_fiscal_end_year(fiscal_month_override)

    # AI 推定月と override の照合
    if ai_month is not None and ai_month != fiscal_month_override:
        warning = (
            f' ⚠ 決算月の不一致: ユーザー指定={fiscal_month_override}月 / '
            f'AI推定={ai_month}月。決算書PDFを目視確認してください'
            f'（ユーザー指定値を優先しました）。'
        )

    # 期首 = 期末の翌月 - 12ヶ月前
    end_ym = f'{end_year:04d}-{fiscal_month_override:02d}'
    if fiscal_month_override == 12:
        start_year = end_year
        start_month = 1
    else:
        start_year = end_year - 1
        start_month = fiscal_month_override + 1
    start_ym = f'{start_year:04d}-{start_month:02d}'

    return f'{start_ym}〜{end_ym}', warning


def _calc_wage_plan_from_ledger(
    detector: FileDetector,
    financial: 'FinancialData',
    extractor=None,
    fiscal_month_override: int | None = None,
) -> tuple[dict[str, float] | None, list, str]:
    """
    賃金台帳から給与支給総額を算出し、年3%成長の計画値を返す。

    extractor が渡されると AI 抽出を優先する（USE_AI_WAGE_EXTRACTION=true 時）。
    AI失敗時は決定論パーサーにフォールバック。

    fiscal_month_override (1〜12) が指定された場合、ユーザー指定の決算月を
    優先して fiscal_period_hint を構築する。AI 推定とズレていれば警告ログ。

    Returns:
        (plan_dict_or_None, employees_raw_list, status_message)
        status_message:
          - '': 正常
          - 'no_ledger': 賃金台帳なし
          - 'no_data': 賃金台帳はあるがデータ抽出失敗
          - 'zero_total': 給与支給総額が0以下
          - 'error': 例外発生
    """
    from .wage_reader import read_wage_ledgers

    ledger_paths = detector.get_all('wage_ledger')
    if not ledger_paths:
        logger.info('賃金台帳が見つかりません → 計画値転記をスキップ')
        return None, [], 'no_ledger'

    fiscal_hint, fiscal_warning = _resolve_fiscal_period(financial, fiscal_month_override)
    if fiscal_warning:
        logger.warning(f'決算月の不一致警告: {fiscal_warning.strip()}')

    try:
        employees_raw = read_wage_ledgers(
            ledger_paths,
            extractor=extractor,
            fiscal_period_hint=fiscal_hint,
        )
        if not employees_raw:
            logger.warning(
                f'賃金台帳からデータを読み取れませんでした '
                f'(ファイル: {[p.name for p in ledger_paths]}, '
                f'fiscal_hint={fiscal_hint})'
            )
            return None, [], 'no_data'

        logger.info(f'賃金台帳: {len(employees_raw)}名読取 ({len(ledger_paths)}ファイル)')

        # WageEmployee → PayrollEmployee に変換
        payroll_list = []
        total_annual_hours = 0.0
        for emp in employees_raw:
            is_officer = '役員' in emp.employment_type
            emp_type = emp.employment_type if emp.employment_type else '正社員'

            # 全月分の給与を受けたか判定
            full_year = emp.is_full_year

            monthly_salary = [
                w if w is not None else 0.0 for w in emp.monthly_wages
            ]

            # 労働時間: 月別実績データがあればそれを優先。なければ月平均で補完
            has_monthly_hours = any(
                h is not None and h > 0 for h in emp.monthly_hours
            )
            if has_monthly_hours:
                monthly_hours = [
                    h if (h is not None and h > 0) else 0.0
                    for h in emp.monthly_hours
                ]
            elif emp.monthly_avg_hours > 0:
                # 月別データが取れないフォーマットは、在籍月数×月平均で概算
                months_with_wage = sum(
                    1 for w in emp.monthly_wages if w is not None
                )
                months = months_with_wage if months_with_wage > 0 else 12
                monthly_hours = [emp.monthly_avg_hours] * months + [0.0] * (12 - months)
            else:
                monthly_hours = []

            payroll_list.append(PayrollEmployee(
                name=emp.name,
                employment_type=emp_type,
                monthly_salary=monthly_salary,
                monthly_hours=monthly_hours,
                is_officer=is_officer,
                full_year=full_year,
            ))

            # 役員を除く全従業員の年間総労働時間を集計
            if not is_officer and monthly_hours:
                total_annual_hours += sum(monthly_hours)

        result = calculate_per_capita_wage(payroll_list)

        if result.total_salary <= 0:
            logger.warning('給与支給総額が0以下 → 計画値転記をスキップ')
            return None, employees_raw, 'zero_total'

        # ── 直近事業年度との整合性チェック (hard stop) ─────────────────
        # 賃金台帳に複数名の記録があるのに、12スロット全埋まり (full_year=True) と
        # 判定される従業員が極端に少ない場合は、賃金台帳の記録期間と直近事業年度が
        # ズレている疑いが濃い (例: 賃金台帳が決算期より新しい月だけ記録、または
        # 決算期内に在籍した人が事実上いないなど)。
        #
        # 公募要領上、給与支給総額は「直近事業年度に全月分の給与支給を受けた従業員」
        # を分子・分母とも算出対象とする。賃金台帳の任意12ヶ月で full_year を判定する
        # 現行ロジックは決算期との整合を保証できないため、対象人数が会社規模と乖離した
        # ケースでは自動転記をスキップし、ユーザに手動入力を促す方が安全。
        # (将来対応: WageEmployee に YYYY-MM 情報を保持して決算期フィルタを
        #  かけてから full_year 判定する根本修正が必要)
        non_officer_count = sum(1 for p in payroll_list if not p.is_officer)
        included_count = len(result.included)
        FISCAL_MISMATCH_RATIO = 0.5  # 全月在籍者が非役員数の50%未満なら乖離扱い
        if non_officer_count >= 2 and included_count < non_officer_count * FISCAL_MISMATCH_RATIO:
            excluded_n = non_officer_count - included_count
            logger.warning(
                f'賃金台帳の全月在籍者({included_count}名)が会社規模({non_officer_count}名)と乖離。'
                f'{excluded_n}名が中途入退社扱いで除外されました。'
                f'賃金台帳の記録期間と直近事業年度がズレている疑いがあるため、'
                f'R215/R216 の自動転記をスキップします（手動入力が必要）'
            )
            return None, employees_raw, 'fiscal_year_mismatch'

        # 給与支給総額ベースで年3%成長の計画値を算出
        base = result.total_salary
        rate = result.GROWTH_RATE
        plan = {
            'employee_count_fte': result.employee_count_fte,
            'wage_total_base': base,
            'wage_total_y1': base * (1 + rate),
            'wage_total_y2': base * (1 + rate) ** 2,
            'wage_total_y3': base * (1 + rate) ** 3,
        }
        if total_annual_hours > 0:
            plan['total_annual_hours'] = round(total_annual_hours, 1)
        logger.info(
            f'給与支給総額: {base:,.0f}円 '
            f'(従業員FTE: {result.employee_count_fte:.1f}人, 年3%成長, '
            f'総労働時間: {total_annual_hours:,.0f}時間)'
        )
        return plan, employees_raw, ''

    except Exception as e:
        # API残高切れは pipeline で全体停止する必要があるので再 raise
        from .ai_extractor import APICreditExhaustedError
        if isinstance(e, APICreditExhaustedError):
            raise
        logger.warning(f'賃金台帳処理エラー（申請書作成は続行）: {e}', exc_info=True)
        return None, [], 'error'


def _build_confidence_warnings(financial) -> list[dict]:
    """FinancialData.confidence から「確認キュー」用の警告リストを構築（Phase 4）。

    UI で「📋 確認キュー」セクションに表示する：項目・元値・根拠・警告理由。
    level='low' のフィールドだけを抽出（high/medium はスルー）。
    """
    if financial is None:
        return []
    conf = getattr(financial, 'confidence', None) or {}
    if not conf:
        return []
    # フィールド名 → 表示ラベル
    label_map = {
        'fiscal_year_start': '事業年度開始日',
        'fiscal_year_end': '事業年度終了日',
        'revenue': '売上高',
        'cost_of_sales': '売上原価',
        'gross_profit': '売上総利益',
        'operating_profit': '営業利益',
        'ordinary_profit': '経常利益',
        'net_profit': '当期純利益',
        'salary': '給料手当',
        'misc_wages': '雑給',
        'bonus': '賞与',
        'officer_compensation': '役員報酬',
        'legal_welfare': '法定福利費',
        'welfare': '福利厚生費',
        'depreciation': '減価償却費',
        'travel_expense': '旅費交通費',
    }
    # source_component → 表示ラベル
    source_map = {
        'basic': '損益計算書本表',
        'PL': '販管費明細',
        'cost': '原価報告書',
        'PL+cost': '販管費 + 原価報告書',
        'unknown': '不明',
    }
    warnings = []
    for field, c in conf.items():
        if not c or getattr(c, 'level', 'high') != 'low':
            continue
        value = getattr(financial, field, None)
        warnings.append({
            'field': field,
            'label': label_map.get(field, field),
            'value': str(value) if value is not None else '(空)',
            'source': source_map.get(getattr(c, 'source_component', ''), getattr(c, 'source_component', '')),
            'reason': getattr(c, 'reason', '') or '抽出失敗',
        })
    return warnings


def _check_industry_code_format(ai_judgment) -> str:
    """AI 生成の業種コードが日本標準産業分類（令和5年6月改定）の細分類4桁形式かチェック。

    旧分類（3桁体系）や AI の自己流コードを検出して警告する。
    プロンプトを強化してもなお古いコードが返るケースを救う。
    AI が int を返してきたケースも考慮（str 化 + NFKC 正規化）。
    """
    if ai_judgment is None:
        return ''
    raw = getattr(ai_judgment, 'industry_code', '')
    if raw is None or raw == '':
        return ''
    # int で返ってきた場合や全角数字対策のため str 化 + NFKC 正規化
    code = unicodedata.normalize('NFKC', str(raw)).strip()
    if not code:
        return ''
    # 細分類は ASCII 半角の4桁数字（NFKC 後は全角数字 → 半角に揃う）
    if not (len(code) == 4 and code.isascii() and code.isdigit()):
        return (
            f' ⚠ 業種コード「{raw}」が日本標準産業分類（令和5年6月改定）の'
            f'細分類4桁形式と異なります。e-Statで再確認してください。'
        )
    return ''


def _check_pl_accounting_consistency(financial) -> str:
    """損益計算書の会計式整合（売上高 − 売上原価 = 売上総利益）をチェック。

    AI が決算書の `(1,234)` `△1,234` `▲1,234` をマイナスとして読み損ねたケースを
    機械的に検出する。整合 or 比較不能なら空文字列、不整合なら警告文字列を返す。

    判定方針:
      - 売上高・売上原価・売上総利益のいずれも 0/None なら判定不能（''）
      - 売上の 0.5% 以内の差は端数誤差として許容
      - それ以上の差は AI の符号読み違いを疑って警告
    """
    if financial is None:
        return ''
    revenue = financial.revenue or 0
    cost = financial.cost_of_sales or 0
    gross = financial.gross_profit or 0
    if revenue <= 0 and cost == 0 and gross == 0:
        return ''
    expected_gross = revenue - cost
    diff = abs(expected_gross - gross)
    diff_ratio = diff / revenue if revenue > 0 else (1.0 if diff > 0 else 0.0)
    if diff_ratio < 0.005:
        return ''
    return (
        f' ⚠ 決算書の会計式不整合: '
        f'売上高({revenue:,}) − 売上原価({cost:,}) = {expected_gross:,} のはずですが、'
        f'抽出された売上総利益は {gross:,}（差 {diff:,}）。'
        f'AIが括弧書きや△/▲記号をマイナスとして読み損ねている可能性があります。'
        f'決算書PDFを目視確認してください。'
    )


def _check_wage_pl_consistency(
    wage_plan: dict | None,
    financial: 'FinancialData',
    tolerance: float = 0.20,
    severe_threshold: float = 0.50,
) -> str:
    """賃金台帳合計と損益計算書の人件費を照合し、不整合があれば警告文字列を返す。

    Args:
        wage_plan: _calc_wage_plan_from_ledger の戻り値（基準年の給与支給総額を含む）
        financial: 損益計算書から抽出した財務データ
        tolerance: 通常警告の許容差（デフォルト ±20%）
        severe_threshold: 強警告に格上げする差分比（デフォルト 50%）

    Returns:
        警告文字列（不整合あり）または空文字列（整合 / 比較不能）

    判定方針:
        - wage_plan['wage_total_base'] は wage_calculator.calculate_per_capita_wage で
          **役員報酬を除外** して算出されている（[wage_calculator.py:49] 参照）。
          そのため PL 側からも officer_compensation を除いて比較する必要がある。
        - 比較対象 = PL の (給料手当 + 雑給 + 賞与) ※役員報酬除く
        - PL に給与系の計上が無い場合は判定不能 → '' を返す
        - 差が severe_threshold を超えたら強警告（⛔ + 「賃金台帳が著しく〜」）
        - 差が tolerance を超えたら通常警告（⚠ + 「〜の差があります」）
        - いずれも処理は続行する
    """
    if not wage_plan or 'wage_total_base' not in wage_plan:
        return ''
    # 役員報酬は除外（賃金台帳側も役員除外で集計しているため）
    pl_personnel = (
        (financial.salary or 0)
        + (financial.misc_wages or 0)
        + (financial.bonus or 0)
    )
    if pl_personnel <= 0:
        return ''  # PL データなし → 比較不能
    ledger_total = wage_plan['wage_total_base']
    if ledger_total <= 0:
        return ''
    diff_ratio = abs(ledger_total - pl_personnel) / pl_personnel
    if diff_ratio <= tolerance:
        return ''
    direction = '多い' if ledger_total > pl_personnel else '少ない'
    if diff_ratio >= severe_threshold:
        # 著しい乖離 — AI 抽出の打ち切り / 雇用区分誤認 / 中途退職者多数 等の可能性大
        return (
            f' ⛔ 強警告: 賃金台帳合計({ledger_total:,.0f}円)と損益計算書の人件費'
            f'({pl_personnel:,.0f}円, 役員報酬除く)に{diff_ratio*100:.0f}%の著しい差があります'
            f'（賃金台帳が著しく{direction}）。賃金台帳の抽出漏れがないか必ず確認してください'
        )
    return (
        f' ⚠ 賃金台帳合計({ledger_total:,.0f}円)と損益計算書の人件費'
        f'({pl_personnel:,.0f}円, 役員報酬除く)に{diff_ratio*100:.0f}%の差があります'
        f'（賃金台帳が{direction}）。AI抽出が月数や雇用区分を取り違えていないか確認してください'
    )


def _classify_emp_type(emp_type: str) -> str:
    """賃金台帳の雇用形態文字列を4分類に正規化"""
    t = (emp_type or '').strip()
    if '役員' in t or '取締役' in t:
        return '役員'
    if 'パート' in t or 'アルバイト' in t or '非常勤' in t:
        return 'パート・アルバイト'
    if '契約' in t:
        return '契約社員'
    # 雇用形態が空/不明な場合は正社員として集計（台帳に区分列が無い一般的なケース）
    return '正社員'


def _fiscal_month_order(fiscal_period_hint: str | None) -> list[int]:
    """fiscal_period_hint から「事業年度内の月の時系列順 Index リスト」を返す。

    例: '2024-05〜2025-04' → [4,5,6,7,8,9,10,11,0,1,2,3]（5月始まり、12月の次は1月）
    None や形式不明な場合は [0,1,...,11]（カレンダー順 = 1月始まり）にフォールバック。
    """
    if not fiscal_period_hint:
        return list(range(12))
    # 'YYYY-MM〜YYYY-MM' / 'YYYY-MM-DD〜...' / 'YYYY-MM' から開始月を抽出
    m = re.search(r'(\d{4})-(\d{1,2})', fiscal_period_hint)
    if not m:
        return list(range(12))
    start_month = int(m.group(2))  # 1〜12
    if not 1 <= start_month <= 12:
        return list(range(12))
    start_idx = start_month - 1  # 0〜11
    return [(start_idx + i) % 12 for i in range(12)]


def _build_employees_detail_from_ledger(
    employees,
    fiscal_period_hint: str | None = None,
) -> list[dict]:
    """賃金台帳の読取結果 → create_wage_calculation が期待する employees_detail 形式に変換。
    役員はカウント対象外として除外する。

    直近3ヶ月の判定方針:
        fiscal_period_hint があれば「事業年度内の月の時系列順」で末尾3つを採用。
        例: 5月開始の年度なら 5,6,7,...,2,3,4 の順序で並べた中の最後3つ。
        fiscal_period_hint がなければカレンダー順末尾3つ（フォールバック、従来動作）。
        この区別が無いと、12月始まりの台帳で「直近3 = 配列末尾の10/11/12月」になり、
        実際の時系列上の直近（9/10/11月）とズレる。
    """
    detail = []
    month_order = _fiscal_month_order(fiscal_period_hint)
    for emp in employees:
        classified = _classify_emp_type(emp.employment_type)
        if classified == '役員':
            continue

        # 事業年度内の時系列順で「データのある月」を取り、末尾3つを直近として採用
        ordered_months_with_data = [
            idx for idx in month_order
            if idx < len(emp.monthly_wages) and emp.monthly_wages[idx] is not None
        ]
        tenure_months = len(ordered_months_with_data)
        full_year = tenure_months >= 12

        last_three = ordered_months_with_data[-3:]
        m_vals = [emp.monthly_wages[m] or 0 for m in last_three]
        # 中途者向け: 表示列に対応する暦月ラベル（実体が分かるように）
        last_three_labels = [f'{m + 1}月' for m in last_three]
        while len(m_vals) < 3:
            m_vals.append(0)
            last_three_labels.append('')

        detail.append({
            'no': len(detail) + 1,
            'name': emp.name,
            'type': classified,
            'm1': m_vals[0],
            'm2': m_vals[1],
            'm3': m_vals[2],
            'hr': emp.hourly_rate,
            'monthly_hours': emp.monthly_avg_hours,
            'judge': '',
            # 中途入退社の扱いを正しくするための追加情報
            'tenure_months': tenure_months,
            'full_year': full_year,
            'last_three_labels': last_three_labels,
        })
    return detail


def _read_wage_report(path: Path) -> tuple[list[dict], int, int, int]:
    """
    賃金状況報告シートから従業員データを読取。
    Returns: (employees_detail, seishain_count, part_count, yakuin_hoshu_3m)
    """
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)

    # シート名を探す
    ws = None
    for name in wb.sheetnames:
        if '賃金' in name and 'マスタ' not in name and '元データ' not in name:
            ws = wb[name]
            break
    if ws is None:
        ws = wb[wb.sheetnames[0]]

    def _to_num(v) -> float:
        # 賃金状況報告シートのセルは数値想定だが、ユーザー入力で
        # '300,000円' / '― ' / 'N/A' / '' のような文字列が混じることがあり
        # そのまま比較すると TypeError: '>' not supported between str and int で落ちる。
        if v is None:
            return 0.0
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).replace(',', '').replace(' ', '').replace('円', '').replace('時間', '').strip()
        try:
            return float(s)
        except (ValueError, TypeError):
            return 0.0

    # 役員報酬（行13, D列）
    yakuin_hoshu_3m = _to_num(ws.cell(13, 4).value)

    employees = []
    for row in ws.iter_rows(min_row=19, max_row=200):
        no = row[1].value
        name = row[2].value
        if name is None or no is None:
            continue

        m1_base = _to_num(row[5].value)
        m1_hr = _to_num(row[6].value)
        m2_base = _to_num(row[8].value)
        m2_hr = _to_num(row[9].value)
        m3_base = _to_num(row[11].value)
        m3_hr = _to_num(row[12].value)
        judge = row[14].value if len(row) > 14 else ''

        # 時間推定
        hours = []
        for base, hr in [(m1_base, m1_hr), (m2_base, m2_hr), (m3_base, m3_hr)]:
            if hr > 0 and base > 0:
                hours.append(base / hr)
        avg_hours = sum(hours) / len(hours) if hours else 0

        # 正社員/パート判定: 時給1300円以上 and 月給18万以上 → 正社員の傾向
        avg_base = (m1_base + m2_base + m3_base) / 3
        avg_hr = (m1_hr + m2_hr + m3_hr) / 3
        emp_type = '正社員' if avg_base >= 180000 and avg_hr >= 1200 else 'パート・アルバイト'

        employees.append({
            'no': no,
            'name': str(name).strip(),
            'type': emp_type,
            'm1': m1_base,
            'm2': m2_base,
            'm3': m3_base,
            'hr': round(avg_hr),
            'monthly_hours': round(avg_hours, 1),
            'judge': judge or '',
            # 賃金状況報告シート由来は在籍中の社員のみが載る前提のため、
            # 全員 12ヶ月在籍扱いとする（賃金台帳の中途入退社検出とは別系統）
            'tenure_months': 12,
            'full_year': True,
            'last_three_labels': ['', '', ''],
        })

    wb.close()

    seishain = [e for e in employees if e['type'] == '正社員']
    part = [e for e in employees if e['type'] != '正社員']
    return employees, len(seishain), len(part), yakuin_hoshu_3m


def run_full_pipeline(
    resource_folder: Path,
    template_path: Path,
    template_type: str,
    company_name: str,
    fiscal_month_override: int | None = None,
) -> list[ProcessingStatus]:
    """タスク1 + タスク2 を一括実行"""
    extractor = create_extractor(CLAUDE_API_KEY)
    results = []

    # タスク1: 申請書
    output_app = resource_folder / f'{company_name}_{template_type.replace("_", "_")}_AI版.xlsx'
    s1 = run_application_transfer(
        resource_folder, template_path, template_type, output_app, extractor,
        fiscal_month_override=fiscal_month_override,
    )
    results.append(s1)

    # タスク2: 給与計算
    output_wage = resource_folder / f'{company_name}_給与支給総額計算.xlsx'
    s2 = run_wage_calculation(
        resource_folder, company_name, output_wage, extractor,
        fiscal_month_override=fiscal_month_override,
    )
    results.append(s2)

    return results
