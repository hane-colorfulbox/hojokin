# -*- coding: utf-8 -*-
"""処理パイプライン: ファイル検出 → 抽出 → 転記 → 出力"""
from __future__ import annotations

import logging
import tempfile
import unicodedata
from pathlib import Path

from .config import get_mapping, CLAUDE_API_KEY
from .models import ExtractionResult, ProcessingStatus
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
        'wage_ledger': ['賃金台帳'],
        'wage_data': ['支給控除一覧', '給与データ'],
    }

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
        'wage_ledger': {'.xlsx', '.xlsm', '.pdf'},
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
                if p.is_dir() and not p.name.startswith('.'):
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

    def get_pl_latest(self) -> Path | None:
        """損益計算書の直近期を返す（第2期 > 第1期）"""
        pls = self.files.get('pl', [])
        if not pls:
            return None
        # ファイル名で「第2期」「2期」等が含まれるものを優先
        for p in pls:
            if '第2期' in p.name or '2期' in p.name:
                return p
        # なければファイルサイズが最大のもの（内容が多い=直近期の可能性）
        return max(pls, key=lambda p: p.stat().st_size)

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
) -> ProcessingStatus:
    """
    タスク1: 申請書転記の実行

    Args:
        resource_folder: 資料フォルダ
        template_path: テンプレートExcelパス
        template_type: '通常枠_2026' or 'インボイス枠_2026'
        output_path: 出力ファイルパス
        extractor: AI抽出器（省略時は自動選択）
    """
    status = ProcessingStatus(
        company_name=resource_folder.name,
        template_type=template_type,
        status='処理中',
    )

    try:
        mapping = get_mapping(template_type)

        if extractor is None:
            extractor = create_extractor(CLAUDE_API_KEY)

        # ファイル検出
        detector = FileDetector(resource_folder)
        logger.info(detector.summary())

        extraction = ExtractionResult()

        # ヒアリングシート読取（Excel直接読取 - API不要）
        hearing_path = detector.get('hearing')
        hearing_data = {}
        if hearing_path:
            hearing_data = read_hearing_sheet(hearing_path)
            logger.info(f'ヒアリングシート: {len(hearing_data)}行読込')
        else:
            logger.warning('ヒアリングシートが見つかりません')

        # 履歴事項PDF → CompanyInfo
        registry_path = detector.get('registry')
        if registry_path:
            images = pdf_to_images(registry_path)
            extraction.company = extractor.extract_registry(images)
            logger.info(f'履歴事項: {extraction.company.name}')

        # 損益計算書PDF → FinancialData
        # 製造原価報告書がある場合は画像を結合してAIに送る
        pl_path = detector.get_pl_latest()
        if pl_path:
            images = pdf_to_images(pl_path)
            cost_report_path = detector.get('cost_report')
            if cost_report_path:
                images += pdf_to_images(cost_report_path)
                logger.info(f'製造原価報告書も読取: {cost_report_path.name}')
            extraction.financial = extractor.extract_pl(images)
            logger.info(f'損益計算書: 売上{extraction.financial.revenue:,}')

        # 納税証明書PDF
        tax_path = detector.get('tax')
        if tax_path:
            images = pdf_to_images(tax_path)
            extraction.tax = extractor.extract_tax(images)

        # 見積書
        estimate_path = detector.get('estimate')
        if estimate_path:
            if estimate_path.suffix == '.xlsx':
                # Excelの見積書は直接読取
                import openpyxl
                wb_est = openpyxl.load_workbook(estimate_path, data_only=True)
                ws = wb_est[wb_est.sheetnames[0]]
                # 「件名」「品名」「ツール名」等のラベル横のセルからツール名を取得
                tool_name_keywords = ['件名', '品名', 'ツール名', '商品名', 'サービス名']
                found = False
                for row in ws.iter_rows(min_row=1, max_row=30):
                    for i, cell in enumerate(row):
                        v = cell.value
                        if v and isinstance(v, str):
                            if any(kw in v for kw in tool_name_keywords):
                                # 同じ行の次のセルを取得
                                if i + 1 < len(row) and row[i + 1].value:
                                    extraction.estimate.tool_name = str(row[i + 1].value)
                                    found = True
                                    break
                    if found:
                        break
                # 見つからなければファイル名から推測
                if not found:
                    import re
                    name = estimate_path.stem
                    # 「様」「お見積り」等を除去
                    for remove in ['お見積り', 'お見積', '見積り', '見積', '御見積', '_', '様']:
                        name = name.replace(remove, '')
                    # 日付パターン除去 (20250519, 2025-05-19 等)
                    name = re.sub(r'\d{8}', '', name)
                    name = re.sub(r'\d{4}[-/]\d{2}[-/]\d{2}', '', name)
                    name = name.strip()
                    if len(name) > 2:
                        extraction.estimate.tool_name = name
                wb_est.close()
            else:
                images = pdf_to_images(estimate_path)
                extraction.estimate = extractor.extract_estimate(images)

        # AI判断（ヒアリングデータも渡してIT投資状況等の矛盾を防ぐ）
        extraction.ai_judgment = extractor.generate_ai_judgment(
            extraction.company,
            extraction.financial,
            extraction.estimate.tool_name,
            hearing_data=hearing_data,
        )

        # 賃金台帳 → 1人当たり給与支給総額の計画値 + 一覧Excel出力
        # AI抽出経路を活かすため extractor を渡す。失敗時は決定論パーサーにフォールバック。
        wage_plan, ledger_employees, wage_status = _calc_wage_plan_from_ledger(
            detector, extraction.financial, extractor=extractor,
        )

        # テンプレート転記
        empty_cells = fill_template(
            template_path=template_path,
            output_path=output_path,
            mapping=mapping,
            hearing_data=hearing_data,
            extraction=extraction,
            wage_plan=wage_plan,
        )

        status.status = '完了'
        status.output_files = [output_path.name]
        status.empty_cells = empty_cells
        # 賃金台帳の読み取り状況に応じて完了メッセージに警告を追記（処理は続行）
        wage_warning = ''
        if wage_status == 'no_data':
            wage_warning = ' ⚠ 賃金台帳が読み取れませんでした（給与支給総額は空欄）'
        elif wage_status == 'zero_total':
            wage_warning = ' ⚠ 賃金台帳の給与支給総額が0でした'
        elif wage_status == 'error':
            wage_warning = ' ⚠ 賃金台帳処理中にエラーが発生しました'
        status.message = f'完了。空欄{len(empty_cells)}件{wage_warning}'
        logger.info(f'申請書作成完了: {output_path.name} (空欄{len(empty_cells)}件{wage_warning})')

        # 賃金台帳一覧Excel出力（チェック用）— AI抽出結果をそのまま再利用してAPI呼出しの2重化を防ぐ
        if ledger_employees:
            company = output_path.stem.split('_')[0]
            ledger_output = output_path.parent / f'{company}_賃金台帳一覧.xlsx'
            export_wage_ledger_summary(ledger_employees, ledger_output, company)
            status.output_files.append(ledger_output.name)

    except Exception as e:
        status.status = 'エラー'
        status.message = str(e)
        logger.error(f'エラー: {e}', exc_info=True)

    return status


def run_wage_calculation(
    resource_folder: Path,
    company_name: str,
    output_path: Path,
    extractor: BaseExtractor | None = None,
) -> ProcessingStatus:
    """
    タスク2: 給与支給総額計算の実行
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

        # 損益計算書（任意: あれば精度向上）
        financial = None
        pl_path = detector.get_pl_latest()
        if pl_path:
            images = pdf_to_images(pl_path)
            financial = extractor.extract_pl(images)

        if financial is None or financial.revenue == 0:
            from .models import FinancialData
            if financial is None:
                financial = FinancialData()
            logger.info('損益計算書なし → 賃金台帳ベースで計算')

        # 賃金状況報告シートから従業員データ読取（あれば）
        employees_detail = None
        seishain_count = 0
        part_count = 0
        yakuin_count = 1
        yakuin_hoshu_3m = 0

        wage_report_path = detector.get('wage_report')
        if wage_report_path:
            employees_detail, seishain_count, part_count, yakuin_hoshu_3m = \
                _read_wage_report(wage_report_path)
            logger.info(f'賃金状況報告シート: 正社員{seishain_count}, パート{part_count}')

        # フォールバック: 賃金状況報告シートで人数が取れなかった場合、賃金台帳から補完
        if seishain_count + part_count == 0:
            ledger_paths = detector.get_all('wage_ledger')
            if ledger_paths:
                fiscal_hint = _format_fiscal_period(financial)
                ledger_emps = read_wage_ledgers(
                    ledger_paths,
                    extractor=extractor,
                    fiscal_period_hint=fiscal_hint,
                )
                if ledger_emps:
                    employees_detail = _build_employees_detail_from_ledger(ledger_emps)
                    seishain_count = sum(
                        1 for e in employees_detail if e['type'] == '正社員'
                    )
                    part_count = sum(
                        1 for e in employees_detail
                        if e['type'] in ('パート・アルバイト', '契約社員')
                    )
                    logger.info(
                        f'賃金台帳フォールバック: 正社員{seishain_count}, '
                        f'パート・契約{part_count} ({len(ledger_paths)}ファイル)'
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

        fiscal_label = f'{financial.fiscal_year_start} ～ {financial.fiscal_year_end}'

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
        )

        status.status = '完了'
        status.output_files = [output_path.name]
        status.message = '給与支給総額計算 完了'
        logger.info(f'給与計算完了: {output_path.name}')

    except Exception as e:
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


def _calc_wage_plan_from_ledger(
    detector: FileDetector,
    financial: 'FinancialData',
    extractor=None,
) -> tuple[dict[str, float] | None, list, str]:
    """
    賃金台帳から給与支給総額を算出し、年3%成長の計画値を返す。

    extractor が渡されると AI 抽出を優先する（USE_AI_WAGE_EXTRACTION=true 時）。
    AI失敗時は決定論パーサーにフォールバック。

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

    fiscal_hint = _format_fiscal_period(financial)

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
        logger.warning(f'賃金台帳処理エラー（申請書作成は続行）: {e}', exc_info=True)
        return None, [], 'error'


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


def _build_employees_detail_from_ledger(employees) -> list[dict]:
    """賃金台帳の読取結果 → create_wage_calculation が期待する employees_detail 形式に変換。
    役員はカウント対象外として除外する。"""
    detail = []
    for emp in employees:
        classified = _classify_emp_type(emp.employment_type)
        if classified == '役員':
            continue

        # 直近データのある3ヶ月を m1/m2/m3 に割り当て
        months_with_data = [m for m, w in enumerate(emp.monthly_wages) if w is not None]
        last_three = months_with_data[-3:]
        m_vals = [emp.monthly_wages[m] or 0 for m in last_three]
        while len(m_vals) < 3:
            m_vals.append(0)

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

    # 役員報酬（行13, D列）
    yakuin_hoshu_3m = ws.cell(13, 4).value or 0

    employees = []
    for row in ws.iter_rows(min_row=19, max_row=200):
        no = row[1].value
        name = row[2].value
        if name is None or no is None:
            continue

        m1_base = row[5].value or 0
        m1_hr = row[6].value or 0
        m2_base = row[8].value or 0
        m2_hr = row[9].value or 0
        m3_base = row[11].value or 0
        m3_hr = row[12].value or 0
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
) -> list[ProcessingStatus]:
    """タスク1 + タスク2 を一括実行"""
    extractor = create_extractor(CLAUDE_API_KEY)
    results = []

    # タスク1: 申請書
    output_app = resource_folder / f'{company_name}_{template_type.replace("_", "_")}_AI版.xlsx'
    s1 = run_application_transfer(
        resource_folder, template_path, template_type, output_app, extractor
    )
    results.append(s1)

    # タスク2: 給与計算
    output_wage = resource_folder / f'{company_name}_給与支給総額計算.xlsx'
    s2 = run_wage_calculation(resource_folder, company_name, output_wage, extractor)
    results.append(s2)

    return results
