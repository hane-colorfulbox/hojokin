# -*- coding: utf-8 -*-
"""データモデル定義"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class CompanyInfo:
    """履歴事項全部証明書から抽出するデータ"""
    name: str = ''
    name_kana: str = ''
    address: str = ''
    postal_code: str = ''
    established_date: datetime | None = None
    capital: int = 0
    representative_name: str = ''
    representative_title: str = ''
    representative_kana: str = ''
    officers: list[dict] = field(default_factory=list)
    # officers = [{'title': '取締役', 'name': '氏名', 'kana': 'フリガナ'}, ...]
    business_purposes: list[str] = field(default_factory=list)


@dataclass
class FinancialData:
    """損益計算書から抽出するデータ"""
    fiscal_year_start: str = ''
    fiscal_year_end: str = ''
    fiscal_month: str = ''  # 例: '3月'
    revenue: int = 0                # 売上高
    cost_of_sales: int = 0          # 売上原価
    gross_profit: int = 0           # 売上総利益
    operating_profit: int = 0       # 営業利益
    ordinary_profit: int = 0        # 経常利益
    net_profit: int = 0             # 当期純利益
    # 販管費内訳
    salary: int = 0                 # 給料手当
    misc_wages: int = 0             # 雑給
    bonus: int = 0                  # 賞与
    officer_compensation: int = 0   # 役員報酬
    legal_welfare: int = 0          # 法定福利費
    welfare: int = 0                # 福利厚生費
    depreciation: int = 0           # 減価償却費
    travel_expense: int = 0         # 旅費交通費


@dataclass
class TaxCertificate:
    """納税証明書から抽出するデータ"""
    tax_type: str = ''              # その1(法人税) 等
    tax_amount: int = 0
    fiscal_year: str = ''


@dataclass
class Employee:
    """従業員の給与データ"""
    name: str = ''
    department: str = ''
    employee_id: str = ''
    employment_type: str = ''       # 正社員 / パート・アルバイト
    working_days: float = 0
    scheduled_hours: float = 0
    base_salary: int = 0
    hourly_rate: int = 0
    taxable_total: int = 0          # 課税支給合計
    total_pay: int = 0              # 支給合計
    deductions: int = 0             # 控除合計
    net_pay: int = 0                # 差引支給額


@dataclass
class MonthlyWageData:
    """月次の給与データ"""
    year_month: str = ''            # 例: '2025-03'
    employees: list[Employee] = field(default_factory=list)
    director_compensation_total: int = 0


@dataclass
class EstimateData:
    """見積書から抽出するデータ"""
    vendor_name: str = ''
    tool_name: str = ''
    items: list[dict] = field(default_factory=list)
    # items = [{'name': '項目名', 'amount': 金額}, ...]
    total_amount: int = 0
    tax_amount: int = 0


@dataclass
class AIJudgment:
    """AIが判断・生成する項目"""
    industry_code: str = ''
    industry_text: str = ''
    business_description: str = ''  # 事業内容(255文字以内)
    management_intent: str = ''     # 経営意欲
    future_goals: str = ''          # 将来目標
    security_status: str = ''       # セキュリティ状況
    business_types: str = ''        # 行っている事業
    min_wage_text: str = ''         # 地域別最低賃金
    it_investment_status: str = ''  # IT投資状況
    it_utilization_status: str = '' # IT活用状況
    it_utilization_scope: str = ''  # IT電子化範囲（インボイス枠）
    invoice_related_work: str = ''  # インボイス対応業務（インボイス枠）


@dataclass
class ExtractionResult:
    """全抽出データの統合"""
    company: CompanyInfo = field(default_factory=CompanyInfo)
    financial: FinancialData = field(default_factory=FinancialData)
    financial_prev: FinancialData | None = None  # 前期（あれば）
    tax: TaxCertificate = field(default_factory=TaxCertificate)
    wages: list[MonthlyWageData] = field(default_factory=list)
    estimate: EstimateData = field(default_factory=EstimateData)
    ai_judgment: AIJudgment = field(default_factory=AIJudgment)


@dataclass
class ProcessingStatus:
    """処理ステータス"""
    company_name: str = ''
    template_type: str = ''
    status: str = '未処理'  # 未処理 / 処理中 / 完了 / エラー
    message: str = ''
    output_files: list[str] = field(default_factory=list)
    empty_cells: list[str] = field(default_factory=list)
    # AI抽出結果（後続タスクで再利用してAPI重複呼出しを防ぐ）
    # 'WageEmployee' は循環import回避のため文字列で型注釈
    financial: FinancialData | None = None
    ledger_employees: list = field(default_factory=list)
