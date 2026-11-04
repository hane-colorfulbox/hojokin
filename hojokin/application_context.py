"""申請の適用回と、確認済みの申告・給与計画。

台帳から算出した前期実績を、未確定の当期見込みとして流用しない。
このモジュールは外部APIやファイルを使用しない。
"""
from __future__ import annotations

import calendar
import math
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_CEILING
from typing import Any


class ApplicationValidationError(ValueError):
    """申請前提や確認済み入力が不足しているため確定出力できない。"""

    def __init__(self, issues: list[str]):
        self.issues = list(dict.fromkeys(issues))
        super().__init__('申請書を確定できません：\n' + '\n'.join(f'・{x}' for x in self.issues))


def _date(value: Any) -> date | None:
    if value in (None, ''):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return date.fromisoformat(str(value))
    except (ValueError, TypeError):
        raise ApplicationValidationError([f'日付の形式が不正です：{value}（YYYY-MM-DD）'])


def _finite_number(value: Any) -> bool:
    return isinstance(value, (int, float, Decimal)) and not isinstance(value, bool) and math.isfinite(value)


def _frame(value: str) -> str:
    text = str(value).strip()
    if text.startswith('通常'):
        return '通常'
    if text.startswith('インボイス'):
        return 'インボイス'
    if text.startswith('セキュリティ'):
        return 'セキュリティ'
    return text


def _text(value: Any) -> str:
    return value.strip() if isinstance(value, str) else ''


def _shift_year(value: date, years: int) -> date:
    year = value.year + years
    return date(year, value.month, min(value.day, calendar.monthrange(year, value.month)[1]))


@dataclass(frozen=True)
class GrantRecord:
    year: int
    frame: str
    decided: bool | None = None


@dataclass(frozen=True)
class WageForecastYear:
    total_salary: int | float | None = None
    employee_count_fte: float | None = None
    source: str = ''
    confirmed: bool = False

    @property
    def per_capita(self) -> Decimal:
        return Decimal(str(self.total_salary)) / Decimal(str(self.employee_count_fte))


@dataclass(frozen=True)
class ProductivityForecastYear:
    operating_profit: int | float | None = None
    personnel_cost: int | float | None = None
    depreciation: int | float | None = None
    total_annual_hours: float | None = None
    source: str = ''
    confirmed: bool = False

    @property
    def productivity(self) -> Decimal:
        value_added = sum((Decimal(str(value)) for value in
                           (self.operating_profit, self.personnel_cost, self.depreciation)), Decimal('0'))
        return value_added / Decimal(str(self.total_annual_hours))


@dataclass
class ApplicationContext:
    call_round: int | None = None
    application_date: date | None = None
    fiscal_year_start: date | None = None
    fiscal_year_end: date | None = None
    frame: str = ''
    subsidy_amount: int | None = None
    # None は未確認。空リストは「交付なし」を確認済み。
    past_grants: list[GrantRecord] | None = None
    optional_wage_bonus: bool | None = None
    wage_exempt: bool | None = None
    wage_exemption_reason: str = ''
    payroll_scope_confirmed: bool = False
    forecast_years: list[WageForecastYear] = field(default_factory=list)
    productivity_forecast_years: list[ProductivityForecastYear] = field(default_factory=list)
    declaration_confirmed: bool | None = None
    declaration_method: str = ''
    declaration_date: date | None = None
    planned_min_wage_increase: int | None = None
    workplace_prefecture: str = ''
    workplace_confirmed: bool = False
    regional_min_wage_yen: int | None = None
    # 読取り・年次計画を自動処理しない個別確認ケース。
    manual_review_reasons: list[str] = field(default_factory=list)

    @classmethod
    def from_dict(cls, values: dict[str, Any]) -> 'ApplicationContext':
        if not isinstance(values, dict):
            raise ApplicationValidationError(['案件設定はオブジェクト形式で入力してください'])
        allowed = set(cls.__dataclass_fields__)
        unknown = set(values) - allowed
        if unknown:
            raise ApplicationValidationError(['未対応の案件設定：' + '、'.join(sorted(unknown))])
        args = dict(values)
        for key in ('application_date', 'fiscal_year_start', 'fiscal_year_end', 'declaration_date'):
            if key in args:
                args[key] = _date(args[key])
        try:
            if args.get('past_grants') is not None:
                args['past_grants'] = [x if isinstance(x, GrantRecord) else GrantRecord(**x)
                                       for x in args['past_grants']]
            if args.get('forecast_years') is not None:
                args['forecast_years'] = [x if isinstance(x, WageForecastYear) else WageForecastYear(**x)
                                         for x in args['forecast_years']]
            else:
                args['forecast_years'] = []
            if args.get('productivity_forecast_years') is not None:
                args['productivity_forecast_years'] = [
                    x if isinstance(x, ProductivityForecastYear) else ProductivityForecastYear(**x)
                    for x in args['productivity_forecast_years']]
            else:
                args['productivity_forecast_years'] = []
            return cls(**args)
        except (TypeError, ValueError) as exc:
            if isinstance(exc, ApplicationValidationError):
                raise
            raise ApplicationValidationError([f'案件設定の形式を確認してください：{exc}']) from exc

    @property
    def is_round7(self) -> bool:
        return isinstance(self.call_round, int) and not isinstance(self.call_round, bool) and self.call_round >= 7

    @property
    def application_ym(self) -> tuple[int, int] | None:
        return (self.application_date.year, self.application_date.month) if self.application_date else None

    @property
    def productivity_period(self) -> tuple[date, date] | None:
        if not self.fiscal_year_start or not self.fiscal_year_end:
            return None
        return (_shift_year(self.fiscal_year_start, -1), self.fiscal_year_start - timedelta(days=1))

    def past_grant_text(self) -> str | None:
        if self.past_grants is None:
            return None
        records = [r for r in self.past_grants if r.decided is True]
        return 'なし' if not records else '\n'.join(f'{r.year}年度 {r.frame}（交付決定）' for r in records)


@dataclass(frozen=True)
class WageRequirements:
    mandatory: bool
    selected: bool
    required_rate: Decimal
    first_year_required: bool
    min_wage_increase: int
    productivity_rate: Decimal


def _productivity_grant_status(record: GrantRecord) -> bool | None:
    if record.decided is not True or record.year not in (2023, 2024, 2025):
        return False
    normalize = lambda value: unicodedata.normalize('NFKC', value).replace(' ', '').replace('　', '')
    frame = normalize(record.frame)
    qualifying = {
        '通常枠', '通常枠A類型', '通常枠B類型', '通常枠（A類型）', '通常枠（B類型）',
        '通常枠（A・B類型）', '複数社連携IT導入枠', '複数社連携IT導入類型',
        'デジタル化基盤導入枠（複数社連携IT導入類型）',
    }
    non_qualifying = {
        'インボイス枠', 'インボイス枠（インボイス対応類型）', 'インボイス対応類型',
        'インボイス枠（電子取引類型）', '電子取引類型',
        'セキュリティ対策推進枠', 'セキュリティ枠',
        'デジタル化基盤導入枠（デジタル化基盤導入類型）', 'デジタル化基盤導入類型',
        'デジタル化基盤導入枠（商流一括インボイス対応類型）', '商流一括インボイス対応類型',
    }
    if frame in {normalize(x) for x in qualifying}:
        return True
    if frame in {normalize(x) for x in non_qualifying}:
        return False
    # 「デジタル化基盤導入枠」だけでは、複数社連携類型か区別できない。
    return None


def wage_requirements(context: ApplicationContext) -> WageRequirements:
    grants = context.past_grants or []
    wage_past = any(r.decided is True and 2022 <= r.year <= 2025 for r in grants)
    prod_past = any(_productivity_grant_status(r) is True for r in grants)
    frame = _frame(context.frame)
    large_normal = frame == '通常' and (context.subsidy_amount or 0) >= 1_500_000
    mandatory = (wage_past or large_normal) and context.wage_exempt is not True
    selected = mandatory or context.optional_wage_bonus is True
    # 第7次：通常150万円未満・インボイス・セキュリティは基本加点+30円、
    # +50円以上は更なる加点。通常150万円以上の任意加点は+50円。
    # 任意加点を選ばない通常150万円以上の申請必須要件は+30円のまま。
    if context.optional_wage_bonus is True:
        minimum = 50 if large_normal or not context.is_round7 else 30
    else:
        minimum = 30 if large_normal and mandatory else 0
    return WageRequirements(
        mandatory=mandatory,
        selected=selected,
        required_rate=Decimal('0.035') if wage_past else Decimal('0.03'),
        first_year_required=context.is_round7 and mandatory,
        min_wage_increase=minimum,
        productivity_rate=Decimal('0.04') if prod_past else Decimal('0.03'),
    )


def validate_application_context(context: ApplicationContext | None, template_type: str | None = None) -> None:
    if not isinstance(context, ApplicationContext):
        raise ApplicationValidationError(['公募回・申告確認・給与計画の案件設定が未入力です'])
    issues: list[str] = []
    if not isinstance(context.call_round, int) or isinstance(context.call_round, bool) or not 1 <= context.call_round <= 8:
        issues.append('公募回を1〜8から指定してください')
    if context.application_date is None:
        issues.append('実際の交付申請日が未確認です')
    if _frame(context.frame) not in {'通常', 'インボイス', 'セキュリティ'}:
        issues.append('対応する申請枠を確認してください（電子取引等はこの4原本の対象外です）')
    if '電子取引' in str(context.frame) or '複数' in str(context.frame):
        issues.append('電子取引・複数者連携等は現行4原本の対象外です。個別確認してください')
    if template_type and _frame(context.frame) != _frame(template_type):
        issues.append('案件設定の申請枠と申請書テンプレートが一致しません')
    if not _finite_number(context.subsidy_amount) or context.subsidy_amount <= 0 or int(context.subsidy_amount) != context.subsidy_amount:
        issues.append('補助金申請額（円）が未確認、または不正です')
    if context.past_grants is None:
        issues.append('IT導入補助金の過去交付歴が未確認です（交付なしの場合も明示確認が必要）')
    else:
        for rec in context.past_grants:
            if not isinstance(rec, GrantRecord) or not isinstance(rec.year, int) or isinstance(rec.year, bool) or not 2000 <= rec.year <= 2100 or not isinstance(rec.frame, str) or not rec.frame.strip() or type(rec.decided) is not bool:
                issues.append('過去交付歴の年度・枠・交付決定の有無を確認してください')
                break
            if _productivity_grant_status(rec) is None:
                issues.append('2023〜2025年の過去交付について、枠・類型が不明なため生産性要件を決められません。'
                              '通常枠・複数社連携か、それ以外の類型かを確認してください')
    if type(context.optional_wage_bonus) is not bool:
        issues.append('任意の賃上げ計画加点を希望するか未確認です')
    if type(context.wage_exempt) is not bool:
        issues.append('賃上げ要件の適用除外に該当するか未確認です')
    elif context.wage_exempt and not _text(context.wage_exemption_reason):
        issues.append('適用除外の法令上の根拠・事業内容を記録してください')
    if context.declaration_confirmed is not None and type(context.declaration_confirmed) is not bool:
        issues.append('表明実施の有無は確認済みの真偽値で指定してください')
    if context.wage_exempt and context.optional_wage_bonus:
        issues.append('適用除外と任意賃上げ加点を併用する案件は個別確認が必要です')
    if not context.fiscal_year_start or not context.fiscal_year_end:
        issues.append('申請日の属する事業年度の開始日・終了日が未確認です')
    else:
        expected_end = _shift_year(context.fiscal_year_start, 1) - timedelta(days=1)
        if context.fiscal_year_start.day != 1 or context.fiscal_year_end != expected_end:
            issues.append('設立初年度・短期決算・期変更は自動処理せず個別確認してください')
        if context.application_date and not context.fiscal_year_start <= context.application_date <= context.fiscal_year_end:
            issues.append('交付申請日と給与0年度目の対象事業年度が一致しません')
        if template_type and '個人' in template_type and (context.fiscal_year_start.month != 1 or context.fiscal_year_end.month != 12):
            issues.append('個人事業主の対象事業年度は暦年（1月1日〜12月31日）で確認してください')
    if context.manual_review_reasons:
        issues.extend('個別確認が必要：' + str(x) for x in context.manual_review_reasons)
    if context.is_round7:
        if context.workplace_confirmed is not True or not _text(context.workplace_prefecture):
            issues.append('主たる事業場の都道府県が未確認です（本店住所からは推測しません）')
        if not _finite_number(context.regional_min_wage_yen) or context.regional_min_wage_yen <= 0 or int(context.regional_min_wage_yen) != context.regional_min_wage_yen:
            issues.append('申請画面に入力する改定後の地域別最低賃金額を確認してください')
        else:
            from .subsidy_rules import rules_for_round
            expected = rules_for_round(context.call_round).min_wage_after.get(context.workplace_prefecture)
            if expected is None:
                issues.append('主たる事業場の都道府県を正式名称で確認してください')
            elif context.regional_min_wage_yen != expected:
                issues.append('申請画面用の改定後最低賃金額が第7次用マスタと一致しません。'
                              '過去月の法定額とは別に確認してください')
    # GrantRecord 等の構造が壊れている場合は、ここで適用要件を計算しない。
    if issues:
        raise ApplicationValidationError(issues)
    req = wage_requirements(context)
    if req.selected:
        if context.declaration_confirmed is not True:
            issues.append('対象の賃上げ計画を従業員に表明したことが未確認・未実施です')
        if not _text(context.declaration_method) or not context.declaration_date:
            issues.append('実際の表明方法・表明日が未確認です')
        elif context.application_date and context.declaration_date > context.application_date:
            issues.append('表明日が交付申請日より後になっています')
        if req.min_wage_increase and (not _finite_number(context.planned_min_wage_increase) or context.planned_min_wage_increase < req.min_wage_increase):
            issues.append(f'事業場内最低賃金の計画は地域別最低賃金＋{req.min_wage_increase}円以上の確認が必要です')
    # 率義務の有無と計画入力の要否を分ける。第7次は適用外でも給与0〜3年度目の
    # 計画値が必要（通常枠FAQ No.39／各枠の交付申請マニュアル画面指示）。
    forecast_required = context.is_round7
    if context.forecast_years or forecast_required:
        if context.payroll_scope_confirmed is not True:
            issues.append('給与見込の対象従業員・FTE・算入範囲が未確認です')
        if len(context.forecast_years) != 4:
            issues.append('0〜3年度目の確認済み給与見込総額・対象FTEを4年分入力してください')
        else:
            for idx, year in enumerate(context.forecast_years):
                if not isinstance(year, WageForecastYear):
                    issues.append(f'{idx}年度目の給与見込入力形式を確認してください')
                    continue
                if not _finite_number(year.total_salary) or year.total_salary <= 0 or int(year.total_salary) != year.total_salary:
                    issues.append(f'{idx}年度目の給与見込総額は確認済みの正の円単位整数が必要です')
                if not _finite_number(year.employee_count_fte) or year.employee_count_fte <= 0:
                    issues.append(f'{idx}年度目の対象FTEは確認済みの正数が必要です（従業員ゼロ等は個別確認）')
                if year.confirmed is not True or not _text(year.source):
                    issues.append(f'{idx}年度目の見込総額・FTEの根拠と担当確認が必要です')
            if not issues and req.selected:
                base = context.forecast_years[0].per_capita
                one = context.forecast_years[1].per_capita
                last = context.forecast_years[3].per_capita
                if req.first_year_required and one < base * (1 + req.required_rate):
                    issues.append(f'1年後の1人当たり給与支給総額が{req.required_rate * 100}%要件を満たしません')
                # 毎年必要率で複利増した水準を下回る案は受け付けない。
                # 年率×3で置き換えず、公式ポータルの表示も最終確認する。
                if last < base * (1 + req.required_rate) ** 3:
                    issues.append(f'3年間の1人当たり給与支給総額の年平均{req.required_rate * 100}%要件を満たしません')
    if context.is_round7 and _frame(context.frame) == '通常':
        if len(context.productivity_forecast_years) != 3:
            issues.append('生産性の1〜3年度目について、確認済みの営業利益・人件費・償却・年間総時間を3年分入力してください')
        else:
            for idx, year in enumerate(context.productivity_forecast_years, start=1):
                if not isinstance(year, ProductivityForecastYear):
                    issues.append(f'生産性{idx}年度目の入力形式を確認してください')
                    continue
                if not all(_finite_number(x) for x in (year.operating_profit, year.personnel_cost, year.depreciation)):
                    issues.append(f'生産性{idx}年度目の営業利益・人件費・償却は有限の数値を確認してください')
                elif year.personnel_cost < 0 or year.depreciation < 0:
                    issues.append(f'生産性{idx}年度目の負の人件費・償却は個別確認してください')
                elif sum((Decimal(str(x)) for x in
                          (year.operating_profit, year.personnel_cost, year.depreciation)), Decimal('0')) <= 0:
                    issues.append(f'生産性{idx}年度目の付加価値額が0以下のため個別確認してください')
                if not _finite_number(year.total_annual_hours) or year.total_annual_hours <= 0:
                    issues.append(f'生産性{idx}年度目の年間総労働時間は確認済みの正数が必要です')
                if year.confirmed is not True or not _text(year.source):
                    issues.append(f'生産性{idx}年度目の計画値の根拠と担当確認が必要です（参考式やAI予測を自動採用しません）')
    if issues:
        raise ApplicationValidationError(issues)


def build_application_wage_plan(context: ApplicationContext) -> dict[str, Any] | None:
    """確認済みの明示入力だけを申請用の給与計画に変換する。"""
    validate_application_context(context)
    if not context.forecast_years:
        return None
    req = wage_requirements(context)
    years = context.forecast_years
    return {
        'employee_count_fte': float(years[0].employee_count_fte),
        'employee_count_fte_by_year': [float(x.employee_count_fte) for x in years],
        'wage_total_base': int(years[0].total_salary),
        'wage_total_y1': int(years[1].total_salary),
        'wage_total_y2': int(years[2].total_salary),
        'wage_total_y3': int(years[3].total_salary),
        'per_capita_by_year': [float(x.per_capita) for x in years],
        'forecast_sources': [x.source for x in years],
        'basis_kind': '当期見込み' if context.is_round7 else '確認済み入力',
        'growth_rate': float(req.required_rate),
        'source': '担当者確認済みの明示入力（前期実績の自動流用なし）',
    }


def validate_productivity_inputs(context: ApplicationContext, financial,
                                 total_annual_hours: float | None) -> dict[str, Any] | None:
    """給与見込の期・FTEを、生産性の前期実績へ混ぜない。"""
    if not context.is_round7 or _frame(context.frame) != '通常':
        return
    validate_application_context(context)
    expected_start, expected_end = context.productivity_period
    actual_start = str(getattr(financial, 'fiscal_year_start', '') or '')[:7]
    actual_end = str(getattr(financial, 'fiscal_year_end', '') or '')[:7]
    issues = []
    if (actual_start, actual_end) != (expected_start.strftime('%Y-%m'), expected_end.strftime('%Y-%m')):
        issues.append('生産性に用いる決算書の年度が案件設定の前事業年度と一致しない、'
                      'または対象年度を確認できません。前事業年度の確定資料を確認してください')
    if not _finite_number(total_annual_hours) or total_annual_hours <= 0:
        issues.append('生産性の前事業年度の総労働時間を確認できません。'
                      '給与0年度目のFTEや見込み時間では補いません')
    base_fields = ('operating_profit', 'salary', 'misc_wages', 'bonus', 'travel_expense', 'depreciation')
    for key in base_fields:
        if not _finite_number(getattr(financial, key, None)):
            issues.append('生産性の前事業年度の営業利益・人件費・償却の実績数値が不足しています')
            break
    confidence = getattr(financial, 'confidence', {}) or {}
    if any(getattr(confidence.get(key), 'level', '') == 'low' for key in base_fields):
        issues.append('生産性の基準実績に低信頼の抽出値があるため、確定資料で確認してください')
    if issues:
        raise ApplicationValidationError(issues)
    personnel = sum((Decimal(str(getattr(financial, key)))
                     for key in ('salary', 'misc_wages', 'bonus', 'travel_expense')), Decimal('0'))
    base_value_added = Decimal(str(financial.operating_profit)) + personnel + Decimal(str(financial.depreciation))
    if base_value_added <= 0:
        raise ApplicationValidationError(['生産性の基準付加価値額が0以下のため、成長率を自動判定せず個別確認してください'])
    base = base_value_added / Decimal(str(total_annual_hours))
    req = wage_requirements(context)
    plans = [year.productivity for year in context.productivity_forecast_years]
    if len(plans) != 3:
        raise ApplicationValidationError(['生産性の確認済み3年計画が不足しています'])
    if plans[0] < base * (1 + req.productivity_rate):
        issues.append(f'生産性の1年後向上率が{req.productivity_rate * 100}%要件を満たしません')
    if plans[2] < base * (1 + req.productivity_rate) ** 3:
        issues.append(f'生産性の3年間の年平均向上率が{req.productivity_rate * 100}%要件を満たしません')
    if issues:
        raise ApplicationValidationError(issues)
    return {
        'baseline_value_added': float(base_value_added),
        'baseline_personnel_cost': float(personnel),
        'baseline_productivity': float(base),
        'forecast_productivity': [float(x) for x in plans],
        'required_rate': float(req.productivity_rate),
        'first_year_growth': float(plans[0] / base - 1),
        'second_year_average': float(plans[1] / base) ** 0.5 - 1,
        'third_year_average': float(plans[2] / base) ** (1 / 3) - 1,
    }


def propose_constant_fte_plan(base: WageForecastYear, required_rate: Decimal | float, *,
                              constant_fte_confirmed: bool = False) -> list[WageForecastYear]:
    """人数一定が確認された案件だけの未確定計画案。案を担当確認済みにはしない。"""
    if constant_fte_confirmed is not True:
        raise ApplicationValidationError(['人数一定の計画であることを確認してください'])
    if not (_finite_number(base.total_salary) and base.total_salary > 0 and
            _finite_number(base.employee_count_fte) and base.employee_count_fte > 0 and
            base.confirmed is True and _text(base.source)):
        raise ApplicationValidationError(['0年度目の確認済み見込総額・FTE・根拠が必要です'])
    rate = Decimal(str(required_rate))
    if rate not in (Decimal('0.03'), Decimal('0.035')):
        raise ApplicationValidationError(['対応する給与計画率は3%または3.5%です'])
    out = [base]
    total = Decimal(str(base.total_salary))
    for _ in range(3):
        total = (total * (1 + rate)).to_integral_value(rounding=ROUND_CEILING)
        out.append(WageForecastYear(int(total), base.employee_count_fte,
                                   '人数一定・前年から必要率で切り上げた未承認案', False))
    return out
