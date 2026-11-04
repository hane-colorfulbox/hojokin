"""申請者・担当者が確認した前提を入力するStreamlit画面。予測・外部通信は行わない。"""
from __future__ import annotations


def render_application_context(template_type: str, call_round: int | None):
    import streamlit as st
    from .application_context import ApplicationContext, validate_application_context

    st.subheader('申請の前提と確認済みの給与計画')
    st.caption('確認が取れた内容を入力してください。未確認の項目は推測で補いません。'
               '給与0年度目は申請年度の見込み値、生産性は前事業年度の実績を使います。')
    c1, c2, c3 = st.columns(3)
    application_date = c1.text_input('実際の申請予定日（YYYY-MM-DD）', key='ctx_application_date')
    fiscal_start = c2.text_input('給与0年度目の開始日（YYYY-MM-DD）', key='ctx_fiscal_start')
    fiscal_end = c3.text_input('給与0年度目の終了日（YYYY-MM-DD）', key='ctx_fiscal_end')
    amount = st.number_input('補助金の申請額（円・消費税を含めない）', min_value=0,
                             step=10000, value=0, key='ctx_subsidy_amount')
    from .config import MIN_WAGE_MAP
    wc1, wc2 = st.columns(2)
    workplace_prefecture = wc1.selectbox('主たる事業場の都道府県', [''] + list(MIN_WAGE_MAP),
                                         key='ctx_workplace_prefecture')
    regional_min_wage = wc2.number_input('申請画面へ入力する地域別最低賃金（円）', min_value=0,
                                         value=0, step=1, key='ctx_regional_min_wage',
                                         help='画面の改定後額の指示に従い確認した額を入力します。過去月の法定額の判定とは別です。')
    workplace_confirmed = st.checkbox('本店住所から推測せず、主たる事業場と入力する最賃額を確認済み',
                                      key='ctx_workplace_confirmed')
    past_state = st.selectbox('過去のIT導入補助金の交付決定',
                             ['未確認', '確認済み・交付決定なし', '確認済み・交付決定あり'],
                             key='ctx_past_state')
    past_grants = None
    if past_state == '確認済み・交付決定なし':
        past_grants = []
    elif past_state == '確認済み・交付決定あり':
        st.caption('採択予定や入金の有無ではなく、年度・枠ごとの交付決定を記録してください。')
        st.caption('枠・類型の例：通常枠、インボイス枠（インボイス対応類型）、'
                   'デジタル化基盤導入枠（複数社連携IT導入類型）。類型が不明なら確定しません。')
        records = st.data_editor([{'year': 2024, 'frame': '', 'decided': False}],
                                 num_rows='dynamic', hide_index=True, key='ctx_past_grants',
                                 column_config={'year': '交付年度',
                                                'frame': st.column_config.SelectboxColumn('交付決定の枠・類型', options=[
                                                    '通常枠', '通常枠A類型', '通常枠B類型',
                                                    '複数社連携IT導入枠',
                                                    'デジタル化基盤導入枠（複数社連携IT導入類型）',
                                                    'デジタル化基盤導入枠（デジタル化基盤導入類型）',
                                                    'インボイス枠（インボイス対応類型）',
                                                    'インボイス枠（電子取引類型）', 'セキュリティ対策推進枠']),
                                                'decided': '交付決定を確認済み'})
        # An incomplete record must not disappear and make the remaining records look complete.
        past_grants = [dict(row) for row in records
                       if any(value is not None and value is not False and str(value).strip()
                              for value in dict(row).values())]
        if not past_grants or not all(row.get('decided') is True
                                     and str(row.get('frame') or '').strip()
                                     for row in past_grants):
            past_grants = None
    exempt_state = st.selectbox('賃上げ要件の適用除外',
                               ['未確認', '適用除外ではない', '適用除外に該当すると確認済み'],
                               key='ctx_exempt')
    exempt = None if exempt_state == '未確認' else exempt_state.startswith('適用除外に')
    exemption_reason = st.text_input('適用除外の確認根拠（該当する場合）', key='ctx_exempt_reason') if exempt else ''
    optional_state = st.selectbox('任意の賃上げ計画加点', ['未確認', '希望しない', '希望する'],
                                 key='ctx_optional_bonus')
    optional = None if optional_state == '未確認' else optional_state == '希望する'
    min_wage_increase = st.number_input('計画する事業場内最低賃金の上乗せ額（円）', min_value=0,
                                       value=0, step=1, key='ctx_mw_increase',
                                       help='通常150万円未満・インボイス枠は任意の基本加点が＋30円、'
                                            '更なる加点が＋50円です。通常150万円以上の任意加点は＋50円です。')
    payroll_scope_confirmed = st.checkbox('給与に算入する範囲と対象人数を担当者が確認済み',
                                          key='ctx_payroll_scope')
    st.caption('0〜3年度目の見込み総額と対象人数を入力してください。'
               '第7次以降は全4原本で、賃上げ率の義務がない場合も4年分の確認済み計画が必要です。')
    forecasts = st.data_editor(
        [{'year': i, 'total_salary': 0, 'employee_count_fte': 0.0, 'source': '', 'confirmed': False}
         for i in range(4)], hide_index=True, disabled=['year'], key='ctx_forecasts',
        column_config={'year': '年度目', 'total_salary': '給与総額の見込み（円）',
                       'employee_count_fte': '対象人数（正社員換算）', 'source': '算定根拠・確認記録',
                       'confirmed': '担当者確認済み'})
    forecast_rows = [{k: v for k, v in dict(row).items() if k != 'year'} for row in forecasts]
    if not any(row.get('total_salary') or row.get('employee_count_fte')
               or str(row.get('source') or '').strip() or row.get('confirmed')
               for row in forecast_rows):
        forecast_rows = []
    productivity_rows = []
    if call_round is not None and call_round >= 7 and template_type.startswith('通常枠'):
        st.subheader('労働生産性の確認済み計画')
        st.caption('生産性の基準は前事業年度の実績です。1年度目は交付申請時点の翌事業年度、'
                   '2年度目は翌々事業年度、3年度目はその次の事業年度で、給与1〜3年度目と同じ期間です。'
                   '営業利益・人件費・減価償却費・総労働時間の確認済み値を入力します。'
                   '過去交付歴に応じた3%／4%を別に検算します。')
        productivity = st.data_editor(
            [{'year': i, 'operating_profit': 0, 'personnel_cost': 0, 'depreciation': 0,
              'total_annual_hours': 0.0, 'source': '', 'confirmed': False} for i in range(1, 4)],
            hide_index=True, disabled=['year'], key='ctx_productivity_forecasts',
            column_config={'year': '年度目', 'operating_profit': '営業利益（円）',
                           'personnel_cost': '人件費（円）', 'depreciation': '減価償却費（円）',
                           'total_annual_hours': '事業者全体の年間労働時間',
                           'source': '算定根拠・確認記録', 'confirmed': '担当者確認済み'})
        productivity_rows = [{k: v for k, v in dict(row).items() if k != 'year'} for row in productivity]
    declaration_state = st.selectbox('該当する賃上げ計画の従業員への表明',
                                     ['未確認', '未実施', '実施済み・確認済み'], key='ctx_declaration')
    declaration_confirmed = None if declaration_state == '未確認' else declaration_state == '実施済み・確認済み'
    dc1, dc2 = st.columns(2)
    declaration_method = dc1.text_input('実際の表明方法', key='ctx_declaration_method')
    declaration_date = dc2.text_input('表明実施日（YYYY-MM-DD）', key='ctx_declaration_date')
    st.info('効果報告での0年度目の比較基準は未確定です。申請時の見込み値・根拠・版を保存してください。')
    payload = {
        'call_round': call_round, 'application_date': application_date,
        'fiscal_year_start': fiscal_start, 'fiscal_year_end': fiscal_end,
        'frame': '通常' if template_type.startswith('通常枠') else 'インボイス',
        'subsidy_amount': int(amount), 'past_grants': past_grants,
        'workplace_prefecture': workplace_prefecture, 'workplace_confirmed': workplace_confirmed,
        'regional_min_wage_yen': int(regional_min_wage) or None,
        'wage_exempt': exempt, 'wage_exemption_reason': exemption_reason,
        'optional_wage_bonus': optional, 'planned_min_wage_increase': int(min_wage_increase),
        'payroll_scope_confirmed': payroll_scope_confirmed,
        'declaration_confirmed': declaration_confirmed, 'declaration_method': declaration_method,
        'declaration_date': declaration_date or None,
        'forecast_years': forecast_rows,
        'productivity_forecast_years': productivity_rows,
    }
    try:
        context = ApplicationContext.from_dict(payload)
        validate_application_context(context, template_type)
        return context, []
    except (ValueError, TypeError, KeyError) as exc:
        return None, getattr(exc, 'issues', [str(exc)])
