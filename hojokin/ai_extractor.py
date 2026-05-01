# -*- coding: utf-8 -*-
"""
Claude APIによるPDFデータ抽出
- StubExtractor: APIキーなしで動作するスタブ（プレースホルダ値を返す）
- ClaudeExtractor: 実際のAPI呼出し（APIキー必要）
"""
from __future__ import annotations

import json
import logging
import time
from abc import ABC, abstractmethod
from typing import Callable, Optional

from .models import (
    CompanyInfo, FinancialData, TaxCertificate,
    Employee, MonthlyWageData, EstimateData, AIJudgment,
)

logger = logging.getLogger(__name__)


# ── リトライ設定 ──
# 初回+2回リトライ = 最大3回試行、バックオフは 2s → 5s
MAX_API_ATTEMPTS = 3
API_BACKOFF_SECONDS = [2, 5]

# リトライ対象のHTTPステータス（APIStatusError系の一時的失敗）
# 422: "context reduction is suggested" 等のflakyエラー
# 429: rate_limit
# 500/502/503/504: サーバ側の一時障害
# 529: overloaded_error
RETRYABLE_STATUS_CODES = {422, 429, 500, 502, 503, 504, 529}

# 残高切れ判定用の文字列（400 invalid_request_error の message に含まれる）
CREDIT_BALANCE_MARKER = 'credit balance is too low'

# 進捗コールバックの型: (attempt, max_attempts, wait_seconds, error_summary) -> None
RetryCallback = Callable[[int, int, float, str], None]


class APICreditExhaustedError(RuntimeError):
    """API残高切れ（400 credit_balance_too_low）を表す専用例外"""
    pass


# ── プロンプトテンプレート ──

PROMPT_REGISTRY = """この履歴事項全部証明書の画像から、以下の情報をJSON形式で抽出してください。
読み取れない項目はnullにしてください。

重要ルール:
- 履歴事項には役員の就任・退任・重任の履歴が記録されています。同一人物が複数回登場する場合は、最新の役職のみを採用してください。
- 下線が引かれた（抹消された）情報は過去のものなので無視してください。
- 代表者はofficersには含めないでください（representative_name/representative_titleに記載）。
- 退任済みの役員は含めないでください。

```json
{
  "name": "法人名（株式会社等含む）",
  "name_kana": "法人名フリガナ（カタカナ）",
  "address": "本店所在地",
  "postal_code": "郵便番号（わかれば）",
  "established_date": "設立年月日 yyyy-mm-dd形式",
  "capital": 資本金（円、整数）,
  "representative_name": "代表者氏名",
  "representative_title": "代表者役職",
  "officers": [
    {"title": "役職", "name": "氏名", "kana": "フリガナ（推定でOK）"}
  ],
  "business_purposes": ["目的1", "目的2"]
}
```"""

PROMPT_PL = """この損益計算書・販管費内訳書の画像から、以下をJSON形式で抽出してください。
製造原価報告書の画像が含まれている場合は、そこからも減価償却費を読み取り、販管費の減価償却費と合算してください。
該当項目がない場合はnullにしてください。金額は円単位の整数で。

個人事業主の「所得税の青色申告決算書」または「収支内訳書」の場合:
- revenue = 売上（収入）金額
- gross_profit = 売上（収入）金額 - 売上原価
- operating_profit / ordinary_profit = 所得金額（青色申告特別控除前）
- salary = 給料賃金
- 役員報酬・賞与・経常利益といった法人特有の項目は null
- 専従者給与がある場合は misc_wages に計上

```json
{
  "fiscal_year_start": "事業年度開始日 yyyy-mm-dd",
  "fiscal_year_end": "事業年度終了日 yyyy-mm-dd",
  "revenue": 売上高,
  "cost_of_sales": 売上原価,
  "gross_profit": 売上総利益,
  "operating_profit": 営業利益（損失ならマイナス）,
  "ordinary_profit": 経常利益（損失ならマイナス）,
  "net_profit": 当期純利益（損失ならマイナス）,
  "salary": 給料手当,
  "misc_wages": 雑給,
  "bonus": 賞与,
  "officer_compensation": 役員報酬,
  "legal_welfare": 法定福利費,
  "welfare": 福利厚生費,
  "depreciation": 減価償却費,
  "travel_expense": 旅費交通費
}
```"""

PROMPT_TAX = """この納税証明書の画像から、以下をJSON形式で抽出してください。

```json
{
  "tax_type": "証明書の種類（その1、その2等）",
  "tax_amount": 納税額（円、整数）,
  "fiscal_year": "事業年度"
}
```"""

PROMPT_WAGES = """この給与支給控除一覧表の画像から、従業員ごとのデータをJSON配列で抽出してください。
全従業員を漏れなく抽出してください。

```json
[
  {
    "name": "氏名",
    "department": "所属（例: 総本店）",
    "employee_id": "社員番号",
    "employment_type": "正社員 または パート・アルバイト",
    "working_days": 出勤日数,
    "scheduled_hours": 所定労働時間,
    "base_salary": 基本給,
    "taxable_total": 課税支給合計,
    "total_pay": 支給合計,
    "deductions": 控除合計,
    "net_pay": 差引支給額
  }
]
```

判定ヒント:
- 社員番号100xxx台 → 正社員、200xxx台 → パート・アルバイト
- 所属欄に「正社員」「アルバイト」の記載があればそれを使う"""

PROMPT_ESTIMATE = """この見積書の画像から、以下をJSON形式で抽出してください。

```json
{
  "vendor_name": "発行元の会社名",
  "tool_name": "ツール/サービス名",
  "items": [
    {"name": "項目名", "amount": 金額}
  ],
  "total_amount": 合計金額（税抜）,
  "tax_amount": 消費税額
}
```"""

PROMPT_AI_JUDGMENT = """以下の会社情報に基づいて、補助金申請に必要な判断項目を埋めてください。

会社情報:
- 会社名: {company_name}
- 事業内容: {business_purposes}
- 所在地: {address}
- 営業利益: {operating_profit}円
- ツール名: {tool_name}

ヒアリングシート回答:
- 主な事業内容: {main_business}
- 強み: {strength}
- 課題（時間がかかっている業務）: {challenge}
- 月間所要時間: {monthly_hours}
- ツールで楽にしたいこと: {tool_usage}
- 削減見込み: {reduction}
- 浮いた時間の活用: {freed_time}
- 3年後の売上目標: {sales_target}
- IT投資実績: {it_investment_answer}
- IT投資金額: {it_investment_amount}
- IT投資プロセス: {it_investment_process}

※ヒアリングシートの回答が空欄の項目がある場合は、履歴事項の事業目的や決算書の情報から合理的に推定してください。

以下をJSON形式で回答してください。
重要: ヒアリングシートの回答を最優先で参照し、矛盾しないようにしてください。

```json
{{
  "industry_code": "日本標準産業分類の細分類コード（4桁）",
  "industry_text": "大分類 X xxx / 中分類 xx xxx / 小分類 xxx xxx / 細分類 xxxx xxx",
  "business_description": "事業内容の説明文。250-255文字。会社の現状・課題・ツール導入による解決策・期待効果を含む",
  "management_intent": "営業利益がプラスなら '事業の拡大に積極的'、マイナスなら '事業の維持に注力'",
  "future_goals": "営業利益がプラスなら '事業の拡大'、マイナスなら '利益の確保'",
  "security_status": "パソコンやサーバなどには、IDやパスワードを設け情報セキュリティ管理を行っている",
  "business_types": "履歴事項の目的から該当する日本標準産業分類の大分類をカンマ区切りで",
  "it_investment_status": "ヒアリングのIT投資実績が「はい」なら過去にIT投資を行ったことがある旨を記載。「いいえ」なら今までIT投資を行っていなかった",
  "it_utilization_status": "ヒアリングのIT投資実績に基づき適切に選択",
  "it_utilization_scope": "ITツールの導入により電子化する事務の範囲（例: '会計', '受発注', '決済' 等から該当するものをカンマ区切りで）",
  "invoice_related_work": "ITツールの導入によりインボイス対応に資する業務（例: '請求書の発行・受領', '仕入税額控除の計算' 等）"
}}
```"""


PROMPT_WAGE_LEDGER = """以下は賃金台帳のExcelをTSV形式に変換したテキストです。
各従業員の月別給与データ・労働時間データを抽出し、JSON形式で返してください。

【抽出ルール】
1. 全シート・全テーブルを横断して、登場するすべての従業員を抽出してください。
2. monthly_wages: 月別の課税支給合計（または支給合計、税込支給額、給与+賞与の合算）。各月の値を Index 0=1月, Index 11=12月 として格納してください。データがない月は null。
3. monthly_hours: 月別の労働時間（または所定労働時間、勤務時間）。同じく1月〜12月のIndex順。データがない月は null。
4. employment_type: 雇用形態（正社員・パート・アルバイト・役員等）。元の表記をそのまま入れてください。役員は「役員」を含む表記に。
5. 給与と賞与が別行・別シートに分かれている場合は、同月分を**合算**してください。
6. 月の判定は以下のいずれかを使用:
   - 列ヘッダ「1月」「5月」等のプレーン表記
   - 「R6.5月」「R7.4月」「令和6年5月」等の和暦付き表記（年は無視して月だけ使用）
   - 「2024年5月」「2024/05」「202405」等の西暦表記
   - 「対象年月」「給与年月」列の値

{fiscal_period_section}

【出力形式（厳密に従ってください）】
```json
[
  {{
    "name": "従業員名",
    "employment_type": "正社員",
    "monthly_wages": [430000, 316000, null, null, null, null, null, null, null, null, null, null],
    "monthly_hours": [160, 160, null, null, null, null, null, null, null, null, null, null]
  }}
]
```

【重要な注意】
- monthly_wages / monthly_hours は **必ず12要素** の配列にしてください。データがない月は null を入れる。
- 金額は **円単位の整数**。コンマや「円」記号は付けないでください。
- 役員報酬は役員として抽出してください（employment_type に「役員」を含める）。
- 名前のフリガナや空欄行は無視してください。
- JSON以外のコメント・説明文は一切含めないでください。

【賃金台帳データ】
{tsv_data}
"""

PROMPT_WAGE_LEDGER_FISCAL_FILTER = """【前事業年度フィルタ】
納税証明書から判定された前事業年度の決算期は **{fiscal_period}** です。
賃金台帳に複数年度のデータが含まれている場合は、この期間に該当する12ヶ月分のデータのみ抽出してください。
それ以外の月のデータは monthly_wages / monthly_hours に含めないでください（該当月のセルはあっても null）。

ただし、賃金台帳が既に前事業年度の12ヶ月分のみで構成されている場合（例: 「R6.5月」〜「R7.4月」の12列のみ）は、すべてのデータを抽出してください。
"""

PROMPT_WAGE_LEDGER_NO_FILTER = """【期間フィルタ】
納税証明書からの決算期情報は提供されていません。賃金台帳に登場するすべての月のデータを抽出してください。
複数年度に渡る場合は、各従業員について **直近12ヶ月** のデータを優先してください。
"""


class BaseExtractor(ABC):
    """データ抽出の基底クラス"""

    @abstractmethod
    def extract_registry(self, images: list[bytes]) -> CompanyInfo:
        """履歴事項全部証明書から会社情報を抽出"""
        ...

    @abstractmethod
    def extract_pl(self, images: list[bytes]) -> FinancialData:
        """損益計算書から財務データを抽出"""
        ...

    @abstractmethod
    def extract_tax(self, images: list[bytes]) -> TaxCertificate:
        """納税証明書からデータを抽出"""
        ...

    @abstractmethod
    def extract_wages(self, images: list[bytes], year_month: str) -> MonthlyWageData:
        """給与支給控除一覧から従業員データを抽出"""
        ...

    @abstractmethod
    def extract_estimate(self, images: list[bytes]) -> EstimateData:
        """見積書からデータを抽出"""
        ...

    @abstractmethod
    def generate_ai_judgment(self, company: CompanyInfo, financial: FinancialData,
                              tool_name: str, hearing_data: dict | None = None) -> AIJudgment:
        """AI判断項目を生成"""
        ...

    @abstractmethod
    def extract_wage_ledger(
        self,
        tsv_data: str,
        fiscal_period_hint: str | None = None,
    ) -> list[dict]:
        """賃金台帳のTSVテキストから従業員データを抽出。

        Args:
            tsv_data: 全シートをTSV形式で結合したテキスト
            fiscal_period_hint: 前事業年度の決算期（例: 'R6.5-R7.4' または '2024-05〜2025-04'）

        Returns:
            従業員データのリスト。各要素は {name, employment_type, monthly_wages[12], monthly_hours[12]}
        """
        ...


class StubExtractor(BaseExtractor):
    """
    APIキーなしで動作するスタブ。
    全フィールドにプレースホルダ値 '[要API: xxx]' を設定。
    """

    STUB_MARKER = '[要API]'

    def extract_registry(self, images: list[bytes]) -> CompanyInfo:
        logger.warning(f'{self.STUB_MARKER} 履歴事項の読取にはClaude APIが必要です')
        return CompanyInfo(
            name=f'{self.STUB_MARKER} 法人名',
            name_kana=f'{self.STUB_MARKER} フリガナ',
            address=f'{self.STUB_MARKER} 所在地',
            established_date=None,
            capital=0,
            representative_name=f'{self.STUB_MARKER} 代表者名',
            representative_title='代表取締役',
            officers=[],
            business_purposes=[],
        )

    def extract_pl(self, images: list[bytes]) -> FinancialData:
        logger.warning(f'{self.STUB_MARKER} 損益計算書の読取にはClaude APIが必要です')
        return FinancialData()

    def extract_tax(self, images: list[bytes]) -> TaxCertificate:
        logger.warning(f'{self.STUB_MARKER} 納税証明書の読取にはClaude APIが必要です')
        return TaxCertificate()

    def extract_wages(self, images: list[bytes], year_month: str) -> MonthlyWageData:
        logger.warning(f'{self.STUB_MARKER} 給与データの読取にはClaude APIが必要です')
        return MonthlyWageData(year_month=year_month)

    def extract_estimate(self, images: list[bytes]) -> EstimateData:
        logger.warning(f'{self.STUB_MARKER} 見積書の読取にはClaude APIが必要です')
        return EstimateData()

    def generate_ai_judgment(self, company, financial, tool_name, hearing_data=None) -> AIJudgment:
        logger.warning(f'{self.STUB_MARKER} AI判断にはClaude APIが必要です')

        # 営業利益の符号だけで判定できる部分はスタブでも埋める
        is_profitable = financial.operating_profit > 0 if financial.operating_profit else False
        return AIJudgment(
            industry_code=f'{self.STUB_MARKER}',
            industry_text=f'{self.STUB_MARKER}',
            business_description=f'{self.STUB_MARKER} 事業内容（250-255文字）',
            management_intent=(
                '■事業の拡大に積極的\n□事業の維持に注力\n□事業の売却・整備・廃業を考えている\n□特に意識したことは無い'
                if is_profitable else
                '□事業の拡大に積極的\n■事業の維持に注力\n□事業の売却・整備・廃業を考えている\n□特に意識したことは無い'
            ),
            future_goals=(
                '■事業の拡大\n□利益の確保' if is_profitable else '□事業の拡大\n■利益の確保'
            ),
            security_status=(
                '□緊急時の対応マニュアルや手順を定め、定期的に訓練を行っている\n'
                '■パソコンやサーバなどには、IDやパスワードを設け情報セキュリティ管理を行っている\n'
                '□セキュリティ対策は講じていないため、対策を講じていく\n'
                '□セキュリティ対策を講じておらず、今後もその予定はない'
            ),
            business_types=f'{self.STUB_MARKER}',
            it_investment_status='■今までIT投資を行っていなかった',
            it_utilization_status='■ITツールを導入しておらず、今回が初めてである',
        )

    def extract_wage_ledger(
        self,
        tsv_data: str,
        fiscal_period_hint: str | None = None,
    ) -> list[dict]:
        logger.warning(f'{self.STUB_MARKER} 賃金台帳のAI抽出にはClaude APIが必要です')
        return []


class ClaudeExtractor(BaseExtractor):
    """Claude API による実データ抽出"""

    def __init__(
        self,
        api_key: str,
        model: str = 'claude-sonnet-4-6',
        retry_callback: Optional[RetryCallback] = None,
    ):
        try:
            import anthropic
        except ImportError:
            raise ImportError('anthropic パッケージが必要です: pip install anthropic')

        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model
        self.retry_callback = retry_callback
        logger.info(f'Claude API 初期化完了 (model={model})')

    def _messages_create_with_retry(self, *, caller: str, stats: str, **kwargs):
        """messages.create を指数バックオフ付きで呼び出す。

        - 422/429/5xx/529/timeout/connection エラーは最大3回まで再試行
        - 400 credit_balance_too_low は APICreditExhaustedError に変換して即失敗
        - その他の 400/401/403/404/413 は即失敗
        - 再試行時は retry_callback(attempt, max_attempts, wait, err_summary) を呼ぶ
        """
        import anthropic

        last_error: Optional[Exception] = None
        for attempt in range(1, MAX_API_ATTEMPTS + 1):
            try:
                return self.client.messages.create(**kwargs)

            except anthropic.BadRequestError as e:
                # 400: 残高切れだけは専用例外に、それ以外は即失敗（リトライしても無駄）
                if CREDIT_BALANCE_MARKER in str(e).lower():
                    logger.error(f'[API残高切れ] caller={caller} {stats}')
                    raise APICreditExhaustedError(
                        'APIの残高が不足しています。村上さんにチャージを依頼してください。'
                    ) from e
                logger.error(f'[API失敗/400] caller={caller} {stats} error={e}')
                raise

            except (anthropic.AuthenticationError,
                    anthropic.PermissionDeniedError,
                    anthropic.NotFoundError) as e:
                # 401/403/404: 設定ミス系、リトライ無意味
                logger.error(f'[API失敗/非リトライ] caller={caller} {stats} error={type(e).__name__}: {e}')
                raise

            except anthropic.APIStatusError as e:
                # 422/429/5xx/529 などステータスコード付きエラー
                status = getattr(e, 'status_code', None)
                if status in RETRYABLE_STATUS_CODES and attempt < MAX_API_ATTEMPTS:
                    last_error = e
                    wait = API_BACKOFF_SECONDS[attempt - 1]
                    err_summary = f'{status} {type(e).__name__}'
                    logger.warning(
                        f'[API再試行] caller={caller} {attempt}/{MAX_API_ATTEMPTS} '
                        f'wait={wait}s error={err_summary}: {e}'
                    )
                    if self.retry_callback:
                        try:
                            self.retry_callback(attempt, MAX_API_ATTEMPTS, wait, err_summary)
                        except Exception as cb_err:
                            logger.warning(f'retry_callback実行失敗: {cb_err}')
                    time.sleep(wait)
                    continue
                logger.error(f'[API失敗/確定] caller={caller} {stats} status={status} error={e}')
                raise

            except (anthropic.APITimeoutError, anthropic.APIConnectionError) as e:
                # ネットワーク系の一時障害もリトライ
                if attempt < MAX_API_ATTEMPTS:
                    last_error = e
                    wait = API_BACKOFF_SECONDS[attempt - 1]
                    err_summary = type(e).__name__
                    logger.warning(
                        f'[API再試行] caller={caller} {attempt}/{MAX_API_ATTEMPTS} '
                        f'wait={wait}s error={err_summary}: {e}'
                    )
                    if self.retry_callback:
                        try:
                            self.retry_callback(attempt, MAX_API_ATTEMPTS, wait, err_summary)
                        except Exception as cb_err:
                            logger.warning(f'retry_callback実行失敗: {cb_err}')
                    time.sleep(wait)
                    continue
                logger.error(f'[API失敗/確定] caller={caller} {stats} error={e}')
                raise

        # ループを抜けた = リトライ全敗（通常到達しない。安全網）
        if last_error:
            raise last_error
        raise RuntimeError('API呼出しリトライが想定外に終了しました')

    def _call_api(self, images: list[bytes], prompt: str, max_tokens: int = 4096) -> str:
        """画像+プロンプトでAPIを呼び出し、テキストを返す"""
        import base64
        import traceback
        content = []

        raw_sizes = []
        b64_sizes = []
        for img in images:
            b64 = base64.standard_b64encode(img).decode('ascii')
            raw_sizes.append(len(img))
            b64_sizes.append(len(b64))
            content.append({
                'type': 'image',
                'source': {'type': 'base64', 'media_type': 'image/png', 'data': b64}
            })

        content.append({'type': 'text', 'text': prompt})

        # 送信直前のペイロード統計（422/413/529 の原因切り分け用）
        n = len(images)
        raw_mb = sum(raw_sizes) / 1_000_000
        b64_mb = sum(b64_sizes) / 1_000_000
        raw_max = max(raw_sizes) / 1_000_000 if raw_sizes else 0
        prompt_chars = len(prompt)
        caller = traceback.extract_stack()[-2].name  # extract_tax 等、どのメソッドからの呼び出しか
        stats = (
            f'images={n}枚 raw合計={raw_mb:.2f}MB raw最大={raw_max:.2f}MB '
            f'base64合計={b64_mb:.2f}MB prompt={prompt_chars}chars max_tokens={max_tokens}'
        )
        logger.warning(f'[API送信] caller={caller} {stats}')

        response = self._messages_create_with_retry(
            caller=caller,
            stats=stats,
            model=self.model,
            max_tokens=max_tokens,
            messages=[{'role': 'user', 'content': content}],
        )

        text = response.content[0].text
        logger.warning(
            f'[API成功] caller={caller} '
            f'応答={len(text)}chars '
            f'tokens={response.usage.input_tokens}in+{response.usage.output_tokens}out'
        )
        return text

    def _parse_json(self, text: str) -> dict | list:
        """API応答からJSONを抽出・パース"""
        # ```json ... ``` ブロックがあれば中身を取り出す
        if '```json' in text:
            start = text.index('```json') + 7
            end = text.index('```', start)
            text = text[start:end].strip()
        elif '```' in text:
            start = text.index('```') + 3
            end = text.index('```', start)
            text = text[start:end].strip()

        return json.loads(text)

    def extract_registry(self, images: list[bytes]) -> CompanyInfo:
        text = self._call_api(images, PROMPT_REGISTRY)
        data = self._parse_json(text)

        # 役員リスト（同一人物の重複を排除）
        officers = []
        seen_names = set()
        rep_name = data.get('representative_name', '')
        for o in data.get('officers', []):
            name = o.get('name', '').strip()
            if not name or name in seen_names or name == rep_name:
                continue
            seen_names.add(name)
            officers.append({
                'title': o.get('title', ''),
                'name': name,
                'kana': o.get('kana', ''),
            })

        from datetime import datetime
        est = None
        if data.get('established_date'):
            try:
                est = datetime.strptime(data['established_date'], '%Y-%m-%d')
            except ValueError:
                pass

        return CompanyInfo(
            name=data.get('name') or '',
            name_kana=data.get('name_kana') or '',
            address=data.get('address') or '',
            postal_code=data.get('postal_code') or '',
            established_date=est,
            capital=data.get('capital', 0) or 0,
            representative_name=data.get('representative_name') or '',
            representative_title=data.get('representative_title') or '',
            officers=officers,
            business_purposes=data.get('business_purposes') or [],
        )

    def extract_pl(self, images: list[bytes]) -> FinancialData:
        text = self._call_api(images, PROMPT_PL)
        d = self._parse_json(text)

        # 決算月を事業年度終了日から推定
        fiscal_month = ''
        if d.get('fiscal_year_end'):
            month = d['fiscal_year_end'].split('-')[1] if '-' in d['fiscal_year_end'] else ''
            month_names = {'01': '1月', '02': '2月', '03': '3月', '04': '4月',
                          '05': '5月', '06': '6月', '07': '7月', '08': '8月',
                          '09': '9月', '10': '10月', '11': '11月', '12': '12月'}
            fiscal_month = month_names.get(month, '')

        return FinancialData(
            fiscal_year_start=d.get('fiscal_year_start', ''),
            fiscal_year_end=d.get('fiscal_year_end', ''),
            fiscal_month=fiscal_month,
            revenue=d.get('revenue', 0) or 0,
            cost_of_sales=d.get('cost_of_sales', 0) or 0,
            gross_profit=d.get('gross_profit', 0) or 0,
            operating_profit=d.get('operating_profit', 0) or 0,
            ordinary_profit=d.get('ordinary_profit', 0) or 0,
            net_profit=d.get('net_profit', 0) or 0,
            salary=d.get('salary', 0) or 0,
            misc_wages=d.get('misc_wages', 0) or 0,
            bonus=d.get('bonus', 0) or 0,
            officer_compensation=d.get('officer_compensation', 0) or 0,
            legal_welfare=d.get('legal_welfare', 0) or 0,
            welfare=d.get('welfare', 0) or 0,
            depreciation=d.get('depreciation', 0) or 0,
            travel_expense=d.get('travel_expense', 0) or 0,
        )

    def extract_tax(self, images: list[bytes]) -> TaxCertificate:
        text = self._call_api(images, PROMPT_TAX)
        d = self._parse_json(text)
        return TaxCertificate(
            tax_type=d.get('tax_type', ''),
            tax_amount=d.get('tax_amount', 0) or 0,
            fiscal_year=d.get('fiscal_year', ''),
        )

    def extract_wages(self, images: list[bytes], year_month: str) -> MonthlyWageData:
        text = self._call_api(images, PROMPT_WAGES, max_tokens=8192)
        data = self._parse_json(text)

        employees = []
        for e in data:
            employees.append(Employee(
                name=e.get('name', ''),
                department=e.get('department', ''),
                employee_id=e.get('employee_id', ''),
                employment_type=e.get('employment_type', ''),
                working_days=e.get('working_days', 0) or 0,
                scheduled_hours=e.get('scheduled_hours', 0) or 0,
                base_salary=e.get('base_salary', 0) or 0,
                taxable_total=e.get('taxable_total', 0) or 0,
                total_pay=e.get('total_pay', 0) or 0,
                deductions=e.get('deductions', 0) or 0,
                net_pay=e.get('net_pay', 0) or 0,
            ))

        return MonthlyWageData(year_month=year_month, employees=employees)

    def extract_estimate(self, images: list[bytes]) -> EstimateData:
        text = self._call_api(images, PROMPT_ESTIMATE)
        d = self._parse_json(text)

        items = [{'name': i.get('name', ''), 'amount': i.get('amount', 0)}
                 for i in d.get('items', [])]

        return EstimateData(
            vendor_name=d.get('vendor_name', ''),
            tool_name=d.get('tool_name', ''),
            items=items,
            total_amount=d.get('total_amount', 0) or 0,
            tax_amount=d.get('tax_amount', 0) or 0,
        )

    def generate_ai_judgment(self, company, financial, tool_name, hearing_data=None) -> AIJudgment:
        # ヒアリングデータから各種情報を取得
        hearing_fields = {
            'it_investment_answer': '不明',
            'it_investment_amount': '不明',
            'it_investment_process': '不明',
            'main_business': '',
            'strength': '',
            'challenge': '',
            'monthly_hours': '',
            'tool_usage': '',
            'reduction': '',
            'freed_time': '',
            'sales_target': '',
        }
        if hearing_data:
            FIELD_KEYWORDS = {
                'main_business': ['主な事業内容'],
                'strength': ['強み'],
                'challenge': ['時間がかかっている'],
                'monthly_hours': ['月間何時間'],
                'tool_usage': ['どの機能'],
                'reduction': ['何％', '何時間'],
                'freed_time': ['浮いた時間'],
                'sales_target': ['売上目標'],
            }
            for row_num, item in hearing_data.items():
                label = str(item.get('label', ''))
                value = item.get('value')
                if 'IT投資' in label and '金額' in label:
                    hearing_fields['it_investment_answer'] = 'はい' if value else 'いいえ'
                    hearing_fields['it_investment_amount'] = str(value) if value else 'なし'
                elif 'IT投資' in label and 'プロセス' in label:
                    hearing_fields['it_investment_process'] = str(value) if value else 'なし'
                else:
                    for field_key, keywords in FIELD_KEYWORDS.items():
                        if any(kw in label for kw in keywords):
                            hearing_fields[field_key] = str(value) if value else ''
                            break

        prompt = PROMPT_AI_JUDGMENT.format(
            company_name=company.name,
            business_purposes=', '.join(company.business_purposes),
            address=company.address,
            operating_profit=financial.operating_profit,
            tool_name=tool_name,
            **hearing_fields,
        )

        # AI判断はテキストのみ（画像なし）
        stats = f'images=0枚 prompt={len(prompt)}chars max_tokens=4096'
        logger.warning(f'[API送信] caller=generate_ai_judgment {stats}')
        response = self._messages_create_with_retry(
            caller='generate_ai_judgment',
            stats=stats,
            model=self.model,
            max_tokens=4096,
            messages=[{'role': 'user', 'content': prompt}],
        )
        text = response.content[0].text
        logger.warning(
            f'[API成功] caller=generate_ai_judgment '
            f'応答={len(text)}chars '
            f'tokens={response.usage.input_tokens}in+{response.usage.output_tokens}out'
        )
        d = self._parse_json(text)

        # 最低賃金はconfig.pyから取得
        from .config import get_min_wage
        mw = get_min_wage(company.address)
        min_wage_text = f'{mw[0]}/{mw[1]}円' if mw else d.get('min_wage', '')

        return AIJudgment(
            industry_code=d.get('industry_code', ''),
            industry_text=d.get('industry_text', ''),
            business_description=d.get('business_description', ''),
            management_intent=d.get('management_intent', ''),
            future_goals=d.get('future_goals', ''),
            security_status=d.get('security_status', ''),
            business_types=d.get('business_types', ''),
            min_wage_text=min_wage_text,
            it_investment_status=d.get('it_investment_status', ''),
            it_utilization_status=d.get('it_utilization_status', ''),
            it_utilization_scope=d.get('it_utilization_scope', ''),
            invoice_related_work=d.get('invoice_related_work', ''),
        )

    def extract_wage_ledger(
        self,
        tsv_data: str,
        fiscal_period_hint: str | None = None,
    ) -> list[dict]:
        if fiscal_period_hint:
            fiscal_section = PROMPT_WAGE_LEDGER_FISCAL_FILTER.format(
                fiscal_period=fiscal_period_hint
            )
        else:
            fiscal_section = PROMPT_WAGE_LEDGER_NO_FILTER

        prompt = PROMPT_WAGE_LEDGER.format(
            tsv_data=tsv_data,
            fiscal_period_section=fiscal_section,
        )

        # 出力JSONは従業員数に比例して大きくなるため最大16Kトークン
        max_tokens = 16384

        stats = f'images=0枚 prompt={len(prompt)}chars max_tokens={max_tokens}'
        logger.warning(f'[API送信] caller=extract_wage_ledger {stats}')
        response = self._messages_create_with_retry(
            caller='extract_wage_ledger',
            stats=stats,
            model=self.model,
            max_tokens=max_tokens,
            messages=[{'role': 'user', 'content': prompt}],
        )
        text = response.content[0].text
        logger.warning(
            f'[API成功] caller=extract_wage_ledger '
            f'応答={len(text)}chars '
            f'tokens={response.usage.input_tokens}in+{response.usage.output_tokens}out'
        )

        try:
            data = self._parse_json(text)
        except json.JSONDecodeError as e:
            logger.error(f'[extract_wage_ledger] JSON解析失敗: {e}, 応答先頭500文字: {text[:500]}')
            return []

        if not isinstance(data, list):
            logger.error(f'[extract_wage_ledger] 応答がリストではありません: type={type(data).__name__}')
            return []

        return data


def create_extractor(
    api_key: str = '',
    retry_callback: Optional[RetryCallback] = None,
) -> BaseExtractor:
    """APIキーの有無に応じて適切なExtractorを返す"""
    if api_key:
        logger.info('Claude API Extractor を使用')
        return ClaudeExtractor(api_key, retry_callback=retry_callback)
    else:
        logger.warning('APIキー未設定 → StubExtractor を使用（PDF読取不可）')
        return StubExtractor()
