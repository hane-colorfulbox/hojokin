# -*- coding: utf-8 -*-
"""賃金台帳 AI 抽出の record / replay 基盤（回帰テスト & フィクスチャ更新用）。

`ClaudeExtractor._messages_create_with_retry` に recorder を差し込むと、
実 API を呼ぶ代わりに記録済み応答を返す（replay）か、実応答を保存する（record）。

- ReplayRecorder: テストで使う。ネットワーク・課金ゼロで過去の抽出結果を再生する。
- RecordingRecorder: refresh_fixtures（実 API でフィクスチャを録り直す）で使う。

recorder 未指定（None）時は従来どおり実 API を直接叩くため、本番挙動は不変。
"""
from __future__ import annotations

from types import SimpleNamespace
from typing import Callable


class ReplayMessage:
    """anthropic.Message の最小互換オブジェクト。

    extract_* / _call_api が参照する属性だけを持つ:
      - content[0].text
      - usage.input_tokens / output_tokens / cache_creation_input_tokens /
        cache_read_input_tokens
      - stop_reason
    """

    def __init__(
        self,
        text: str,
        *,
        input_tokens: int = 0,
        output_tokens: int = 0,
        cache_creation_input_tokens: int = 0,
        cache_read_input_tokens: int = 0,
        stop_reason: str = 'end_turn',
    ):
        self.content = [SimpleNamespace(type='text', text=text)]
        self.usage = SimpleNamespace(
            input_tokens=input_tokens,
            output_tokens=output_tokens,
            cache_creation_input_tokens=cache_creation_input_tokens,
            cache_read_input_tokens=cache_read_input_tokens,
        )
        self.stop_reason = stop_reason


class ReplayRecorder:
    """記録済み応答テキストを呼び出し順に返す（実 API・ネットワーク不使用）。

    用意した応答数を超えて API が呼ばれた場合は AssertionError を送出する。
    これにより「録画した想定より多く API を叩いた（=分割再抽出などの想定外経路に
    入った）」回帰を検出できる。
    """

    def __init__(self, responses: list[str] | str):
        if isinstance(responses, str):
            responses = [responses]
        self._responses = list(responses)
        self._idx = 0
        self.calls: list[dict] = []  # caller / model を記録（検証用）

    def intercept(self, *, caller: str, real_call: Callable, **kwargs) -> ReplayMessage:
        self.calls.append({'caller': caller, 'model': kwargs.get('model')})
        if self._idx >= len(self._responses):
            raise AssertionError(
                f'ReplayRecorder: 想定外の追加 API 呼び出し caller={caller} '
                f'（用意した応答 {len(self._responses)} 件を超過）'
            )
        text = self._responses[self._idx]
        self._idx += 1
        return ReplayMessage(text)


class RecordingRecorder:
    """実 API を呼びつつ応答テキストを保存する（refresh_fixtures 用）。

    保存した recorded[i] をフィクスチャの ai_response として書き出すことで、
    現行モデル/プロンプトでの抽出結果を録り直してドリフトを検知できる。
    """

    def __init__(self):
        self.recorded: list[str] = []

    def intercept(self, *, caller: str, real_call: Callable, **kwargs):
        msg = real_call()
        try:
            self.recorded.append(msg.content[0].text)
        except (AttributeError, IndexError):
            pass
        return msg
