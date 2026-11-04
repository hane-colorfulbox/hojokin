"""画面の前提が変わった後、古い申請結果を現案件の提出物として使わせない。"""
from __future__ import annotations

import hashlib
import json
from dataclasses import asdict, is_dataclass


def processing_identity(**settings) -> str:
    values = {key: asdict(value) if is_dataclass(value) else value for key, value in settings.items()}
    payload = json.dumps(values, ensure_ascii=False, sort_keys=True, default=str, separators=(',', ':'))
    return hashlib.sha256(payload.encode('utf-8')).hexdigest()
