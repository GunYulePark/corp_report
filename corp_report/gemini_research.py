from __future__ import annotations

"""Gemini-backed, source-constrained issue analysis.

The model receives only facts already collected by the local pipeline. It must
cite supplied source IDs for every item; outputs without a valid source ID are
dropped before they can enter the Fact Pack.
"""

import json
import os
from typing import Any

import requests

from .models import MatterFact


DEFAULT_MODEL = "gemini-3.5-flash-lite"
API_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
USER_AGENT = {"User-Agent": "CorporateReportAutomation/1.0", "Content-Type": "application/json"}

ISSUE_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "items": {
            "type": "array",
            "maxItems": 6,
            "items": {
                "type": "object",
                "properties": {
                    "category": {"type": "string"},
                    "fact": {"type": "string"},
                    "interpretation": {"type": "string"},
                    "source_ids": {"type": "array", "items": {"type": "string"}, "minItems": 1, "maxItems": 3},
                    "verification_status": {"type": "string", "enum": ["verified", "needs_review"]},
                },
                "required": ["category", "fact", "interpretation", "source_ids", "verification_status"],
            },
        },
    },
    "required": ["items"],
}


def _source_payload(matters: list[MatterFact]) -> list[dict[str, str]]:
    return [
        {
            "source_id": item.source_document_id,
            "category": item.category,
            "fact": item.fact,
            "interpretation": item.interpretation,
            "verification_status": item.verification_status,
            "source_title": item.source_title,
            "date": item.disclosure_date,
            "url": item.url,
        }
        for item in matters[:12]
    ]


def _prompt(company_name: str, issue_query: str, matters: list[MatterFact]) -> str:
    sources = json.dumps(_source_payload(matters), ensure_ascii=False, separators=(",", ":"))
    return f"""당신은 한국 상장사 기업분석 보고서의 이슈 리서치 애널리스트다.
회사: {company_name}
사용자 요청: {issue_query}

아래 SOURCE 목록만 근거로 사용해 분석 JSON을 작성하라. 외부 지식, 추정, 숫자 보완은 금지한다.
각 items 항목은 source_ids에 SOURCE의 source_id를 하나 이상 정확히 넣어야 한다.
SOURCE의 verification_status가 needs_review이면 그 항목도 needs_review으로 유지한다.
계약 최대금액과 선급금, 확정 매출을 혼동하지 말고, 비교 대상의 후보 수·개발 단계·권리범위가 다르면 비교 한계를 해석에 명시한다.
사실은 짧고 검증 가능하게, 해석은 사업·재무상 시사점과 한계를 포함해 작성한다.
출처로 뒷받침되지 않는 문장은 items에 넣지 말라.

SOURCE:
{sources}"""


def _response_json(response: requests.Response) -> dict[str, Any] | None:
    try:
        payload = response.json()
        text = payload["candidates"][0]["content"]["parts"][0]["text"]
        parsed = json.loads(text)
    except (KeyError, IndexError, TypeError, ValueError, json.JSONDecodeError):
        return None
    return parsed if isinstance(parsed, dict) else None


def _to_matters(payload: dict[str, Any], sources: list[MatterFact]) -> list[MatterFact]:
    source_index = {item.source_document_id: item for item in sources}
    results: list[MatterFact] = []
    for item in payload.get("items", []):
        if not isinstance(item, dict):
            continue
        source_ids = [value for value in item.get("source_ids", []) if value in source_index]
        if not source_ids:
            continue
        source = source_index[source_ids[0]]
        status = item.get("verification_status", "needs_review")
        if source.verification_status != "verified":
            status = "needs_review"
        if status not in {"verified", "needs_review"}:
            status = "needs_review"
        fact = str(item.get("fact", "")).strip()
        interpretation = str(item.get("interpretation", "")).strip()
        category = str(item.get("category", "웹 이슈 조사")).strip()
        if not fact or not interpretation:
            continue
        results.append(
            MatterFact(
                category=f"Gemini 분석·{category}",
                fact=fact,
                interpretation=interpretation,
                verification_status=status,
                source_document_id=source.source_document_id,
                source_title=source.source_title,
                disclosure_date=source.disclosure_date,
                url=source.url,
            )
        )
    return results


def analyze_issue(company_name: str, issue_query: str, sources: list[MatterFact], api_key: str = "") -> list[MatterFact]:
    """Return structured Gemini analysis or an empty list on any API failure."""
    key = api_key.strip() or os.getenv("GEMINI_API_KEY", "").strip()
    if not key or not issue_query.strip() or not sources:
        return []
    model = os.getenv("GEMINI_MODEL", DEFAULT_MODEL).strip() or DEFAULT_MODEL
    request_body = {
        "contents": [{"role": "user", "parts": [{"text": _prompt(company_name, issue_query, sources)}]}],
        "generationConfig": {"responseMimeType": "application/json", "responseSchema": ISSUE_SCHEMA, "temperature": 0.1},
    }
    try:
        response = requests.post(
            API_URL.format(model=model),
            params={"key": key},
            headers=USER_AGENT,
            json=request_body,
            timeout=45,
        )
        response.raise_for_status()
    except requests.RequestException:
        return []
    payload = _response_json(response)
    return _to_matters(payload, sources) if payload else []
