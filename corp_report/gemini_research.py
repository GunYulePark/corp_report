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

# Google Search grounding returns the source URLs in groundingMetadata rather
# than in the fixed source pack.  Keep the model output compact and validate
# every URL against that metadata before it can become a report fact.
GROUNDED_ISSUE_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "items": {
            "type": "array",
            "maxItems": 5,
            "items": {
                "type": "object",
                "properties": {
                    "category": {"type": "string"},
                    "fact": {"type": "string"},
                    "interpretation": {"type": "string"},
                    "source_urls": {"type": "array", "items": {"type": "string"}, "minItems": 1, "maxItems": 3},
                },
                "required": ["category", "fact", "interpretation", "source_urls"],
            },
        },
    },
    "required": ["items"],
}

CORPORATE_CITED_TEXT: dict[str, Any] = {
    "type": "object",
    "properties": {
        "text": {"type": "string"},
        "source_ids": {"type": "array", "items": {"type": "string"}, "minItems": 1, "maxItems": 2},
    },
    "required": ["text", "source_ids"],
}

CORPORATE_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "business_summary": CORPORATE_CITED_TEXT,
        "growth_strategy": CORPORATE_CITED_TEXT,
        "core_competencies": {"type": "array", "items": CORPORATE_CITED_TEXT, "maxItems": 4},
        "chronology": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "date": {"type": "string"},
                    "event": {"type": "string"},
                    "source_ids": {"type": "array", "items": {"type": "string"}, "minItems": 1, "maxItems": 2},
                },
                "required": ["date", "event", "source_ids"],
            },
            "maxItems": 6,
        },
    },
    "required": ["business_summary", "growth_strategy", "core_competencies", "chronology"],
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


def _response_payload(response: requests.Response) -> dict[str, Any] | None:
    """Read a raw Gemini response while retaining grounding metadata."""
    try:
        payload = response.json()
    except (TypeError, ValueError):
        return None
    return payload if isinstance(payload, dict) else None


def _grounding_sources(payload: dict[str, Any]) -> dict[str, dict[str, str]]:
    """Normalize Gemini Google Search citations from the documented metadata."""
    try:
        metadata = payload["candidates"][0].get("groundingMetadata", {})
    except (KeyError, IndexError, TypeError):
        return {}
    sources: dict[str, dict[str, str]] = {}
    for chunk in metadata.get("groundingChunks", []):
        if not isinstance(chunk, dict):
            continue
        web = chunk.get("web", {})
        if not isinstance(web, dict):
            continue
        url = str(web.get("uri", "")).strip()
        if not url.startswith(("http://", "https://")):
            continue
        sources[url.rstrip("/")] = {
            "url": url,
            "title": str(web.get("title", "웹 검색 근거")),
            "date": str(web.get("publishedDate", "")),
        }
    return sources


def _grounded_to_matters(payload: dict[str, Any]) -> list[MatterFact]:
    """Convert a grounded JSON response to reviewable source-linked facts only."""
    try:
        raw_text = payload["candidates"][0]["content"]["parts"][0]["text"]
        structured = json.loads(raw_text)
    except (KeyError, IndexError, TypeError, ValueError, json.JSONDecodeError):
        return []
    if not isinstance(structured, dict):
        return []
    source_index = _grounding_sources(payload)
    if not source_index:
        return []
    results: list[MatterFact] = []
    for index, item in enumerate(structured.get("items", []), start=1):
        if not isinstance(item, dict):
            continue
        fact = str(item.get("fact", "")).strip()
        interpretation = str(item.get("interpretation", "")).strip()
        category = str(item.get("category", "이슈 분석")).strip()
        if not fact or not interpretation:
            continue
        citation = next(
            (source_index[url.rstrip("/")] for url in item.get("source_urls", []) if isinstance(url, str) and url.rstrip("/") in source_index),
            None,
        )
        if citation is None:
            continue
        # Search grounding gives a citable web lead, not a filing-level audit
        # opinion. Keep it reviewable until the linked original is confirmed.
        results.append(
            MatterFact(
                category=f"Gemini 검색 기반·{category}",
                fact=fact,
                interpretation=interpretation,
                verification_status="needs_review",
                source_document_id=f"gemini-grounded-{index}",
                source_title=citation["title"],
                disclosure_date=citation["date"],
                url=citation["url"],
            )
        )
    return results


def _grounded_issue_prompt(company_name: str, issue_query: str) -> str:
    return f"""당신은 한국 상장사 기업분석 보고서의 이슈 리서치 애널리스트다.
회사: {company_name}
사용자 요청: {issue_query}

Google Search로 요청과 직접 관련된 최근 기사, 회사 IR·보도자료, DART 공시를 조사하라.
계약 종료·판매권 변동·매출·영업이익의 관계는 수치와 기간이 인용 자료에 명시된 경우에만 서술하라.
확정 매출, 계약 규모, 예상 영향, 기사 해석을 혼동하지 말고 비교 기준·한계를 해석에 포함하라.
반환값은 JSON만 사용한다. 각 item의 source_urls에는 Google Search grounding 결과에 실제 포함된 원문 URL만 1~3개 넣어라.
출처 URL이 없는 추정·의견은 item에 포함하지 말라.
"""


def _analyze_grounded_issue(company_name: str, issue_query: str, api_key: str) -> list[MatterFact]:
    """Use Gemini's managed Google Search and retain only returned citations."""
    if not api_key.strip() or not issue_query.strip():
        return []
    model = os.getenv("GEMINI_GROUNDED_MODEL", os.getenv("GEMINI_MODEL", DEFAULT_MODEL)).strip() or DEFAULT_MODEL
    body = {
        "contents": [{"role": "user", "parts": [{"text": _grounded_issue_prompt(company_name, issue_query)}]}],
        "tools": [{"google_search": {}}],
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseSchema": GROUNDED_ISSUE_SCHEMA,
            "temperature": 0.1,
        },
    }
    try:
        response = requests.post(
            API_URL.format(model=model),
            params={"key": api_key},
            headers=USER_AGENT,
            json=body,
            timeout=60,
        )
        response.raise_for_status()
    except requests.RequestException:
        return []
    payload = _response_payload(response)
    return _grounded_to_matters(payload) if payload else []


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
    """Prefer Gemini Google Search grounding; preserve source-pack fallback."""
    key = api_key.strip() or os.getenv("GEMINI_API_KEY", "").strip()
    if not key or not issue_query.strip():
        return []
    grounded = _analyze_grounded_issue(company_name, issue_query, key)
    if grounded:
        return grounded
    if not sources:
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


def analyze_corporate_profile(company_name: str, sources: list[MatterFact], api_key: str = "") -> dict[str, Any]:
    """Summarize business content and timeline, constrained to annual-report text."""
    key = api_key.strip() or os.getenv("GEMINI_API_KEY", "").strip()
    if not key or not sources:
        return {}
    model = os.getenv("GEMINI_MODEL", DEFAULT_MODEL).strip() or DEFAULT_MODEL
    source_text = json.dumps(_source_payload(sources), ensure_ascii=False, separators=(",", ":"))
    prompt = f"""당신은 한국 기업분석 보고서 작성자다. 회사는 {company_name}이다.
아래 SOURCE는 사업보고서 또는 감사보고서의 회사 개요·사업의 내용·회사 연혁 및 공식 홈페이지 원문에서 수집했다. SOURCE만 근거로 JSON을 작성하라.
사업개요는 2문장 이내, 성장전략은 2문장 이내, 핵심역량은 최대 4개로 작성한다.
business_summary·growth_strategy·core_competencies·연혁의 모든 항목에 정확한 source_ids를 하나 이상 넣어야 한다.
성장전략은 신약개발, 투자, 증설, 시장확대처럼 명시된 전략만 쓴다. 사내 업무도구·일반 규정·무관한 운영 문구는 전략이나 핵심역량으로 쓰지 말고, 근거가 부족하면 text에 '확인 필요'라고 쓴다.
SOURCE에 없는 수치·사업·연도는 쓰지 말라.

SOURCE:
{source_text}"""
    body = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"responseMimeType": "application/json", "responseSchema": CORPORATE_SCHEMA, "temperature": 0.1},
    }
    try:
        response = requests.post(API_URL.format(model=model), params={"key": key}, headers=USER_AGENT, json=body, timeout=45)
        response.raise_for_status()
    except requests.RequestException:
        return {}
    payload = _response_json(response)
    if not payload:
        return {}
    known_ids = {source.source_document_id for source in sources}
    def cited_text(field_name: str) -> str:
        item = payload.get(field_name, {})
        if not isinstance(item, dict) or not any(source_id in known_ids for source_id in item.get("source_ids", [])):
            return ""
        text = str(item.get("text", "")).strip()
        return text if text and text != "확인 필요" else ""

    competencies = []
    for item in payload.get("core_competencies", []):
        if not isinstance(item, dict) or not any(source_id in known_ids for source_id in item.get("source_ids", [])):
            continue
        text = str(item.get("text", "")).strip()
        if text and text != "확인 필요":
            competencies.append(text)

    chronology = []
    for item in payload.get("chronology", []):
        if not isinstance(item, dict) or not any(source_id in known_ids for source_id in item.get("source_ids", [])):
            continue
        date_value, event = str(item.get("date", "")).strip(), str(item.get("event", "")).strip()
        if date_value and event:
            chronology.append({"date": date_value, "event": event, "source_ids": [source_id for source_id in item["source_ids"] if source_id in known_ids]})
    return {
        "business_summary": cited_text("business_summary"),
        "growth_strategy": cited_text("growth_strategy"),
        "core_competencies": competencies[:4],
        "chronology": chronology[:6],
        "source": "Gemini 구조화 요약 · 최신 사업보고서 원문 기반",
    }
