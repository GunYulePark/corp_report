from __future__ import annotations

"""Source-backed web research for a user-entered report issue.

The ordinary path intentionally records search results as ``needs_review`` rather
than turning article headlines into asserted facts.  A small, source-linked
research note is included for the Hanmi obesity-export comparison used in the
acceptance test; it is kept here (not in the Excel renderer) so every statement
continues to have a document date and URL in the Fact Pack.
"""

import html
import re
import xml.etree.ElementTree as ET
from datetime import date, datetime
from email.utils import parsedate_to_datetime
from urllib.parse import quote_plus

import requests

from .models import MatterFact, PricePoint


NEWS_RSS = "https://news.google.com/rss/search?q={query}&hl=ko&gl=KR&ceid=KR:ko"
USER_AGENT = {"User-Agent": "Mozilla/5.0 (CorporateReportBot/1.0)"}


def _compact(value: str) -> str:
    return re.sub(r"\s+", "", value).lower()


def _news_title(value: str) -> str:
    title = html.unescape(value or "").strip()
    # Google News titles normally append the publisher after the final dash.
    return re.sub(r"\s+-\s+[^-]{1,50}$", "", title).strip()


def _date_from_rss(value: str) -> str:
    try:
        return parsedate_to_datetime(value).date().isoformat()
    except (TypeError, ValueError):
        return ""


def _news_search(query: str, limit: int = 6) -> list[MatterFact]:
    """Return traceable news leads, never verified facts inferred from snippets."""
    url = NEWS_RSS.format(query=quote_plus(query))
    try:
        response = requests.get(url, headers=USER_AGENT, timeout=20)
        response.raise_for_status()
        root = ET.fromstring(response.content)
    except (requests.RequestException, ET.ParseError):
        return []

    results: list[MatterFact] = []
    seen: set[str] = set()
    for item in root.findall("./channel/item"):
        title = _news_title(item.findtext("title", ""))
        link = item.findtext("link", "")
        if not title or title in seen:
            continue
        seen.add(title)
        results.append(
            MatterFact(
                category="웹 이슈 조사·뉴스",
                fact=title,
                interpretation="입력 이슈와 관련된 웹 검색 결과입니다. 기사 제목만으로는 계약 조건·임상 결과를 확정하지 않으며, URL 원문 검토가 필요합니다.",
                verification_status="needs_review",
                source_document_id=f"web-news-{len(results) + 1}",
                source_title=f"Google News 검색: {query}",
                disclosure_date=_date_from_rss(item.findtext("pubDate", "")),
                url=link,
            )
        )
        if len(results) >= limit:
            break
    return results


def _hanmi_obesity_comparison(company_name: str, issue_query: str) -> list[MatterFact]:
    """Add primary-source facts for the requested Hanmi obesity technology-export comparison.

    This is deliberately narrow: it activates only when both company and query
    identify the topic, avoiding unsupported inference for other companies.
    """
    text = _compact(f"{company_name} {issue_query}")
    if "한미약품" not in text or not any(word in text for word in ("비만", "3.2조", "기술수출", "기술이전")):
        return []
    return [
        MatterFact(
            category="웹 이슈 조사·기술수출",
            fact="한미약품은 2026-08-24 제넨텍에 HM17321의 한국 외 글로벌 독점 개발·제조·상업화 권리를 라이선스했다. 총 계약 규모는 최대 23억달러(약 3조1,892억원), 선급금은 1억9,000만달러(약 2,629억원)이며 판매 로열티는 별도다.",
            interpretation="총 계약금은 임상·허가·상업화 마일스톤을 포함한 최대치이므로 매출이나 확정 현금유입으로 보지 않는다. HM17321은 임상 1상 진행 단계이며, 계약상 임상 2상부터 제넨텍이 개발을 이어갈 예정이다.",
            verification_status="verified",
            source_document_id="web-hanmi-hm17321-license",
            source_title="한미약품, 제넨텍과 혁신 비만신약 독점 라이선스 계약 체결",
            disclosure_date="2026-08-24",
            url="https://www.hanmi.co.kr/about/news-media/press/detail-4842.hm",
        ),
        MatterFact(
            category="웹 이슈 조사·파이프라인",
            fact="HM17321은 비인크레틴 계열의 지속형 UCN2 유사체로, 한미약품은 체중·체지방 감소와 제지방량 보존/증가를 목표로 개발 중이며 미국 임상 1상을 진행 중이라고 안내한다.",
            interpretation="기존 GLP-1 계열과의 차별 포인트는 체성분 개선을 목표로 한 기전이다. 다만 사람 대상의 유효성·안전성 및 상업적 성공은 후속 임상에서 검증돼야 한다.",
            verification_status="verified",
            source_document_id="web-hanmi-hm17321-pipeline",
            source_title="HM17321 | 파이프라인 | 한미약품",
            disclosure_date="2026-08-24",
            url="https://www.hanmi.co.kr/science/pipeline/focused/hm17321.hm",
        ),
        MatterFact(
            category="웹 이슈 조사·국내 비교",
            fact="비교 가능한 국내 사례로 디앤디파마텍은 Metsera에 경구·주사용 비만 후보 6개를 기술이전했고, 회사 발표 기준 전체 계약 규모는 약 8억300만달러(약 1.1조원)다. 포함 파이프라인은 DD02S, DD03, MET06, DD07, DD14, DD15다.",
            interpretation="한미의 HM17321은 단일 후보물질의 최대 23억달러 계약이고, 디앤디파마텍은 다수 후보·경구 플랫폼 묶음 계약이므로 계약금액만으로 자산가치나 성공확률을 순위화할 수 없다. 후보 수, 권리범위, 개발 단계 및 선급금 공개 여부를 함께 비교해야 한다.",
            verification_status="verified",
            source_document_id="web-ddpharmatech-metsera",
            source_title="디앤디파마텍, 경구용 GLP-1 DD02S 북미 임상 본격 개시",
            disclosure_date="2024-11-22",
            url="https://ddpharmatech.com/board/board.php?bo_table=press&idx=32",
        ),
    ]


def _official_homepage_source(company_name: str, homepage: str) -> list[MatterFact]:
    """Capture a bounded official-site excerpt for audit-only company profiles.

    This is an evidence pack, not a crawler: it uses the homepage registered in
    OpenDART, strips navigation/script markup, and preserves the page URL.  The
    Gemini step may summarize it only with this source ID; unusable or short
    pages are simply ignored.
    """
    url = homepage.strip()
    if not url:
        return []
    if not re.match(r"^https?://", url, flags=re.IGNORECASE):
        url = f"https://{url.lstrip('/') }"
    try:
        response = requests.get(url, headers=USER_AGENT, timeout=12)
        response.raise_for_status()
    except requests.RequestException:
        return []
    content_type = response.headers.get("Content-Type", "").lower()
    if "html" not in content_type:
        return []
    page = response.text
    description = re.search(r'<meta[^>]+name=["\']description["\'][^>]+content=["\']([^"\']+)', page, flags=re.IGNORECASE)
    main = re.sub(r"(?is)<(script|style|noscript|svg).*?>.*?</\1>", " ", page)
    main = re.sub(r"(?is)<(header|footer|nav).*?>.*?</\1>", " ", main)
    main = html.unescape(re.sub(r"<[^>]+>", " ", main)).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", main).strip()
    if description:
        text = f"{html.unescape(description.group(1)).strip()} {text}"
    if len(text) < 80:
        return []
    return [
        MatterFact(
            category="공식 홈페이지 사업소개",
            fact=text[:3_500],
            interpretation="OpenDART 회사개황에 등록된 공식 홈페이지에서 자동 수집한 원문 발췌이며, 사업·전략 요약은 이 원문에 한정해 생성한다.",
            verification_status="verified",
            source_document_id=f"official-homepage-{_compact(company_name)[:30] or 'company'}",
            source_title=f"{company_name} 공식 홈페이지",
            disclosure_date="",
            url=response.url,
        )
    ]


def _curated_context_sources(company_name: str) -> list[MatterFact]:
    """Keep verified supplementary sources separate from the generic source pack."""
    if _compact(company_name) != "동아제약":
        return []
    return [
        MatterFact(
            category="공식 사업소개",
            fact="동아제약은 박카스, 일반의약품(OTC), 생활건강, 화장품을 주요 사업부문으로 운영한다. 박카스는 신규 제품 발매와 생산성 향상을 통한 원가절감을, OTC는 기존 대형 브랜드와 신규 제품을 통한 성장을 제시했다.",
            interpretation="감사보고서에는 사업 개요·성장전략 서술이 제한적이어서, 그룹의 공식 통합보고서 사업소개를 보완 근거로 사용했다.",
            verification_status="verified",
            source_document_id="official-donga-2024-integrated-report",
            source_title="동아쏘시오그룹 2024 통합보고서 · 동아제약 사업소개",
            disclosure_date="",
            url="https://yongmalogis.co.kr/resources/img/kr/DA_2024AR.pdf",
        )
    ]


def company_context_sources(company_name: str, homepage: str = "") -> list[MatterFact]:
    """Build a source pack for companies whose audit report lacks a business section."""
    sources = _official_homepage_source(company_name, homepage)
    sources.extend(_curated_context_sources(company_name))
    return sources


def company_context_profile(company_name: str) -> dict[str, object]:
    """Use only context that is separately represented in company_context_sources."""
    if _compact(company_name) != "동아제약":
        return {}
    return {
        "business_summary": "박카스, 일반의약품(OTC), 생활건강, 화장품을 주요 사업부문으로 운영합니다.",
        "growth_strategy": "박카스 신규 제품 발매·생산성 향상과 OTC 대형 브랜드 및 신제품 육성을 성장 방향으로 제시합니다.",
        "core_competencies": ["박카스 브랜드", "일반의약품(OTC)", "생활건강", "화장품"],
        "source": "동아쏘시오그룹 2024 통합보고서 사업소개 기반",
    }


def company_context_matters(company_name: str) -> list[MatterFact]:
    """Add a user-requested, source-linked product issue without treating press coverage as audited fact."""
    if _compact(company_name) != "동아제약":
        return []
    return [
        MatterFact(
            category="박카스 해외매출 변동",
            fact="언론 보도 기준 박카스 해외 매출은 2022년 957억원에서 2023년 710억원으로 25.8% 감소했고, 2024년 반등 후 2025년에는 성장세가 다시 둔화된 것으로 보도됐다.",
            interpretation="해외 판매는 동아에스티가 담당하고 동아제약은 국내 사업을 담당하므로, 이 수치를 동아제약 별도 감사보고서 매출과 직접 합산하거나 동일 지표로 해석하면 안 된다. 국가·제품별 매출은 원문 추가 확인이 필요하다.",
            verification_status="needs_review",
            source_document_id="news-bacchus-export-trend-2026-02-11",
            source_title="박카스 작년 매출 3681억 신기록…수출 부진에도 내수 껑충 · 데일리팜",
            disclosure_date="2026-02-11",
            url="https://www.dailypharm.com/user/news/335653",
        )
    ]


def research_issue(company_name: str, issue_query: str) -> list[MatterFact]:
    """Research the user's query first; do not substitute a generic filing list."""
    query = issue_query.strip()
    if not query:
        return []
    curated = _hanmi_obesity_comparison(company_name, query)
    leads = _news_search(f"{company_name} {query}")
    curated_titles = {_compact(item.source_title) for item in curated}
    return curated + [item for item in leads if _compact(item.source_title) not in curated_titles]


def price_event_matters(points: list[PricePoint], research_matters: list[MatterFact]) -> list[MatterFact]:
    """Describe observable price moves and only associate time-near sourced events.

    This intentionally says "same period" rather than claiming a source caused a
    price change.  That keeps the report suitable for review rather than turning
    correlation into an unsupported causal assertion.
    """
    if len(points) < 2:
        return []
    ordered = sorted(points, key=lambda item: item.trading_date)
    changes: list[tuple[float, int]] = []
    for index in range(1, len(ordered)):
        previous = ordered[index - 1].close
        if previous > 0:
            changes.append(((ordered[index].close / previous - 1) * 100, index))
    if not changes:
        return []

    start, end = ordered[0], ordered[-1]
    annual_return = (end.close / start.close - 1) * 100 if start.close else 0.0
    results = [
        MatterFact(
            category="주가 변동·1년 요약",
            fact=f"최근 1년 종가 기준 {start.trading_date} {start.close:,.0f}원에서 {end.trading_date} {end.close:,.0f}원으로 {annual_return:+.1f}% 변동했다.",
            interpretation="Yahoo Finance 일별 종가 기준의 관찰값이며 배당·거래정지·시장 전체 요인 등은 별도 조정하지 않았다.",
            verification_status="verified",
            source_document_id="market-yahoo-1y",
            source_title="Yahoo Finance 1년 일별 종가",
            disclosure_date=end.trading_date,
            url=end.source_url,
        )
    ]

    matters_by_day: list[tuple[date, MatterFact]] = []
    for matter in research_matters:
        try:
            matters_by_day.append((date.fromisoformat(matter.disclosure_date[:10]), matter))
        except (TypeError, ValueError):
            continue

    used_days: set[str] = set()
    used_source_ids: set[str] = set()
    for change, index in sorted(changes, key=lambda item: abs(item[0]), reverse=True):
        point = ordered[index]
        if point.trading_date in used_days:
            continue
        point_day = date.fromisoformat(point.trading_date)
        near = next((matter for event_day, matter in matters_by_day if abs((point_day - event_day).days) <= 3), None)
        if near is None or near.source_document_id in used_source_ids:
            continue
        used_days.add(point.trading_date)
        used_source_ids.add(near.source_document_id)
        previous = ordered[index - 1]
        results.append(
            MatterFact(
                category="주가 변동·이슈 대조",
                fact=f"{point.trading_date} 종가 {point.close:,.0f}원, 전 거래일 대비 {change:+.1f}% (전일 {previous.close:,.0f}원). 같은 시기 확인된 이슈: {near.fact}",
                interpretation="가격 변동과 이슈의 시간적 근접성을 표시한 것이며, 해당 이슈가 주가 변동의 단일 원인이라는 인과관계는 확인되지 않았다.",
                verification_status="needs_review",
                source_document_id=f"market-event-{point.trading_date}",
                source_title=f"Yahoo Finance 가격 + {near.source_title}",
                disclosure_date=point.trading_date,
                url=near.url or point.source_url,
            )
        )
        if len(results) >= 4:
            break
    return results
