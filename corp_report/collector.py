from __future__ import annotations

import importlib.util
import math
import os
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from types import ModuleType
from typing import Any
from uuid import uuid4

import requests

from .models import FactPack, FinancialFact, PricePoint, ReportRequest, SourceDocument, now_iso
from .gemini_research import analyze_issue
from .web_research import price_event_matters, research_issue


DEFAULT_PIPELINE_PATH = Path(r"C:\Users\CKD\dart-fss\dart_app_4.py")
REPORT_LABELS = {"annual": "사업보고서", "half": "반기보고서", "quarter1": "분기보고서", "quarter3": "분기보고서"}
FS_LABELS = {"OFS": "별도", "CFS": "연결"}
ISSUE_CATEGORY_MAP = {
    "계약": "계약·수주", "수주": "계약·수주", "공급": "계약·수주", "투자": "투자", "시설": "투자",
    "양수": "M&A", "합병": "M&A", "분할": "M&A", "소송": "소송·리스크", "횡령": "리스크",
    "유상증자": "자금조달", "전환사채": "자금조달", "신주인수권": "자금조달", "자사주": "주주환원",
    "인허가": "인허가", "임상": "임상·인허가",
}


def load_pipeline(path: str | None = None) -> ModuleType:
    pipeline_path = Path(path or os.getenv("DART_PIPELINE_PATH", DEFAULT_PIPELINE_PATH))
    if not pipeline_path.exists():
        raise FileNotFoundError(f"기존 수집 코드를 찾지 못했습니다: {pipeline_path}")
    spec = importlib.util.spec_from_file_location("dart_existing_pipeline", pipeline_path)
    if spec is None or spec.loader is None:
        raise RuntimeError("기존 수집 코드를 불러올 수 없습니다.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _clean(value: object) -> str:
    return re.sub(r"\s+", "", str(value or "")).lower()


def _optional_int(value: object) -> int | None:
    """Convert reported numeric values without turning missing values into errors."""
    if value is None:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return int(number) if math.isfinite(number) else None


def _optional_float(value: object) -> float | None:
    if value is None:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def standard_account(account_name: object, account_id: object) -> str:
    name = _clean(account_name)
    concept = _clean(account_id)
    if "매출" in name or "revenue" in concept or "sales" in concept:
        return "매출액"
    if "매출원가" in name or "costofsales" in concept:
        return "매출원가"
    if "영업이익" in name or "operatingincome" in concept or "operatingprofit" in concept:
        return "영업이익"
    if "당기순이익" in name or "profitloss" == concept or "netincome" in concept:
        return "당기순이익"
    if "유동자산" in name or "currentassets" in concept:
        return "유동자산"
    if "자산총계" in name or name == "자산" or "assets" == concept:
        return "자산총계"
    if "유동부채" in name or "currentliabilities" in concept:
        return "유동부채"
    if "부채총계" in name or name == "부채" or "liabilities" == concept:
        return "부채총계"
    if "자본총계" in name or name == "자본" or "equity" == concept:
        return "자본총계"
    if "이자비용" in name or "interestexpense" in concept or "financecosts" in concept:
        return "이자비용"
    if any(word in name for word in ["단기차입금", "장기차입금", "사채", "차입부채"]) or "borrowings" in concept:
        return "차입금"
    if "현금및현금성자산" in name or "cashandcashequivalents" in concept:
        return "현금및현금성자산"
    return str(account_name or "").strip() or "미분류"


def _period(report_type: str, year: int) -> tuple[str, str, str, bool]:
    if report_type == "annual":
        return str(year), f"{year}-01-01", f"{year}-12-31", True
    if report_type == "half":
        return f"{year} 1H", f"{year}-01-01", f"{year}-06-30", True
    if report_type == "quarter1":
        return f"{year} 1Q", f"{year}-01-01", f"{year}-03-31", True
    return f"{year} 3Q", f"{year}-01-01", f"{year}-09-30", True


def _latest_report_type() -> tuple[int, str]:
    today = date.today()
    if today.month >= 8:
        return today.year, "half"
    if today.month >= 5:
        return today.year, "quarter1"
    return today.year - 1, "annual"


class DartFactPackCollector:
    def __init__(self, api_key: str, pipeline_path: str | None = None, gemini_api_key: str = "") -> None:
        if not api_key.strip():
            raise ValueError("OpenDART API 키를 corp_report/local_settings.py 또는 DART_API_KEY 환경변수에 설정하세요.")
        self.core = load_pipeline(pipeline_path)
        self.api_key = self.core.ensure_api_key(api_key)
        self.gemini_api_key = gemini_api_key

    def resolve_company(self, identifier: str) -> dict[str, str]:
        companies = self.core.fetch_corp_codes(self.api_key)
        compact = re.sub(r"\D", "", identifier)
        if compact and len(compact) == 6:
            for company in companies:
                if str(company.get("stock_code", "")).zfill(6) == compact:
                    return {**company, "input_name": identifier, "match_type": "stock_code"}
        resolved = self.core.resolve_company_rows(companies, [identifier])[0]
        if not resolved.get("corp_code"):
            raise ValueError(f"회사명 또는 종목코드를 찾지 못했습니다: {identifier}")
        return resolved

    def _documents(self, corp_code: str, years: list[int], report_type: str) -> list[SourceDocument]:
        documents: list[SourceDocument] = []
        seen: set[str] = set()
        for year in years:
            try:
                filings = self.core.search_filings(self.api_key, corp_code, year, report_type)
            except Exception:
                continue
            for _, filing in filings.iterrows():
                rcept_no = str(filing.get("rcept_no", ""))
                if not rcept_no or rcept_no in seen:
                    continue
                seen.add(rcept_no)
                documents.append(
                    SourceDocument(
                        document_id=f"dart-{rcept_no}",
                        title=str(filing.get("report_nm", REPORT_LABELS[report_type])),
                        disclosure_date=str(filing.get("rcept_dt", "")),
                        url=str(filing.get("disclosure_url", self.core.disclosure_url(rcept_no))),
                        rcept_no=rcept_no,
                        document_type=report_type,
                    )
                )
        return documents

    @staticmethod
    def _document_for(documents: list[SourceDocument], report_type: str, year: int) -> SourceDocument | None:
        year_text = str(year + 1 if report_type == "annual" else year)
        candidates = [item for item in documents if item.document_type == report_type and item.disclosure_date.startswith(year_text)]
        return candidates[0] if candidates else next((item for item in documents if item.document_type == report_type), None)

    def _company_profile(self, corp_code: str) -> dict[str, Any]:
        try:
            result = self.core.call_open_dart_json("company.json", self.api_key, {"corp_code": corp_code})
        except Exception:
            return {}
        return {
            "company_name": result.get("corp_name", ""),
            "ceo_name": result.get("ceo_nm", "확인 필요"),
            "address": result.get("adres", "확인 필요"),
            "establishment_date": result.get("est_dt", "확인 필요"),
            "industry": result.get("induty_code", "확인 필요"),
            "homepage": result.get("hm_url", ""),
        }

    @staticmethod
    def _price_history(stock_code: str) -> list[PricePoint]:
        if not stock_code:
            return []
        period1 = int((datetime.now() - timedelta(days=370)).timestamp())
        period2 = int(datetime.now().timestamp())
        for suffix in ("KS", "KQ"):
            url = f"https://query2.finance.yahoo.com/v8/finance/chart/{stock_code}.{suffix}?period1={period1}&period2={period2}&interval=1d"
            try:
                payload = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20).json()["chart"]["result"][0]
                closes = payload["indicators"]["quote"][0]["close"]
                volumes = payload["indicators"]["quote"][0].get("volume", [])
                timestamps = payload["timestamp"]
            except (KeyError, IndexError, TypeError, requests.RequestException, ValueError):
                continue
            points = []
            for index, close in enumerate(closes):
                close_value = _optional_float(close)
                if close_value is None:
                    continue
                points.append(
                    PricePoint(
                        trading_date=datetime.fromtimestamp(timestamps[index]).date().isoformat(),
                        close=round(close_value, 2),
                        volume=_optional_int(volumes[index]) if index < len(volumes) else None,
                        source_url=url,
                    )
                )
            return points
        return []

    def collect(self, request: ReportRequest) -> FactPack:
        company = self.resolve_company(request.identifier)
        years = sorted(set(request.years))
        if not years:
            years = list(range(date.today().year - 4, date.today().year))
        periods = [(year, "annual") for year in years]
        if request.include_latest_interim:
            interim = _latest_report_type()
            if interim not in periods:
                periods.append(interim)

        documents: list[SourceDocument] = []
        for report_type in sorted({kind for _, kind in periods}):
            documents.extend(self._documents(company["corp_code"], [year for year, kind in periods if kind == report_type], report_type))
        document_index = {document.document_id: document for document in documents}
        facts: list[FinancialFact] = []
        for year, report_type in periods:
            try:
                raw = self.core.fetch_financial_statement_all(self.api_key, company["corp_code"], year, report_type, request.fs_div)
            except Exception:
                continue
            label, start, end, cumulative = _period(report_type, year)
            document = self._document_for(documents, report_type, year)
            source_id = document.document_id if document else ""
            for index, (_, row) in enumerate(raw.iterrows(), start=1):
                statement = self.core.canonicalize_section(row.get("sj_div", ""))
                amount = row.get("thstrm_amount")
                facts.append(
                    FinancialFact(
                        fact_id=f"fin-{year}-{report_type}-{index}",
                        company_name=str(company.get("corp_name", request.identifier)),
                        stock_code=str(company.get("stock_code", "")),
                        fiscal_year=year,
                        period_label=label,
                        report_type=report_type,
                        period_start=start,
                        period_end=end,
                        is_cumulative=cumulative,
                        fs_div=request.fs_div,
                        statement=statement,
                        stock_or_flow="stock" if statement == "BS" else "flow",
                        standard_account=standard_account(row.get("account_nm"), row.get("account_id")),
                        source_account=str(row.get("account_nm", "")),
                        account_id=str(row.get("account_id", "")),
                        value_krw=_optional_int(amount),
                        currency=str(row.get("currency", "KRW")),
                        source_document_id=source_id,
                        source_location="OpenDART fnlttSinglAcntAll.json",
                    )
                )

        profile = self._company_profile(company["corp_code"])
        entity = {
            "input_identifier": request.identifier,
            "company_name": company.get("corp_name", request.identifier),
            "corp_code": company.get("corp_code", ""),
            "stock_code": company.get("stock_code", ""),
            "listing_status": "listed" if company.get("stock_code") else "unlisted",
        }
        price_history = self._price_history(str(company.get("stock_code", ""))) if request.include_price_chart else []
        company_name = str(company.get("corp_name", request.identifier))
        source_matters = research_issue(company_name, request.issue_query)
        gemini_matters = analyze_issue(company_name, request.issue_query, source_matters, self.gemini_api_key)
        research_matters = gemini_matters or source_matters
        # User-entered issue research is the primary source for this section;
        # generic recent disclosures are not substituted when no query is given.
        major_matters = research_matters + price_event_matters(price_history, research_matters)
        return FactPack(
            pack_id=str(uuid4()),
            generated_at=now_iso(),
            entity=entity,
            reporting_policy={
                "primary_fs_basis": request.fs_div,
                "display_unit": "억원",
                "currency": "KRW",
                "issue_query": request.issue_query,
                "issue_analysis_provider": "Gemini" if gemini_matters else "source_only",
            },
            documents=list(document_index.values()),
            financial_facts=facts,
            corporate_profile=profile,
            major_matters=major_matters,
            price_history=price_history,
        )
