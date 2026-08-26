from __future__ import annotations

import importlib.util
import html
import math
import os
import re
from datetime import date, datetime, timedelta
from io import StringIO
from pathlib import Path
from types import ModuleType
from typing import Any
from uuid import uuid4

import pandas as pd
import requests

from .models import FactPack, FinancialFact, MatterFact, PricePoint, ReportRequest, SourceDocument, now_iso
from .gemini_research import analyze_corporate_profile, analyze_issue
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


def _dart_number(value: object) -> float | None:
    """Parse OpenDART numeric strings such as '1,234' or '73.7'."""
    text = re.sub(r"[^0-9.\-]", "", str(value or ""))
    return _optional_float(text)


def standard_account(account_name: object, account_id: object) -> str:
    name = _clean(account_name)
    concept = _clean(account_id)
    # Specific revenue-cost accounts must be checked before the broad "매출"
    # condition; otherwise 매출원가 is incorrectly displayed as 매출액.
    if "매출원가" in name or "costofsales" in concept:
        return "매출원가"
    # Do not treat receivables, gross profit, product-level revenue or cash-flow
    # disposal proceeds as the total top-line.  The report uses only a disclosed
    # total revenue concept (or an explicitly named total-revenue account).
    if any(word in name for word in ("매출채권", "매출총이익")):
        return str(account_name or "").strip() or "미분류"
    if concept in {"ifrs-full_revenue", "ifrs-full_salesrevenue", "dart_revenue"} or name in {"매출액", "매출", "영업수익", "수익(매출액)"}:
        return "매출액"
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

    def _governance_snapshot(self, corp_code: str, year: int) -> dict[str, Any]:
        """Collect employee, major-holder and issued-share facts from OpenDART."""
        params = {"corp_code": corp_code, "bsns_year": str(year), "reprt_code": "11011"}
        result: dict[str, Any] = {"source": "OpenDART 사업보고서 정기공시 API", "fiscal_year": year}
        try:
            employees = self.core.call_open_dart_json("empSttus.json", self.api_key, params).get("list", [])
            regular = sum(_dart_number(row.get("rgllbr_co")) or 0 for row in employees)
            contract = sum(_dart_number(row.get("cnttk_co")) or 0 for row in employees)
            if regular or contract:
                result["employee_count"] = int(regular + contract)
                result["regular_employee_count"] = int(regular)
                result["contract_employee_count"] = int(contract)
        except Exception:
            pass
        try:
            holders = self.core.call_open_dart_json("hyslrSttus.json", self.api_key, params).get("list", [])
            # hyslrSttus includes a group-total row ("계").  It is useful for
            # reconciliation but must never be shown as the largest shareholder.
            individual_holders = [
                row for row in holders
                if _clean(row.get("nm")) not in {"계", "합계", "소계", "-", "nan", "none"}
            ]
            ranked = sorted(individual_holders, key=lambda row: _dart_number(row.get("trmend_posesn_stock_qota_rt")) or -1, reverse=True)
            if ranked:
                holder = ranked[0]
                result["largest_holder"] = str(holder.get("nm", "확인 필요"))
                result["largest_holder_relation"] = str(holder.get("relate", ""))
                result["largest_holder_ratio"] = _dart_number(holder.get("trmend_posesn_stock_qota_rt"))
        except Exception:
            pass
        try:
            stocks = self.core.call_open_dart_json("stockTotqySttus.json", self.api_key, params).get("list", [])
            ordinary = next((row for row in stocks if "보통" in str(row.get("se", ""))), stocks[0] if stocks else None)
            if ordinary:
                result["issued_shares"] = int(_dart_number(ordinary.get("istc_totqy")) or 0) or None
        except Exception:
            pass
        return result

    @staticmethod
    def _flat_columns(frame: pd.DataFrame) -> list[str]:
        columns = []
        for column in frame.columns:
            values = column if isinstance(column, tuple) else (column,)
            labels = [str(value) for value in values if str(value).strip() and not str(value).lower().startswith("unnamed")]
            # With a two-level header the lower level gives the actual field
            # (for example, 당기말 / 지분율(%)), not a merged parent heading.
            text = labels[-1] if labels else ""
            columns.append(re.sub(r"\s+", "", text))
        return columns

    @staticmethod
    def _column_index(columns: list[str], keywords: tuple[str, ...]) -> int | None:
        return next((index for index, value in enumerate(columns) if any(keyword in value for keyword in keywords)), None)

    @staticmethod
    def _subsidiary_group(business: str, location: str) -> str:
        text = _clean(f"{business} {location}")
        if any(word in text for word in ("연구", "임상", "바이오", "신약")):
            return "연구개발·바이오"
        if any(word in text for word in ("제조", "생산", "공장")):
            return "제조"
        if any(word in text for word in ("유통", "판매", "도매")):
            return "유통·판매"
        if any(word in text for word in ("미국", "중국", "일본", "유럽", "해외")):
            return "해외사업"
        return "기타"

    @staticmethod
    def _subsidiary_name_key(value: object) -> str:
        """Normalize a disclosed entity name enough to join adjacent note tables."""
        name = re.sub(r"\(\*?\d+\)", "", str(value or ""))
        return _clean(name).replace("㈜", "").replace("주식회사", "")

    def _subsidiaries(self, corp_code: str, years: list[int]) -> list[dict[str, str]]:
        """Extract subsidiary rows from the latest annual-report note table when available.

        This is intentionally independent of the selected OFS/CFS financial basis:
        subsidiary coverage is a group disclosure, so it is sourced from the
        annual report rather than inferred from an individual financial statement.
        """
        if not years:
            return []
        for year in sorted(years, reverse=True):
            try:
                filings = self.core.search_filings(self.api_key, corp_code, year, "annual")
            except Exception:
                continue
            if filings.empty:
                continue
            ordered = filings[filings["report_role"] == "사업보고서"]
            records = (ordered if not ordered.empty else filings).to_dict("records")
            for filing in records:
                receipt = str(filing.get("rcept_no", ""))
                if not receipt:
                    continue
                try:
                    nodes = self.core.parse_dart_tree_nodes(self.core.fetch_report_main_html(receipt))
                except Exception:
                    continue
                note_nodes = [
                    node for node in nodes
                    if any(hint in _clean(node.get("text", "")) for hint in ("종속기업", "연결대상", "타법인출자"))
                ]
                for node in note_nodes[:8]:
                    try:
                        tables = pd.read_html(StringIO(self.core.fetch_viewer_html(node)))
                    except Exception:
                        continue
                    ownership_rows: list[dict[str, str]] = []
                    financials: dict[str, dict[str, str]] = {}
                    for table_index, table in enumerate(tables, start=1):
                        if table.empty or len(table.columns) < 2:
                            continue
                        if any(isinstance(column, tuple) and len(column) > 2 for column in table.columns):
                            continue
                        columns = self._flat_columns(table)
                        name_index = self._column_index(columns, ("회사명", "법인명", "종속기업", "피투자회사"))
                        if name_index is None:
                            continue
                        stake_index = self._column_index(columns, ("지분율", "소유지분", "지분"))
                        business_index = self._column_index(columns, ("주요사업", "주요영업", "영업활동", "사업내용", "업종"))
                        location_index = self._column_index(columns, ("소재지", "주소", "국가"))
                        note_index = self._column_index(columns, ("비고", "관계", "구분"))
                        asset_index = self._column_index(columns, ("자산",))
                        revenue_index = self._column_index(columns, ("매출액", "매출"))
                        profit_index = self._column_index(columns, ("당기순손익", "당기순이익", "순손익", "순이익"))
                        # A summary-financial table is linked to the ownership
                        # table by entity name.  Preserve the disclosure's text
                        # as-is: DART note tables do not expose a consistent unit
                        # field that would make an automatic conversion reliable.
                        if asset_index is not None and (revenue_index is not None or profit_index is not None) and stake_index is None:
                            for values in table.itertuples(index=False, name=None):
                                name = str(values[name_index] if name_index < len(values) else "").strip()
                                key = self._subsidiary_name_key(name)
                                if not name or key in {"회사명", "법인명", "합계", "nan", "none"}:
                                    continue
                                def field(index: int | None) -> str:
                                    value = str(values[index] if index is not None and index < len(values) else "").strip()
                                    return "" if value.lower() in {"nan", "none"} else value
                                financials[key] = {
                                    "자산(공시 표기)": field(asset_index),
                                    "매출액(공시 표기)": field(revenue_index),
                                    "당기순이익(공시 표기)": field(profit_index),
                                }
                            continue
                        # Summary-financial tables and narrative containers can
                        # also contain a company-name column; only use a genuine
                        # ownership table that exposes stake, business and location.
                        if stake_index is None or business_index is None or location_index is None:
                            continue
                        rows: list[dict[str, str]] = []
                        for values in table.itertuples(index=False, name=None):
                            name = str(values[name_index] if name_index < len(values) else "").strip()
                            compact_name = _clean(name)
                            if not name or compact_name in {"회사명", "법인명", "종속기업", "합계", "nan", "none"} or len(name) < 2 or len(name) > 120:
                                continue
                            business = str(values[business_index] if business_index is not None and business_index < len(values) else "확인 필요").strip()
                            location = str(values[location_index] if location_index is not None and location_index < len(values) else "확인 필요").strip()
                            stake = str(values[stake_index] if stake_index is not None and stake_index < len(values) else "확인 필요").strip()
                            note = str(values[note_index] if note_index is not None and note_index < len(values) else "").strip()
                            ownership_rows.append(
                                {
                                    "사업군": self._subsidiary_group(business, location),
                                    "자회사명": name,
                                    "지분율": stake or "확인 필요",
                                    "사업영역": business or "확인 필요",
                                    "소재지": location or "확인 필요",
                                    "주요 생산시설 또는 핵심 역량": "확인 필요",
                                    "비고": note if note and note.lower() not in {"nan", "none"} else "",
                                    "출처 문서": str(filing.get("report_nm", "사업보고서")),
                                    "출처일": str(filing.get("rcept_dt", "")),
                                    "출처 URL": self.core.disclosure_url(receipt),
                                    "출처 위치": f"{node.get('text', '주석')} · 표 {table_index}",
                                }
                            )
                    if ownership_rows:
                        unique: dict[str, dict[str, str]] = {}
                        for row in ownership_rows:
                            key = self._subsidiary_name_key(row["자회사명"])
                            if key not in unique:
                                unique[key] = row
                            if key in financials:
                                unique[key].update(financials[key])
                        return list(unique.values())[:50]
        return []

    @staticmethod
    def _plain_text(value: str, limit: int = 12_000) -> str:
        text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value)
        text = re.sub(r"(?i)<br\s*/?>", "\n", text)
        text = re.sub(r"<[^>]+>", " ", text)
        text = html.unescape(text).replace("\xa0", " ")
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n\s*\n+", "\n", text).strip()
        return text[:limit]

    def _corporate_report_sources(self, corp_code: str, years: list[int]) -> list[MatterFact]:
        """Collect only named annual-report sections used for company overview."""
        for year in sorted(years, reverse=True):
            try:
                filings = self.core.search_filings(self.api_key, corp_code, year, "annual")
            except Exception:
                continue
            if filings.empty:
                continue
            report = filings[filings["report_role"] == "사업보고서"]
            filing = (report.iloc[0] if not report.empty else filings.iloc[0]).to_dict()
            receipt = str(filing.get("rcept_no", ""))
            if not receipt:
                continue
            try:
                nodes = self.core.parse_dart_tree_nodes(self.core.fetch_report_main_html(receipt))
            except Exception:
                continue
            sections: list[MatterFact] = []
            for node in nodes:
                node_text = str(node.get("text", ""))
                normalized = _clean(node_text)
                category = ""
                if "사업의내용" in normalized:
                    category = "사업의 내용"
                elif any(hint in normalized for hint in ("회사의연혁", "회사연혁", "연혁")):
                    category = "회사 연혁"
                if not category:
                    continue
                try:
                    text = self._plain_text(self.core.fetch_viewer_html(node))
                except Exception:
                    continue
                if len(text) < 100:
                    continue
                sections.append(
                    MatterFact(
                        category=category,
                        fact=text,
                        interpretation="사업보고서 원문 발췌",
                        verification_status="verified",
                        source_document_id=f"dart-section-{receipt}-{len(sections) + 1}",
                        source_title=str(filing.get("report_nm", "사업보고서")),
                        disclosure_date=str(filing.get("rcept_dt", "")),
                        url=self.core.disclosure_url(receipt),
                    )
                )
                if len(sections) >= 2:
                    return sections
            if sections:
                return sections
        return []

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
        governance = self._governance_snapshot(company["corp_code"], max(years))
        profile.update({key: value for key, value in governance.items() if value not in (None, "")})
        corporate_sources = self._corporate_report_sources(company["corp_code"], years)
        corporate_analysis = analyze_corporate_profile(str(company.get("corp_name", request.identifier)), corporate_sources, self.gemini_api_key)
        if corporate_analysis:
            profile.update({key: value for key, value in corporate_analysis.items() if key != "chronology" and value not in (None, "", [])})
        subsidiaries = self._subsidiaries(company["corp_code"], years)
        entity = {
            "input_identifier": request.identifier,
            "company_name": company.get("corp_name", request.identifier),
            "corp_code": company.get("corp_code", ""),
            "stock_code": company.get("stock_code", ""),
            "listing_status": "listed" if company.get("stock_code") else "unlisted",
        }
        price_history = self._price_history(str(company.get("stock_code", ""))) if request.include_price_chart else []
        if price_history and governance.get("issued_shares"):
            entity["market_cap_krw"] = round(price_history[-1].close * governance["issued_shares"])
            entity["market_price_date"] = price_history[-1].trading_date
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
            chronology=corporate_analysis.get("chronology", []) if corporate_analysis else [],
            corporate_sources=corporate_sources,
            subsidiaries=subsidiaries,
            major_matters=major_matters,
            price_history=price_history,
        )
