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
from .web_research import company_context_matters, company_context_profile, company_context_sources, price_event_matters, research_issue


DEFAULT_PIPELINE_PATH = Path(r"C:\Users\CKD\dart-fss\dart_app_4.py")
REPORT_LABELS = {"annual": "사업보고서", "half": "반기보고서", "quarter1": "분기보고서", "quarter3": "분기보고서"}
FS_LABELS = {"OFS": "별도", "CFS": "연결"}
ISSUE_CATEGORY_MAP = {
    "계약": "계약·수주", "수주": "계약·수주", "공급": "계약·수주", "투자": "투자", "시설": "투자",
    "양수": "M&A", "합병": "M&A", "분할": "M&A", "소송": "소송·리스크", "횡령": "리스크",
    "유상증자": "자금조달", "전환사채": "자금조달", "신주인수권": "자금조달", "자사주": "주주환원",
    "인허가": "인허가", "임상": "임상·인허가",
}

# OpenDART's financial-statement API does not cover every audit-report-only
# filer.  These are the standard items retrieved from the existing pipeline's
# audit-report HTML parser only when that API produces no rows.
AUDIT_ACCOUNT_SPECS = (
    ("매출액", ("매출",)), ("매출원가", ("매출원가",)),
    ("영업이익", ("영업이익",)), ("당기순이익", ("당기순이익",)),
    ("자산총계", ("자산총계",)), ("유동자산", ("유동자산",)),
    ("현금및현금성자산", ("현금및현금성자산",)), ("매출채권", ("매출채권",)),
    ("재고자산", ("재고자산",)), ("부채총계", ("부채총계",)),
    ("유동부채", ("유동부채",)), ("자본총계", ("자본총계",)),
    ("자본금", ("자본금",)),
    ("영업활동현금흐름", ("영업활동으로인한현금흐름", "영업활동으로부터창출된현금흐름")),
    ("이자비용", ("이자비용", "금융비용")),
)

AUDIT_ACCOUNT_LABELS = {
    "매출액": {"매출", "매출액", "영업수익", "수익매출액"}, "매출원가": {"매출원가"},
    "영업이익": {"영업이익", "영업손실"},
    "당기순이익": {"당기순이익", "당기순손실", "당기순이익손실"},
    "자산총계": {"자산총계"}, "유동자산": {"유동자산"},
    "부채총계": {"부채총계"}, "유동부채": {"유동부채"},
    "자본총계": {"자본총계"}, "자본금": {"자본금"}, "재고자산": {"재고자산"},
    "매출채권": {"매출채권", "매출채권및기타채권"}, "현금및현금성자산": {"현금및현금성자산"},
    "영업활동현금흐름": {"영업활동으로인한현금흐름", "영업활동으로부터창출된현금흐름"},
    "이자비용": {"이자비용", "금융비용", "이자비용및금융비용"},
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


def _audit_label(value: object) -> str:
    """Normalize headings such as 'I. 매출' and '자 산 총 계'."""
    text = re.sub(r"\s+", "", str(value or "")).lower()
    text = re.sub(r"^[ivxlcdm]+[.)]", "", text, flags=re.IGNORECASE)
    return re.sub(r"[^0-9a-z가-힣]", "", text)


def standard_account(account_name: object, account_id: object) -> str:
    name = _clean(account_name)
    concept = _clean(account_id)
    # Equity-and-liabilities is a balance-sheet check total, not a liability
    # (or equity) line.  Keep it out of the report's selected financial facts.
    if any(label in name for label in ("자본과부채총계", "부채와자본총계", "자본및부채총계")) or "equityandliabilities" in concept:
        return str(account_name or "").strip() or "미분류"
    # Specific revenue-cost accounts must be checked before the broad "매출"
    # condition; otherwise 매출원가 is incorrectly displayed as 매출액.
    if "매출원가" in name or "costofsales" in concept:
        return "매출원가"
    if "매출채권" in name or "tradereceivables" in concept:
        return "매출채권"
    if "재고자산" in name or "inventories" in concept:
        return "재고자산"
    if name == "자본금" or "issuedcapital" in concept:
        return "자본금"
    if "영업활동으로인한현금흐름" in name or "영업활동으로부터창출된현금흐름" in name or "cashflowsfromusedinoperatingactivities" in concept or "cashflowsfromoperatingactivities" in concept:
        return "영업활동현금흐름"
    # Do not treat receivables, gross profit, product-level revenue or cash-flow
    # disposal proceeds as the total top-line.  The report uses only a disclosed
    # total revenue concept (or an explicitly named total-revenue account).
    if any(word in name for word in ("매출채권", "매출총이익")):
        return str(account_name or "").strip() or "미분류"
    if concept in {"ifrs-full_revenue", "ifrs-full_salesrevenue", "dart_revenue"} or name in {"매출액", "매출", "영업수익", "수익(매출액)"}:
        return "매출액"
    if "영업이익" in name or "operatingincome" in concept or "operatingprofit" in concept:
        return "영업이익"
    if any(label in name for label in ("당기순이익", "분기순이익", "반기순이익", "당기순손익")) or "profitloss" in concept or "netincome" in concept:
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
    def _audit_matches(frame: pd.DataFrame, standard_account: str) -> list[dict[str, Any]]:
        """Choose only a total-line match from audit statement HTML results."""
        if frame.empty:
            return []
        labels = AUDIT_ACCOUNT_LABELS[standard_account]
        candidates = []
        for record in frame.to_dict("records"):
            label = _audit_label(record.get("재무항목", ""))
            if label not in labels or _optional_int(record.get("당기금액")) is None:
                continue
            candidates.append(record)
        candidates.sort(key=lambda item: (not bool(item.get("matched_exact")), int(item.get("테이블번호", 0)), int(item.get("행번호", 0))))
        return candidates[:1]

    def _audit_fallback_facts(
        self,
        company: dict[str, str],
        year: int,
        fs_div: str,
        documents: list[SourceDocument],
        period_label: str,
        period_start: str,
        period_end: str,
        is_cumulative: bool,
    ) -> list[FinancialFact]:
        """Read core facts from a standalone/connected audit report once API data is absent."""
        try:
            filings = self.core.search_filings(self.api_key, company["corp_code"], year, "annual")
        except Exception:
            return []
        if filings.empty:
            return []
        preferred_roles = ["연결감사보고서", "사업보고서"] if fs_div == "CFS" else ["감사보고서", "사업보고서"]
        ordered: list[dict[str, Any]] = []
        for role in preferred_roles:
            ordered.extend(filings[filings["report_role"] == role].to_dict("records"))
        if not ordered:
            ordered = filings.to_dict("records")
        document_ids = {document.rcept_no: document.document_id for document in documents}
        for filing in ordered:
            receipt = str(filing.get("rcept_no", ""))
            try:
                nodes = self.core.parse_dart_tree_nodes(self.core.fetch_report_main_html(receipt))
                node = self.core.select_financial_statement_node(nodes)
                if node is None:
                    continue
                viewer_html = self.core.fetch_viewer_html(node)
            except Exception:
                continue
            facts: list[FinancialFact] = []
            for index, (standard, queries) in enumerate(AUDIT_ACCOUNT_SPECS, start=1):
                matches: list[dict[str, Any]] = []
                for query in queries:
                    try:
                        found = self.core.search_account_in_viewer_html(
                            viewer_html=viewer_html,
                            account_query=query,
                            exact=False,
                            company_name=str(company.get("corp_name", "")),
                            year=year,
                            report_type="annual",
                            fs_div=fs_div,
                            rcept_no=receipt,
                        )
                    except Exception:
                        continue
                    matches = self._audit_matches(found, standard)
                    if matches:
                        break
                for match in matches:
                    amount = _optional_int(match.get("당기금액"))
                    if amount is None:
                        continue
                    statement = "CF" if standard == "영업활동현금흐름" else "PL" if standard in {"매출액", "매출원가", "영업이익", "당기순이익", "이자비용"} else "BS"
                    facts.append(
                        FinancialFact(
                            fact_id=f"audit-{year}-annual-{index}",
                            company_name=str(company.get("corp_name", "")),
                            stock_code=str(company.get("stock_code", "")),
                            fiscal_year=year,
                            period_label=period_label,
                            report_type="annual",
                            period_start=period_start,
                            period_end=period_end,
                            is_cumulative=is_cumulative,
                            fs_div=fs_div,
                            statement=statement,
                            stock_or_flow="flow" if statement in {"PL", "CF"} else "stock",
                            standard_account=standard,
                            source_account=str(match.get("재무항목", standard)),
                            account_id="audit_report_html",
                            value_krw=amount,
                            currency="KRW",
                            source_document_id=document_ids.get(receipt, f"dart-{receipt}"),
                            source_location=f"감사보고서 재무제표 · 표 {match.get('테이블번호', '')} · 행 {match.get('행번호', '')}",
                            value_type="collected",
                        )
                    )
            # Require the six core totals before accepting a parsed statement;
            # a partial HTML table must not masquerade as a complete audit source.
            core = {fact.standard_account for fact in facts}
            if {"매출액", "영업이익", "당기순이익", "자산총계", "부채총계", "자본총계"}.issubset(core):
                return facts
        return []

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

    def _governance_snapshot(self, corp_code: str, year: int, ceo_name: str = "") -> dict[str, Any]:
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
        try:
            executives = self.core.call_open_dart_json("exctvSttus.json", self.api_key, params).get("list", [])
            normalized_ceo = _clean(ceo_name)
            matching = [row for row in executives if normalized_ceo and _clean(row.get("nm")) in normalized_ceo]
            representatives = [row for row in matching if "대표" in str(row.get("ofcps", ""))] or [row for row in executives if "대표" in str(row.get("ofcps", ""))]
            careers = []
            for representative in representatives[:2]:
                career = re.sub(r"\s+", " ", str(representative.get("main_career", ""))).strip()
                if career and career.lower() not in {"nan", "none", "-"}:
                    careers.append(f"{representative.get('nm', '대표이사')}: {career}")
            if careers:
                result["ceo_bio"] = " / ".join(careers)[:500]
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
        if any(word in text for word in ("진단", "분석", "검사")):
            return "진단·분석"
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

    def _audit_related_subsidiaries(self, tables: list[pd.DataFrame], filing: dict[str, Any], node: dict[str, str], receipt: str) -> list[dict[str, str]]:
        """Read a subsidiary name from an audit report's related-party note.

        Audit-only filers often disclose the entity in the related-party note
        (for example, note 33) rather than in a business-report ownership table.
        The note does not necessarily disclose stake, activity or location, so
        those fields remain explicit "확인 필요" rather than being inferred.
        """
        results: dict[str, dict[str, str]] = {}
        for table_index, table in enumerate(tables, start=1):
            if table.empty:
                continue
            for values in table.fillna("").itertuples(index=False, name=None):
                cells = [str(value).strip() for value in values]
                relation_index = next((index for index, value in enumerate(cells) if _clean(value) in {"종속기업", "자회사"}), None)
                if relation_index is None:
                    continue
                name = next((value for value in cells[relation_index + 1:] if value and _clean(value) not in {"합계", "nan", "none"}), "")
                key = self._subsidiary_name_key(name)
                if not name or len(name) < 2 or key in results:
                    continue
                results[key] = {
                    "사업군": "기타",
                    "자회사명": name,
                    "지분율": "확인 필요",
                    "사업영역": "감사보고서 특수관계자 주석상 종속기업",
                    "소재지": "확인 필요",
                    "주요 생산시설 또는 핵심 역량": "확인 필요",
                    "비고": "종속기업 · 지분율 및 사업영역은 원문 별도 확인 필요",
                    "출처 문서": str(filing.get("report_nm", "감사보고서")),
                    "출처일": str(filing.get("rcept_dt", "")),
                    "출처 URL": self.core.disclosure_url(receipt),
                    "출처 위치": f"{node.get('text', '주석')} · 특수관계자 표 {table_index}",
                }
        return list(results.values())[:50]

    @staticmethod
    def _table_unit(tables: list[pd.DataFrame]) -> str:
        for table in tables[:3]:
            text = " ".join(str(value) for value in table.fillna("").to_numpy().flatten())
            match = re.search(r"단위\s*:\s*(천원|백만원|원)", text)
            if match:
                return match.group(1)
        return "공시 단위 확인 필요"

    def _subsidiary_investment_stakes(self, nodes: list[dict[str, str]], filing: dict[str, Any], receipt: str) -> dict[str, dict[str, str]]:
        """Join ending ownership percentages from the business-report investment table."""
        results: dict[str, dict[str, str]] = {}
        for node in [item for item in nodes if "타법인출자" in _clean(item.get("text", ""))]:
            try:
                tables = pd.read_html(StringIO(self.core.fetch_viewer_html(node)))
            except Exception:
                continue
            for table_index, table in enumerate(tables, start=1):
                if table.empty or len(table.columns) < 2:
                    continue
                columns = self._flat_columns(table)
                name_index = self._column_index(columns, ("법인명", "회사명", "상호"))
                stake_indices = [index for index, column in enumerate(columns) if "지분율" in column]
                if name_index is None or not stake_indices:
                    continue
                stake_index = stake_indices[-1]  # ending ownership, not opening ownership
                for values in table.itertuples(index=False, name=None):
                    name = str(values[name_index] if name_index < len(values) else "").strip()
                    stake = str(values[stake_index] if stake_index < len(values) else "").strip()
                    key = self._subsidiary_name_key(name)
                    if not name or key in {"법인명", "회사명", "합계", "nan", "none"} or not stake or stake.lower() in {"nan", "none", "-"}:
                        continue
                    results[key] = {
                        "지분율": stake if "%" in stake else f"{stake}%",
                        "지분율 출처": f"{node.get('text', '타법인출자 현황')} · 표 {table_index}",
                    }
        return results

    def _subsidiaries(self, corp_code: str, years: list[int]) -> list[dict[str, str]]:
        """Extract subsidiary rows from the latest business/audit-report note table when available.

        This is intentionally independent of the selected OFS/CFS financial basis:
        subsidiary coverage is a group disclosure, so it is sourced from the
        business or audit report rather than inferred from an individual financial statement.
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
                investment_stakes = self._subsidiary_investment_stakes(nodes, filing, receipt)
                note_nodes = [
                    node for node in nodes
                    if any(hint in _clean(node.get("text", "")) for hint in ("종속기업", "연결대상", "타법인출자"))
                ]
                # Audit reports typically expose all notes under a single
                # "주석" node, including the related-party/subsidiary note.
                if not note_nodes:
                    note_nodes = [node for node in nodes if "주석" in _clean(node.get("text", ""))]
                for node in note_nodes[:8]:
                    try:
                        tables = pd.read_html(StringIO(self.core.fetch_viewer_html(node)))
                    except Exception:
                        continue
                    disclosed_unit = self._table_unit(tables)
                    ownership_rows: list[dict[str, str]] = []
                    financials: dict[str, dict[str, str]] = {}
                    for table_index, table in enumerate(tables, start=1):
                        if table.empty or len(table.columns) < 2:
                            continue
                        if any(isinstance(column, tuple) and len(column) > 2 for column in table.columns):
                            continue
                        columns = self._flat_columns(table)
                        name_index = self._column_index(columns, ("회사명", "법인명", "종속기업", "피투자회사", "상호"))
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
                        # A detailed connected-subsidiary table in a business
                        # report often gives the entity, address and main
                        # business but no ownership percentage.  It is still a
                        # valid group-coverage source; keep its missing stake
                        # explicit instead of dropping every subsidiary.
                        if stake_index is None and business_index is not None and location_index is not None and "연결대상종속회사" in _clean(node.get("text", "")):
                            for values in table.itertuples(index=False, name=None):
                                name = str(values[name_index] if name_index < len(values) else "").strip()
                                compact_name = _clean(name)
                                if not name or compact_name in {"상호", "회사명", "법인명", "합계", "nan", "none"} or len(name) < 2 or len(name) > 120:
                                    continue
                                business = str(values[business_index] if business_index < len(values) else "확인 필요").strip()
                                location = str(values[location_index] if location_index < len(values) else "확인 필요").strip()
                                assets = str(values[asset_index] if asset_index is not None and asset_index < len(values) else "").strip()
                                establishment_index = self._column_index(columns, ("설립일",))
                                establishment = str(values[establishment_index] if establishment_index is not None and establishment_index < len(values) else "").strip()
                                investment = investment_stakes.get(self._subsidiary_name_key(name), {})
                                ownership_rows.append(
                                    {
                                        "사업군": self._subsidiary_group(business, location),
                                        "자회사명": name,
                                        "지분율": investment.get("지분율", "확인 필요"),
                                        "사업영역": business or "확인 필요",
                                        "소재지": location or "확인 필요",
                                        "자산(공시 표기)": f"{assets}{disclosed_unit}" if assets and assets.lower() not in {"nan", "none"} else "",
                                        "주요 생산시설 또는 핵심 역량": "연결대상 종속회사 상세표상 주요사업",
                                        "비고": " · ".join(value for value in ["연결대상 종속회사", investment.get("지분율 출처", "지분율은 별도 공시 표에서 확인 필요"), f"설립일 {establishment}" if establishment else ""] if value),
                                        "출처 문서": str(filing.get("report_nm", "사업보고서")),
                                        "출처일": str(filing.get("rcept_dt", "")),
                                        "출처 URL": self.core.disclosure_url(receipt),
                                        "출처 위치": f"{node.get('text', '연결대상 종속회사')} · 표 {table_index}",
                                    }
                                )
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
                    audit_rows = self._audit_related_subsidiaries(tables, filing, node, receipt)
                    if audit_rows:
                        return audit_rows
        return []

    def _audit_parent_snapshot(self, corp_code: str, years: list[int]) -> tuple[dict[str, Any], MatterFact | None]:
        """Fill an audit-only filer's parent relationship from an audited note.

        The business-report shareholder APIs are usually unavailable for an
        unlisted audit filer.  This deliberately reads only a table that labels
        a party as ``지배기업`` and leaves the field blank when the note is not
        explicit; it never guesses a group affiliation from the company name.
        """
        for year in sorted(years, reverse=True):
            try:
                filings = self.core.search_filings(self.api_key, corp_code, year, "annual")
            except Exception:
                continue
            if filings.empty:
                continue
            for filing in filings.to_dict("records"):
                receipt = str(filing.get("rcept_no", ""))
                if not receipt:
                    continue
                try:
                    nodes = self.core.parse_dart_tree_nodes(self.core.fetch_report_main_html(receipt))
                except Exception:
                    continue
                for node in [item for item in nodes if "주석" in _clean(item.get("text", ""))][:2]:
                    try:
                        tables = pd.read_html(StringIO(self.core.fetch_viewer_html(node)))
                    except Exception:
                        continue
                    for table_index, table in enumerate(tables, start=1):
                        for values in table.fillna("").itertuples(index=False, name=None):
                            cells = [str(value).strip() for value in values if str(value).strip()]
                            text = " ".join(cells)
                            if "지배기업" not in text:
                                continue
                            ratio_match = re.search(r"(\d+(?:\.\d+)?)\s*%", text)
                            candidates = [
                                value for value in cells
                                if len(value) >= 2 and len(value) <= 80
                                and not re.search(r"\d", value)
                                and _clean(value) not in {"지배기업", "관계", "구분", "당사", "회사"}
                                and "지배기업" not in value
                            ]
                            parent = candidates[0] if candidates else ""
                            if not parent:
                                continue
                            ratio = float(ratio_match.group(1)) if ratio_match else None
                            description = f"감사보고서 특수관계자 주석에서 지배기업을 {parent}로 기재"
                            if ratio is not None:
                                description += f"하고 지분율 {ratio:.2f}%를 표시한다."
                            else:
                                description += "."
                            return (
                                {
                                    "parent_company": parent,
                                    "parent_company_ratio": ratio,
                                    "largest_holder": parent,
                                    "largest_holder_relation": "지배기업(감사보고서 주석)",
                                    "largest_holder_ratio": ratio,
                                },
                                MatterFact(
                                    category="지배관계·감사보고서 주석",
                                    fact=description,
                                    interpretation="감사보고서만 제출하는 기업의 경우 사업보고서 주주현황 API 대신 감사보고서 주석의 명시 정보를 사용한다.",
                                    verification_status="verified",
                                    source_document_id=f"dart-parent-{receipt}",
                                    source_title=str(filing.get("report_nm", "감사보고서")),
                                    disclosure_date=str(filing.get("rcept_dt", "")),
                                    url=self.core.disclosure_url(receipt),
                                ),
                            )
        return {}, None

    @staticmethod
    def _plain_text(value: str, limit: int = 12_000) -> str:
        text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value)
        text = re.sub(r"(?i)<br\s*/?>", "\n", text)
        text = re.sub(r"<[^>]+>", " ", text)
        text = html.unescape(text).replace("\xa0", " ")
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n\s*\n+", "\n", text).strip()
        return text[:limit]

    @staticmethod
    def _audit_overview_profile(sources: list[MatterFact]) -> dict[str, str]:
        """Use an explicit audit-note purpose clause without generative inference."""
        source = next((item for item in sources if item.category == "회사 개요·감사보고서 주석"), None)
        if source is None:
            return {}
        text = re.sub(r"\s+", " ", source.fact)
        purpose = re.search(r"당사는\s+(.{2,260}?(?:목적사업으로|사업을)\s*영위하고[^.]*\.)", text)
        profile: dict[str, str] = {"source": "감사보고서 주석 일반사항 원문 기반"}
        if purpose:
            profile["business_summary"] = purpose.group(0).strip()
        location = re.search(r"([가-힣A-Za-z0-9\-()·,\s]{5,100})에\s*위치하고\s*있습니다", text)
        if location:
            profile["audit_note_location"] = location.group(1).strip()
        return profile

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
                elif any(hint in normalized for hint in ("회사의개요", "회사개요", "영업의개요", "영업활동")):
                    # Audit reports commonly label their only business
                    # description as "I. 회사의 개요" rather than the
                    # business-report section "사업의 내용".
                    category = "회사 개요·감사보고서"
                if not category:
                    continue
                try:
                    text = self._plain_text(self.core.fetch_viewer_html(node), limit=3_500)
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
                if len(sections) >= 3:
                    return sections
            # Audit-only filings expose a single top-level "주석" viewer.  Its
            # opening "일반사항" normally contains the legal entity's business
            # description, so retain just that bounded note rather than the
            # entire financial-statement body.
            if not any(item.category in {"사업의 내용", "회사 개요·감사보고서"} for item in sections):
                note_node = next((node for node in nodes if "주석" in _clean(node.get("text", ""))), None)
                if note_node:
                    try:
                        note_text = self._plain_text(self.core.fetch_viewer_html(note_node), limit=18_000)
                    except Exception:
                        note_text = ""
                    general = re.search(r"(?:^|\n)\s*1\s*[.)]\s*(?:일반사항|회사.?개요).*?(?=(?:\n\s*2\s*[.)])|\Z)", note_text, flags=re.DOTALL)
                    excerpt = (general.group(0) if general else note_text[:3_500]).strip()
                    if len(excerpt) >= 100 and any(word in excerpt for word in ("사업", "제조", "판매", "의약", "서비스", "제품")):
                        sections.append(
                            MatterFact(
                                category="회사 개요·감사보고서 주석",
                                fact=excerpt[:3_500],
                                interpretation="감사보고서 주석의 일반사항 원문 발췌",
                                verification_status="verified",
                                source_document_id=f"dart-audit-overview-{receipt}",
                                source_title=str(filing.get("report_nm", "감사보고서")),
                                disclosure_date=str(filing.get("rcept_dt", "")),
                                url=self.core.disclosure_url(receipt),
                            )
                        )
            if sections:
                return sections
        return []

    @staticmethod
    def _corporate_highlights(profile: dict[str, Any], chronology: list[dict[str, Any]], sources: list[MatterFact]) -> list[MatterFact]:
        """Create a compact, source-linked default issue block from filings.

        This prevents a listed company report with an empty user query from
        showing only price movement.  The facts are disclosed business scope and
        dated corporate events; no unverified market interpretation is added.
        """
        source_index = {source.source_document_id: source for source in sources}
        results: list[MatterFact] = []
        business = str(profile.get("business_summary", "")).strip()
        business_source = next((source for source in sources if source.category in {"사업의 내용", "회사 개요·감사보고서", "공식 사업소개", "공식 홈페이지 사업소개"}), None)
        if business and business != "확인 필요" and business_source:
            results.append(
                MatterFact(
                    category="사업 포트폴리오",
                    fact=business,
                    interpretation="사업보고서 또는 공식 홈페이지에 기재된 주요 사업 기준입니다. 제품·지역별 매출 및 수익성은 재무제표 기준과 구분해 확인합니다.",
                    verification_status="verified",
                    source_document_id=business_source.source_document_id,
                    source_title=business_source.source_title,
                    disclosure_date=business_source.disclosure_date,
                    url=business_source.url,
                )
            )
        growth_strategy = str(profile.get("growth_strategy", "")).strip()
        if growth_strategy and growth_strategy != "확인 필요" and business_source:
            results.append(
                MatterFact(
                    category="성장 전략·글로벌 진출",
                    fact=growth_strategy,
                    interpretation="사업보고서·공식 홈페이지에 명시된 허가, 수출 또는 생산기반 확장 내용입니다. 계약금액·매출 기여는 개별 공시와 재무제표로 별도 검증합니다.",
                    verification_status="verified",
                    source_document_id=business_source.source_document_id,
                    source_title=business_source.source_title,
                    disclosure_date=business_source.disclosure_date,
                    url=business_source.url,
                )
            )
        for item in sorted(chronology, key=lambda value: str(value.get("date", "")), reverse=True):
            event = str(item.get("event", "")).strip()
            source = next((source_index[source_id] for source_id in item.get("source_ids", []) if source_id in source_index), None)
            if not event or source is None:
                continue
            category = "생산시설·투자" if any(word in event for word in ("공장", "시설", "증설", "투자", "인증", "허가")) else "핵심 연혁"
            results.append(
                MatterFact(
                    category=category,
                    fact=f"{item.get('date', '')} {event}".strip(),
                    interpretation="사업보고서 회사 연혁에 기재된 사실입니다. 실적·생산능력에 미치는 영향은 후속 공시와 재무 수치로 확인이 필요합니다.",
                    verification_status="verified",
                    source_document_id=source.source_document_id,
                    source_title=source.source_title,
                    disclosure_date=source.disclosure_date,
                    url=source.url,
                )
            )
            if len(results) >= 4:
                break
        return results

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
            label, start, end, cumulative = _period(report_type, year)
            try:
                raw = self.core.fetch_financial_statement_all(self.api_key, company["corp_code"], year, report_type, request.fs_div)
            except Exception:
                raw = pd.DataFrame()
            if raw.empty and report_type == "annual":
                facts.extend(self._audit_fallback_facts(company, year, request.fs_div, documents, label, start, end, cumulative))
                continue
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
        governance = self._governance_snapshot(company["corp_code"], max(years), str(profile.get("ceo_name", "")))
        profile.update({key: value for key, value in governance.items() if value not in (None, "")})
        parent_snapshot, parent_source = self._audit_parent_snapshot(company["corp_code"], years)
        if not profile.get("largest_holder"):
            profile.update({key: value for key, value in parent_snapshot.items() if value not in (None, "")})
        corporate_sources = [
            *self._corporate_report_sources(company["corp_code"], years),
            *([parent_source] if parent_source else []),
            *company_context_sources(str(company.get("corp_name", request.identifier)), str(profile.get("homepage", ""))),
        ]
        audit_overview = self._audit_overview_profile(corporate_sources)
        for key, value in audit_overview.items():
            if value and not profile.get(key):
                profile[key] = value
        analysis_sources = [
            source for source in corporate_sources
            if source.category in {"사업의 내용", "회사 연혁", "회사 개요·감사보고서", "공식 홈페이지 사업소개", "공식 사업소개"}
        ]
        corporate_analysis = analyze_corporate_profile(str(company.get("corp_name", request.identifier)), analysis_sources, self.gemini_api_key)
        if corporate_analysis:
            profile.update({key: value for key, value in corporate_analysis.items() if key != "chronology" and value not in (None, "", [])})
        chronology = corporate_analysis.get("chronology", []) if corporate_analysis else []
        if not profile.get("growth_strategy"):
            operational_event = next(
                (item for item in sorted(chronology, key=lambda value: str(value.get("date", "")), reverse=True)
                 if any(word in str(item.get("event", "")) for word in ("공장", "시설", "증설", "인증", "허가", "투자"))),
                None,
            )
            if operational_event:
                profile["growth_strategy"] = f"최근 사업보고서 연혁상 {operational_event.get('date', '')} {operational_event.get('event', '')}이 기재되어 있습니다. 중장기 성장전략의 정량 목표는 추가 확인이 필요합니다."
        for key, value in company_context_profile(str(company.get("corp_name", request.identifier))).items():
            if not profile.get(key):
                profile[key] = value
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
        corporate_highlights = self._corporate_highlights(profile, chronology, corporate_sources)
        issue_sources = [*company_context_matters(company_name), *research_issue(company_name, request.issue_query)]
        gemini_matters = analyze_issue(company_name, request.issue_query, issue_sources, self.gemini_api_key)
        issue_matters = gemini_matters or issue_sources
        if any(item.category.startswith("Gemini 검색 기반·") for item in gemini_matters):
            issue_provider = "Gemini Google Search grounding"
        elif gemini_matters:
            issue_provider = "Gemini source-pack synthesis"
        else:
            issue_provider = "source_only"
        # Keep default business/production highlights alongside a user issue;
        # the latter must not erase source-linked corporate context.
        major_matters = [*corporate_highlights, *issue_matters]
        major_matters.extend(price_event_matters(price_history, issue_matters))
        return FactPack(
            pack_id=str(uuid4()),
            generated_at=now_iso(),
            entity=entity,
            reporting_policy={
                "primary_fs_basis": request.fs_div,
                "display_unit": "억원",
                "currency": "KRW",
                "issue_query": request.issue_query,
                "issue_analysis_provider": issue_provider,
            },
            documents=list(document_index.values()),
            financial_facts=facts,
            corporate_profile=profile,
            chronology=chronology,
            corporate_sources=corporate_sources,
            subsidiaries=subsidiaries,
            major_matters=major_matters,
            price_history=price_history,
        )
