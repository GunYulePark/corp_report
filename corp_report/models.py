from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from typing import Any, Literal


Basis = Literal["OFS", "CFS"]


@dataclass(frozen=True)
class ReportRequest:
    identifier: str
    fs_div: Basis = "OFS"
    years: list[int] = field(default_factory=list)
    issue_query: str = ""
    include_latest_interim: bool = True
    include_price_chart: bool = True


@dataclass
class SourceDocument:
    document_id: str
    title: str
    disclosure_date: str
    url: str
    rcept_no: str = ""
    document_type: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class FinancialFact:
    fact_id: str
    company_name: str
    stock_code: str
    fiscal_year: int
    period_label: str
    report_type: str
    period_start: str
    period_end: str
    is_cumulative: bool
    fs_div: Basis
    statement: str
    stock_or_flow: Literal["stock", "flow"]
    standard_account: str
    source_account: str
    account_id: str
    value_krw: int | None
    currency: str
    source_document_id: str
    source_location: str
    value_type: Literal["collected", "converted", "calculated"] = "collected"

    @property
    def value_eok(self) -> float | None:
        return None if self.value_krw is None else self.value_krw / 100_000_000

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["value_eok"] = self.value_eok
        return data


@dataclass
class MatterFact:
    category: str
    fact: str
    interpretation: str
    verification_status: Literal["verified", "needs_review"]
    source_document_id: str
    source_title: str
    disclosure_date: str
    url: str

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class PricePoint:
    trading_date: str
    close: float
    volume: int | None
    source_url: str

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class FactPack:
    pack_id: str
    generated_at: str
    entity: dict[str, Any]
    reporting_policy: dict[str, Any]
    documents: list[SourceDocument] = field(default_factory=list)
    financial_facts: list[FinancialFact] = field(default_factory=list)
    corporate_profile: dict[str, Any] = field(default_factory=dict)
    chronology: list[dict[str, Any]] = field(default_factory=list)
    corporate_sources: list[MatterFact] = field(default_factory=list)
    subsidiaries: list[dict[str, Any]] = field(default_factory=list)
    major_matters: list[MatterFact] = field(default_factory=list)
    price_history: list[PricePoint] = field(default_factory=list)
    validation_results: list[dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": "1.0",
            "pack_id": self.pack_id,
            "generated_at": self.generated_at,
            "entity": self.entity,
            "reporting_policy": self.reporting_policy,
            "documents": [item.to_dict() for item in self.documents],
            "financial_facts": [item.to_dict() for item in self.financial_facts],
            "corporate_profile": self.corporate_profile,
            "chronology": self.chronology,
            "corporate_sources": [item.to_dict() for item in self.corporate_sources],
            "subsidiaries": self.subsidiaries,
            "major_matters": [item.to_dict() for item in self.major_matters],
            "price_history": [item.to_dict() for item in self.price_history],
            "validation_results": self.validation_results,
        }


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def current_years(count: int = 4) -> list[int]:
    return list(range(date.today().year - count, date.today().year))
