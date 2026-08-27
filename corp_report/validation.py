from __future__ import annotations

from collections import defaultdict
from typing import Any

from .models import FactPack, FinancialFact


_STATEMENT_BY_ACCOUNT = {
    "매출액": "PL", "매출원가": "PL", "영업이익": "PL", "당기순이익": "PL", "이자비용": "PL",
    "영업활동현금흐름": "CF",
    "자산총계": "BS", "유동자산": "BS", "부채총계": "BS", "유동부채": "BS", "자본총계": "BS",
    "현금및현금성자산": "BS", "매출채권": "BS", "재고자산": "BS", "자본금": "BS", "차입금": "BS",
}


def _latest_fact(facts: list[FinancialFact], account: str) -> FinancialFact | None:
    candidates = [item for item in facts if item.standard_account == account and item.value_krw is not None]
    preferred_statement = _STATEMENT_BY_ACCOUNT.get(account)
    if preferred_statement:
        candidates = [item for item in candidates if item.statement == preferred_statement] or candidates
    return candidates[0] if candidates else None


def validate_fact_pack(pack: FactPack) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    required = ["매출액", "영업이익", "당기순이익", "자산총계", "부채총계", "자본총계"]
    primary_basis = pack.reporting_policy["primary_fs_basis"]
    if not pack.financial_facts:
        results.append({"severity": "warning", "rule": "financial_facts_present", "message": "재무 수집값이 없습니다. 원천 공시와 수집 기준을 확인해야 합니다."})
    for fact in pack.financial_facts:
        if fact.fs_div != primary_basis:
            results.append({"severity": "error", "rule": "basis_consistency", "fact_id": fact.fact_id, "message": "선택한 재무제표 기준과 다릅니다."})
        # OpenDART includes blank comparative/detail rows.  Warn only when a
        # report metric that can be displayed is missing, not for every raw
        # disclosure row retained for auditability.
        if fact.value_krw is None and fact.standard_account in required:
            results.append({"severity": "warning", "rule": "value_present", "fact_id": fact.fact_id, "message": "수집값이 없습니다."})
        if fact.value_krw is not None and fact.standard_account in _STATEMENT_BY_ACCOUNT and not fact.source_document_id:
            results.append({"severity": "warning", "rule": "source_document", "fact_id": fact.fact_id, "message": "공시 문서 연결을 확인해야 합니다."})

    by_period: dict[str, list[FinancialFact]] = defaultdict(list)
    for fact in pack.financial_facts:
        by_period[fact.period_label].append(fact)
    for period, facts in by_period.items():
        values = {name: _latest_fact(facts, name) for name in ["자산총계", "부채총계", "자본총계"]}
        if all(values.values()):
            asset, liability, equity = (values[name].value_krw or 0 for name in ["자산총계", "부채총계", "자본총계"])
            tolerance = max(abs(asset) * 0.01, 1)
            if abs(asset - liability - equity) > tolerance:
                results.append({"severity": "warning", "rule": "balance_sheet_tie_out", "period": period, "message": "자산 = 부채 + 자본 대사가 허용오차를 벗어났습니다."})
        for account in required:
            if not _latest_fact(facts, account):
                results.append({"severity": "warning", "rule": "required_metric", "period": period, "message": f"{account}을 찾지 못했습니다."})

    if pack.entity.get("listing_status") == "listed" and not pack.price_history:
        results.append({"severity": "warning", "rule": "price_history", "message": "상장사이나 주가 데이터를 수집하지 못했습니다."})
    return results


def growth_display(previous: float | None, current: float | None, profit_metric: bool = False) -> str | float:
    if previous is None or current is None:
        return "N/A"
    if profit_metric:
        if previous < 0 < current:
            return "흑자전환"
        if previous > 0 > current:
            return "적자전환"
        if previous < 0 and current < 0:
            return "적자축소" if current > previous else "적자확대"
    if previous <= 0 or current == 0:
        return "N.M."
    return current / previous - 1
