from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .models import FactPack, FinancialFact


NAVY = "002060"
LIGHT_BLUE = "D9EAF7"
GREY = "F2F2F2"
WHITE = "FFFFFF"
THIN = Side(style="thin", color="B7C3D0")
HAIR = Side(style="hair", color="A6A6A6")
TITLE = Font(name="맑은 고딕", size=18, bold=True)
SECTION = Font(name="맑은 고딕", size=12, bold=True, color=WHITE)
HEADER = Font(name="맑은 고딕", size=10, bold=True)
BODY = Font(name="맑은 고딕", size=10)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
NUMBER_FORMAT = '#,##0.0;[Red](#,##0.0);-'
PERCENT_FORMAT = '0.0%;[Red](0.0%);-'


def _basis_label(fs_div: str) -> str:
    return "별도" if fs_div == "OFS" else "연결"


def _safe_sheet_name(name: str) -> str:
    return name.replace("[", "_").replace("]", "_").replace("/", "_")[:31]


def _style_section(ws, row: int, start_col: int, end_col: int, text: str) -> None:
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    cell = ws.cell(row, start_col, text)
    cell.font = SECTION
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[row].height = 21


def _apply_table_header(ws, row: int, columns: list[str]) -> None:
    for col, label in enumerate(columns, start=1):
        cell = ws.cell(row, col, label)
        cell.font = HEADER
        cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        cell.alignment = CENTER


def _write_data_sheet(ws, pack: FactPack) -> None:
    headers = [
        "회사명", "종목코드", "기준연도/분기", "보고서유형", "기간시작", "기간종료", "누적여부", "연결·별도",
        "계정과목 표준명", "원문 계정과목명", "계정ID", "재무제표", "시점·기간", "값(원)", "통화", "원 단위",
        "표시 단위", "값 구분", "출처 문서명", "공시일", "URL", "페이지 또는 표 위치", "표시 사용",
    ]
    ws.title = "재무 data"
    _apply_table_header(ws, 1, headers)
    doc_index = {document.document_id: document for document in pack.documents}
    display_fact_ids = _display_fact_ids(pack.financial_facts)
    for row_index, fact in enumerate(pack.financial_facts, start=2):
        document = doc_index.get(fact.source_document_id)
        values = [
            fact.company_name, fact.stock_code, fact.period_label, fact.report_type, fact.period_start, fact.period_end,
            "Y" if fact.is_cumulative else "N", _basis_label(fact.fs_div), fact.standard_account, fact.source_account,
            fact.account_id, fact.statement, fact.stock_or_flow, fact.value_krw, fact.currency, "원", "억원", fact.value_type,
            document.title if document else "확인 필요", document.disclosure_date if document else "", document.url if document else "", fact.source_location,
            "Y" if fact.fact_id in display_fact_ids else "N",
        ]
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row_index, col, value)
            cell.font = BODY
            cell.border = Border(left=HAIR, right=HAIR, top=HAIR, bottom=HAIR)
            cell.alignment = LEFT_WRAP if col in {1, 9, 10, 19, 21, 22} else CENTER
            if col == 14 and isinstance(value, (int, float)):
                cell.number_format = '#,##0;[Red](#,##0);-'
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:W{max(ws.max_row, 2)}"
    widths = [18, 11, 12, 12, 12, 12, 10, 10, 18, 28, 28, 12, 10, 16, 10, 10, 10, 12, 28, 12, 48, 30, 10]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width


def _display_fact_ids(facts: Iterable[FinancialFact]) -> set[str]:
    """Select one reported total per metric; borrowing components remain additive."""
    selected: set[str] = set()
    selected_keys: set[tuple[str, str, str]] = set()
    aggregate_accounts = {"차입금"}
    for fact in facts:
        if fact.value_krw is None:
            continue
        if fact.standard_account in aggregate_accounts:
            selected.add(fact.fact_id)
            continue
        key = (fact.period_label, fact.fs_div, fact.standard_account)
        if key not in selected_keys:
            selected.add(fact.fact_id)
            selected_keys.add(key)
    return selected


def _periods(pack: FactPack) -> list[str]:
    labels = []
    for fact in pack.financial_facts:
        if fact.period_label not in labels:
            labels.append(fact.period_label)
    annuals = sorted([label for label in labels if label.isdigit()], key=int)
    interims = [label for label in labels if not label.isdigit()]
    return annuals + interims


def _amount_formula(period_cell: str, account_cell: str, basis: str) -> str:
    criteria = f"'재무 data'!$C:$C,{period_cell},'재무 data'!$H:$H,\"{basis}\",'재무 data'!$I:$I,{account_cell},'재무 data'!$W:$W,\"Y\""
    return f'=IF(COUNTIFS({criteria})=0,"N/A",SUMIFS(\'재무 data\'!$N:$N,\'재무 data\'!$C:$C,{period_cell},\'재무 data\'!$H:$H,"{basis}",\'재무 data\'!$I:$I,{account_cell},\'재무 data\'!$W:$W,"Y")/100000000)'


def _write_financial_sheet(ws, pack: FactPack) -> None:
    basis = _basis_label(pack.reporting_policy["primary_fs_basis"])
    periods = _periods(pack)
    ws.title = "재무"
    ws["B2"] = "(단위 : 억원)"
    ws["B2"].font = Font(name="Arial", size=9)
    ws["B3"] = f"재무제표 기준: {basis}"
    ws["B3"].font = HEADER
    ws["B4"] = "구분"
    ws["B4"].font = HEADER
    ws["B4"].fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    ws["B4"].alignment = CENTER
    ws["B4"].border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    for col, label in enumerate(periods, start=3):
        cell = ws.cell(4, col, label)
        cell.font = HEADER
        cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        cell.alignment = CENTER
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        ws.column_dimensions[get_column_letter(col)].width = 14

    rows = [
        (6, "매출액"), (7, "영업이익"), (8, "당기순이익"), (10, "자산총계"), (11, "유동자산"),
        (12, "부채총계"), (13, "유동부채"), (14, "자본총계"), (15, "이자비용"), (16, "차입금"), (17, "현금및현금성자산"),
    ]
    for row, label in rows:
        ws.cell(row, 2, label).font = HEADER if label in {"매출액", "영업이익", "당기순이익", "자산총계", "부채총계", "자본총계"} else BODY
        for col in range(3, 3 + len(periods)):
            cell = ws.cell(row, col)
            cell.value = _amount_formula(f"{get_column_letter(col)}$4", f"$B{row}", basis)
            cell.number_format = NUMBER_FORMAT
            cell.font = Font(name="Arial", size=10, color="008000")
            cell.alignment = CENTER
            cell.border = Border(bottom=HAIR)
    for row, label in [(19, "성장성"), (23, "안정성")]:
        ws.cell(row, 2, label).font = HEADER
        ws.cell(row, 2).fill = PatternFill("solid", fgColor=GREY)
    ratio_rows = [(20, "매출액증감률", 6, False), (21, "영업이익증감률", 7, True), (24, "영업이익률", 7, None), (25, "순이익률", 8, None), (26, "부채비율", 12, "equity"), (27, "유동비율", 11, "current_liability"), (28, "순차입금", 16, "net_debt"), (29, "이자보상배율", 7, "coverage")]
    for row, label, source_row, kind in ratio_rows:
        ws.cell(row, 2, label).font = BODY
        for col in range(3, 3 + len(periods)):
            letter = get_column_letter(col)
            cell = ws.cell(row, col)
            if kind is False:
                prev = get_column_letter(col - 1)
                cell.value = "=\"N/A\"" if col == 3 else f'=IF(OR(NOT(ISNUMBER({prev}{source_row})),NOT(ISNUMBER({letter}{source_row}))),"N/A",IF({prev}{source_row}<=0,"N.M.",{letter}{source_row}/{prev}{source_row}-1))'
                cell.number_format = PERCENT_FORMAT
            elif kind is True:
                prev = get_column_letter(col - 1)
                cell.value = "=\"N/A\"" if col == 3 else f'=IF(OR(NOT(ISNUMBER({prev}{source_row})),NOT(ISNUMBER({letter}{source_row}))),"N/A",IF(AND({prev}{source_row}<0,{letter}{source_row}>0),"흑자전환",IF(AND({prev}{source_row}>0,{letter}{source_row}<0),"적자전환",IF(AND({prev}{source_row}<0,{letter}{source_row}<0),IF({letter}{source_row}>{prev}{source_row},"적자축소","적자확대"),IF({prev}{source_row}<=0,"N.M.",{letter}{source_row}/{prev}{source_row}-1)))))'
                cell.number_format = PERCENT_FORMAT
            elif kind is None:
                cell.value = f'=IF(OR(NOT(ISNUMBER({letter}{source_row})),NOT(ISNUMBER({letter}6)),{letter}6<=0),"N.M.",{letter}{source_row}/{letter}6)'
                cell.number_format = PERCENT_FORMAT
            elif kind == "equity":
                cell.value = f'=IF(OR(NOT(ISNUMBER({letter}12)),NOT(ISNUMBER({letter}14)),{letter}14<=0),"N.M.",{letter}12/{letter}14)'
                cell.number_format = PERCENT_FORMAT
            elif kind == "current_liability":
                cell.value = f'=IF(OR(NOT(ISNUMBER({letter}11)),NOT(ISNUMBER({letter}13)),{letter}13<=0),"N.M.",{letter}11/{letter}13)'
                cell.number_format = PERCENT_FORMAT
            elif kind == "net_debt":
                cell.value = f'=IF(OR(NOT(ISNUMBER({letter}16)),NOT(ISNUMBER({letter}17))),"N/A",{letter}16-{letter}17)'
                cell.number_format = NUMBER_FORMAT
            else:
                cell.value = f'=IF(OR(NOT(ISNUMBER({letter}7)),NOT(ISNUMBER({letter}15)),{letter}15<=0),"N.M.",{letter}7/{letter}15)'
                cell.number_format = '0.0x;[Red](0.0x);-'
            cell.alignment = CENTER
            cell.border = Border(bottom=HAIR)
    ws.column_dimensions["B"].width = 20
    ws.freeze_panes = "C5"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1


def _write_summary_sheet(ws, pack: FactPack) -> None:
    company = str(pack.entity.get("company_name", "기업"))
    ws.title = "본장"
    ws.merge_cells("C1:AL1")
    ws["C1"] = f"◐ {company}"
    ws["C1"].font = TITLE
    ws["C3"] = f"< {datetime.now().strftime('%Y. %m. %d')} >"
    ws["C3"].alignment = CENTER
    _style_section(ws, 5, 4, 38, "가. 회사 및 사업개요")
    profile = pack.corporate_profile
    profile_fields = [
        ("상호", company), ("대표자", profile.get("ceo_name", "확인 필요")), ("소재지", profile.get("address", "확인 필요")),
        ("설립일자", profile.get("establishment_date", "확인 필요")), ("종목코드", pack.entity.get("stock_code") or "비상장"),
        ("기업형태", "상장" if pack.entity.get("listing_status") == "listed" else "비상장"), ("재무제표 기준", _basis_label(pack.reporting_policy["primary_fs_basis"])),
    ]
    for index, (label, value) in enumerate(profile_fields):
        row = 7 + index // 2
        col = 5 if index % 2 == 0 else 20
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 3)
        ws.merge_cells(start_row=row, start_column=col + 4, end_row=row, end_column=col + 13)
        label_cell = ws.cell(row, col, label)
        value_cell = ws.cell(row, col + 4, value)
        for cell in (label_cell, value_cell):
            cell.border = Border(left=HAIR, right=HAIR, top=HAIR, bottom=HAIR)
            cell.alignment = CENTER
            cell.font = BODY
        label_cell.fill = PatternFill("solid", fgColor=GREY)
        value_cell.font = Font(name="맑은 고딕", size=10, bold=label in {"상호", "종목코드"})

    _style_section(ws, 14, 4, 38, "나. 재무 현황")
    periods = _periods(pack)[-5:]
    ws.merge_cells("E16:H16")
    ws["E16"] = "지표"
    ws["E16"].font = HEADER
    ws["E16"].fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    ws["E16"].alignment = CENTER
    for period_index, period in enumerate(periods):
        start_col = 9 + period_index * 5
        end_col = start_col + 4
        ws.merge_cells(start_row=16, start_column=start_col, end_row=16, end_column=end_col)
        cell = ws.cell(16, start_col, period)
        cell.font = HEADER
        cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        cell.alignment = CENTER
    summary_rows = [(17, "매출액", 6), (18, "영업이익", 7), (19, "당기순이익", 8), (20, "자산총계", 10), (21, "부채총계", 12), (22, "자본총계", 14), (23, "부채비율", 26)]
    for row, label, financial_row in summary_rows:
        ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=8)
        ws.cell(row, 5, label)
        for period_index, _ in enumerate(periods):
            start_col = 9 + period_index * 5
            end_col = start_col + 4
            source_col = get_column_letter(3 + len(_periods(pack)) - len(periods) + period_index)
            ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
            ws.cell(row, start_col, f"='재무'!{source_col}{financial_row}")
            ws.cell(row, start_col).number_format = PERCENT_FORMAT if label == "부채비율" else NUMBER_FORMAT
        for col in [5] + [9 + period_index * 5 for period_index in range(len(periods))]:
            ws.cell(row, col).border = Border(left=HAIR, right=HAIR, top=HAIR, bottom=HAIR)
            ws.cell(row, col).alignment = CENTER
    ws.merge_cells("E25:AG25")
    ws["E25"] = f"재무제표 기준: {_basis_label(pack.reporting_policy['primary_fs_basis'])} · 표시값은 재무 data 원천값을 재무 시트 수식으로 연결"
    ws["E25"].alignment = LEFT_WRAP

    _style_section(ws, 27, 4, 38, "다. 주요 사항")
    if pack.major_matters:
        for index, matter in enumerate(pack.major_matters[:5], start=29):
            ws.merge_cells(start_row=index, start_column=5, end_row=index, end_column=34)
            ws.cell(index, 5, f"• [{matter.category}] {matter.fact} — {matter.interpretation}")
            ws.cell(index, 5).alignment = LEFT_WRAP
    else:
        ws.merge_cells("E29:AH30")
        ws["E29"] = "조회 기간 내 주요 공시를 찾지 못했습니다."
        ws["E29"].alignment = LEFT_WRAP

    _style_section(ws, 36, 4, 38, "라. 자회사 등 현황")
    ws.merge_cells("E38:AH39")
    ws["E38"] = "자회사 정보는 사업보고서의 타법인 출자·연결대상 종속기업 주석 추출 단계에서 보강됩니다. 현재: 확인 필요"
    ws["E38"].alignment = LEFT_WRAP

    if pack.entity.get("listing_status") == "listed" and pack.price_history:
        for row, point in enumerate(pack.price_history, start=2):
            ws.cell(row, 40, point.trading_date)
            ws.cell(row, 41, point.close)
        chart = LineChart()
        chart.title = "최근 1년 주가 추이"
        chart.y_axis.title = "종가"
        chart.x_axis.title = "거래일"
        chart.height = 7.5
        chart.width = 13
        data = Reference(ws, min_col=41, min_row=1, max_row=1 + len(pack.price_history))
        categories = Reference(ws, min_col=40, min_row=2, max_row=1 + len(pack.price_history))
        chart.add_data(data, titles_from_data=False)
        chart.set_categories(categories)
        chart.legend = None
        chart.dataLabels = DataLabelList()
        ws.add_chart(chart, "T41")
        ws.column_dimensions["AN"].hidden = True
        ws.column_dimensions["AO"].hidden = True
    elif pack.entity.get("listing_status") != "listed":
        ws["T41"] = "주가 추이: 비상장으로 시장가격 정보 없음"

    for column in range(1, 39):
        ws.column_dimensions[get_column_letter(column)].width = 3.2
    ws.column_dimensions["C"].width = 12
    ws.row_breaks.append(__import__("openpyxl").worksheet.pagebreak.Break(id=47))
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 2


def _write_subsidiary_sheet(ws, pack: FactPack) -> None:
    ws.title = "자회사 등"
    headers = ["사업군", "자회사명", "지분율", "사업영역", "소재지", "주요 생산시설 또는 핵심 역량", "비고", "출처 문서"]
    _apply_table_header(ws, 2, headers)
    if not pack.subsidiaries:
        ws.merge_cells("A3:H3")
        ws["A3"] = "자동 추출 대상 공시가 확보되지 않았습니다. 사업보고서 주석 확인 필요"
        ws["A3"].alignment = LEFT_WRAP
    else:
        for row, subsidiary in enumerate(pack.subsidiaries, start=3):
            for col, header in enumerate(headers, start=1):
                ws.cell(row, col, subsidiary.get(header, "확인 필요"))
    for col, width in enumerate([18, 24, 10, 24, 18, 44, 24, 32], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def _write_matters_sheet(ws, pack: FactPack) -> None:
    ws.title = "주요사항"
    headers = ["분류", "사실", "해석", "검증 상태", "출처 문서명", "공시일", "URL"]
    _apply_table_header(ws, 2, headers)
    for row, matter in enumerate(pack.major_matters, start=3):
        values = [matter.category, matter.fact, matter.interpretation, matter.verification_status, matter.source_title, matter.disclosure_date, matter.url]
        for col, value in enumerate(values, start=1):
            ws.cell(row, col, value)
            ws.cell(row, col).alignment = LEFT_WRAP if col in {2, 3, 5, 7} else CENTER
            ws.cell(row, col).border = Border(left=HAIR, right=HAIR, top=HAIR, bottom=HAIR)
    if not pack.major_matters:
        ws.merge_cells("A3:G3")
        ws["A3"] = "요청 이슈 공시가 없습니다. 이슈 검색어 또는 조회 기간을 확인하세요."
    widths = [14, 38, 48, 14, 34, 14, 55]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def render_report(pack: FactPack, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    summary = workbook.active
    financial = workbook.create_sheet()
    financial_data = workbook.create_sheet()
    subsidiaries = workbook.create_sheet()
    matters = workbook.create_sheet()
    _write_data_sheet(financial_data, pack)
    _write_financial_sheet(financial, pack)
    _write_summary_sheet(summary, pack)
    _write_subsidiary_sheet(subsidiaries, pack)
    _write_matters_sheet(matters, pack)
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    safe_name = str(pack.entity.get("company_name", "기업")).replace("/", "_")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"{safe_name}_기업분석보고서_{stamp}.xlsx"
    workbook.save(output_path)
    (output_path.with_suffix(".factpack.json")).write_text(json.dumps(pack.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path
