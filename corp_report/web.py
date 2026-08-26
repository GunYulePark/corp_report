from __future__ import annotations

import asyncio
import os
from datetime import date
from pathlib import Path

from nicegui import app, run, ui
from dotenv import load_dotenv

from .collector import DartFactPackCollector
from .models import ReportRequest
from .report_renderer import render_report
from .validation import validate_fact_pack


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = PROJECT_ROOT / "outputs"
load_dotenv(PROJECT_ROOT / ".env")

try:
    from .local_settings import DART_API_KEY as EMBEDDED_DART_API_KEY
except ImportError:
    EMBEDDED_DART_API_KEY = ""

DEFAULT_DART_API_KEY = os.getenv("DART_API_KEY", EMBEDDED_DART_API_KEY)


def _parse_years(text: str) -> list[int]:
    values = []
    for part in text.replace(" ", "").split(","):
        if not part:
            continue
        year = int(part)
        if not 2000 <= year <= date.today().year:
            raise ValueError("재무연도는 2000년부터 현재 연도 사이여야 합니다.")
        values.append(year)
    if not values:
        raise ValueError("최소 한 개의 재무연도를 입력하세요.")
    return sorted(set(values))


def create_app() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    app.add_static_files("/reports", str(OUTPUT_DIR))
    ui.page_title("기업 분석 보고서 자동화")
    ui.colors(primary="#002060", secondary="#146c94", accent="#f3b31b")
    with ui.column().classes("w-full max-w-6xl mx-auto p-6 gap-5"):
        ui.label("기업 분석 보고서 자동화").classes("text-3xl font-bold text-slate-800")
        ui.label("OpenDART 재무정보와 공시를 바탕으로 별도/연결 기준 Excel 보고서를 생성합니다.").classes("text-slate-600")
        with ui.card().classes("w-full p-5"):
            ui.label("보고서 조건").classes("text-lg font-bold")
            with ui.row().classes("w-full gap-4"):
                identifier = ui.input("회사명 또는 종목코드", placeholder="예: 삼성전자 또는 005930").classes("w-72")
                basis = ui.select({"OFS": "별도 (기본)", "CFS": "연결"}, value="OFS", label="재무제표 기준").classes("w-48")
                years = ui.input("재무연도", value=",".join(str(year) for year in range(date.today().year - 4, date.today().year)), placeholder="예: 2022,2023,2024,2025").classes("w-72")
            with ui.row().classes("w-full gap-4 items-center"):
                issue = ui.input("확인할 이슈", placeholder="예: 투자, 계약, 인허가, 소송, M&A").classes("w-[32rem]")
                latest_interim = ui.checkbox("최신 분기·반기 포함", value=True)
                price_chart = ui.checkbox("상장사 주가 추이 포함", value=True)
            api_key = ui.input("OpenDART API 키", password=True, password_toggle_button=True, value=DEFAULT_DART_API_KEY).classes("w-full")
            ui.label("API 키는 브라우저에 저장하지 않으며, 이번 생성 작업에만 사용합니다.").classes("text-xs text-slate-500")

        status = ui.label("조건을 입력한 뒤 보고서를 생성하세요.").classes("text-slate-600")
        result = ui.column().classes("w-full")

        async def generate() -> None:
            result.clear()
            try:
                report_request = ReportRequest(
                    identifier=str(identifier.value or "").strip(),
                    fs_div=str(basis.value),
                    years=_parse_years(str(years.value or "")),
                    issue_query=str(issue.value or "").strip(),
                    include_latest_interim=bool(latest_interim.value),
                    include_price_chart=bool(price_chart.value),
                )
                if not report_request.identifier:
                    raise ValueError("회사명 또는 종목코드를 입력하세요.")
                status.text = "공시·재무정보를 수집하고 있습니다…"
                collector = DartFactPackCollector(str(api_key.value or ""))
                pack = await run.io_bound(collector.collect, report_request)
                pack.validation_results = validate_fact_pack(pack)
                status.text = "Excel 보고서를 생성하고 있습니다…"
                output_path = await run.io_bound(render_report, pack, OUTPUT_DIR)
            except Exception as exc:
                status.text = f"생성 실패: {exc}"
                ui.notify(str(exc), type="negative")
                return

            warnings = [item for item in pack.validation_results if item.get("severity") == "warning"]
            status.text = f"완료: {pack.entity.get('company_name')} 보고서를 생성했습니다. 검증 경고 {len(warnings)}건"
            with result:
                ui.link("Excel 보고서 다운로드", f"/reports/{output_path.name}", new_tab=True).classes("text-lg font-bold text-primary")
                ui.label(f"Fact Pack JSON도 같은 폴더에 저장되었습니다. 재무제표 기준: {'별도' if report_request.fs_div == 'OFS' else '연결'}")
                if warnings:
                    with ui.expansion(f"검증 경고 {len(warnings)}건", icon="warning"):
                        for warning in warnings[:20]:
                            ui.label(f"• {warning.get('message', '')}")
            ui.notify("보고서 생성이 완료되었습니다.", type="positive")

        ui.button("Excel 기업분석 보고서 생성", on_click=generate, icon="description").props("unelevated size=lg color=primary")

    ui.run(title="기업 분석 보고서 자동화", port=8080, reload=False)
