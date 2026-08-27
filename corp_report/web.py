from __future__ import annotations

import asyncio
import importlib.util
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

def _embedded_secret(name: str) -> str:
    """Load a local-only secret on each report run without exposing it to the browser."""
    settings_path = PROJECT_ROOT / "corp_report" / "local_settings.py"
    if not settings_path.exists():
        return ""
    try:
        spec = importlib.util.spec_from_file_location("corp_report_runtime_settings", settings_path)
        if spec is None or spec.loader is None:
            return ""
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return str(getattr(module, name, "")).strip()
    except Exception:
        return ""


def _embedded_api_key() -> str:
    """Backward-compatible OpenDART key accessor used by local tooling and tests."""
    return _embedded_secret("DART_API_KEY")


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
                issue = ui.input("이슈 조사 요청", placeholder="예: 최근 3.2조 기술수출과 국내 비만치료제 기술수출 금액·파이프라인 비교").classes("w-[32rem]")
                latest_interim = ui.checkbox("최신 분기·반기 포함", value=True)
                price_chart = ui.checkbox("상장사 주가 추이 포함", value=True)
            ui.label("이슈 조사 방식: 뉴스·공식 출처의 제한 원문을 수집한 뒤 Gemini가 사실·비교·시사점 JSON으로 정리합니다. 기사 기반 결과는 ‘검토 필요’로 표시합니다.").classes("text-xs text-slate-500")
            ui.label("OpenDART API 키는 로컬 설정 파일에서만 읽습니다.").classes("text-xs text-slate-500")

        status = ui.label("조건을 입력한 뒤 보고서를 생성하세요.").classes("text-slate-600")
        with ui.row().classes("items-center gap-2") as loading_row:
            ui.spinner(size="lg", color="primary")
            ui.label("작업을 시작하는 중입니다…").classes("text-primary font-medium")
        loading_row.visible = False
        result = ui.column().classes("w-full")

        async def generate() -> None:
            result.clear()
            loading_row.visible = True
            generate_button.disable()
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
                status.text = "1/2 공시·재무정보와 요청 이슈 출처를 수집하고 Gemini 분석을 실행하고 있습니다…"
                await asyncio.sleep(0)
                resolved_api_key = os.getenv("DART_API_KEY", "").strip() or _embedded_api_key()
                resolved_gemini_key = os.getenv("GEMINI_API_KEY", "").strip() or _embedded_secret("GEMINI_API_KEY")
                collector = DartFactPackCollector(resolved_api_key, gemini_api_key=resolved_gemini_key)
                pack = await run.io_bound(collector.collect, report_request)
                pack.validation_results = validate_fact_pack(pack)
                status.text = "2/2 Excel 보고서를 생성하고 있습니다. 시트·수식·주가 그래프를 작성 중입니다…"
                await asyncio.sleep(0)
                output_path = await run.io_bound(render_report, pack, OUTPUT_DIR)
            except Exception as exc:
                status.text = f"생성 실패: {exc}"
                ui.notify(str(exc), type="negative")
                return
            finally:
                loading_row.visible = False
                generate_button.enable()

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

        generate_button = ui.button("Excel 기업분석 보고서 생성", on_click=generate, icon="description").props("unelevated size=lg color=primary")

    ui.run(title="기업 분석 보고서 자동화", port=8080, reload=False)
