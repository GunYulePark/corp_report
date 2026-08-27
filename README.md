# 기업 분석 보고서 자동화

회사명 또는 종목코드를 입력하면 OpenDART 공시 재무정보를 바탕으로 Excel 기업 분석 보고서를 생성하는 로컬 웹 앱입니다.

## 실행

```powershell
py -m pip install -r requirements.txt
$env:DART_API_KEY = "발급받은_OpenDART_API_키"
python app.py
```

반복 사용 시에는 프로젝트 최상위에 `.env` 파일을 만들고 아래처럼 저장할 수 있습니다. `.env`는 Git에 포함되지 않습니다.

```text
DART_API_KEY=발급받은_OpenDART_API_키
```

프로그램 내부 설정 파일을 쓰려면 `corp_report/local_settings.py.example`을 같은 폴더의 `local_settings.py`로 복사한 뒤 `DART_API_KEY` 값에 키를 입력합니다. 이 파일도 Git에서 제외됩니다. 단, 실행 파일을 다른 사람에게 배포하면 키가 노출될 수 있으므로 개인용으로만 사용하세요.

브라우저에서 `http://localhost:8080`을 열어 회사명 또는 종목코드, 별도/연결 기준, 재무연도, **이슈 조사 요청**을 입력합니다. 기본 기준은 별도(OFS)입니다. 이슈를 입력하면 요청 문장을 기준으로 웹 검색 결과와 확인 가능한 공식 자료를 `주요사항` 시트에 사실·해석·출처 URL로 남깁니다. 기사 제목만 확인된 결과는 `needs_review`로 표시하며, 빈 이슈란을 최근 공시 목록으로 대체하지 않습니다.

`corp_report/local_settings.py`에 `GEMINI_API_KEY`를 추가하면, 이슈 조사 시 Gemini Google Search grounding을 우선 사용합니다. 반환된 실제 인용 URL과 일치하는 내용만 Fact Pack에 저장하고, 본장에는 사용자 이슈를 일반 회사소개보다 먼저 표시합니다. 검색 근거는 기사·웹 출처이므로 `검토 필요`로 표시하며, 키 미설정·할당량 초과·호출 실패 시에는 Google News 제목·URL 기반의 출처 결과로 자동 대체합니다. 키는 서버 프로세스에서만 읽고 Excel, Fact Pack JSON, 브라우저에 기록하지 않습니다.

## 기존 수집기 연결

기본적으로 `C:\Users\CKD\dart-fss\dart_app_4.py`를 사용합니다. 다른 위치에 있으면 환경변수로 지정합니다.

```powershell
$env:DART_PIPELINE_PATH = "D:\path\to\dart_app_4.py"
```

## 결과물

생성 파일은 `outputs/`에 저장됩니다.

- `본장`: 임원 검토용 요약, 재무 연결표, 상장사의 1년 주가 추이
- `재무`: 표시용 재무 분석 및 비율
- `재무 data`: 출처가 있는 표준 원천 재무 Fact
- `자회사 등`: 자회사 근거 표
- `주요사항`: 입력 이슈의 웹 조사 사실·해석·출처 및 주가 변동 구간 대조

주가 데이터는 Yahoo Finance 차트 API를 보조 출처로 사용합니다. 주가와 이슈의 날짜가 가까운 경우에도 보고서에는 시간적 근접성만 표시하며, 단일 원인이라는 인과관계로 단정하지 않습니다.
