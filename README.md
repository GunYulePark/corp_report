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

브라우저에서 `http://localhost:8080`을 열어 회사명 또는 종목코드, 별도/연결 기준, 재무연도, 확인할 이슈를 입력합니다. 기본 기준은 별도(OFS)입니다.

## 기존 수집기 연결

기본적으로 `C:\Users\CKD\dart-fss\dart_app_4.py`를 사용합니다. 다른 위치에 있으면 환경변수로 지정합니다.

```powershell
$env:DART_PIPELINE_PATH = "D:\path\to\dart_app_4.py"
```

## 결과물

생성 파일은 `outputs/`에 저장됩니다.

- `●[회사명]`: 요약 보고서 및 상장사의 1년 주가 추이
- `재무`: 표시용 재무 분석 및 비율
- `재무 data`: 출처가 있는 표준 원천 재무 Fact
- `자회사 등`: 자회사 근거 표
- `주요사항`: 요청 이슈 및 공시 근거

주가 데이터는 Yahoo Finance 차트 API를 보조 출처로 사용하며, 상장 여부 또는 가격 조회가 확인되지 않으면 보고서에 `N/A`로 표시됩니다.
