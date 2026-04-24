# usdrate

CSV 파일의 날짜 컬럼을 기준으로 한국수출입은행 환율 API에서 USD 매매기준율을 조회해 `환율` 컬럼을 추가하는 경량 도구입니다.

브라우저 서버 없이 동작하며, 다음 두 방식으로 사용할 수 있습니다.

- Python 코드 실행
- Windows `exe` 실행

## API 키

이 프로그램을 사용하려면 한국수출입은행 환율 API 키가 반드시 필요합니다.

API 키 발급 페이지:

- [공공데이터포털 한국수출입은행 환율 정보](https://www.data.go.kr/data/3068846/openapi.do)

위 페이지에서 활용신청 후 받은 API 키를 `--api-key` 옵션이나 GUI 입력창에 넣어 사용하면 됩니다.

## 다운로드

Windows 실행 파일은 GitHub 릴리즈 페이지에서 받을 수 있습니다.

- [Releases 페이지](https://github.com/simdorei/usdrate/releases)
- [최신 릴리즈 바로가기](https://github.com/simdorei/usdrate/releases/latest)

릴리즈 페이지의 `Assets`에서 `usdrate.exe`를 다운로드한 뒤 바로 실행하면 됩니다.

## 포함 파일

- `usdrate.py`: GUI와 CLI를 함께 제공하는 메인 스크립트
- `build_usdrate.ps1`: Windows용 `exe` 빌드 스크립트

## 실행 방법

GUI:

```powershell
py -3 .\usdrate.py
```

CLI:

```powershell
py -3 .\usdrate.py --input .\input.csv --api-key YOUR_API_KEY
```

옵션:

- `--output`: 결과 CSV 경로 지정
- `--date-column`: 날짜 컬럼명 직접 지정

## exe 빌드

```powershell
powershell -ExecutionPolicy Bypass -File .\build_usdrate.ps1
```

빌드 결과:

- `dist_light\usdrate.exe`

## 출력 파일

기본 출력 파일명은 원본 파일명 뒤에 `_환율추가.csv`를 붙입니다.

예:

- `input.csv` -> `input_환율추가.csv`

## 참고

- CSV만 지원합니다.
- 결과 파일은 UTF-8 BOM(`utf-8-sig`)으로 저장합니다.
- 날짜 컬럼은 자동 탐지하며, 필요하면 직접 지정할 수 있습니다.
