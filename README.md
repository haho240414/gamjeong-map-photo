# 감정평가 지도 사진자료 생성기

물건지 주소를 입력하면 감정평가 사진자료용 `광역위치도`, `상세위치도`, `지적도` 이미지를 만들고 Word 문서로 배치하는 초기 모델입니다.

현재 모델은 기존 자체감정 코드에서 분리한 지도 사진자료 엔진입니다. 물건지는 빨간색, 비교표준지는 파란색, 평가사례는 주황색, 거래사례는 초록색으로 표시합니다.

## 주요 기능

- 물건지 주소 기반 지도 사진자료 생성
- 광역위치도, 상세위치도, 지적도 1페이지 1장 Word 양식 생성
- 여러 필지 물건지 주소 처리
- 비교표준지 수동 입력 또는 기존 `gamjeong` 이력 재사용
- 평가사례 주소 수동 입력
- 거래사례 수동 입력 및 공공데이터포털 키가 있을 때 자동조회 시도
- 경기부동산포털 화면 자동 캡처 모드
- API가 없을 때 양식 검증용 대체 도면 생성

## 설치

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Chrome 자동 캡처를 쓰려면 Windows에 Google Chrome이 설치되어 있어야 합니다.

## 환경변수

필요하면 `.env.example`을 참고해 `.env`를 만들고 키를 넣습니다. `.env`는 GitHub에 올리지 않습니다.

```text
GAMJEONG_VWORLD_KEY=...
GAMJEONG_INDVD_KEY=...
GAMJEONG_RTMS_KEY=...
```

## 웹 입력 화면 실행

```powershell
.\run_webapp.ps1
```

브라우저에서 `http://127.0.0.1:8877/`을 열어 물건지, 비교표준지, 평가사례를 입력합니다.

## 명령행 실행 예시

```powershell
python .\map_photo_prototype.py `
  --subject-address "경기도 양주시 봉양동 92, 96, 96-1, 97, 산 13-8" `
  --standard-address "경기도 양주시 봉양동 산 129" `
  --standard-address "경기도 양주시 봉양동 109" `
  --capture-portal `
  --skip-trades `
  --output-dir .\outputs
```

평가사례를 넣는 경우:

```powershell
python .\map_photo_prototype.py `
  --subject-address "경기도 의정부시 의정부동 158-7" `
  --standard-address "경기도 의정부시 의정부동 159-10" `
  --case-address "경기도 의정부시 의정부동 178-4" `
  --capture-portal `
  --skip-trades `
  --output-dir .\outputs
```

API 없이 양식만 확인하는 경우:

```powershell
python .\map_photo_prototype.py `
  --subject-address "경기도 양주시 남면 상수리 365-1" `
  --standard-address "경기도 양주시 남면 상수리 360" `
  --case-address "경기도 양주시 남면 상수리 366" `
  --no-network `
  --output-dir .\outputs
```

## 파일 구성

- `map_photo_prototype.py`: 주소 파싱, 지도/도면 이미지 생성, Word 사진자료 생성 엔진
- `portal_capture_chrome.py`: 경기부동산포털 지도서비스를 Chrome DevTools로 캡처하는 모듈
- `webapp_server.py`: 로컬 웹 입력 화면용 작은 HTTP 서버
- `static/index.html`: 입력 화면
- `compose_arcgis_overview.py`: 포털 항공사진 타일 누락 시 별도 배경과 필지 geometry를 합성하는 보조 스크립트

## 법적/운영 주의

현재 `--capture-portal` 방식은 경기부동산포털 화면을 자동 캡처합니다. 내부 검토와 개인 업무 자동화에는 유용하지만, 외부 공개 서비스나 유료 서비스로 제공하려면 지도 제공자 약관, 저작권, 출처 표기, 캐시/재배포 제한을 별도로 확인해야 합니다.

전국 단위 공개 서비스로 확장할 때는 화면 캡처보다 VWorld, 주소기반산업지원서비스, 공공데이터포털 등 정식 API와 사용승인 범위가 명확한 지도/지적도 소스를 사용하는 구조로 바꾸는 것이 안전합니다.
