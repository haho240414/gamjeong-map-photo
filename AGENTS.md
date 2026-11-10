# 작업 지침 (Claude / Codex 등 AI 에이전트용)

이 저장소는 감정평가 **지도 사진자료**(광역위치도·상세위치도·지적도 JPG 3장 + 3페이지 Word)를 만드는 도구다.
사용자가 주소를 주면 아래 절차대로 **실제 지도 캡처**로 자료를 만들고, 검수한 뒤 Word와 JPG를 전달한다.
사람이 읽는 설명은 [`docs/USAGE.md`](docs/USAGE.md), 이 문서는 에이전트가 따라야 할 규칙이다.

## 0. 준비 (최초 1회)

```bash
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt     # Pillow, python-docx
# Google Chrome 필요 (헤드리스로 실행). macOS: /Applications/Google Chrome.app
```

- **Windows:** `run_windows.bat`(웹 화면) 또는 `py -3 -m venv .venv` 후 `.venv\Scripts\pip install -r requirements.txt`. 이후 명령의 `./.venv/bin/python`은 `.venv\Scripts\python`으로 바꾼다. Chrome(없으면 Edge)은 자동 탐색하며 `CHROME_PATH`로 지정할 수 있다. 콘솔 출력은 자동으로 UTF-8로 맞춘다.
- 검수용으로 Word→PDF→PNG 변환이 필요하면 `pymupdf`를 추가 설치한다(`pip install pymupdf`). macOS는 LibreOffice가 없으면 Pages로 docx→PDF 변환.
- API 키는 **필수 아님**(`.env.example` 참고). 키가 없어도 `--capture-portal`로 실제 지도 캡처가 된다. `.env`는 절대 커밋하지 않는다.

## 1. 입력 해석

| 사용자 표현 | 인자 |
|---|---|
| 물건지 (기존토지) | `--subject-address` (쉼표로 여러 필지 가능) |
| 비교표준지 | `--standard-address` (여러 개면 인자를 반복 → 비교표준지1, 2…) |
| 평가사례지 | `--case-address` (인자 1개 = 사례 1건. **쉼표로 나열한 필지는 한 건**으로 묶여 라벨이 하나) |
| 거래사례는 쓰지 않음 | `--skip-trades` 항상 붙인다 |

- 같은 리/동이면 앞 필지에만 전체 주소를 쓰고 나머지는 지번만 써도 된다 (`"경기도 양주시 덕정동 100-1, 100-2"`).
- 사용자가 `/` 로 구분해 여러 사례를 주면 **각각 별도의 `--case-address`** 로 넘긴다.
- "산-22-5"처럼 쓴 지번은 "산 22-5"로 정규화된다(산=임야 대지구분).
- 새로운 법정동(리)이면 `map_photo_prototype.py`의 `DONG_MAP`/`SIGUNGU_MAP`에 코드가 없을 수 있다 → §6 참고.

## 2. 실행

```bash
./.venv/bin/python map_photo_prototype.py \
  --subject-address "경기도 양주시 덕정동 366-12" \
  --standard-address "경기도 양주시 덕정동 190-5" \
  --case-address "경기도 양주시 덕정동 138-14, 138-88, 138-37" \
  --capture-portal --skip-trades \
  --appraisal-no "사건번호" --output-dir ./outputs
```

- 결과: `outputs/photo_material_YYYYMMDD_HHMMSS/` 아래 JPG 3장 + `metadata.json` + `경기부동산포털_캡처_메모.txt`, 그리고 Word 1개(기본 `~/Documents/Codex/지적도 자료/`, `--word-output-dir` 또는 환경변수 `MAP_PHOTO_WORD_DIR`로 변경).
- 경기부동산포털이 느리면 1~5분 걸리거나 실패할 수 있다. **실패 시 같은 명령을 다시 실행**한다(§5).
- **채무자 기존토지 + 신규 매입지 모드:** `--existing-new`. 물건지 인자=기존토지(빨강, 라벨 "기존토지"), `--case-address`=신규 매입지(주황, 라벨 "신규 매입지"). 광역·상세·지적도 모두에 둘 다 표기하고 비교표준지 자동탐색은 하지 않는다.
- 웹 입력 화면: `./.venv/bin/python webapp_server.py --open` (포트 8877).

## 3. 지도 규칙 (절대 지킬 것)

1. **광역위치도 = 카카오맵 스카이뷰 실제 캡처**, 물건지 중심, 약 50m 눈금(반경 약 500m). 물건지 필지 경계를 빨강으로, 라벨 "본건위치". **광역위치도에는 물건지만** 표시한다(`--existing-new`일 때만 신규 매입지도 표시).
2. **상세위치도 = 경기부동산포털 캡처.** 물건지(빨강)·비교표준지(파랑)·평가사례(주황)를 모두 화면에 넣는다. 멀리 있는 참고 필지가 있으면 화면이 넓어지고 물건지가 작아질 수 있다 → 사용자에게 알린다.
3. **지적도 = 경기부동산포털 연속지적도.** 물건지 경계와 지번이 선명해야 한다.
4. 색: 물건지 빨강 / 비교표준지 파랑 / 평가사례 주황 / 거래사례 초록. 필지와 라벨의 대응이 분명해야 한다.
5. **가짜·도식 지도 금지.** 산출 경고에 `경기부동산포털 자동 캡처 실패` 또는 `VWorld 배경 타일…`이 있으면 도식 지도로 대체된 것이다 → **그 결과는 버리고 재실행**한다.
6. 지도 중심을 비교표준지·평가사례로 옮기지 않는다.

## 4. 검수 (전달 전 반드시)

1. `outputs/.../portal_capture/portal_raw_detail.json`의 `result.value`에서 `found` 개수가 입력 필지 수와 같고 `missing`이 `[]`인지 확인. **누락 필지가 있으면 전달하지 말고** 주소/PNU를 재확인한다.
2. JPG 3장을 직접 열어 확인: 핀·경계·라벨 색/대응, 잘림, 빈 공간, 남은 검색패널/팝업.
3. Word를 PDF/PNG로 렌더링해 3페이지·이미지 비율·넘침을 확인한다(생략했다면 생략했다고 사용자에게 말한다).
4. 존재하지 않는 지번이 있으면(지적도에 없음) **임의로 그리지 말고** 사용자에게 알려 정확한 지번을 받는다. 예: 연속지적도 조회로 해당 리의 지번 목록을 확인할 수 있다(§6).

## 5. 알려진 함정

- 카카오맵 UI가 바뀌면 스카이뷰/확대 버튼 셀렉터가 깨진다 (`kakao_capture_chrome.py`의 `click_skyview`, `find_zoom_*`는 옛 `title` 속성과 새 `ico_scale_*`/`skyview` 클래스 둘 다 찾는다).
- 카카오 줌은 불연속(30→50→100→250→500m)이라 중간 배율은 못 만든다. 창 크기를 키워서 넓히려 하지 말 것(좌우 팝업이 크롭에 섞인다).
- 오버레이는 항상 `m/px`(축척 막대를 DOM에서 읽어 계산)로 환산한다. 줌마다 비율이 달라 하드코딩 금지.
- `portal_capture_chrome.py`의 JS는 파이썬 f-string 안에 있어 중괄호를 `{{ }}`로 이중화해야 한다.
- 같은 색 라벨이 둘 이상이면 첫 번째는 필지 아래, 두 번째·세 번째는 좌우로 배치된다(`labelOffsets`).
- 헤드리스 Chrome은 임시 프로필로 실행되고 끝나면 정리된다. 디버그 포트는 자동 선택.

## 6. 자주 하는 확장

- **새 법정동 추가:** 연속지적도 REST로 BCD(법정동코드 10자리)를 확인해 `DONG_MAP`에 뒤 5자리를 넣는다.
  ```bash
  curl -sk -G "https://gris.gg.go.kr:8888/grisgis/rest/services/bdsMap_Cbnd/MapServer/5/query" \
    --data-urlencode "where=JIBUN_NM LIKE '%덕정동 366-12%'" \
    --data-urlencode "outFields=PNU,JIBUN_NM,BCD" --data-urlencode "returnGeometry=false" --data-urlencode "f=json"
  ```
  PNU = 법정동코드(10) + 대지구분(1: 일반, 2: 산) + 본번(4) + 부번(4).
  `PNU LIKE '…%'` 조회로 특정 리의 지번 목록도 볼 수 있다(페이지네이션 미지원 → 짧은 접두어로 나눠 조회).
- 시군구가 경기도 밖이면 경기부동산포털 자체가 지원하지 않으므로 이 도구의 범위 밖이다.

- **평가사례가 멀 때(수 km):** 한 화면에 다 넣으면 본건이 점이 된다. 사용자가 "멀어도 전부 나오게"를 택했으면 기본 3쪽(상세위치도에 사례 포함) 뒤에 본건·비교표준지만 확대한 상세도를 덧붙인다: `--extra-detail "<확대 상세 배경 이미지>|상세위치도(본건·비교표준지 확대)"` (여러 번 가능). 확대용 배경은 사례 없이 한 번 더 `--capture-portal`로 만들어 `portal_capture/portal_bg_detail.png`를 쓴다.
- **한글 서식(위치도 34-580·건물개황도 34-581) 재현:** `form_documents.py`(`build_location_form`, `build_building_form`)가 A4·표 치수까지 서식을 Word로 재현한다. 건물개황도 도면 칸 그림은 `building_sketch.py`(층별 외곽+치수선 PNG)로 만든다. 면적·용도·구조는 **건축물대장 값**을 쓰고, 현황도(스캔) 치수 계산값은 대장과 3~5% 어긋날 수 있으니 대장이 없으면 반드시 그 한계를 고지한다. 건별 데이터는 `jobs/<건명>/`에 두며 `jobs/`는 `.gitignore` 대상이다(의뢰 건 정보 보호).
- docx→PDF 검수가 필요하면 맥은 Pages(`open -a Pages` 후 AppleScript export), PDF→PNG는 pymupdf. QA 렌더용 폰트는 `FORM_FONT` 환경변수로 바꿀 수 있다(납품 기본은 맑은 고딕).

## 7. 개인정보·저작권

- `outputs/`, `.env`, 생성된 docx/jpg는 `.gitignore` 대상이다. **실제 의뢰 건(채무자·소유자 정보가 있는 결과물)을 저장소에 올리지 않는다.**
- 카카오맵·경기부동산포털 화면 캡처는 내부 검토·업무 자료 용도다. 외부 공개·유료 서비스·대량 재배포는 각 제공자의 약관을 먼저 확인한다.
