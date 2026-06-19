# RED HOUR — 대표 보고용 기획 보고서

마왕족발 브랜드 경험 프로젝트 **"RED HOUR(가제)"** 기획안을 **A4 가로 22장** 슬라이드로 정리한 보고서.
무드: **블랙 · 레드 · 호러** / 브랜드 레드 `#C4141B`.

## 구성 (22장)
- 01 표지 / 02 프로젝트 개요 / 03 기획 방향 재정립 / 04 컨셉 / 05 공간 / 06 드레스코드 / 07 스태프
- 08 일정(타임테이블) / 09 VIP / 10 RED HOUR EXPERIENCE / 11 NIGHT PARTY / 12 푸드 / 13 굿즈
- 14 콘텐츠 전략 / 15 행사 이후 / 16 기대효과 / 17 KEY MESSAGE
- 부록 A1 공간 비교 / A2 어덜트온리 / A3 로우키 / A4 견적①(뉴리디파인) / A5 견적②(브랜다즐)

## 파일
- `RED_HOUR_기획안.pdf` — **대표님께 바로 전달/인쇄 가능한 보고본** (A4 가로 22장)
- `index.html` — 디자인 원본(브라우저에서 열람). 폰트는 `fonts/` 자체 포함이라 오프라인에서도 정상 표시
- `assets/img/` — 본문 이미지(드레스코드·스태프·푸드), `assets/img/venue/` — 공간 이미지
- `render.mjs` — HTML → PDF 재생성 스크립트 (puppeteer)

## 보는 법 / 다시 인쇄
- 그냥 보기: `index.html`을 크롬에서 열기
- PDF로 저장: 크롬에서 `index.html` 열고 **인쇄 → 대상 PDF → 용지 A4, 가로, 배경 그래픽 켜기**
- **Chrome 헤드리스로 재생성 (Node 불필요):**
  ```bash
  chrome --headless=new --disable-gpu --no-pdf-header-footer \
    --print-to-pdf="RED_HOUR_기획안.pdf" "file:///<경로>/index.html"
  ```
- puppeteer 스크립트로 재생성 (Node 설치 시):
  ```bash
  npm install        # 최초 1회 (puppeteer)
  node render.mjs     # RED_HOUR_기획안.pdf 생성
  ```

## 공식 로고로 교체하는 법
현재 머리말의 마왕 캐릭터는 **임시 엠블럼**(SVG)이고, "마왕족발" 워드마크는 폰트로 **흰색** 재현한 버전입니다.
공식 로고/CI 캐릭터 이미지를 그대로 쓰려면:

1. 받은 파일을 `assets/`에 저장 — `assets/logo.png`(워드마크), `assets/mascot.png`(데빌 캐릭터)
2. `index.html`의 `<svg ...><use href="#mask"/></svg>` 엠블럼과 `.wm` 워드마크 영역을 `<img>`로 교체
3. Chrome 헤드리스 또는 `node render.mjs` 재실행

> 채팅으로 보내주신 이미지는 자동 삽입이 어렵습니다. 위 이름으로 레포에 올려주시면 반영해 드립니다.
