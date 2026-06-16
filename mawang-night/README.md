# MAWANG NIGHT — 대표 보고용 기획 보고서

마왕족발 브랜드 이슈메이킹 프로젝트 "MAWANG NIGHT(가제)" 기획안을 **A4 가로 11장** 슬라이드로 정리한 보고서.
무드: **블랙 · 레드 · 호러** / 브랜드 레드 `#C4141B`.

## 파일
- `MAWANG_NIGHT_기획보고.pdf` — **대표님께 바로 전달/인쇄 가능한 보고본** (A4 가로)
- `index.html` — 디자인 원본(브라우저에서 열람). 폰트는 `fonts/` 자체 포함이라 오프라인에서도 정상 표시
- `assets/` — 브랜드 이미지 자리 (`mascot.svg` = 임시 데빌 엠블럼)
- `render.mjs` — HTML → PDF 재생성 스크립트

## 보는 법 / 다시 인쇄
- 그냥 보기: `index.html`을 크롬에서 열기
- PDF로 저장: 크롬에서 `index.html` 열고 **인쇄 → 대상 PDF → 용지 A4, 가로, 배경 그래픽 켜기**
- 스크립트로 재생성:
  ```bash
  npm install        # 최초 1회 (puppeteer)
  node render.mjs     # MAWANG_NIGHT_기획보고.pdf 생성
  ```

## 공식 로고로 교체하는 법
현재 표지/머리말의 마왕 캐릭터는 **임시 엠블럼**(`assets/mascot.svg`)이고, "마왕족발" 워드마크는
공식 컬러(`#C4141B`)로 **폰트로 재현**한 버전입니다. 공식 로고 이미지를 그대로 쓰려면:

1. 받은 로고 파일을 `assets/`에 아래 이름으로 저장
   - `assets/logo.png` (가로형 풀 로고), `assets/mascot.png` (데빌 캐릭터)
2. `index.html`에서 `<svg ...><use href="#mask"/></svg>` 엠블럼을
   `<img src="assets/mascot.png">`로, 워드마크 영역을 `<img src="assets/logo.png">`로 교체
3. `node render.mjs` 재실행

> 채팅으로 보내주신 이미지는 디스크에 파일로 저장되지 않아 자동 삽입이 불가합니다.
> 위 이름으로 레포에 올려주시면 그대로 반영해 드립니다.
