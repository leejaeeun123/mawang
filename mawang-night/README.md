# RED HOUR — 마왕족발 브랜드 경험 프로젝트

마왕족발 브랜드 경험 프로젝트 **"RED HOUR(레드 아워, 가제)"** 기획안.
본편 17장 + 부록 5장(공간 대여 상세 · 굿즈 견적) = **A4 가로 22장** 슬라이드.
무드: **블랙 · 레드** / 브랜드 레드 `#C4141B` / 폰트: **Pretendard**.

## 산출물
- `RED_HOUR_기획안.pdf` — 바로 열람/인쇄/공유 가능한 PDF (A4 가로)
- `index.html` — 디자인 원본(브라우저 열람). 폰트는 `fonts/`에 자체 포함이라 오프라인에서도 정상 표시
- `assets/` — 표지 로고·캐릭터·공간·굿즈 등 실제 이미지

## 구성
- **본편 P01–P17** — 표지 / 프로젝트 개요 / 기획 방향 재정립 / 행사 컨셉 / 공간 방향성 /
  드레스코드 / 스태프 컨셉 / 행사 일정 / VIP·RED HOUR·NIGHT PARTY / 푸드 / 굿즈 /
  콘텐츠 전략 / 행사 이후 활용 / 기대 효과 / KEY MESSAGE
- **부록 A1–A5** — 공간 대여 검토(어덜트 온리 / 로우키 상세 + 소통 요약) ·
  굿즈 견적(① 뉴리디파인 표 · ② 브랜다즐 표)

## 다시 만들기 / 인쇄
```bash
npm install        # 최초 1회 (puppeteer)
node render.mjs    # -> RED_HOUR_기획안.pdf
```
- 그냥 보기: `index.html`을 크롬에서 열기
- PDF로 저장: 크롬에서 `index.html` → 인쇄 → 대상 PDF → 용지 A4, 가로, 배경 그래픽 켜기

## 사진 / 로고 교체
이미지는 `assets/`에 들어있고 `index.html`에서 `<img src="assets/...">`로 연결됩니다.
교체하려면 같은 파일명으로 `assets/`의 이미지를 바꾸거나, `index.html`의 경로를 수정한 뒤 `node render.mjs`를 다시 실행하세요.
- 표지 로고: `assets/logo_white.png` (원본 로고의 ‘마왕’을 화이트, ‘족발’은 레드로 변환)
- 캐릭터(BI): `assets/character.png` · 공간: `assets/venue_*.jpg` · 굿즈: `assets/goods_*.jpg`
