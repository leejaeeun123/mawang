# MAWANG NIGHT RUN : GHOST COURSE — 기획안 덱

마왕족발 야간 러닝 브랜드 행사 **"MAWANG NIGHT RUN : GHOST COURSE"** 기획안.
가로 A4 · PPT 형식 · HTML 기반 → PDF. 다크 호러 네온 레드 톤앤매너.

## 구성 (12 슬라이드)
표지 · INTRO · WHY · 프로젝트 목표 · 행사 컨셉 · 보상&굿즈 · 코스 흐름 · 행사 스펙 · 톤앤매너 · 무드보드 1·2 · 클로징

## 파일
- `index.html` — 단일 파일 덱(인라인 CSS 디자인 시스템)
- `fonts/` — Pretendard(본문) · Bebas Neue · Cinzel(디스플레이)
- `assets/` — 로고 · 캐릭터 · 무드보드 이미지(`img/mood/`)
- `MAWANG_NIGHT_RUN_기획안.pdf` — 가로 A4 고화질 산출물

## PDF 렌더 방법
Chrome 헤드리스로 렌더(한글 출력 경로는 exit 21 실패 → ASCII 임시경로 렌더 후 복사):

```bash
chrome --headless=new --disable-gpu --no-pdf-header-footer \
  --user-data-dir=<임시프로필> \
  --print-to-pdf=<ASCII경로>/nightrun.pdf \
  "file:///.../nightrun/index.html"
# 이후 한글 파일명으로 Copy-Item
```

## 원본
`마왕 나이트런.pdf`(Notion 5p 초안) 기반. 무드보드 이미지는 원본 4~5p에서 추출.
