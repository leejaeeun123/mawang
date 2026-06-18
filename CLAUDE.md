# CLAUDE.md — 프로젝트 메모리

## 프로젝트
- **RED HOUR (레드 아워, 가제)** — 마왕족발 브랜드 경험 프로젝트 기획안 덱
- 위치: `mawang-night/`
- 무드: 블랙 · 레드(`#C4141B`) / 폰트: **Pretendard**(woff2 자체 번들, `mawang-night/fonts/`)
- 작업 브랜치: **`claude/zen-hamilton-8v7nf7`** (여기에만 커밋·푸시) / PR: **#1**

## 산출물 (단일 체계 — PPTX는 만들지 않음)
- `mawang-night/RED_HOUR_기획안.pdf` — 최종 PDF (A4 가로, 22장)
- `mawang-night/index.html` — 디자인 원본
- `mawang-night/render.mjs` — `node render.mjs` 로 HTML→PDF 재생성 (puppeteer, `npm install` 최초 1회)
- `mawang-night/assets/` — 실제 이미지
- ⚠️ 사용자 요청으로 **편집형 .pptx 산출 제거**(`build_pptx.py`, `RED_HOUR.pptx` 삭제됨). 다시 만들지 말 것.

## 덱 구성 (본편 17 + 부록 5 = 22장)
- P01 표지 / P02 개요 / P03 기획 방향 재정립 / P04 컨셉 / P05 공간 방향성 /
  P06 드레스코드 / P07 스태프 / P08 일정 / P09 VIP / P10 RED HOUR EXP / P11 NIGHT PARTY /
  P12 푸드 / P13 굿즈 / P14 콘텐츠 / P15 행사 이후 / P16 기대효과 / P17 KEY MESSAGE
- 부록 A1 공간비교 / A2 어덜트 온리 / A3 로우키 / A4 굿즈견적(뉴리디파인) / A5 굿즈견적(브랜다즐)

## 반영 완료된 데이터/이미지
- 로고: `assets/logo_white.png` (원본 로고에서 ‘마왕’→화이트, ‘족발’→레드 변환)
- 캐릭터 BI: `assets/character.png` (원본 그대로, 변형 금지) → P07
- 공간: 어덜트온리 `venue_adult_*.jpg` (A2/P05), 로우키 `venue_lowkey_*.jpg` (A3/P05)
- 굿즈: `goods_led.jpg`, `goods_box.jpg` → P13
- 견적: A4 뉴리디파인(확정 표), A5 브랜다즐(주식회사 에이블러, brandazzle_cs@brandazzle.kr, PDF 기반 실표)

## 확정 규칙 / 사용자 선호
- "대표님 보고용" 등 보고 대상 단어 사용 금지
- 무한도전 키캡 키링 MZ 어필 레퍼런스 사용 금지
- 소통 요약(A2/A3): **우리(MAWANG)=오른쪽(레드 말풍선)**, **상대=왼쪽(그레이)**
- 채팅 인라인 붙여넣기 이미지는 디스크 저장 안 됨 → **파일/zip 첨부**로 받아야 사용 가능

## 남은(선택) 자리표
- P10 현장 경험컷 / P11 파티·DJ컷 / P12 푸드(DEVIL COLD SKEWER, FIRE CUP) — 사진 받으면 `.phimg`로 교체

## 공간 정보 요약
- ① 어덜트 온리(가로수길, @adultonlybar): 평일 600만/주말 800만(VAT별도, 조정가능), 초대제 중심,
  서버·바텐더 시급 16,000원, DJ 15~20만, 지하1·2층 복층 단독출입구
- ② 로우키(압구정로데오, @lowkeyseoul, 대표 주현성): 단순대관 시간당 30만(팝업+파티 시 보장매출 따라 조정),
  조명 조절 가능, 푸드트럭 자리 협의 필요, 장마철만 회피
