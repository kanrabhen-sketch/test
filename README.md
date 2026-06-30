# 일매출 관리 시스템 (MVP)

이관 문서 기준으로 실제 구현한 결과물입니다.

## 분석 결과 요약

- 전체 구조(Form → sales_logs → monthly_summary / missing_check)는 설계상 문제 없음.
- 이관 문서에 두 가지 버전(초안 + GPT 검수 후 최종본)이 있었는데, 내용은 동일하며
  `missing_check` 기준 셀이 설계상 C1이었으나 실제 파일은 **B1**으로 굳어진 점만 차이가 있음.
  → 이 구현에서는 **B1 기준으로 통일**했습니다 (`확인일` 레이블은 A1, 값은 B1).
- 수식 자체(ARRAYFORMULA/MAP/LET/HSTACK/FILTER/COUNTUNIQUE)는 모두 Google Sheets 전용이며
  Excel에서는 작동하지 않습니다. xlsx 파일에는 텍스트로 기록되어 있고, **Google Sheets로 열어야**
  정상 계산됩니다.
- `missing_check`의 `COUNTIFS(... , stores!$A$2:$A$200, ...)` 부분은 두 번째 인자에 배열을 넘기는
  비표준 사용법이라 일부 환경에서 `#N/A`/`#ERROR!`가 날 수 있습니다. 이를 대비해 Apps Script
  기반 대체 구현(`refreshMissingCheck`)을 추가했습니다.

## 만든 것

```
scripts/generate_template.py   xlsx 템플릿 생성 스크립트 (4개 탭, 수식, 조건부 서식, 데이터 검증 포함)
output/일매출관리시스템_템플릿.xlsx   생성된 템플릿 파일 (S001~S018 매장 데이터 포함)
apps_script/Code.gs             Apps Script 모음
```

### `output/일매출관리시스템_템플릿.xlsx`

이관 문서 3장(탭별 컬럼 구조)·5장(확정 수식)·6장(조건부 서식) 그대로 구현:

- **stores**: A:H, store_id(S001~S018) + 상태 active/inactive 드롭다운 검증
- **sales_logs**: A:K, G/H/I 수식 자동 채움(2행), D열 영업/휴무 드롭다운, J열 TRUE/FALSE 드롭다운,
  조건부 서식(중복 빨간 배경, 영업인데 0원 빨간 텍스트)
- **monthly_summary**: C1=기준연도(2026)/E1=기준월(6), 3행 헤더, 4~21행(18개 매장)에
  A~J 및 K~AO(1~31일) 수식 전부 채움, 조건부 서식(휴무/미입력/중복확인/--- 색상, 오늘 날짜 열 강조,
  전월대비 빨강/파랑)
- **missing_check**: B1=확인일(`=TODAY()`), 2행 요약 통계, A5에 HSTACK 수식

### `apps_script/Code.gs`

- `generateStoreLinks()` — stores H열에 매장별 pre-filled Form 링크 자동 생성
  (사용 전 `FORM_ENTRY_ID`, `FORM_BASE_URL`을 실제 값으로 교체 필요)
- `onFormSubmit(e)` — Form 제출 시 J열(확인여부) 기본값 FALSE 자동 채움 + 중복 입력 로그 알림
- `refreshMissingCheck()` — `missing_check`의 HSTACK/LET 수식이 오류를 낼 경우 쓰는 스크립트 기반 대체본
- `onOpen()` — 스프레드시트 메뉴에 "일매출관리" 메뉴 추가

## 사용자가 해야 할 작업 (Google 계정에서만 가능)

1. `output/일매출관리시스템_템플릿.xlsx`를 Google Drive에 업로드 → "Google Sheets로 열기"
2. `stores` 탭 매장명 오타 확인 (이미지 기반 입력이라 원본 대조 필요)
3. `sales_logs`에 이관 문서 11장 테스트 데이터 3~5줄 입력 → `monthly_summary`/`missing_check` 결과 확인
4. Google Form 생성 (질문 순서: 매장코드→영업일→영업여부→일매출→특이사항, 섹션 분기 설정)
5. Form 응답 시트를 `sales_logs`로 연결 (또는 응답 탭 이름을 `sales_logs`로 변경 후 G2/H2/I2 수식 재입력)
6. Apps Script 편집기에 `apps_script/Code.gs` 붙여넣기 → `FORM_ENTRY_ID`/`FORM_BASE_URL` 교체 →
   `generateStoreLinks` 실행 → stores H열에 매장별 링크 생성됨 → 점주에게 카카오톡 공유
7. Form 트리거 연결: Apps Script 편집기 > 트리거 > `onFormSubmit` 추가 (이벤트: 양식 제출 시)

## 알려진 제약 / 확인 필요 사항

- xlsx → Google Sheets 변환 시 조건부 서식 일부가 깨질 수 있음 (특히 FormulaRule 기반 규칙) → 업로드 후 재확인 필요.
- `missing_check` A5 수식이 오류 나면 `refreshMissingCheck()` Apps Script로 대체.
- `monthly_summary`는 현재 매장 18개 기준 4~21행까지만 생성됨. 매장이 50개로 늘어나면
  21행 이후로 수식을 복사(드래그)해서 확장해야 함.
