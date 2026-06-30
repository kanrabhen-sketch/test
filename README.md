# 일매출 관리 시스템 v2

Google Sheets + Apps Script 강화판. v1(legacy/ 폴더)에서 쓰던 LET/HSTACK/MAP 같은
고급 배열 수식을 전부 걷어내고, 집계·중복체크·미제출체크·대시보드를 Apps Script가
값으로 계산해서 채우는 방식으로 다시 짰습니다.

## v1 → v2 핵심 변경

| 항목 | v1 | v2 |
|---|---|---|
| 중복여부(sales_logs I열) | MAP+LAMBDA 수식 | Apps Script가 값으로 계산 |
| monthly_summary 31일 그리드 | 셀마다 LET/COUNTIFS 중첩 수식 | `rebuildMonthlySummary()`가 전부 계산해서 값으로 씀 |
| missing_check | HSTACK+LET+배열 COUNTIFS | `refreshMissingCheck()`가 계산 |
| 한눈에 보기 | 없음 | `dashboard` 탭 신규 — 오늘 현황 KPI + 컨설턴트별 표 |
| 자동화 | 수동 메뉴 실행 | Form 제출 시 자동 트리거 + 매일 23:50 자동 새로고침 |
| Form 생성 | 수동 안내 | `setupAll()` 한 번 실행으로 Form 생성+트리거 설치+초기 계산까지 자동 |

수식을 값으로 바꾼 이유: LET/HSTACK/MAP은 Google Sheets 버전·지역 설정에 따라
지원 여부가 달라 오류가 잦았습니다. 스크립트로 계산하면 항상 같은 결과가 나오고,
Form 제출마다 자동으로 다시 계산되어 "새로고침을 잊어서 숫자가 안 맞는" 문제도 없습니다.

## 구성

```
scripts/generate_template.py   xlsx 템플릿 생성 스크립트 (구조/예시데이터/조건부서식만 담음)
output/일매출관리시스템_템플릿.xlsx   생성된 템플릿
apps_script/Code.gs             전체 자동화 스크립트 (집계/대시보드/Form생성/트리거)
legacy/                         v1 (수식 기반) 백업 — 참고용, 더 이상 사용 안 함
```

## 탭 구성

- **dashboard**(신규) — 기준일, 활성 매장수/오늘 제출/오늘 미제출/오늘 매출합/이번달 매출합/입력률
  KPI 6개, 컨설턴트별 현황 표. `refreshDashboard()`가 채움.
- **stores** — 매장 마스터 (변경 없음)
- **sales_logs** — A~F는 Form 응답, G~J는 `recomputeSalesLogs_()`가 값으로 채움 (수식 없음)
- **monthly_summary** — C1/E1 기준 연월, `rebuildMonthlySummary()`가 전체 계산
- **missing_check** — B1 기준일, `refreshMissingCheck()`가 계산
- **사용예시_가이드** — 맨 앞 탭, 설정 순서와 예시 데이터 설명

## 자동화 흐름

```
Form 제출
  → onFormSubmitHandler 트리거
  → recomputeSalesLogs_() (log_id/업체명/중복여부/확인여부)
  → rebuildMonthlySummary() / refreshMissingCheck() / refreshDashboard()
  → (중복이면) 로그 기록

매일 23:50
  → dailyAutoRefresh 시간 트리거
  → missing_check 기준일을 오늘로 갱신 + 전체 재계산
  → ADMIN_EMAIL이 설정돼 있으면 미제출 매장 요약 이메일 발송
```

## 사용자가 해야 할 작업

1. `output/일매출관리시스템_템플릿.xlsx`를 Google Drive에 업로드 → "Google Sheets로 열기"
2. `사용예시_가이드` 탭에서 안내 확인, `stores` 탭 매장명 오타 확인
3. 확장 프로그램 > Apps Script 편집기 → `apps_script/Code.gs` 전체를 붙여넣고 저장
4. (선택) `ADMIN_EMAIL` 상수에 본사 이메일을 넣으면 매일 미제출 요약을 받음
5. 시트로 돌아와 새로고침 → 메뉴 **일매출관리** 표시 확인
6. **[일매출관리 > 0. 전체 자동 설정]** 실행
   - 처음 실행 시 권한 승인 팝업이 뜸 (정상)
   - 실행 전 `sales_logs`의 예시 데이터(2~5행)는 백업하거나 지울 것
     (Form 연결 시 응답 탭이 새로 생기며 `sales_logs`를 대체하기 때문)
   - 실행 후: 실제 Google Form 생성, `stores!H열`에 진짜 pre-filled 링크 생성,
     `onFormSubmit`/매일 23:50 트리거 설치, dashboard/monthly_summary/missing_check
     초기 계산까지 한 번에 완료됨
7. `stores!H열` 링크를 매장별로 점주에게 카카오톡 공유
8. 이후로는 그냥 두면 됨 — Form 제출마다, 그리고 매일 밤 자동으로 전부 갱신됨

수동으로 다시 계산하고 싶을 때는 메뉴의 1~4번 항목을 사용하면 됩니다.

## 알려진 제약

- `createSalesForm()`/`setupAll()`은 AI가 아니라 동진님 Google 계정에서 실행되어야
  실제 Form이 생성됩니다(이 환경은 구글 로그인을 할 수 없음).
- xlsx → Google Sheets 변환 시 조건부 서식 일부가 깨질 수 있음 → 업로드 후 확인 필요.
- 매장이 50개로 늘어나도 `rebuildMonthlySummary()`는 active 매장 전체를 자동으로
  다시 깔아주므로 행 추가를 수동으로 할 필요 없음 (stores 탭에 매장만 추가하면 됨).
