"""
일매출 관리 시스템 v2 - Google Sheets용 xlsx 템플릿 생성기

v1과의 차이:
- monthly_summary/missing_check의 LET/HSTACK/MAP 같은 고급 배열 수식을 모두 제거하고,
  Apps Script(apps_script/Code.gs)가 값으로 직접 계산해서 채우는 방식으로 바꿈.
  → 환경별 수식 호환성 문제가 없고, Form 제출 시 자동으로 갱신됨(트리거 기반).
- sales_logs의 중복여부(I열)도 MAP/LAMBDA 대신 Apps Script가 값으로 채움.
- "dashboard" 탭을 추가해서 오늘 현황을 한눈에 볼 수 있게 함(총매장/제출/미제출/매출/컨설턴트별).
- 모든 계산 결과는 "값"이며, 조건부 서식 규칙(휴무/미입력/중복확인 등 텍스트 매칭)은 동일하게 유지.

이 스크립트가 만드는 xlsx는 구조와 헤더, 예시 데이터, 조건부 서식만 담고 있고
실제 계산은 Google Sheets에 연결된 Apps Script가 수행한다.
"""
import datetime as dt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter

STORES = [
    ("S001", "김한성", "굴지막 영천본점", "", "", "경북", "active", ""),
    ("S002", "김한성", "백동소풍가는길", "", "", "경기", "active", ""),
    ("S003", "김한성", "원시민족갈비 부평구청점", "", "", "인천", "active", ""),
    ("S004", "김한성", "미미냉삼(신증동)", "", "", "경기", "active", ""),
    ("S005", "김왕걸", "참숯삼인조 대구총로점", "", "", "대구", "active", ""),
    ("S006", "김왕걸", "고솜 고양덕은본점", "", "", "경기", "active", ""),
    ("S007", "김왕걸", "다용차반", "", "", "경기", "active", ""),
    ("S008", "김왕걸", "든볼닉 대전톤산동직영점", "", "", "대전", "active", ""),
    ("S009", "김왕걸", "화로갈비", "", "", "경기", "active", ""),
    ("S010", "김왕걸", "양치기소녀 신설동점", "", "", "서울", "active", ""),
    ("S011", "김왕걸", "일산황금설렁탕", "", "", "경기", "active", ""),
    ("S012", "김왕걸", "흥능축발 왕십리", "", "", "서울", "active", ""),
    ("S013", "김왕걸", "태갈공명", "", "", "경기", "active", ""),
    ("S014", "박석현", "스마일부대찌개", "", "", "경기", "active", ""),
    ("S015", "박석현", "만세갈비", "", "", "경기", "active", ""),
    ("S016", "박석현", "부치공", "", "", "경기", "active", ""),
    ("S017", "박석현", "소랑", "", "", "경기", "active", ""),
    ("S018", "송다영", "한계령코다리찜동태찌개", "", "", "강원", "active", ""),
]
NUM_STORES = len(STORES)
SUMMARY_START = 4
SUMMARY_END = SUMMARY_START + NUM_STORES - 1

HEADER_FILL = PatternFill("solid", fgColor="2F5496")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SETTING_FILL = PatternFill("solid", fgColor="FFF2CC")
KPI_FILL = PatternFill("solid", fgColor="E8F0FE")
THIN = Side(style="thin", color="CCCCCC")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = openpyxl.Workbook()
wb.remove(wb.active)


def style_header_row(ws, row, max_col):
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")


# ---------------------------------------------------------------------------
# 0. dashboard — 한눈에 보는 오늘 현황 (전부 Apps Script refreshDashboard()가 채움)
# ---------------------------------------------------------------------------
ws = wb.create_sheet("dashboard")
ws["A1"] = "일매출 대시보드"
ws["A1"].font = Font(bold=True, size=16)

ws["A3"] = "기준일"
ws["B3"] = dt.date(2026, 6, 1)
ws["B3"].number_format = "yyyy-mm-dd"
ws["B3"].fill = SETTING_FILL
ws["A3"].font = Font(bold=True)
ws["C3"] = "(refreshDashboard 실행 시 missing_check!B1과 같은 날짜로 자동 동기화됨)"
ws["C3"].font = Font(italic=True, size=9, color="888888")

kpi_labels = ["활성 매장 수", "오늘 제출", "오늘 미제출", "오늘 매출합", "이번달 매출합(누적)", "오늘 입력률"]
for i, label in enumerate(kpi_labels):
    col = 1 + i
    ws.cell(row=5, column=col, value=label).font = Font(bold=True)
    ws.cell(row=5, column=col).fill = HEADER_FILL
    ws.cell(row=5, column=col).font = Font(bold=True, color="FFFFFF")
    ws.cell(row=5, column=col).alignment = Alignment(horizontal="center")
    v = ws.cell(row=6, column=col)
    v.fill = KPI_FILL
    v.border = BOX
    v.alignment = Alignment(horizontal="center")
    v.font = Font(size=14, bold=True)

ws.cell(row=8, column=1, value="컨설턴트별 현황").font = Font(bold=True, size=12)
headers = ["담당컨설턴트", "담당 매장수", "오늘 제출", "오늘 미제출", "이번달 매출합"]
for i, h in enumerate(headers, start=1):
    ws.cell(row=9, column=i, value=h)
style_header_row(ws, 9, len(headers))

widths = [16, 14, 12, 12, 18, 40]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A10"

# ---------------------------------------------------------------------------
# 1. stores
# ---------------------------------------------------------------------------
ws = wb.create_sheet("stores")
headers = ["store_id", "담당컨설턴트", "업체명", "점주명", "연락처", "지역", "상태", "입력링크"]
ws.append(headers)
style_header_row(ws, 1, len(headers))
for row in STORES:
    ws.append(list(row))

dv_status = DataValidation(type="list", formula1='"active,inactive"', allow_blank=False)
ws.add_data_validation(dv_status)
dv_status.add("G2:G200")

for i, row in enumerate(STORES, start=2):
    ws.cell(row=i, column=8, value=(
        f"(예시-형식만, createSalesForm() 실행 후 자동 교체) "
        f"https://docs.google.com/forms/d/e/FORM_ID/viewform?usp=pp_url&entry.123456789={row[0]}"
    ))

widths = [10, 14, 30, 12, 14, 8, 10, 55]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A2"

# ---------------------------------------------------------------------------
# 2. sales_logs
#   G(log_id)/H(업체명)는 그대로 둬도 안전한 단순 수식이라 유지.
#   I(중복여부)는 Apps Script가 값으로 기록(수식 제거) → 환경 호환성 문제 차단.
# ---------------------------------------------------------------------------
ws = wb.create_sheet("sales_logs")
headers = ["제출시간", "store_id", "영업일", "영업여부", "일매출", "메모",
           "log_id", "업체명", "중복여부", "확인여부", "수정메모"]
ws.append(headers)
style_header_row(ws, 1, len(headers))

TEST_ROWS = [
    (dt.datetime(2026, 6, 1, 22, 0), "S001", dt.date(2026, 6, 1), "영업", 900000, "테스트1: 정상 입력", "L0001", "굴지막 영천본점", False, False, ""),
    (dt.datetime(2026, 6, 2, 22, 0), "S001", dt.date(2026, 6, 2), "휴무", 0, "테스트2: 휴무", "L0002", "굴지막 영천본점", False, False, ""),
    (dt.datetime(2026, 6, 3, 22, 0), "S001", dt.date(2026, 6, 1), "영업", 950000, "테스트3: 중복 입력(6/1 재입력)", "L0003", "굴지막 영천본점", True, False, ""),
    (dt.datetime(2026, 6, 1, 21, 30), "S003", dt.date(2026, 6, 1), "영업", 650000, "예시: 다른 매장 정상 입력", "L0004", "원시민족갈비 부평구청점", False, False, ""),
]
for r_idx, row in enumerate(TEST_ROWS, start=2):
    for c_idx, val in enumerate(row, start=1):
        ws.cell(row=r_idx, column=c_idx, value=val)
    ws.cell(row=r_idx, column=1).number_format = "yyyy-mm-dd hh:mm"
    ws.cell(row=r_idx, column=3).number_format = "yyyy-mm-dd"

dv_status2 = DataValidation(type="list", formula1='"영업,휴무"', allow_blank=False)
ws.add_data_validation(dv_status2)
dv_status2.add("D2:D1000")

dv_bool = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=False)
ws.add_data_validation(dv_bool)
dv_bool.add("J2:J1000")

ws.conditional_formatting.add(
    "A2:K1000",
    FormulaRule(formula=['AND($D2="영업",$E2=0)'], font=Font(color="CC0000")),
)
ws.conditional_formatting.add(
    "I2:I1000",
    CellIsRule(operator="equal", formula=["TRUE"], fill=PatternFill("solid", fgColor="FFCCCC")),
)

widths = [18, 10, 12, 10, 12, 20, 10, 30, 10, 10, 24]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A2"

# ---------------------------------------------------------------------------
# 3. monthly_summary — 헤더/포맷만 유지, 모든 값은 rebuildMonthlySummary()가 채움
# ---------------------------------------------------------------------------
ws = wb.create_sheet("monthly_summary")
ws["B1"] = "기준연도"
ws["C1"] = 2026
ws["D1"] = "기준월"
ws["E1"] = 6
ws["B1"].font = Font(bold=True)
ws["D1"].font = Font(bold=True)
ws["C1"].fill = SETTING_FILL
ws["E1"].fill = SETTING_FILL
ws["G1"] = "(C1/E1 변경 후 메뉴 > monthly_summary 새로고침 실행)"
ws["G1"].font = Font(italic=True, size=9, color="888888")

headers = ["store_id", "담당컨설턴트", "업체명", "입력률", "미입력일수", "휴무일수",
           "영업일수", "월매출", "일평균", "전월대비"]
for i, h in enumerate(headers, start=1):
    ws.cell(row=3, column=i, value=h)
style_header_row(ws, 3, len(headers))

for day in range(1, 32):
    col = 11 + day - 1
    cell = ws.cell(row=3, column=col, value=day)
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = Alignment(horizontal="center")

# 예시 미리보기 값(스크립트 실행 전 참고용) — S001 1행만 채워서 형태를 보여줌
ws.cell(row=4, column=1, value="S001")
ws.cell(row=4, column=2, value="김한성")
ws.cell(row=4, column=3, value="굴지막 영천본점")
ws.cell(row=4, column=4, value="(자동계산)")
ws.cell(row=4, column=5, value="(자동계산)")
ws.cell(row=4, column=6, value=1)
ws.cell(row=4, column=7, value=1)
ws.cell(row=4, column=8, value=900000)
ws.cell(row=4, column=9, value=900000)
ws.cell(row=4, column=10, value="데이터부족")
ws.cell(row=4, column=11, value="중복확인")
ws.cell(row=4, column=12, value="휴무")
for day in range(3, 31):
    ws.cell(row=4, column=10 + day, value="미입력")
ws.cell(row=4, column=10 + 31, value="---")

day_range = f"K4:AO{SUMMARY_END}"
ws.conditional_formatting.add(day_range, CellIsRule(operator="equal", formula=['"휴무"'],
                               fill=PatternFill("solid", fgColor="CCCCCC")))
ws.conditional_formatting.add(day_range, CellIsRule(operator="equal", formula=['"미입력"'],
                               fill=PatternFill("solid", fgColor="FFF3CD")))
ws.conditional_formatting.add(day_range, CellIsRule(operator="equal", formula=['"중복확인"'],
                               fill=PatternFill("solid", fgColor="FF0000"), font=Font(color="FFFFFF")))
ws.conditional_formatting.add(day_range, CellIsRule(operator="equal", formula=['"---"'],
                               fill=PatternFill("solid", fgColor="AAAAAA")))

j_range = f"J4:J{SUMMARY_END}"
ws.conditional_formatting.add(j_range, CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="CC0000")))
ws.conditional_formatting.add(j_range, CellIsRule(operator="lessThan", formula=["0"], font=Font(color="0000CC")))
ws.conditional_formatting.add("K3:AO3", FormulaRule(formula=["K$3=DAY(TODAY())"],
                              fill=PatternFill("solid", fgColor="E3F2FD")))

ws.column_dimensions["A"].hidden = True
widths = {"A": 10, "B": 14, "C": 30, "D": 10, "E": 12, "F": 10, "G": 10, "H": 14, "I": 12, "J": 12}
for col, w in widths.items():
    ws.column_dimensions[col].width = w
for day in range(1, 32):
    ws.column_dimensions[get_column_letter(11 + day - 1)].width = 9
ws.freeze_panes = "K4"

# ---------------------------------------------------------------------------
# 4. missing_check — 헤더만 유지, A5 이하 값은 refreshMissingCheck()가 채움
# ---------------------------------------------------------------------------
ws = wb.create_sheet("missing_check")
ws["A1"] = "확인일"
ws["B1"] = dt.date(2026, 6, 1)
ws["A1"].font = Font(bold=True)
ws["B1"].fill = SETTING_FILL
ws["B1"].number_format = "yyyy-mm-dd"

ws["A2"] = "오늘 제출 수"
ws["B2"] = 2
ws["C2"] = "미제출 수"
ws["D2"] = 16
ws["E2"] = "오늘 매출합"
ws["F2"] = 1550000
for c in ["A2", "C2", "E2"]:
    ws[c].font = Font(bold=True)
ws["G2"] = "(값은 refreshMissingCheck() 실행 시 자동 갱신됨)"
ws["G2"].font = Font(italic=True, size=9, color="888888")

headers = ["담당컨설턴트", "업체명", "연락처", "미입력일", "입력링크", "상태"]
for i, h in enumerate(headers, start=1):
    ws.cell(row=4, column=i, value=h)
style_header_row(ws, 4, len(headers))

widths = [14, 30, 14, 12, 40, 10]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A5"

# ---------------------------------------------------------------------------
# 5. 사용예시_가이드
# ---------------------------------------------------------------------------
ws = wb.create_sheet("사용예시_가이드", 0)
ws.column_dimensions["A"].width = 100
guide_lines = [
    "■ v2 변경점: 모든 집계(monthly_summary/missing_check)와 중복체크(sales_logs I열)는",
    "  더 이상 LET/HSTACK/MAP 같은 고급 수식이 아니라 Apps Script가 계산해서 값으로 채웁니다.",
    "  → 수식 호환성 문제가 없고, Form 제출 시 트리거가 자동으로 전부 다시 계산합니다.",
    "",
    "■ 이 xlsx 파일 자체는 '구조 + 예시 데이터 + 조건부 서식'만 담고 있습니다.",
    "  실제 계산은 Google Sheets로 연 다음 Apps Script(apps_script/Code.gs)를 붙여넣고",
    "  메뉴 [일매출관리 > 0. 전체 자동 설정]을 한 번 실행해야 시작됩니다.",
    "",
    "■ sales_logs에 예시 데이터 4줄이 들어가 있습니다 (S001 정상/휴무/중복, S003 정상).",
    "  monthly_summary 4행(S001)에는 '이렇게 보여야 한다'는 미리보기 값을 손으로 채워뒀습니다.",
    "  실제로는 메뉴 [1. monthly_summary 새로고침]을 누르면 18개 매장 전체가 이 형태로 계산됩니다.",
    "",
    "■ 자동화 흐름",
    "  Form 제출 → onFormSubmit 트리거 → 중복여부/확인여부 기록",
    "             → monthly_summary, missing_check, dashboard 자동 재계산",
    "  매일 23:50 → dailyAutoRefresh 시간 트리거 → 미제출 매장 자동 갱신 + 관리자 이메일 발송(선택)",
    "",
    "■ 처음 설정 순서",
    "  1) Google Drive 업로드 → Google Sheets로 열기",
    "  2) 확장 프로그램 > Apps Script 편집기 → Code.gs 내용 전체 붙여넣기 → 저장",
    "  3) 시트로 돌아와 새로고침 → 메뉴 [일매출관리] 표시 확인",
    "  4) [일매출관리 > 0. 전체 자동 설정] 실행 (Form 생성 + 트리거 설치 한 번에 처리)",
    "     - 실행 전 sales_logs 예시 데이터(2~5행)는 지우거나 백업할 것",
    "  5) stores!H열에 생성된 실제 pre-filled 링크를 매장별로 점주에게 공유",
]
for i, line in enumerate(guide_lines, start=1):
    ws.cell(row=i, column=1, value=line)
ws["A1"].font = Font(bold=True, size=13)
ws.freeze_panes = "A2"

# ---------------------------------------------------------------------------
import sys
out_path = sys.argv[1] if len(sys.argv) > 1 else "output/일매출관리시스템_템플릿.xlsx"
wb.save(out_path)
print(f"Saved: {out_path}")
