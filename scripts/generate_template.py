"""
일매출 관리 시스템 - Google Sheets용 xlsx 템플릿 생성기
이관 문서(업무 이관 문서 — 일매출 관리 시스템) 기준으로 구현.

주의: MAP / LET / HSTACK / FILTER / COUNTUNIQUE / ARRAYFORMULA 함수는
Google Sheets 전용이며 Excel에서는 작동하지 않는다. 이 스크립트는
xlsx 파일에 "문자열로서" 해당 수식을 기록하며, Google Drive에 업로드 후
Google Sheets로 열어야 정상 작동한다.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
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
SUMMARY_DATA_START_ROW = 4
SUMMARY_DATA_END_ROW = SUMMARY_DATA_START_ROW + NUM_STORES - 1  # 21

HEADER_FILL = PatternFill("solid", fgColor="2F5496")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SETTING_FILL = PatternFill("solid", fgColor="FFF2CC")
HIDDEN_COL_FILL = PatternFill("solid", fgColor="EEEEEE")

wb = openpyxl.Workbook()
wb.remove(wb.active)


def style_header_row(ws, row, max_col):
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")


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
dv_status.add(f"G2:G200")

widths = [10, 14, 30, 12, 14, 8, 10, 40]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A2"

# ---------------------------------------------------------------------------
# 2. sales_logs
# ---------------------------------------------------------------------------
ws = wb.create_sheet("sales_logs")
headers = ["제출시간", "store_id", "영업일", "영업여부", "일매출", "메모",
           "log_id", "업체명", "중복여부", "확인여부", "수정메모"]
ws.append(headers)
style_header_row(ws, 1, len(headers))

ws["G2"] = '=ARRAYFORMULA(IF(A2:A="","","L"&TEXT(ROW(A2:A)-1,"0000")))'
ws["H2"] = '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,stores!$A:$C,3,0),"store_id 오류")))'
ws["I2"] = ('=MAP(B2:B1000,C2:C1000,ROW(B2:B1000),'
            'LAMBDA(store,date,rownum,'
            'IF(OR(store="",date=""),FALSE,'
            'COUNTIFS($B$2:INDEX($B:$B,rownum),store,'
            '$C$2:INDEX($C:$C,rownum),date)>1)))')
ws["J2"] = False

dv_status2 = DataValidation(type="list", formula1='"영업,휴무"', allow_blank=False)
ws.add_data_validation(dv_status2)
dv_status2.add("D2:D1000")

dv_bool = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=False)
ws.add_data_validation(dv_bool)
dv_bool.add("J2:J1000")

# 영업인데 0원 경고
ws.conditional_formatting.add(
    "A2:K1000",
    FormulaRule(formula=['AND($D2="영업",$E2=0)'], font=Font(color="CC0000")),
)
# 중복 행 강조
ws.conditional_formatting.add(
    "I2:I1000",
    CellIsRule(operator="equal", formula=["TRUE"], fill=PatternFill("solid", fgColor="FFCCCC")),
)

widths = [18, 10, 12, 10, 12, 20, 10, 30, 10, 10, 24]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A2"

# ---------------------------------------------------------------------------
# 3. monthly_summary
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

headers = ["store_id", "담당컨설턴트", "업체명", "입력률", "미입력일수", "휴무일수",
           "영업일수", "월매출", "일평균", "전월대비"]
for i, h in enumerate(headers, start=1):
    ws.cell(row=3, column=i, value=h)
style_header_row(ws, 3, len(headers))

# K~AO = 1~31일
for day in range(1, 32):
    col = 11 + day - 1  # K=11
    cell = ws.cell(row=3, column=col, value=day)
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = Alignment(horizontal="center")

for r in range(SUMMARY_DATA_START_ROW, SUMMARY_DATA_END_ROW + 1):
    n = r - SUMMARY_DATA_START_ROW + 1  # 1-based index among active stores
    ws.cell(row=r, column=1, value=(
        f'=IFERROR(INDEX(FILTER(stores!$A$2:$A,stores!$G$2:$G="active"),ROW()-3),"")'
    ))
    ws.cell(row=r, column=2, value=f'=IF($A{r}="","",VLOOKUP($A{r},stores!$A:$B,2,0))')
    ws.cell(row=r, column=3, value=f'=IF($A{r}="","",VLOOKUP($A{r},stores!$A:$C,3,0))')

    ws.cell(row=r, column=4, value=(
        f'=IF($A{r}="","",'
        f'IF(DATE($C$1,$E$1,1)>TODAY(),0,'
        f'IFERROR('
        f'COUNTUNIQUE(FILTER(sales_logs!$C$2:$C$1000,'
        f'sales_logs!$B$2:$B$1000=$A{r},'
        f'sales_logs!$C$2:$C$1000>=DATE($C$1,$E$1,1),'
        f'sales_logs!$C$2:$C$1000<=MIN(TODAY(),EOMONTH(DATE($C$1,$E$1,1),0)),'
        f'sales_logs!$I$2:$I$1000=FALSE))'
        f'/DAY(MIN(TODAY(),EOMONTH(DATE($C$1,$E$1,1),0))),0)))'
    ))

    ws.cell(row=r, column=5, value=(
        f'=IF($A{r}="","",'
        f'IF(DATE($C$1,$E$1,1)>TODAY(),0,'
        f'DAY(MIN(TODAY(),EOMONTH(DATE($C$1,$E$1,1),0)))'
        f'-IFERROR(COUNTUNIQUE(FILTER(sales_logs!$C$2:$C$1000,'
        f'sales_logs!$B$2:$B$1000=$A{r},'
        f'sales_logs!$C$2:$C$1000>=DATE($C$1,$E$1,1),'
        f'sales_logs!$C$2:$C$1000<=MIN(TODAY(),EOMONTH(DATE($C$1,$E$1,1),0)),'
        f'sales_logs!$I$2:$I$1000=FALSE)),0)))'
    ))

    ws.cell(row=r, column=6, value=(
        f'=IF($A{r}="","",COUNTIFS('
        f'sales_logs!$B:$B,$A{r},'
        f'sales_logs!$C:$C,">="&DATE($C$1,$E$1,1),'
        f'sales_logs!$C:$C,"<="&EOMONTH(DATE($C$1,$E$1,1),0),'
        f'sales_logs!$D:$D,"휴무",'
        f'sales_logs!$I:$I,FALSE))'
    ))

    ws.cell(row=r, column=7, value=(
        f'=IF($A{r}="","",COUNTIFS('
        f'sales_logs!$B:$B,$A{r},'
        f'sales_logs!$C:$C,">="&DATE($C$1,$E$1,1),'
        f'sales_logs!$C:$C,"<="&EOMONTH(DATE($C$1,$E$1,1),0),'
        f'sales_logs!$D:$D,"영업",'
        f'sales_logs!$I:$I,FALSE))'
    ))

    ws.cell(row=r, column=8, value=(
        f'=IF($A{r}="","",SUMIFS('
        f'sales_logs!$E:$E,'
        f'sales_logs!$B:$B,$A{r},'
        f'sales_logs!$C:$C,">="&DATE($C$1,$E$1,1),'
        f'sales_logs!$C:$C,"<="&EOMONTH(DATE($C$1,$E$1,1),0),'
        f'sales_logs!$D:$D,"영업",'
        f'sales_logs!$I:$I,FALSE))'
    ))

    ws.cell(row=r, column=9, value=f'=IF(OR($A{r}="",G{r}=0),"",H{r}/G{r})')

    ws.cell(row=r, column=10, value=(
        f'=IF($A{r}="","",'
        f'LET('
        f'cur_end, MIN(TODAY(),EOMONTH(DATE($C$1,$E$1,1),0)),'
        f'cmp_day, DAY(cur_end),'
        f'prev_start, DATE($C$1,$E$1-1,1),'
        f'prev_end, MIN(DATE($C$1,$E$1-1,cmp_day),EOMONTH(DATE($C$1,$E$1-1,1),0)),'
        f'cur_sales, SUMIFS(sales_logs!$E:$E,'
        f'sales_logs!$B:$B,$A{r},'
        f'sales_logs!$C:$C,">="&DATE($C$1,$E$1,1),'
        f'sales_logs!$C:$C,"<="&cur_end,'
        f'sales_logs!$D:$D,"영업",'
        f'sales_logs!$I:$I,FALSE),'
        f'prev_sales, SUMIFS(sales_logs!$E:$E,'
        f'sales_logs!$B:$B,$A{r},'
        f'sales_logs!$C:$C,">="&prev_start,'
        f'sales_logs!$C:$C,"<="&prev_end,'
        f'sales_logs!$D:$D,"영업",'
        f'sales_logs!$I:$I,FALSE),'
        f'IF(prev_sales=0,"데이터부족",(cur_sales-prev_sales)/prev_sales)))'
    ))

    for day in range(1, 32):
        col = 11 + day - 1
        col_letter = get_column_letter(col)
        ws.cell(row=r, column=col, value=(
            f'=IF($A{r}="","",'
            f'IF(MONTH(DATE($C$1,$E$1,{col_letter}$3))<>$E$1,"---",'
            f'IF(DATE($C$1,$E$1,{col_letter}$3)>TODAY(),"",'
            f'IF(COUNTIFS(sales_logs!$B:$B,$A{r},'
            f'sales_logs!$C:$C,DATE($C$1,$E$1,{col_letter}$3),'
            f'sales_logs!$I:$I,TRUE)>0,"중복확인",'
            f'IF(COUNTIFS(sales_logs!$B:$B,$A{r},'
            f'sales_logs!$C:$C,DATE($C$1,$E$1,{col_letter}$3),'
            f'sales_logs!$D:$D,"휴무",'
            f'sales_logs!$I:$I,FALSE)>0,"휴무",'
            f'IF(COUNTIFS(sales_logs!$B:$B,$A{r},'
            f'sales_logs!$C:$C,DATE($C$1,$E$1,{col_letter}$3),'
            f'sales_logs!$I:$I,FALSE)=0,"미입력",'
            f'SUMIFS(sales_logs!$E:$E,'
            f'sales_logs!$B:$B,$A{r},'
            f'sales_logs!$C:$C,DATE($C$1,$E$1,{col_letter}$3),'
            f'sales_logs!$D:$D,"영업",'
            f'sales_logs!$I:$I,FALSE)))))))'
        ))

day_range = f"K4:AO{SUMMARY_DATA_END_ROW}"
ws.conditional_formatting.add(day_range, CellIsRule(operator="equal", formula=['"휴무"'],
                               fill=PatternFill("solid", fgColor="CCCCCC")))
ws.conditional_formatting.add(day_range, CellIsRule(operator="equal", formula=['"미입력"'],
                               fill=PatternFill("solid", fgColor="FFF3CD")))
ws.conditional_formatting.add(day_range, CellIsRule(operator="equal", formula=['"중복확인"'],
                               fill=PatternFill("solid", fgColor="FF0000"), font=Font(color="FFFFFF")))
ws.conditional_formatting.add(day_range, CellIsRule(operator="equal", formula=['"---"'],
                               fill=PatternFill("solid", fgColor="AAAAAA")))

j_range = f"J4:J{SUMMARY_DATA_END_ROW}"
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
# 4. missing_check
# ---------------------------------------------------------------------------
ws = wb.create_sheet("missing_check")

ws["A1"] = "확인일"
ws["B1"] = "=TODAY()"
ws["A1"].font = Font(bold=True)
ws["B1"].fill = SETTING_FILL
ws["B1"].number_format = "yyyy-mm-dd"

ws["A2"] = "오늘 제출 수"
ws["B2"] = "=COUNTIF(sales_logs!$C$2:$C$1000,$B$1)"
ws["C2"] = "미제출 수"
ws["D2"] = '=COUNTIF(stores!$G$2:$G$200,"active")-COUNTIF(sales_logs!$C$2:$C$1000,$B$1)'
ws["E2"] = "오늘 매출합"
ws["F2"] = '=SUMIFS(sales_logs!$E$2:$E$1000,sales_logs!$C$2:$C$1000,$B$1,sales_logs!$D$2:$D$1000,"영업")'
for c in ["A2", "C2", "E2"]:
    ws[c].font = Font(bold=True)

headers = ["담당컨설턴트", "업체명", "연락처", "미입력일", "입력링크", "상태"]
for i, h in enumerate(headers, start=1):
    ws.cell(row=4, column=i, value=h)
style_header_row(ws, 4, len(headers))

ws["A5"] = (
    '=IFERROR('
    'LET('
    'cond,(stores!$G$2:$G$200="active")*'
    '(COUNTIFS(sales_logs!$B$2:$B$1000,stores!$A$2:$A$200,'
    'sales_logs!$C$2:$C$1000,$B$1)=0),'
    'HSTACK('
    'FILTER(stores!$B$2:$B$200,cond),'
    'FILTER(stores!$C$2:$C$200,cond),'
    'FILTER(stores!$E$2:$E$200,cond),'
    'FILTER(IF(stores!$A$2:$A$200<>"",$B$1,""),cond),'
    'FILTER(stores!$H$2:$H$200,cond),'
    'FILTER(stores!$G$2:$G$200,cond)'
    ')'
    '),'
    '"오늘 미제출 매장 없음")'
)

widths = [14, 30, 14, 12, 40, 10]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A5"

# ---------------------------------------------------------------------------
import sys
out_path = sys.argv[1] if len(sys.argv) > 1 else "output/일매출관리시스템_템플릿.xlsx"
wb.save(out_path)
print(f"Saved: {out_path}")
