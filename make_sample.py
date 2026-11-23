# make_sample.py
# 실습용 입력 파일(부서별 매출 엑셀 5개 + 청구서 템플릿)을 만들어 주는 프로그램
# ※ 수업에서 받은 파일이 이미 있으면 이 파일은 실행하지 않아도 됩니다.
import os
from datetime import datetime
import openpyxl as excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SALES_DIR = os.path.join(BASE_DIR, "input", "salesbooks")
TEMPLATE_FILE = os.path.join(BASE_DIR, "input", "invoice-template.xlsx")

# 부서 담당자별 3월 매출 데이터: (일, 거래처, 품목, 수량, 가격)
SALES = {
    "choi": ("최원필", [
        (1, "백현우", "귤", 2, 2500), (2, "안수현", "사과", 1, 1500),
        (3, "도길환", "귤", 3, 2500), (15, "주진호", "귤", 4, 2500),
        (16, "백현우", "사과", 2, 1500), (17, "주진호", "포도", 2, 3000),
        (22, "백현우", "포도", 2, 3000)]),
    "kim": ("김승호", [
        (6, "문정아", "간장", 20, 1500), (11, "강다슬", "소스", 2, 1500),
        (20, "문정아", "된장", 8, 950), (21, "강다슬", "간장", 10, 1500)]),
    "lee": ("이수진", [
        (6, "이문경", "귤 젤리", 20, 1500), (11, "안수현", "복숭아 젤리", 2, 1500),
        (20, "도길환", "귤 젤리", 8, 950), (21, "강다슬", "복숭아 젤리", 10, 1500)]),
    "noh": ("노지훈", [
        (3, "문정아", "두부", 4, 1000), (5, "강다슬", "두부", 1, 1000),
        (5, "정형태", "검은콩", 1, 2000), (7, "문정아", "대두 세트", 2, 2000),
        (15, "강다슬", "팥", 1, 1000), (16, "정형태", "두부", 4, 1000),
        (20, "강다슬", "팥", 5, 1000)]),
    "park": ("박혜리", [
        (3, "이문경", "바나나", 3, 1000), (5, "안수현", "사과", 2, 1500),
        (5, "도길환", "딸기", 2, 2000), (7, "주진호", "바나나", 8, 1000),
        (15, "안수현", "사과", 2, 1500), (16, "주진호", "딸기", 5, 2000),
        (20, "신준영", "망고", 3, 1900)]),
}

thin = Side(style="thin")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
NAVY = PatternFill("solid", fgColor="1F4E79")
WHITE_BOLD = Font(color="FFFFFF", bold=True)
CENTER = Alignment(horizontal="center", vertical="center")


def make_sales_books():
    os.makedirs(SALES_DIR, exist_ok=True)
    for key, (person, rows) in SALES.items():
        book = excel.Workbook()
        sh = book.active
        sh.title = "3월"
        # 1행: 제목
        sh.merge_cells("A1:F1")
        sh["A1"] = "3월 매출 내역"
        sh["A1"].fill = PatternFill("solid", fgColor="70AD47")
        sh["A1"].font = Font(color="FFFFFF", size=16)
        sh["A1"].alignment = CENTER
        # 2행: 담당 / 작성일
        sh["A2"] = "담당"
        sh["B2"] = person
        sh["E2"] = "작성일"
        sh["F2"] = "3월30일"
        # 3행: 머리글
        for col, title in enumerate(["일자", "거래처", "품목", "수량", "가격", "소계"], 1):
            c = sh.cell(3, col, title)
            c.fill = PatternFill("solid", fgColor="4472C4")
            c.font = WHITE_BOLD
            c.alignment = CENTER
        # 4행부터: 데이터
        for i, (day, cust, item, qty, price) in enumerate(rows):
            r = 4 + i
            sh.cell(r, 1, datetime(2020, 3, day)).number_format = 'm"월"d"일"'
            sh.cell(r, 2, cust)
            sh.cell(r, 3, item)
            sh.cell(r, 4, qty)
            sh.cell(r, 5, price).number_format = "#,##0"
            sh.cell(r, 6, f"=D{r}*E{r}").number_format = "#,##0"
        for col, w in zip("ABCDEF", [12, 10, 14, 6, 10, 10]):
            sh.column_dimensions[col].width = w
        book.save(os.path.join(SALES_DIR, f"sales-{key}.xlsx"))


def make_template():
    book = excel.Workbook()
    sh = book.active
    sh.title = "청구서"
    for col, w in zip("ABCDEFG", [2, 18, 30, 4, 7, 10, 14]):
        sh.column_dimensions[col].width = w

    sh.merge_cells("B2:G2")
    sh["B2"] = "청구서"
    sh["B2"].fill = NAVY
    sh["B2"].font = Font(color="FFFFFF", bold=True, size=16)
    sh["B2"].alignment = CENTER
    sh.row_dimensions[2].height = 30

    sh["F3"] = "발행일"
    sh["G3"].number_format = "yyyy/mm/dd"
    sh["B4"].font = Font(bold=True, size=14)
    sh["B4"].border = Border(bottom=thin)
    sh["C4"].border = Border(bottom=thin)
    sh["D4"] = "귀하"
    sh["D4"].font = Font(bold=True)

    sh["G6"] = "주식회사 ●●●●"
    sh["G7"] = "서울시 △△△ ■■빌딩 N층"
    sh["G8"] = "전화 : 02-xxx-xxxx"
    for a in ("G6", "G7", "G8"):
        sh[a].alignment = Alignment(horizontal="right")

    for r, label in ((10, "청구 건명"), (11, "청구 금액")):
        sh.cell(r, 2, label).fill = NAVY
        sh.cell(r, 2).font = WHITE_BOLD
        sh.cell(r, 2).alignment = Alignment(horizontal="right")
        sh.merge_cells(f"C{r}:F{r}")
        sh.cell(r, 3).font = Font(bold=True, size=12)
        sh.cell(r, 3).alignment = CENTER
        for col in range(2, 7):
            sh.cell(r, col).border = BOX
    sh["C11"].number_format = "₩#,##0"

    sh["B13"] = "내역"
    sh["B13"].font = Font(bold=True)
    sh["B13"].alignment = CENTER

    sh.merge_cells("B14:D14")
    for col, title in ((2, "개요"), (5, "수량"), (6, "단가"), (7, "소계")):
        sh.cell(14, col, title)
    for col in range(2, 8):
        sh.cell(14, col).fill = NAVY
        sh.cell(14, col).font = WHITE_BOLD
        sh.cell(14, col).alignment = CENTER

    for r in range(15, 23):            # 내역 8줄 (15~22행)
        sh.merge_cells(f"B{r}:D{r}")
        for col in range(2, 8):
            sh.cell(r, col).border = BOX
        sh.cell(r, 6).number_format = "₩#,##0"
        sh.cell(r, 7).number_format = "₩#,##0"

    sh.merge_cells("B23:F23")
    sh["B23"] = "총계"
    for col in range(2, 7):
        sh.cell(23, col).fill = NAVY
    sh["B23"].font = WHITE_BOLD
    sh["B23"].alignment = CENTER
    sh["G23"] = "=SUM(G15:G22)"
    sh["G23"].number_format = "₩#,##0"
    sh["G23"].border = BOX

    os.makedirs(os.path.dirname(TEMPLATE_FILE), exist_ok=True)
    book.save(TEMPLATE_FILE)


if __name__ == "__main__":
    make_sales_books()
    make_template()
    print("샘플 파일 생성 완료:", os.path.join(BASE_DIR, "input"))
