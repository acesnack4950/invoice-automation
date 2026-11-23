# step3_invoice.py  ③ 집계한 내용을 템플릿에 넣어 고객별 청구서 발행하기

import os
import json
from datetime import datetime
import openpyxl as excel

# 1. 입력 파일 지정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, "input", "invoice-template.xlsx")
JSON_FILE = os.path.join(BASE_DIR, "output", "customers.json")

# 2. 청구 월 및 저장 폴더 지정
SUBJECT = "3월 청구분"   # 샘플 매출 데이터가 3월분이라 3월로 맞춤
SAVE_DIR = os.path.join(BASE_DIR, "output", "invoice")

# 3. 발행일자 지정
ISSUE_DATE = datetime.now()

# 템플릿의 내역 칸: 15행 ~ 22행 (8줄)
START_ROW = 15
MAX_ROWS = 8


# 4. 메인 처리 함수 정의
def fill_template():
    with open(JSON_FILE, "rt", encoding="utf-8") as fp:
        customers = json.load(fp)

    os.makedirs(SAVE_DIR, exist_ok=True)
    for name, data in customers.items():
        make_invoice(name, data)
    print("③ 청구서 발행 완료 →", SAVE_DIR)


# 5. 템플릿에 데이터 채우는 함수 정의
def make_invoice(name, data):
    # 고객마다 템플릿을 새로 열어야 이전 고객 내용이 섞이지 않음
    book = excel.load_workbook(TEMPLATE_FILE)
    sheet = book.active

    # - 템플릿에 데이터 입력
    sheet["B4"] = name
    sheet["C10"] = SUBJECT
    sheet["C11"] = data["total"]
    sheet["G3"] = ISSUE_DATE
    sheet["G3"].number_format = "yyyy/mm/dd"

    items = data["items"]
    if len(items) > MAX_ROWS:
        print(f"  ⚠ {name}: 내역이 {len(items)}건이라 앞의 {MAX_ROWS}건만 들어갑니다.")

    for i, (date, item, count, price) in enumerate(items[:MAX_ROWS]):
        row = START_ROW + i
        sheet.cell(row, 2, f"{item}({date})")   # B열: 개요  예) 소스(03/11)
        sheet.cell(row, 5, count)               # E열: 수량
        sheet.cell(row, 6, price)               # F열: 단가
        sheet.cell(row, 7, count * price)       # G열: 소계

    # - 고객 이름별로 파일 저장
    save_file = os.path.join(SAVE_DIR, f"{name}.xlsx")
    book.save(save_file)
    print("  저장:", os.path.basename(save_file), f"(₩{data['total']:,})")


# 6. 메인 프로그램 실행하기
if __name__ == "__main__":
    fill_template()
