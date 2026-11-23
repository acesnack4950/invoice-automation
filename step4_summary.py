# step4_summary.py  ④ 고객별 청구 금액 요약표 만들기
#   ※ 수업 실습(①~③)에 직접 추가한 기능
#   청구서는 고객마다 파일이 따로라서, 이번 달 전체 청구 현황을 한눈에 볼 수 있게 한 장으로 정리한다.

import os
import json
import openpyxl as excel
from openpyxl.styles import Font, PatternFill

# 1. 입출력 파일 지정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, "output", "customers.json")   # ②의 결과
SAVE_FILE = os.path.join(BASE_DIR, "output", "summary.xlsx")     # ④의 결과

# 청구 금액이 이 금액 이상인 고객은 노란색으로 표시 (엑셀자동화3 조건부 서식 응용)
HIGHLIGHT_LIMIT = 30000
yellow = PatternFill(fill_type="solid", fgColor="ffff00")


# 2. 메인 처리 함수 정의
def make_summary():
    # ②에서 만든 고객별 집계 결과 읽기
    with open(JSON_FILE, "rt", encoding="utf-8") as fp:
        customers = json.load(fp)

    # 새 엑셀 파일 만들고 머리글 넣기
    book = excel.Workbook()
    sheet = book.active
    sheet.title = "요약"
    sheet.append(["고객명", "거래 건수", "청구 금액"])

    # 청구 금액이 큰 고객부터 한 줄씩 넣기
    #   sorted(..., key=금액, reverse=True) : 금액 기준 내림차순 정렬
    grand_total = 0
    for name, data in sorted(customers.items(), key=lambda c: c[1]["total"], reverse=True):
        sheet.append([name, len(data["items"]), data["total"]])
        grand_total += data["total"]

        # 방금 넣은 줄(맨 아래 줄)의 금액이 기준 이상이면 그 줄 전체를 노란색으로
        if data["total"] >= HIGHLIGHT_LIMIT:
            for cell in sheet[sheet.max_row]:
                cell.fill = yellow

    # 맨 아래 합계 줄
    sheet.append(["합계", len(customers), grand_total])

    # 보기 좋게 꾸미기: 머리글·합계 굵게, 금액에 쉼표, 열 너비
    for cell in sheet[1] + sheet[sheet.max_row]:
        cell.font = Font(bold=True)
    for row in sheet.iter_rows(min_row=2, min_col=3, max_col=3):
        row[0].number_format = "#,##0"
    sheet.column_dimensions["A"].width = 12
    sheet.column_dimensions["B"].width = 10
    sheet.column_dimensions["C"].width = 14

    book.save(SAVE_FILE)
    print(f"④ 요약표 완료 → {SAVE_FILE}  (고객 {len(customers)}명, 합계 ₩{grand_total:,})")


# 3. 메인 프로그램 실행하기
if __name__ == "__main__":
    make_summary()
