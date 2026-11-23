# step2_aggregate.py  ② 모은 데이터를 고객별로 나누고 청구 금액 집계하기

import os
import json
import openpyxl as excel

# 1. 입출력 파일 지정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MERGE_FILE = os.path.join(BASE_DIR, "output", "merge_file.xlsx")   # ①의 결과
JSON_FILE = os.path.join(BASE_DIR, "output", "customers.json")     # ②의 결과


# 2. 메인 처리 함수 정의
def split_data():
    sheet = excel.load_workbook(MERGE_FILE, data_only=True).active
    customers = split_sheet(sheet)
    calc_total(customers)

    # 딕셔너리를 JSON 파일로 저장
    with open(JSON_FILE, "wt", encoding="utf-8") as f:
        json.dump(customers, f, ensure_ascii=False, indent=4)  # 한글 그대로, 들여쓰기 4칸
    print("② 집계 완료 →", JSON_FILE)
    return customers


# 3. 고객명을 키로 하는 딕셔너리 반환 함수 정의
#    결과 모양: {"백현우": {"items": [["03/01", "귤", 2, 2500], ...]}, ...}
def split_sheet(sheet):
    customers = {}
    for row in sheet.iter_rows(min_row=1, values_only=True):
        date, name, item, count, price = row[:5]
        if name is None:
            continue
        if name not in customers:          # 처음 나온 고객이면 빈 목록 만들기
            customers[name] = {"items": []}
        date_str = date.strftime("%m/%d")  # datetime → "03/01" (JSON은 날짜형을 저장 못 함)
        customers[name]["items"].append([date_str, item, count, price])
    return customers


# 4. 고객별 집계 결과 반환 함수 정의
def calc_total(customers):
    for name, data in customers.items():
        total = 0
        for date, item, count, price in data["items"]:
            total += count * price
        data["total"] = total
    return customers


# 5. 메인 프로그램 실행하기
if __name__ == "__main__":
    result = split_data()
    print(result)
