# step1_merge.py  ① 담당자별 매출 엑셀을 하나의 시트에 모으기

# 1. 필요한 모듈 불러오기
import os
import glob
import openpyxl as excel

# 2. 대상 폴더명과 저장할 파일명 지정
#    (이 .py 파일이 있는 폴더를 기준으로 경로를 만들어서, 어디서 실행해도 잘 동작하게 함)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_PATTERN = os.path.join(BASE_DIR, "input", "salesbooks", "*.xlsx")
SAVE_FILE = os.path.join(BASE_DIR, "output", "merge_file.xlsx")


# 3. 메인 처리 함수 정의
def merge_files():
    # 결과를 담을 새 엑셀 파일(빈 시트 1개) 만들기
    book = excel.Workbook()
    merge_sheet = book.active
    merge_sheet.title = "통합"

    # 폴더 안의 엑셀 파일을 하나씩 읽어서 merge_sheet에 붙이기
    for fname in enum_files():
        read_book(merge_sheet, fname)

    # output 폴더가 없으면 만들고 저장
    os.makedirs(os.path.dirname(SAVE_FILE), exist_ok=True)
    book.save(SAVE_FILE)
    print("① 통합 완료 →", SAVE_FILE)


# 4. 대상 폴더에서 파일 조회하는 함수 정의
def enum_files():
    files = sorted(glob.glob(INPUT_PATTERN))
    # 엑셀이 열려 있을 때 생기는 임시파일(~$로 시작)은 제외
    return [f for f in files if not os.path.basename(f).startswith("~$")]


# 5. 파일 열어서 내용을 시트에 복사하는 함수 정의
def read_book(merge_sheet, fname):
    print("  읽는 중:", os.path.basename(fname))
    # data_only=True : 수식(=D4*E4) 대신 계산된 값(3000)을 읽음
    book = excel.load_workbook(fname, data_only=True)
    sheet = book.active

    # 1~3행은 제목/담당/머리글이므로 4행부터 읽기
    for row in sheet.iter_rows(min_row=4, values_only=True):
        if row[0] is None:          # 일자가 비어 있으면 데이터 끝
            break
        date, customer, item, count, price, subtotal = row[:6]
        # 소계가 비어 있으면(수식 값이 저장 안 된 경우) 직접 계산
        if subtotal is None:
            subtotal = count * price
        merge_sheet.append([date, customer, item, count, price, subtotal])


# 6. 메인 프로그램 실행하기
#    → 이 파일을 직접 실행할 때만 동작. 다른 파일에서 import 하면 실행되지 않음(모듈로 재사용)
if __name__ == "__main__":
    merge_files()
