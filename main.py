# main.py
# 네 단계를 순서대로 실행하는 프로그램  →  python main.py
import step1_merge, step2_aggregate, step3_invoice, step4_summary

step1_merge.merge_files()        # ① 담당자별 매출 모으기
step2_aggregate.split_data()     # ② 고객별 집계 → JSON
step3_invoice.fill_template()    # ③ 고객별 청구서 발행
step4_summary.make_summary()     # ④ 청구 요약표 만들기 (직접 추가)
