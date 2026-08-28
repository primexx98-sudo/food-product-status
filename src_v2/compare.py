"""기존 category_mapper.py vs category_mapper_v2.py 분류 결과 비교.

사용법:
  python compare.py                  전체 diff 요약
  python compare.py --category 피부   특정 카테고리만 상세 diff (구→신, 신→구 양방향)
  python compare.py --sample 항산화 30   해당 카테고리로 새로 분류된 품목 샘플 30건 출력(원재료 포함)
"""
import sys
import os
import glob
import argparse
from collections import defaultdict

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
sys.path.insert(0, os.path.dirname(__file__))

from category_mapper import categorize_general_food, is_general_excluded  # 기존
from category_mapper_v2 import categorize_general_food_v2, DONE  # 신규

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(ROOT, 'output')


def load_all_items():
    """모든 월별 파일에서 일반식품 원본 데이터를 품목제조번호 기준 중복 제거해서 로드."""
    items = {}
    for path in sorted(glob.glob(os.path.join(OUTPUT_DIR, '*.xlsx'))):
        wb = load_workbook(path, data_only=True)
        ws = wb['일반식품']
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        idx = {h: i for i, h in enumerate(header)}
        for r in rows[1:]:
            report_no = str(r[idx['품목제조번호']] or '')
            name = str(r[idx['품목명']] or '')
            raw = str(r[idx['원재료명']] or '')
            company = str(r[idx['업소명']] or '')
            items[report_no] = (name, raw, company)
    return items


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--category', help='이 카테고리만 상세 diff')
    ap.add_argument('--sample', nargs=2, metavar=('CATEGORY', 'N'), help='이 카테고리로 새로 분류된 품목 샘플 N건')
    args = ap.parse_args()

    items = load_all_items()
    print(f"전체 품목(품목제조번호 기준 중복제거): {len(items)}건\n")

    if args.sample:
        cat, n = args.sample[0], int(args.sample[1])
        n_shown = 0
        for report_no, (name, raw, company) in items.items():
            if is_general_excluded(name, report_no, company):
                continue
            new_cat = categorize_general_food_v2(name, raw, report_no)
            if new_cat == cat:
                print(f"[{cat}] {name}  |  원재료: {raw[:70]}  |  업소: {company}")
                n_shown += 1
                if n_shown >= n:
                    break
        print(f"\n(v2 기준 '{cat}' 총 매칭 —  위는 앞 {n}건만 표시)")
        return

    old_counts = defaultdict(int)
    new_counts = defaultdict(int)
    diff_old_to_new = defaultdict(lambda: defaultdict(int))

    for report_no, (name, raw, company) in items.items():
        if is_general_excluded(name, report_no, company):
            continue
        old_cat = categorize_general_food(name, raw, report_no)
        new_cat = categorize_general_food_v2(name, raw, report_no)
        old_counts[old_cat] += 1
        new_counts[new_cat] += 1
        if old_cat != new_cat:
            diff_old_to_new[old_cat][new_cat] += 1

    if args.category:
        cat = args.category
        print(f"=== '{cat}' 상세 diff ===")
        print(f"기존 로직 '{cat}' 건수: {old_counts.get(cat, 0)}")
        print(f"v2 로직   '{cat}' 건수: {new_counts.get(cat, 0)}")
        print(f"\n[기존={cat} → v2=다른 카테고리로 바뀐 것] (v2가 아직 이 카테고리를 다루지 않았다면 전부 '기타'로 표시됨)")
        for report_no, (name, raw, company) in items.items():
            if is_general_excluded(name, report_no, company):
                continue
            old_cat = categorize_general_food(name, raw, report_no)
            new_cat = categorize_general_food_v2(name, raw, report_no)
            if old_cat == cat and new_cat != cat:
                print(f"  {name}  [{old_cat} -> {new_cat}]  원재료: {raw[:60]}")
        print(f"\n[v2={cat} → 기존=다른 카테고리였던 것] (v2가 새로 잡아낸 것)")
        for report_no, (name, raw, company) in items.items():
            if is_general_excluded(name, report_no, company):
                continue
            old_cat = categorize_general_food(name, raw, report_no)
            new_cat = categorize_general_food_v2(name, raw, report_no)
            if new_cat == cat and old_cat != cat:
                print(f"  {name}  [{old_cat} -> {new_cat}]  원재료: {raw[:60]}")
        return

    print("=== 카테고리별 건수 (기존 vs v2) ===")
    all_cats = sorted(set(old_counts) | set(new_counts))
    for cat in all_cats:
        marker = '' if cat in DONE or cat == '기타' else '  (v2 미착수)'
        print(f"  {cat}: 기존 {old_counts.get(cat,0)} -> v2 {new_counts.get(cat,0)}{marker}")


if __name__ == '__main__':
    main()
