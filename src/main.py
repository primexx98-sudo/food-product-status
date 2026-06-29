import os
import sys
from datetime import datetime

import requests

from api_client import fetch_health_food, fetch_general_food
from category_mapper import categorize_health_food, categorize_general_food
from excel_writer import write_excel


def main():
    api_key = os.environ.get('FOOD_API_KEY', '').strip()
    if not api_key:
        print("오류: FOOD_API_KEY 환경변수가 설정되지 않았습니다.")
        sys.exit(1)

    now = datetime.now()
    year, month = now.year, now.month
    print(f"[{year}-{month:02d}] 데이터 수집 시작")

    print("건강기능식품 수집 중...")
    try:
        health_data = fetch_health_food(api_key, year, month)
    except requests.exceptions.RequestException as e:
        print(f"  [건너뜀] API 서버 연결 실패, 오늘 수집을 건너뜁니다: {e}")
        sys.exit(0)
    for item in health_data:
        item['카테고리'] = categorize_health_food(
            item.get('주된기능성', ''),
            item.get('품목명', ''),
            item.get('원재료', ''),
        )
    print(f"  건강기능식품: {len(health_data)}건")

    print("일반식품 수집 중...")
    general_data = fetch_general_food(api_key, year, month)
    for item in general_data:
        item['카테고리'] = categorize_general_food(
            item.get('품목명', ''), item.get('원재료명', '')
        )
    print(f"  일반식품: {len(general_data)}건")

    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    output_dir = os.path.join(root_dir, 'output')
    os.makedirs(output_dir, exist_ok=True)

    filename = f"{year}-{month:02d}_품목신고보고현황.xlsx"
    filepath = os.path.join(output_dir, filename)

    write_excel(filepath, output_dir, year, month, health_data, general_data)


if __name__ == '__main__':
    main()
