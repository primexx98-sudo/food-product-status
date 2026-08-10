import os
import sys
from datetime import datetime

import requests

from api_client import fetch_health_food, fetch_general_food
from category_mapper import categorize_health_food, categorize_general_food, is_general_excluded
from excel_writer import write_excel


def main():
    api_key = os.environ.get('FOOD_API_KEY', '').strip()
    if not api_key:
        print("오류: FOOD_API_KEY 환경변수가 설정되지 않았습니다.")
        sys.exit(1)

    if len(sys.argv) >= 3:
        year, month = int(sys.argv[1]), int(sys.argv[2])
    else:
        now = datetime.now()
        year, month = now.year, now.month
    print(f"[{year}-{month:02d}] 데이터 수집 시작")

    print("건강기능식품 수집 중...")
    try:
        health_data = fetch_health_food(api_key, year, month)
    except requests.exceptions.HTTPError as e:
        status = e.response.status_code
        if status == 403:
            print("  [건너뜀] HTTP 403 — 해외 IP 차단. GitHub Actions에서는 수집 불가.")
        else:
            print(f"  [건너뜀] HTTP {status} 오류. 오늘 수집을 건너뜁니다.")
        sys.exit(0)
    except requests.exceptions.Timeout:
        print("  [건너뜀] 타임아웃 — 일시적 장애 또는 IP 차단. 내일 재시도합니다.")
        sys.exit(0)
    except requests.exceptions.RequestException as e:
        print(f"  [건너뜀] 연결 실패 ({type(e).__name__}): {e}")
        sys.exit(0)
    except RuntimeError as e:
        print(f"  [오류] API 응답 오류: {e}")
        print("  → 인증키 만료·쿼터 소진·서버 장애 중 하나일 수 있습니다.")
        print("  → 식품안전나라 포털에서 인증키 상태를 확인하세요.")
        sys.exit(1)
    for item in health_data:
        item['카테고리'] = categorize_health_food(
            item.get('주된기능성', ''),
            item.get('품목명', ''),
            item.get('원재료', ''),
        )
    print(f"  건강기능식품: {len(health_data)}건")

    print("일반식품 수집 중...")
    general_data = fetch_general_food(api_key, year, month)
    general_data = [
        r for r in general_data
        if not is_general_excluded(r.get('품목명', ''), r.get('품목제조번호', ''))
    ]
    general_data.sort(key=lambda x: str(x.get('보고일자', '') or ''))
    for item in general_data:
        item['카테고리'] = categorize_general_food(
            item.get('품목명', ''), item.get('원재료명', ''), item.get('품목제조번호', '')
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
