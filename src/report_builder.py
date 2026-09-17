"""output/*.xlsx 전체를 읽어 업소별·카테고리별·원재료(기능성)별 랭킹 JSON을 docs/data/에 생성합니다.

월별 파일 + 전체 누적(cumulative) 두 종류를 만듭니다. 정적 페이지(docs/index.html)가
fetch()로 이 JSON을 읽어 랭킹을 렌더링합니다.
"""

import glob
import json
import os
import re
from collections import Counter

from openpyxl import load_workbook

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(ROOT_DIR, 'output')
DATA_DIR = os.path.join(ROOT_DIR, 'docs', 'data')

TOP_N = 20

# 일반식품 원재료명은 부형제·감미료·용매 등 범용 성분이 섞여 있어 그대로 세면
# "정제수/이산화규소" 같은 게 상위권을 차지함 — 신제품 기획에 의미 없는 성분은 제외.
GENERAL_EXCLUDE_INGREDIENTS = {
    '정제수', '물', '설탕', '정백당', '정제염', '정제소금', '식염', '포도당', '과당', '올리고당',
    '이산화규소', '스테아린산마그네슘', '결정셀룰로스',
    '히드록시프로필메틸셀룰로스', '카복시메틸셀룰로스', '카복시메틸셀룰로스칼슘',
    '카복시메틸셀룰로스나트륨', '가교카복시메틸셀룰로스나트륨',
    '구연산', '구연산삼나트륨', '구연산(무수)', '제이인산칼슘',
    '로커스트콩검', '잔탄검', '펙틴', '아미드펙틴', '카라기난',
    '혼합제제', '향료', '식품첨가물', '식품첨가물혼합제제',
    '탄산수소나트륨', '알긴산', '이산화티타늄', '치자황색소', 'DL-사과산', '시클로덱스트린액',
    # C002 제품형태 필터 10종 — 원재료 목록에도 자기참조로 섞여 나와 "원재료"가 아님
    '기타가공품', '과.채가공품', '액상차', '당류가공품', '고형차',
    '캔디류', '과.채주스', '효소식품', '발효식초', '올리브유',
    # 원재료명에 흔히 등장하는 범용 가공품 분류명(구체 원료가 아닌 상위 분류)
    '기타 농산가공품', '식물성크림', '수크랄로스', '알룰로오스', '효소처리스테비아',
    '물엿', '덱스트린', '유당', '혼합유당',
}

# 사용자가 지정한 관심 제조사 목록(2026-09-17) — 업소명에 공장/지점 suffix가 붙어 여러 줄로
# 흩어지는 경우가 많아(예: "콜마비앤에이치(주)세종3공장"/"...음성공장") 대표 키워드로 부분일치
# 매칭해 한 회사로 합산한다. "서흥"은 계열사 "서흥헬스케어"도 함께 집계(같은 브랜드로 판단).
COMPANY_WATCHLIST = [
    ('노바렉스', '노바렉스'),
    ('한미양행', '한미양행'),
    ('서흥', '서흥'),
    ('우리바이오', '우리바이오'),
    ('콜마비앤에이치', '콜마비앤에이치'),
    ('코스맥스바이오', '코스맥스바이오'),
    ('코스맥스엔비티', '코스맥스엔비티'),
    ('유유헬스케어', '유유헬스케어'),
    ('동서바이오팜', '동서바이오팜'),
    ('대원헬스케어', '대원헬스케어'),
]

_BRACKET_RE = re.compile(r'\[([^\]]+)\]')


def _watchlist_counts(company_counter):
    result = [
        (label, sum(c for name, c in company_counter.items() if keyword in name))
        for label, keyword in COMPANY_WATCHLIST
    ]
    result.sort(key=lambda pair: -pair[1])
    return result


def _normalize_health_ingredient(name):
    # 원본 데이터에 "비타민 B6"/"비타민B6"처럼 공백 유무가 뒤섞여 있어 같은 성분이
    # 다른 이름으로 집계되는 걸 방지 (다른 성분명까지 붙는 걸 막기 위해 알파벳/숫자
    # 앞의 공백만 좁게 제거)
    return re.sub(r'비타민\s+(?=[A-Za-z0-9])', '비타민', name)


def _extract_health_ingredients(functional_text):
    # 주된기능성 필드는 "[성분명]설명..." 형태로 기능성 성분이 괄호로 태깅돼 있음
    return {_normalize_health_ingredient(t) for t in _BRACKET_RE.findall(functional_text or '')}


def _extract_general_ingredients(raw_text):
    tokens = [t.strip() for t in (raw_text or '').split(',')]
    return {
        t for t in tokens
        if t and t not in GENERAL_EXCLUDE_INGREDIENTS
        # "OO가공품"/"OO식품" 형태는 구체 원료가 아니라 상위 가공분류명인 경우가 많음
        # (기타가공품/두류가공품/유함유가공품/효모식품/효소식품 등)
        and not t.endswith('가공품') and not t.endswith('식품')
    }


def _read_month_file(fpath):
    """xlsx 한 파일에서 건강기능식품/일반식품 레코드를 딕셔너리 리스트로 읽음."""
    wb = load_workbook(fpath, read_only=True, data_only=True)
    result = {}
    for sheet_name in ('건강기능식품', '일반식품'):
        if sheet_name not in wb.sheetnames:
            result[sheet_name] = []
            continue
        ws = wb[sheet_name]
        rows = ws.iter_rows(values_only=True)
        headers = list(next(rows, []))
        records = [dict(zip(headers, row)) for row in rows]
        result[sheet_name] = records
    wb.close()
    return result


def _aggregate(records, kind):
    """kind: 'health' 또는 'general'"""
    company = Counter()
    category = Counter()
    ingredient = Counter()

    for r in records:
        if r.get('업소명'):
            company[r['업소명']] += 1
        category[r.get('카테고리') or '기타'] += 1

        if kind == 'health':
            for ing in _extract_health_ingredients(r.get('주된기능성', '')):
                ingredient[ing] += 1
        else:
            for ing in _extract_general_ingredients(r.get('원재료명', '')):
                ingredient[ing] += 1

    return {
        'total': len(records),
        'company': company.most_common(TOP_N),
        'company_watchlist': _watchlist_counts(company),
        'category': category.most_common(),
        'ingredient': ingredient.most_common(TOP_N),
    }


def _build_month_payload(records_by_sheet):
    return {
        'health': _aggregate(records_by_sheet['건강기능식품'], 'health'),
        'general': _aggregate(records_by_sheet['일반식품'], 'general'),
    }


def build():
    os.makedirs(DATA_DIR, exist_ok=True)

    pattern = os.path.join(OUTPUT_DIR, '*_품목신고보고현황.xlsx')
    files = sorted(glob.glob(pattern))
    if not files:
        print('output/*.xlsx 파일이 없습니다.')
        return

    months = []
    cumulative_health = {}  # 품목제조번호 -> record (최신 파일이 덮어씀)
    cumulative_general = {}

    for fpath in files:
        fname = os.path.basename(fpath)
        month = fname[:7]  # 'YYYY-MM'
        try:
            data = _read_month_file(fpath)
        except Exception as e:
            print(f'경고: {fpath} 읽기 실패 — {e}')
            continue

        months.append(month)
        with open(os.path.join(DATA_DIR, f'{month}.json'), 'w', encoding='utf-8') as f:
            json.dump(_build_month_payload(data), f, ensure_ascii=False, indent=2)
        print(f'{month}.json 생성 완료 (건기식 {len(data["건강기능식품"])}건 / 일반 {len(data["일반식품"])}건)')

        for r in data['건강기능식품']:
            key = r.get('품목제조번호')
            if key:
                cumulative_health[key] = r
        for r in data['일반식품']:
            key = r.get('품목제조번호')
            if key:
                cumulative_general[key] = r

    cumulative_payload = _build_month_payload({
        '건강기능식품': list(cumulative_health.values()),
        '일반식품': list(cumulative_general.values()),
    })
    with open(os.path.join(DATA_DIR, 'cumulative.json'), 'w', encoding='utf-8') as f:
        json.dump(cumulative_payload, f, ensure_ascii=False, indent=2)
    print(f'cumulative.json 생성 완료 (건기식 {len(cumulative_health)}건 / 일반 {len(cumulative_general)}건, 품목제조번호 기준 중복제거)')

    with open(os.path.join(DATA_DIR, 'months.json'), 'w', encoding='utf-8') as f:
        json.dump(months, f, ensure_ascii=False, indent=2)


if __name__ == '__main__':
    build()
