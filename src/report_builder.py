"""output/*.xlsx 전체를 읽어 업소별·카테고리별·원재료(기능성)별 랭킹 JSON을 docs/data/에 생성합니다.

월별 파일 + 전체 누적(cumulative) 두 종류를 만듭니다. 정적 페이지(docs/index.html)가
fetch()로 이 JSON을 읽어 랭킹을 렌더링합니다.

2026-09-18 추가: 월별 파일에는 전월 대비 카테고리·원재료 순위 변동(rank delta)과 급상승
원재료 목록을 함께 생성합니다. 누적(cumulative.json)은 "전월"에 대응하는 개념이 없어 이
필드들을 생성하지 않습니다. (카테고리 월간 추이 스파크라인은 같은 날 UI에서 제거되어
trend.json 생성도 함께 중단했습니다.)
"""

import glob
import json
import os
import re
from collections import Counter, defaultdict

from openpyxl import load_workbook

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(ROOT_DIR, 'output')
DATA_DIR = os.path.join(ROOT_DIR, 'docs', 'data')

TOP_N = 20
TRENDING_LIMIT = 8
# 1~2건짜리 등락까지 "급상승"으로 띄우면 노이즈라 최소 건수 기준을 둠
TRENDING_MIN_COUNT = 3

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
    # 2026-09-17 추가 — 실제 신고건수 데이터 기준 비슷한 급의 OEM/ODM 추천 8곳
    ('메디오젠', '메디오젠'),
    ('한풍네이처팜', '한풍네이처팜'),
    ('비오팜', '비오팜'),
    ('엠에스바이오텍', '엠에스바이오텍'),
    ('광동헬스바이오', '광동헬스바이오'),
    ('일동바이오사이언스', '일동바이오사이언스'),
    ('웰레스트', '웰레스트'),
    ('엔피케이', '엔피케이'),
    ('알피바이오', '알피바이오'),
]

_BRACKET_RE = re.compile(r'\[([^\]]+)\]')


def _watchlist_counts(company_counter):
    result = [
        (label, sum(c for name, c in company_counter.items() if keyword in name))
        for label, keyword in COMPANY_WATCHLIST
    ]
    result.sort(key=lambda pair: -pair[1])
    return result


def _watchlist_items(company_items):
    """지정 제조사(라벨)별 실제 신고 품목 리스트 — 공장/지점별로 흩어진 업소명을 키워드
    부분일치로 합쳐서 모은다(`_watchlist_counts`와 동일한 매칭 방식)."""
    return {
        label: [item for comp, items in company_items.items() if keyword in comp for item in items]
        for label, keyword in COMPANY_WATCHLIST
    }


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


def _report_date(r):
    return str(r.get('보고일자') or r.get('신고일자') or '')


_CORP_TOKENS_RE = re.compile(r'\(주\)|주식회사')


def _company_root(name):
    """업소명에서 법인 표기를 지우고 공백으로 구분된 첫 토큰만 남긴다.

    실제 데이터에 같은 회사가 "(주)메디오젠 충주공장"/"(주)메디오젠 제천공장"/
    "(주)메디오젠"처럼 공장·지점명이 공백 뒤에 붙어 서로 다른 문자열로 흩어지는 경우가
    흔해(COMPANY_WATCHLIST가 부분일치로 우회하는 것과 동일한 문제), 정확히 같은 업소명만
    요구하면 재신고 탐지를 놓친다. 첫 토큰만 비교해 보수적으로 넓히되, 품목명·카테고리까지
    동시에 일치해야 병합되므로 우연히 회사명 첫 단어만 같은 별개 업체까지 합쳐질 위험은
    낮다(공백 없이 붙어 있는 "OO(주)세종3공장" 같은 표기는 이 정규화로도 못 잡지만, 그
    경우는 그냥 병합을 건너뛸 뿐 잘못 병합되는 방향의 실수는 아니다).
    """
    n = _CORP_TOKENS_RE.sub('', str(name or '')).strip()
    return n.split(' ')[0] if n else n


def _dedupe_resubmissions(records):
    """동일 업소(첫 토큰 기준)가 동일 품목명·동일 카테고리로 서로 다른 품목제조번호를 받아
    재신고한 경우(2026-09-18 메디오젠 사례로 발견 — 예: 2주~6주 간격으로 재등록, 원재료
    공급처 변경 등으로 추정)를 같은 제품의 갱신으로 보고 최신 보고일자 1건만 남긴다.

    업소명/품목명이 없는 레코드는 중복 판정 자체가 불가능하므로 그대로 둔다.
    """
    best = {}
    passthrough = []
    for r in records:
        company, name = r.get('업소명'), r.get('품목명')
        if not company or not name:
            passthrough.append(r)
            continue
        key = (_company_root(company), name, r.get('카테고리'))
        existing = best.get(key)
        if existing is None or _report_date(r) > _report_date(existing):
            best[key] = r
    removed = len(records) - len(best) - len(passthrough)
    return list(best.values()) + passthrough, removed


def _raw_aggregate(records, kind):
    """kind: 'health' 또는 'general'. 잘라내기 전 전체 Counter를 반환 (전월 대비 비교용).

    2026-09-18 추가: 카드 클릭 드릴다운용으로 카테고리·원재료·업소별 실제 품목 리스트
    ([품목명, 업소명] 쌍)도 함께 모은다. 최종 JSON엔 `_finalize`에서 필요한 항목(표시되는
    상위 N개)만 골라 담아 파일 크기를 억제한다.
    """
    records, removed = _dedupe_resubmissions(records)
    if removed:
        print(f'  [{kind}] 동일 업소·품목명·카테고리 재신고 {removed}건을 최신 1건으로 병합')

    company = Counter()
    category = Counter()
    ingredient = Counter()
    category_items = defaultdict(list)
    ingredient_items = defaultdict(list)
    company_items = defaultdict(list)  # 업소명 원문 기준 (watchlist는 finalize에서 키워드로 재취합)

    for r in records:
        name = r.get('품목명') or ''
        comp = r.get('업소명') or ''
        cat = r.get('카테고리') or '기타'

        if comp:
            company[comp] += 1
            company_items[comp].append([name, cat])
        category[cat] += 1
        category_items[cat].append([name, comp])

        if kind == 'health':
            ings = _extract_health_ingredients(r.get('주된기능성', ''))
        else:
            ings = _extract_general_ingredients(r.get('원재료명', ''))
        for ing in ings:
            ingredient[ing] += 1
            ingredient_items[ing].append([name, comp])

    return {
        'total': len(records), 'company': company, 'category': category, 'ingredient': ingredient,
        'category_items': category_items, 'ingredient_items': ingredient_items, 'company_items': company_items,
    }


def _rank_map(counter):
    return {name: i for i, (name, _) in enumerate(counter.most_common())}


def _rank_deltas(curr_counter, prev_counter, top_items):
    """top_items 중 전월에도 있던 항목은 순위 변동(양수=상승)을, 전월에 없던 항목은
    new_list에 담아 반환. prev_counter가 None이면(수집 첫 달) 둘 다 비워 반환한다."""
    if prev_counter is None:
        return {}, []
    curr_rank = _rank_map(curr_counter)
    prev_rank = _rank_map(prev_counter)
    delta_map = {}
    new_list = []
    for name, _ in top_items:
        if not prev_counter.get(name):
            new_list.append(name)
        elif name in curr_rank and name in prev_rank:
            delta_map[name] = prev_rank[name] - curr_rank[name]
    return delta_map, new_list


def _trending_ingredients(curr_counter, prev_counter):
    """전월 대비 건수가 급증했거나 이번 달 새로 등장한 원재료 TOP N (노이즈 방지용 최소 건수 적용)."""
    if prev_counter is None:
        return []
    items = []
    for name, count in curr_counter.items():
        if count < TRENDING_MIN_COUNT:
            continue
        prev_count = prev_counter.get(name, 0)
        is_new = prev_count == 0
        delta = count - prev_count
        if delta <= 0:
            continue
        items.append({'name': name, 'count': count, 'count_delta': delta, 'is_new': is_new})
    items.sort(key=lambda x: (x['is_new'], x['count_delta']), reverse=True)
    return items[:TRENDING_LIMIT]


def _finalize(raw, prev_raw):
    """_raw_aggregate 결과를 JSON 직렬화 가능한 랭킹 payload(집계 수치만)로 변환.

    2026-09-18: 실제 품목 리스트(드릴다운용)는 용량이 커서(누적 기준 ~2MB) 메인 페이로드에
    안 넣고 `_finalize_items()`로 분리해 별도 `*_items.json`에 저장 — 모바일에서 페이지 첫
    로드 시 무거운 품목 리스트까지 매번 받지 않고, 사용자가 실제로 드릴다운을 열 때만
    지연 로딩(lazy fetch)하기 위함.
    """
    company, category, ingredient = raw['company'], raw['category'], raw['ingredient']
    prev_category = prev_raw['category'] if prev_raw else None
    prev_ingredient = prev_raw['ingredient'] if prev_raw else None

    category_top = category.most_common()
    ingredient_top = ingredient.most_common(TOP_N)

    category_delta, category_new = _rank_deltas(category, prev_category, category_top)
    ingredient_delta, ingredient_new = _rank_deltas(ingredient, prev_ingredient, ingredient_top)

    return {
        'total': raw['total'],
        'company': company.most_common(TOP_N),
        'company_watchlist': _watchlist_counts(company),
        'category': category_top,
        'category_delta': category_delta,
        'category_new': category_new,
        'ingredient': ingredient_top,
        'ingredient_delta': ingredient_delta,
        'ingredient_new': ingredient_new,
        'ingredient_trending': _trending_ingredients(ingredient, prev_ingredient),
    }


def _finalize_items(raw):
    """드릴다운 모달용 실제 품목 리스트. 화면에 표시되는 상위 항목(카테고리 전체,
    원재료·업소 TOP N)에 대해서만 담아 용량을 제한한다."""
    category_top = raw['category'].most_common()
    ingredient_top = raw['ingredient'].most_common(TOP_N)
    return {
        'company_watchlist_items': _watchlist_items(raw['company_items']),
        'category_items': {name: raw['category_items'].get(name, []) for name, _ in category_top},
        'ingredient_items': {name: raw['ingredient_items'].get(name, []) for name, _ in ingredient_top},
    }


def _build_month_payload(raw_by_kind, prev_raw_by_kind):
    prev_raw_by_kind = prev_raw_by_kind or {}
    return {
        'health': _finalize(raw_by_kind['health'], prev_raw_by_kind.get('health')),
        'general': _finalize(raw_by_kind['general'], prev_raw_by_kind.get('general')),
    }


def _build_items_payload(raw_by_kind):
    return {
        'health': _finalize_items(raw_by_kind['health']),
        'general': _finalize_items(raw_by_kind['general']),
    }


def _build_cumulative_raw(records_by_sheet):
    return {
        'health': _raw_aggregate(records_by_sheet['건강기능식품'], 'health'),
        'general': _raw_aggregate(records_by_sheet['일반식품'], 'general'),
    }


def build():
    os.makedirs(DATA_DIR, exist_ok=True)

    pattern = os.path.join(OUTPUT_DIR, '*_품목신고보고현황.xlsx')
    files = sorted(glob.glob(pattern))
    if not files:
        print('output/*.xlsx 파일이 없습니다.')
        return

    months = []
    raw_by_month = []
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
        raw_by_month.append({
            'health': _raw_aggregate(data['건강기능식품'], 'health'),
            'general': _raw_aggregate(data['일반식품'], 'general'),
        })

        for r in data['건강기능식품']:
            key = r.get('품목제조번호')
            if key:
                cumulative_health[key] = r
        for r in data['일반식품']:
            key = r.get('품목제조번호')
            if key:
                cumulative_general[key] = r

    for i, month in enumerate(months):
        prev_raw = raw_by_month[i - 1] if i > 0 else None
        payload = _build_month_payload(raw_by_month[i], prev_raw)
        with open(os.path.join(DATA_DIR, f'{month}.json'), 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        items_payload = _build_items_payload(raw_by_month[i])
        with open(os.path.join(DATA_DIR, f'{month}_items.json'), 'w', encoding='utf-8') as f:
            json.dump(items_payload, f, ensure_ascii=False, indent=2)
        print(f'{month}.json 생성 완료 (건기식 {payload["health"]["total"]}건 / 일반 {payload["general"]["total"]}건)')

    cumulative_raw = _build_cumulative_raw({
        '건강기능식품': list(cumulative_health.values()),
        '일반식품': list(cumulative_general.values()),
    })
    cumulative_payload = {
        'health': _finalize(cumulative_raw['health'], None),
        'general': _finalize(cumulative_raw['general'], None),
    }
    with open(os.path.join(DATA_DIR, 'cumulative.json'), 'w', encoding='utf-8') as f:
        json.dump(cumulative_payload, f, ensure_ascii=False, indent=2)
    cumulative_items_payload = _build_items_payload(cumulative_raw)
    with open(os.path.join(DATA_DIR, 'cumulative_items.json'), 'w', encoding='utf-8') as f:
        json.dump(cumulative_items_payload, f, ensure_ascii=False, indent=2)
    print(f'cumulative.json 생성 완료 (건기식 {len(cumulative_health)}건 / 일반 {len(cumulative_general)}건, 품목제조번호 기준 중복제거)')

    with open(os.path.join(DATA_DIR, 'months.json'), 'w', encoding='utf-8') as f:
        json.dump(months, f, ensure_ascii=False, indent=2)


if __name__ == '__main__':
    build()
