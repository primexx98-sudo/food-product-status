"""일반식품 분류 로직 v2 (실험용).

기존 src/category_mapper.py의 categorize_general_food를 카테고리별 인터뷰로 처음부터
재구축하는 버전. 카테고리 목록은 기존과 동일하게 유지한다.

진행 방식: GENERAL_CAT_LIST 순서대로 카테고리 하나씩 인터뷰를 거쳐 _MAP에 항목을 채운다.
아직 다루지 않은 카테고리는 매칭 대상에서 빠지므로(빈 리스트), 해당 품목은 자동으로 '기타'로 떨어진다
— 즉 인터뷰가 끝난 카테고리만 v2 결과에 반영되고, 나머지는 기존 로직과 비교 시 '아직 미착수'로 표시됨.
"""

GENERAL_CAT_LIST = [
    '피부', '에너지', '단백질', '효소', '항산화', '혈행개선', '관절', '다이어트', '면역',
    '수면', '피로개선', '프로바이오틱스', '구강', '여성건강', '남성건강',
    '모발', '혈당', '기타',
]

# 인터뷰로 확정된 카테고리만 여기 채운다. {카테고리: (품목명 매칭 키워드 리스트, 원재료명 매칭 키워드 리스트)}
# 원재료명 키워드는 품목명 키워드보다 보수적으로(부원료 오매칭 위험 고려) 채운다.
_MAP: dict[str, tuple[list[str], list[str]]] = {
    # 2026-08-28 인터뷰 확정. 기존 대비 '뷰티'/'글로우' 삭제(브랜드명에만 있고 실제 성분과
    # 무관한 오탐 다수 확인 — 뷰티풀농장도라지배즙/터닝뷰티클렌즈/글로우빈헛개차진액/
    # 데이튠글로우온/매쉬드글로우베리즈/뷰티바이옴(실제론 유산균 제품) 등 6건), '뮤신'은 유지
    # (실제 피부보습 성분 신호, 오탐 없음 확인).
    '피부': (
        ['콜라겐', '히알루론산', '글루타치온', '피부탄력', '피부미용', '피부건강',
         'PDRN', '엘라스틴', '병풀', '마데카', '뮤신', '연어이리추출물'],
        # 원재료 매칭은 기존과 동일하게 보수적으로 — 콜라겐/히알루론산/PDRN 등은 관절 등
        # 다른 목적 제품에도 부원료로 흔히 섞여 오매칭 위험 있음(기존 로직에서 이미 확인됨)
        ['피부탄력', '피부미용', '피부건강'],
    ),
}

# 인터뷰 완료 표시 — 여기 있는 카테고리만 v2가 "확정 판단"한 것으로 간주.
# compare.py에서 미완료 카테고리는 비교 대상에서 제외(공정 비교를 위해).
DONE: set[str] = {'피부'}


def categorize_general_food_v2(product_name: str, raw_material: str = '', report_no: str = '') -> str:
    name = str(product_name or '')
    raw = str(raw_material or '')

    for cat, (name_kws, raw_kws) in _MAP.items():
        if any(kw.lower() in name.lower() for kw in name_kws):
            return cat
    for cat, (name_kws, raw_kws) in _MAP.items():
        if any(kw.lower() in raw.lower() for kw in raw_kws):
            return cat
    return '기타'
