import re

HEALTH_CAT_LIST = [
    '멀티팩', '비타민', '다이어트', '이너뷰티', '장', '눈', '두뇌',
    '뼈건강', '관절', '간', '위', '피로개선', '수면', '항산화',
    '혈행개선', '혈당', '여성건강', '남성건강', '프로바이오틱스', '오메가3', '홍삼', '단백질',
]

GENERAL_CAT_LIST = [
    '피부', '에너지', '단백질', '효소', '관절', '다이어트', '면역',
    '수면', '피로개선', '프로바이오틱스', '구강', '여성건강', '남성건강',
    '모발', '혈당', '기타',
]

# ─────────────────────────────────────────────────────────────────
# 건강기능식품 카테고리 분류 맵
# 우선순위: 위에서 아래 순서로 매칭 (더 구체적인 원료가 위에 위치)
# 수정 방법: 키워드 리스트에 단어 추가/제거하면 됨
# ─────────────────────────────────────────────────────────────────
_HEALTH_MAP = [
    # [원료명] 괄호에서 직접 식별되는 고특이도 원료
    ('홍삼',           ['홍삼', '홍삼제품', '진세노사이드', '사포닌', '흑삼', '인삼']),
    ('프로바이오틱스',  ['프로바이오틱스', '프로바이오틱스 제품', '유산균',
                       'Lactobacillus', 'Bifidobacterium']),
    ('오메가3',        ['EPA 및 DHA 함유 유지', 'EPA및DHA함유유지',
                       'EPA 및 DHA함유유지', 'EPA및DHA', '오메가3', '오메가-3']),
    ('간',             ['밀크씨슬']),
    ('여성건강',       ['회화나무열매', '대두이소플라본', '감마리놀렌산', '크랜베리 분말',
                       '석류', '갱년기']),
    ('남성건강',       ['쏘팔메토', '전립선']),
    ('다이어트',       ['가르시니아', '공액리놀레산', '카르니틴', '히비스커스',
                       '난소화성말토덱스트린', '체지방감소', '체중감소']),
    ('혈당',           ['바나바잎', '크롬']),
    ('수면',           ['유단백가수분해물', '락티움']),
    ('이너뷰티',       ['히알루론산', '저분자콜라겐', '콜라겐펩타이드', '피부탄력', '자외선']),
    ('눈',             ['마리골드꽃추출물', '헤마토코쿠스', '루테인', '지아잔틴', '황반']),
    ('두뇌',           ['포스파티딜세린', '은행잎 추출물', '은행잎추출물', '기억력', '인지기능']),
    ('혈행개선',       ['홍국', '혈중중성지방', '혈행개선', '혈액순환']),
    ('항산화',         ['프로폴리스', '코엔자임Q10', '피크노제놀', '항산화']),
    ('관절',           ['엠에스엠', 'MSM', '뮤코다당', '글루코사민', '콘드로이친', 'NAG']),
    ('뼈건강',         ['골다공', '골밀도', '뼈건강']),
    ('피로개선',       ['홍경천', '피로개선', '피로회복']),
    ('장',             ['차전자피', '프락토올리고당', '자일로올리고당',
                       '알로에 겔', '알로에겔', '배변활동']),
    ('위',             ['스페인감초', '매스틱검', '위건강', '위장건강']),
    ('단백질',         ['크레아틴', 'WPI', 'WPC', '유청단백']),
    # 비타민은 맨 마지막 (보조영양소로 다수 카테고리에 등장하므로 fallback)
    ('비타민',         ['비타민A', '비타민 A', '비타민B', '비타민 B',
                       '비타민C', '비타민 C', '비타민D', '비타민 D',
                       '비타민E', '비타민 E', '비타민K', '비타민 K',
                       '나이아신', '엽산', '판토텐산', '비오틴', '비타민']),
]

# 제품명에서 멀티팩 감지용 키워드
# 비타민/미네랄 원료명 집합 (이것들만 있으면 무조건 비타민)
_VITAMIN_MINERAL_SET = {
    '비타민', '나이아신', '엽산', '판토텐산', '비오틴',
    '아연', '마그네슘', '칼슘', '셀레늄', '셀렌', '철',
    '구리', '망간', '크롬', '몰리브덴', '요오드', '칼륨', '인',
}

def _is_vitamin_mineral_only(items: list[str]) -> bool:
    """모든 기능성 원료가 비타민/미네랄 계열인지 확인"""
    if not items:
        return False
    for item in items:
        il = item.lower()
        matched = (
            il.startswith('비타민') or
            any(vm in il for vm in _VITAMIN_MINERAL_SET)
        )
        if not matched:
            return False
    return True


_MULTIPACK_NAME_KEYWORDS = [
    '멀티팩', '올인원', '종합비타민', '멀티비타민', '토탈팩',
    '올팩', '모두팩', '마더스케어', '올인 원', '멀티 팩',
]

# ─────────────────────────────────────────────────────────────────
# 일반식품 카테고리 분류 맵
# 품목명 우선 매칭 → 미매칭 시 원재료명으로 재시도
# 수정 방법: 키워드 리스트에 단어 추가/제거하면 됨
# ─────────────────────────────────────────────────────────────────
_GENERAL_MAP = [
    ('피부',           ['콜라겐', '히알루론산', '글루타치온', '피부탄력', '피부미용', '피부건강']),
    ('에너지',         ['홍경천', '과라나', '에너지드링크', '에너지', '활력', '부스터']),
    ('단백질',         ['단백질', '프로틴', 'WPI', 'WPC', '크레아틴', '유청']),
    ('효소',           ['발효효소', '식물성효소', '효소', '발효식품']),
    ('관절',           ['글루코사민', 'MSM', '콘드로이친', '관절']),
    ('다이어트',       ['다이어트', '체지방', '체중관리', '슬림', '가르시니아', '카르니틴']),
    ('면역',           ['면역력', '베타글루칸', '이뮨', '면역']),
    ('수면',           ['숙면', '테아닌', '멜라토닌', '수면', 'GABA', '트립토판']),
    ('피로개선',       ['피로회복', '타우린', '피로', '자양강장']),
    ('프로바이오틱스', ['프로바이오틱스', '유산균', 'Lactobacillus']),
    ('구강',           ['잇몸', '자일리톨', '치아', '구강']),
    ('여성건강',       ['갱년기', '엽산', '여성건강', '여성', '이소플라본', '월경']),
    ('남성건강',       ['전립선', '쏘팔메토', '남성건강', '남성']),
    ('모발',           ['탈모', '바이오틴', '케라틴', '모발', '두피']),
    ('혈당',           ['혈당조절', '혈당', '당뇨']),
]


def _extract_bracket_items(text: str) -> list[str]:
    """주된기능성에서 [원료명] 형식의 기능성 원료명 추출"""
    return re.findall(r'\[([^\]]{1,40})\]', text)


def _match_items(items: list[str], keyword_map: list) -> str | None:
    """원료명 리스트를 우선순위 맵과 매칭"""
    for cat, keywords in keyword_map:
        for kw in keywords:
            kw_l = kw.lower()
            for item in items:
                if kw_l in item.lower():
                    return cat
    return None


def _match_text(text: str, keyword_map: list) -> str | None:
    """전체 텍스트에서 키워드 매칭 (fallback)"""
    t_lower = text.lower()
    for cat, keywords in keyword_map:
        for kw in keywords:
            if kw.lower() in t_lower:
                return cat
    return None


def categorize_health_food(fnclty_cn: str, product_name: str = '', raw_material: str = '') -> str:
    text = str(fnclty_cn or '')
    name = str(product_name or '')
    raw = str(raw_material or '')

    # 1. 제품명으로 멀티팩 감지
    name_l = name.lower()
    if any(kw in name_l for kw in _MULTIPACK_NAME_KEYWORDS):
        return '멀티팩'

    # 2. [원료명] 괄호 추출
    items = _extract_bracket_items(text)

    # 3. 비타민/미네랄만 있는 경우 → 즉시 비타민 (항산화 오분류 방지)
    if items and _is_vitamin_mineral_only(items):
        return '비타민'

    # 4. 원료명 기준 우선순위 매칭 (가장 정확)
    if items:
        result = _match_items(items, _HEALTH_MAP)
        if result:
            return result

    # 5. 주된기능성 전체 텍스트 매칭 (구형 포맷 대응)
    result = _match_text(text, _HEALTH_MAP)
    if result:
        return result

    # 6. 제품명 → 원재료 순으로 보조 매칭
    result = _match_text(name, _HEALTH_MAP) or _match_text(raw, _HEALTH_MAP)
    if result:
        return result

    return '기타'


def categorize_general_food(product_name: str, raw_material: str = '') -> str:
    name = str(product_name or '')
    raw  = str(raw_material or '')

    # 규칙 1·2: 맥주효모 또는 (효모+비오틴) → 모발
    if '맥주효모' in name or '맥주효모' in raw:
        return '모발'
    if '효모' in raw and '비오틴' in raw:
        return '모발'

    # 규칙 4: 이노시톨 계열 → 여성건강
    if any(kw in name for kw in ('이노시톨', '콜린미오', '콜린미오이노시톨', '미오이노시톨콜린')):
        return '여성건강'

    # 규칙 5: 파우더 → 기타
    if '파우더' in name:
        return '기타'

    # 규칙 3·6: 오분류 유발 원재료명 마스킹 후 매칭
    # - '효소처리스테비아' → '효소' 키워드 오매칭 방지
    # - '피로인산제일철'   → '피로' 키워드 오매칭 방지
    _MASK = ('효소처리스테비아', '피로인산제일철')
    name_c = name
    raw_c  = raw
    for m in _MASK:
        name_c = name_c.replace(m, '')
        raw_c  = raw_c.replace(m, '')

    result = _match_text(name_c, _GENERAL_MAP)
    if result:
        return result
    result = _match_text(raw_c, _GENERAL_MAP)
    return result if result else '기타'
