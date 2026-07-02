import re

HEALTH_CAT_LIST = [
    '멀티팩', '면역', '비타민', '다이어트', '이너뷰티', '장', '눈', '두뇌',
    '뼈건강', '관절', '간', '위', '피로개선', '수면', '항산화',
    '혈행개선', '혈당', '여성건강', '남성건강', '프로바이오틱스', '오메가3', '홍삼', '단백질',
]

GENERAL_CAT_LIST = [
    '피부', '에너지', '단백질', '효소', '항산화', '혈행개선', '관절', '다이어트', '면역',
    '수면', '피로개선', '프로바이오틱스', '구강', '여성건강', '남성건강',
    '모발', '혈당', '기타',
]

_HEALTH_MAP = [
    ('홍삼',           ['홍삼', '홍삼제품', '진세노사이드', '사포닌', '흑삼', '인삼']),
    ('프로바이오틱스',  ['프로바이오틱스', '프로바이오틱스 제품', '유산균',
                       'Lactobacillus', 'Bifidobacterium']),
    ('오메가3',        ['EPA 및 DHA 함유 유지', 'EPA및DHA함유유지',
                       'EPA 및 DHA함유유지', 'EPA및DHA', '오메가3', '오메가-3']),
    # 필수영양소("~에 필요")가 아닌 기능성 원료 증진 제품만 면역으로 분류
    ('면역',           ['AHCC', '표고버섯균사체', '아가리쿠스', '이뮨',
                       '면역기능 증진', '면역력 증진']),
    ('간',             ['밀크씨슬', '실리마린', '헛개나무', '민들레', '간건강', '간기능', '오르니틴']),
    ('여성건강',       ['대두이소플라본', '감마리놀렌산', '크랜베리 분말', '석류', '갱년기']),
    ('남성건강',       ['쏘팔메토', '전립선']),
    ('다이어트',       ['가르시니아', '공액리놀레산', '카르니틴', '히비스커스',
                       '난소화성말토덱스트린', '체지방감소', '체지방 감소', '체중감소']),
    ('혈당',           ['바나바잎']),
    ('수면',           ['유단백가수분해물', '락티움', 'GABA', '감마아미노부티르산', '감태추출물',
                       '수면의 질', '수면']),
    ('이너뷰티',       ['히알루론산', '저분자콜라겐', '콜라겐펩타이드', '피부탄력', '자외선']),
    ('눈',             ['마리골드꽃추출물', '헤마토코쿠스', '루테인', '지아잔틴', '황반',
                       '빌베리', '마키베리', '눈 건강', '눈의 피로', '건조한 눈']),
    ('두뇌',           ['포스파티딜세린', '은행잎 추출물', '은행잎추출물', '기억력', '인지기능']),
    ('혈행개선',       ['홍국', '혈중중성지방', '혈행개선', '혈액순환', '나토키나제', '루틴']),
    ('항산화',         ['프로폴리스', '코엔자임Q10', '피크노제놀', '항산화', '회화나무']),
    ('관절',           ['엠에스엠', 'MSM', '뮤코다당', '글루코사민', '콘드로이친', 'NAG']),
    ('뼈건강',         ['골다공', '골밀도', '뼈건강', '뼈와 치아', '골형성', '칼슘흡수']),
    ('피로개선',       ['홍경천', '피로개선', '피로회복', '옥타코사놀', '피로감소', '지구력 증진']),
    ('장',             ['차전자피', '프락토올리고당', '자일로올리고당',
                       '알로에 겔', '알로에겔', '배변활동']),
    ('위',             ['스페인감초', '매스틱검', '위건강', '위장건강', '꾸지뽕잎',
                       '마누카꿀', '헬리코박터', '위점막', '위 건강', '감초추출물']),
    ('단백질',         ['크레아틴', 'WPI', 'WPC', '유청단백']),
    ('비타민',         ['비타민A', '비타민 A', '비타민B', '비타민 B',
                       '비타민C', '비타민 C', '비타민D', '비타민 D',
                       '비타민E', '비타민 E', '비타민K', '비타민 K',
                       '나이아신', '엽산', '판토텐산', '비오틴', '비타민',
                       # 미네랄 제품명 매칭
                       '칼륨', '포타슘', '마그네슘', '철분',
                       # 서술형 기능성 텍스트 (괄호 없는 구형 포맷)
                       '에너지 이용에 필요', '에너지 생성에 필요',
                       '신경과 근육 기능', '체내 물과 전해질',
                       '체내 산소운반', '혈액생성에 필요',
                       '정상적인 세포분열', '정상적인 면역기능에 필요']),
]

# 비타민/미네랄 원료명 집합 (이것들만 있으면 무조건 비타민)
_VITAMIN_MINERAL_SET = {
    '비타민', '나이아신', '엽산', '판토텐산', '비오틴',
    '아연', '마그네슘', '칼슘', '셀레늄', '셀렌', '철',
    '구리', '망간', '크롬', '몰리브덴', '요오드', '칼륨', '인',
    '포타슘', '철분',
}


def _is_vitamin_mineral_only(items: list[str]) -> bool:
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

_BONE_TEXT_KWS = ('뼈와 치아', '골밀도', '골형성', '뼈건강', '골다공')
_GENERAL_MASK  = ('효소처리스테비아', '피로인산제일철', '마카다미아')

_GENERAL_MAP = [
    # '피부'는 목록 최상위 — 품목명에 콜라겐/PDRN 등 뷰티 브랜딩이 명시되면 다른 카테고리보다 우선
    ('피부',           ['콜라겐', '히알루론산', '글루타치온', '피부탄력', '피부미용', '피부건강',
                       'PDRN', '뷰티', '글로우', '뮤신', '연어이리추출물', '엘라스틴',
                       '병풀', '마데카']),
    ('에너지',         ['홍경천', '과라나', '에너지드링크', '에너지', '활력',
                       '홍삼', '인삼']),
    ('단백질',         ['단백질', '프로틴', 'WPI', 'WPC', '크레아틴', '유청단백']),
    ('효소',           ['발효효소', '식물성효소', '효소', '발효식품']),
    ('항산화',         ['퀘르세틴', '커큐민', '회화나무', '아세로라', '포도종자',
                       '레스베라트롤', '카테킨', '항산화',
                       '스피루리나', '클로렐라',
                       'NMN', 'NAD', '니코티나마이드리보사이드',
                       '아사이베리', '아로니아', '블루베리', '안토시아닌',
                       '모링가', '루이보스', '히비스커스', '오미자', '헤스페리딘', 'PQQ']),
    ('혈행개선',       ['디오스민', '혈행', '나토키나제', '루틴']),
    ('관절',           ['글루코사민', 'MSM', '콘드로이친', '관절', '보스웰리아']),
    ('다이어트',       ['다이어트', '체지방', '체중관리', '슬림', '가르시니아', '카르니틴', '애사비']),
    ('면역',           ['면역력', '베타글루칸', '이뮨', '면역', 'AHCC', '표고버섯', '도라지', '맥문동']),
    ('수면',           ['숙면', '테아닌', '멜라토닌', '수면', 'GABA', '트립토판']),
    ('피로개선',       ['피로회복', '타우린', '피로', '자양강장', '아슈와간다']),
    ('프로바이오틱스', ['프로바이오틱스', '유산균', 'Lactobacillus', '콤부차']),
    ('구강',           ['잇몸', '자일리톨', '치아', '구강']),
    ('여성건강',       ['갱년기', '엽산', '여성건강', '여성', '이소플라본', '월경']),
    ('남성건강',       ['전립선', '쏘팔메토', '남성건강', '남성', '마카', '블랙마카',
                       '호로파', 'BCAA', '아르기닌', '다미아나', '흑마늘']),
    ('모발',           ['탈모', '바이오틴', '케라틴', '모발', '두피']),
    ('혈당',           ['혈당조절', '혈당', '당뇨', '베르베린', '여주', '돼지감자',
                       '바나바', '뽕나무잎']),
]

# 원재료명 매칭 전용 '피부' 키워드 — '콜라겐'/'히알루론산'/'글루타치온'은 관절·에너지 등
# 다른 목적 제품에도 부원료(코팅제·젤화제 등)로 흔히 섞여 오분류를 유발하므로 원재료 매칭에서는
# 제외. 연어이리추출물(PDRN 원료)도 관절 제품(예: 마디젠)에 재생 보조 성분으로 섞이는 사례가
# 확인되어 원재료 매칭에서 제외 — PDRN 계열은 품목명(_GENERAL_MAP)에 브랜딩이 드러나는 경우가
# 대부분이라 명칭 매칭만으로 충분히 커버됨.
#
# '효소'는 원재료 매칭 자체를 제외 — "효소식품"이 다이어트/기능성 정제류 전반에서
# 덱스트린급으로 흔한 부형제 원료명이라, 품목명에 효소/발효/엔자임 신호가 전혀 없는
# 다이어트 브랜딩 제품(알파CD·컷·핏·리셋·버닝 등)까지 원재료 하나로 효소에 매칭되는
# 문제 확인 (2026-07 v2 데이터 재현). 품목명(_GENERAL_MAP) 매칭만으로 충분히 커버됨.
_GENERAL_MAP_RAW_OVERRIDES = {
    '피부': ['피부탄력', '피부미용', '피부건강'],
    '효소': [],
}

_GENERAL_MAP_RAW = [
    (cat, _GENERAL_MAP_RAW_OVERRIDES.get(cat, kws))
    for cat, kws in _GENERAL_MAP
]

# 일반식품에서 건강기능과 무관한 식품군 제외 (품목명 suffix 기준)
_GENERAL_EXCLUDE_SUFFIXES = ('청', '소스', '양념', '드레싱', '나물')

# 품목명 어디에 있어도 제외 (코드/식별자가 붙는 B2B 원료·반제품·시즈닝류)
# '그레인'은 제외하지 않음 — '파라다이스그레인' 등 실제 효소식품 브랜드 존재
_GENERAL_EXCLUDE_CONTAINS = ('시즈닝', '씨즈닝', '베이스', '반제품', '조미액', '조미분')


def is_general_excluded(product_name: str) -> bool:
    name = str(product_name or '').strip()
    if any(name.endswith(s) for s in _GENERAL_EXCLUDE_SUFFIXES):
        return True
    return any(kw in name for kw in _GENERAL_EXCLUDE_CONTAINS)


def _extract_bracket_items(text: str) -> list[str]:
    """주된기능성에서 [원료명] 형식의 기능성 원료명 추출"""
    return re.findall(r'\[([^\]]{1,40})\]', text)


def _match_items(items: list[str], keyword_map: list) -> str | None:
    for cat, keywords in keyword_map:
        for kw in keywords:
            kw_l = kw.lower()
            for item in items:
                if kw_l in item.lower():
                    return cat
    return None


def _match_text(text: str, keyword_map: list) -> str | None:
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

    # 2-1. 원료가 10개 이상이면 복합 멀티팩으로 처리
    if len(items) >= 10:
        return '멀티팩'

    # 3. 비타민/미네랄만 있는 경우 → 즉시 비타민 (뼈건강 명시 시 뼈건강 우선)
    if items and _is_vitamin_mineral_only(items):
        if any(kw in text for kw in _BONE_TEXT_KWS):
            return '뼈건강'
        return '비타민'

    # 4. 원료명 기준 우선순위 매칭 (가장 정확)
    if items:
        result = _match_items(items, _HEALTH_MAP)
        if result:
            return result

    # 5. 주된기능성 전체 텍스트 매칭 (구형/서술형 포맷 대응)
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

    # 규칙 3: 이노시톨 계열 → 여성건강
    if any(kw in name for kw in ('이노시톨', '콜린미오', '콜린미오이노시톨', '미오이노시톨콜린')):
        return '여성건강'

    # 규칙 4: 파우더/분말 suffix → 기타
    if '파우더' in name or name.endswith('분말'):
        return '기타'

    # 규칙 5: 오분류 유발 원재료명 마스킹 후 매칭
    name_c = name
    raw_c  = raw
    for m in _GENERAL_MASK:
        name_c = name_c.replace(m, '')
        raw_c  = raw_c.replace(m, '')

    result = _match_text(name_c, _GENERAL_MAP)
    if result:
        return result
    result = _match_text(raw_c, _GENERAL_MAP_RAW)
    return result if result else '기타'
