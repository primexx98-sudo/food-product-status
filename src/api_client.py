import time
import xml.etree.ElementTree as ET

import requests

BASE_URL = "https://openapi.foodsafetykorea.go.kr/api"
PAGE_SIZE = 1000

HEALTH_SERVICE  = "C003"   # 건강기능식품 품목제조신고(원재료)
GENERAL_SERVICE = "C002"   # 식품(첨가물)품목제조보고

HEALTH_FIELD_MAP = {
    'PRDLST_REPORT_NO': '품목제조번호',
    'PRMS_DT':          '보고일자',
    'PRDT_SHAP_CD_NM':  '제품형태',
    'LAST_UPDT_DTM':    '최종생성일시',
    'PRDLST_NM':        '품목명',
    'BSSH_NM':          '업소명',
    'PRIMARY_FNCLTY':   '주된기능성',
    'POG_DAYCNT':       '소비기한',
    'NTK_MTHD':         '섭취방법',
    'RAWMTRL_NM':       '원재료',
}

# v2 데이터 기준 포함 제품형태 (건강기능성 관련 카테고리만)
GENERAL_ALLOWED_TYPES = {
    '기타가공품', '과.채가공품', '액상차', '당류가공품', '고형차',
    '캔디류', '과.채주스', '효소식품', '발효식초', '올리브유',
}

GENERAL_FIELD_MAP = {
    'PRDLST_REPORT_NO': '품목제조번호',
    'PRMS_DT':          '보고일자',
    'LCNS_NO':          '인허가번호',
    'PRDLST_NM':        '품목명',
    'PRDLST_DCNM':      '제품형태',
    'BSSH_NM':          '업소명',
    'CHNG_DT':          '최근수정일자',
    'RAWMTRL_NM':       '원재료명',
}


def _build_url(api_key: str, service: str, start: int, end: int, date_filter: str) -> str:
    return f"{BASE_URL}/{api_key}/{service}/xml/{start}/{end}/PRMS_DT={date_filter}"


def _parse_xml(xml_text: str, field_map: dict) -> tuple[int, list[dict]]:
    try:
        root = ET.fromstring(xml_text.encode('utf-8'))
    except ET.ParseError:
        root = ET.fromstring(xml_text.encode('utf-8-sig'))

    # 인증키 오류 감지 (JavaScript alert 응답)
    if xml_text.strip().startswith('<script'):
        raise RuntimeError("API 인증키 오류: 서비스 코드를 사용할 권한이 없습니다.")

    # RESULT 형식 (신규)
    result = root.find('RESULT')
    if result is not None:
        code = result.findtext('CODE', 'INFO-000').strip()
        if not code.startswith('INFO'):
            msg = result.findtext('MSG', '')
            raise RuntimeError(f"API 오류 [{code}]: {msg}")

    # header 형식 (구형 호환)
    header = root.find('header')
    if header is not None:
        code = header.findtext('resultCode', '00').strip()
        if code not in ('00', '0'):
            msg = header.findtext('resultMsg', '')
            raise RuntimeError(f"API 오류 [{code}]: {msg}")

    total = int(root.findtext('total_count', '0') or '0')

    rows = []
    for row_el in root.findall('.//row'):
        item = {col: (row_el.findtext(api_f) or '').strip()
                for api_f, col in field_map.items()}
        rows.append(item)

    return total, rows


def _fetch_all(api_key: str, service: str, field_map: dict, year: int, month: int) -> list[dict]:
    date_filter = f"{year}{month:02d}"
    all_rows: list[dict] = []
    start = 1

    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})

    while True:
        end = start + PAGE_SIZE - 1
        url = _build_url(api_key, service, start, end, date_filter)

        for attempt in range(5):
            try:
                resp = session.get(url, timeout=60)
                resp.raise_for_status()
                break
            except requests.RequestException as e:
                if attempt == 4:
                    raise
                wait = 2 ** attempt  # 1, 2, 4, 8초 지수 백오프
                print(f"  재시도 {attempt + 1}/5 ({wait}초 대기): {e}")
                time.sleep(wait)

        try:
            total, rows = _parse_xml(resp.text, field_map)
        except RuntimeError as e:
            if all_rows:
                print(f"  [경고] {service}: 쿼터 초과, {len(all_rows)}건까지만 수집됨")
                break
            raise

        if not rows:
            break

        all_rows.extend(rows)
        print(f"  {service}: {len(all_rows)}/{total} 수집")

        if len(all_rows) >= total or len(rows) < PAGE_SIZE:
            break

        start = end + 1
        time.sleep(1.0)

    date_col = '보고일자'
    filtered = [r for r in all_rows if r.get(date_col, '').startswith(date_filter)]
    return filtered if filtered else all_rows


def fetch_health_food(api_key: str, year: int, month: int) -> list[dict]:
    return _fetch_all(api_key, HEALTH_SERVICE, HEALTH_FIELD_MAP, year, month)


def fetch_general_food(api_key: str, year: int, month: int) -> list[dict]:
    try:
        rows = _fetch_all(api_key, GENERAL_SERVICE, GENERAL_FIELD_MAP, year, month)
    except RuntimeError as e:
        print(f"  [경고] 일반식품 수집 실패: {e}")
        print("  [안내] 식품안전나라 포털에서 C002 서비스 키 활성화 확인 필요")
        return []

    # C002는 원재료 1개당 행 1개 → 품목제조번호 기준으로 합산
    merged: dict[str, dict] = {}
    for row in rows:
        key = row.get('품목제조번호', '')
        if key not in merged:
            merged[key] = row.copy()
        else:
            existing_raw = merged[key].get('원재료명', '')
            new_raw = row.get('원재료명', '')
            if new_raw and new_raw not in existing_raw:
                merged[key]['원재료명'] = f"{existing_raw}, {new_raw}".strip(', ')

    # v2와 동일한 제품형태 필터 적용
    result = [r for r in merged.values() if r.get('제품형태', '') in GENERAL_ALLOWED_TYPES]
    return result
