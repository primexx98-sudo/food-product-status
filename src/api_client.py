import time
import urllib.parse
import xml.etree.ElementTree as ET

import requests

BASE_URL = "http://openapi.foodsafetykorea.go.kr/api"
PAGE_SIZE = 1000

HEALTH_SERVICE = "건강기능식품 품목제조신고(원재료)"
GENERAL_SERVICE = "식품(첨가물)품목제조보고"

# API 필드명 → 엑셀 컬럼명
HEALTH_FIELD_MAP = {
    'PRDLST_REPORT_NO': '품목제조번호',
    'PRMS_DT':          '보고일자',
    'PRDLST_DCNM':      '제품형태',
    'LSTUPD_DTM':       '최종생성일시',
    'PRDLST_NM':        '품목명',
    'BSSH_NM':          '업소명',
    'FNCLTY_CN':        '주된기능성',
    'POG_DAYCNT':       '소비기한',
    'INTAKE_HINT1':     '섭취방법',
    'RAWMTRL_NM':       '원재료',
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
    encoded = urllib.parse.quote(service, safe='')
    return f"{BASE_URL}/{api_key}/{encoded}/xml/{start}/{end}/PRMS_DT={date_filter}"


def _parse_xml(xml_text: str, field_map: dict) -> tuple[int, list[dict]]:
    try:
        root = ET.fromstring(xml_text.encode('utf-8'))
    except ET.ParseError:
        # 일부 응답이 BOM 포함 시 처리
        root = ET.fromstring(xml_text.encode('utf-8-sig'))

    header = root.find('header')
    if header is not None:
        code = header.findtext('resultCode', '').strip()
        if code not in ('00', '0'):
            msg = header.findtext('resultMsg', '')
            raise RuntimeError(f"API 오류 [{code}]: {msg}")
        total = int(header.findtext('total_count', '0') or '0')
    else:
        total = 0

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

    while True:
        end = start + PAGE_SIZE - 1
        url = _build_url(api_key, service, start, end, date_filter)

        for attempt in range(3):
            try:
                resp = session.get(url, timeout=30)
                resp.raise_for_status()
                break
            except requests.RequestException as e:
                if attempt == 2:
                    raise
                print(f"  재시도 {attempt + 1}/3: {e}")
                time.sleep(2)

        total, rows = _parse_xml(resp.text, field_map)

        if not rows:
            break

        all_rows.extend(rows)
        print(f"  {service}: {len(all_rows)}/{total} 수집")

        if len(all_rows) >= total or len(rows) < PAGE_SIZE:
            break

        start = end + 1
        time.sleep(0.3)

    # 클라이언트 측 월 필터 (이중 확인)
    date_col = '보고일자'
    filtered = [r for r in all_rows if r.get(date_col, '').startswith(date_filter)]
    return filtered if filtered else all_rows


def fetch_health_food(api_key: str, year: int, month: int) -> list[dict]:
    return _fetch_all(api_key, HEALTH_SERVICE, HEALTH_FIELD_MAP, year, month)


def fetch_general_food(api_key: str, year: int, month: int) -> list[dict]:
    return _fetch_all(api_key, GENERAL_SERVICE, GENERAL_FIELD_MAP, year, month)
