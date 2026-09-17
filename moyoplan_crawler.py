from __future__ import annotations

import html
import math
import os
import random
import re
import sys
import time
import warnings
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urljoin, urlparse

# requests가 사용하는 정적 CA 묶음 대신 운영체제의 시스템 인증서 저장소를 사용합니다.
# 회사망/VPN/백신의 HTTPS 검사 인증서가 운영체제에 신뢰 인증서로 등록된 환경에서
# "self-signed certificate in certificate chain" 오류를 해결하는 데 유용합니다.
# 반드시 requests/urllib3를 import하기 전에 실행합니다.
TRUSTSTORE_ENABLED = False
TRUSTSTORE_IMPORT_ERROR: Optional[Exception] = None

CUSTOM_CA_BUNDLE = os.environ.get("MOYOPLAN_CA_BUNDLE", "").strip()
ALLOW_INSECURE_SSL = os.environ.get("MOYOPLAN_INSECURE_SSL", "").strip().lower() in {
    "1",
    "true",
    "yes",
    "on",
}

if not CUSTOM_CA_BUNDLE and not ALLOW_INSECURE_SSL:
    try:
        import truststore
    except ImportError as exc:
        TRUSTSTORE_IMPORT_ERROR = exc
    else:
        truststore.inject_into_ssl()
        TRUSTSTORE_ENABLED = True

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from requests.adapters import HTTPAdapter
from urllib3.exceptions import InsecureRequestWarning
from urllib3.util.retry import Retry


BASE_URL = "https://www.moyoplan.com"
LIST_PATH = "/plans"
PAGE_SIZE = 10

# 전체 수집 모드입니다. 전체 건수를 읽을 수 있으면 예상 페이지 수를 계산하고,
# 읽지 못하면 새 요금제가 더 이상 나오지 않을 때까지 자동으로 다음 페이지를 탐색합니다.
# 다시 일부 페이지만 시험하려면 2처럼 원하는 페이지 수를 지정하세요.
MAX_PAGES: Optional[int] = None

# 전체 건수 문구가 HTML에 없을 때의 안전장치입니다.
# 실제 약 2,200개 요금제라면 PAGE_SIZE=10 기준 약 220페이지이므로 충분한 상한입니다.
AUTO_PAGE_LIMIT = 1000

# 전체 건수를 읽지 못한 자동 탐색 모드에서만 사용하는 종료 안전장치입니다.
# 행의 중복은 제거하지 않으며, 빈 페이지 또는 완전히 같은 페이지가 계속 반환될 때만
# 무한 반복을 막기 위해 수집을 종료합니다.
MAX_CONSECUTIVE_EMPTY_PAGES = 2
MAX_CONSECUTIVE_IDENTICAL_PAGES = 2

# 사이트 부하를 줄이기 위한 요청 간 대기 시간입니다.
REQUEST_DELAY_RANGE: Tuple[float, float] = (0.5, 0.9)
TIMEOUT: Tuple[int, int] = (10, 30)
OUTPUT_PREFIX = "moyoplan_parsed_plans"
OUTPUT_XLSX = "moyoplan_parsed_plans.xlsx"

# 첫 번째 시트는 기존 기본 목록(페이백 포함 체감요금),
# 두 번째 시트는 사용자가 전달한 페이백 미포함 URL의 쿼리를 그대로 사용합니다.
SHEET_CONFIGS: List[Dict[str, Any]] = [
    {
        "label": "페이백 포함",
        "sheet_name": "페이백 포함",
        "csv_suffix": "payback_included",
        "params": {},
    },
    {
        "label": "페이백 미포함",
        "sheet_name": "페이백 미포함",
        "csv_suffix": "payback_excluded",
        "params": {
            "applyPerceivedFee": "false",
            "filters.data.includeUnlimited": "true",
            "filters.data.ranges.0.max": "0",
            "filters.data.ranges.0.min": "0",
            "sort": "RECOMMEND",
        },
    },
]

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/152.0.0.0 Safari/537.36"
)

NETWORK_RE = re.compile(r"(LG\s*U\+|KT|SKT)\s*망\s*(LTE|5G)\b", re.I)
MONTHLY_PRICE_RE = re.compile(
    r"(?:페이백\s*포함(?:하면)?\s*)?월\s*([\d,]+)\s*원"
)
AFTER_PRICE_RE = re.compile(
    r"(\d+)\s*개월\s*이후(?:\s*월)?\s*([\d,]+)\s*원"
)
DETAIL_TITLE_RE = re.compile(
    r"^\[(?P<provider>[^\]]+)]\s*(?P<plan>.*?)\s*\|\s*[\d,]+원\s*\|"
)

DATA_VOLUME = r"[\d,.]+\s*(?:KB|MB|GB|TB)"
DATA_SPEED = r"[\d,.]+\s*(?:Kbps|Mbps|Gbps)"
DATA_PATTERN = (
    rf"(?:"
    rf"월\s+(?:{DATA_VOLUME}|무제한)"
    rf"(?:\s*\+\s*(?:매일|일)\s+{DATA_VOLUME})?"
    rf"(?:\s*\+\s*{DATA_SPEED})?"
    rf"|(?:매일|일)\s+{DATA_VOLUME}(?:\s*\+\s*{DATA_SPEED})?"
    rf"|데이터\s*제공안함(?:\s*\+\s*{DATA_SPEED})?"
    rf"|무제한(?:\s*\+\s*{DATA_SPEED})?"
    rf")"
)
DATA_SEGMENT_RE = re.compile(rf"(?P<data>{DATA_PATTERN})$", re.I)
DATA_FULL_RE = re.compile(rf"^{DATA_PATTERN}$", re.I)

CALL_TOKEN = " 통화 "
SMS_TOKEN = " 문자 "


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def number_only(value: Optional[str]) -> Optional[int]:
    if not value:
        return None
    return int(value.replace(",", ""))


def normalize_call(value: str) -> str:
    value = normalize_spaces(value)
    compact = value.replace(" ", "")

    if "무제한" in value:
        return "무제한"
    if "기본제공" in compact:
        return "기본제공"

    match = re.search(r"[\d,]+", value)
    return f"{match.group(0)}분" if match else value


def normalize_sms(value: str) -> str:
    value = normalize_spaces(value)
    compact = value.replace(" ", "")

    if "무제한" in value:
        return "무제한"
    if "기본제공" in compact:
        return "기본제공"

    match = re.search(r"[\d,]+", value)
    return f"{match.group(0)}건" if match else value


def build_session() -> requests.Session:
    retry = Retry(
        total=4,
        connect=4,
        read=4,
        status=4,
        backoff_factor=1.0,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET"}),
        respect_retry_after_header=True,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=4, pool_maxsize=4)

    session = requests.Session()
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(
        {
            "User-Agent": USER_AGENT,
            "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
            "Accept": (
                "text/html,application/xhtml+xml,application/xml;q=0.9,"
                "image/avif,image/webp,*/*;q=0.8"
            ),
            "Referer": BASE_URL + LIST_PATH,
        }
    )

    if CUSTOM_CA_BUNDLE:
        ca_path = Path(CUSTOM_CA_BUNDLE).expanduser().resolve()
        if not ca_path.is_file():
            raise FileNotFoundError(
                "MOYOPLAN_CA_BUNDLE에 지정한 인증서 파일이 없습니다: "
                f"{ca_path}"
            )
        # PEM 형식의 회사/기관 CA 인증서 묶음을 명시적으로 사용합니다.
        session.verify = str(ca_path)
    elif ALLOW_INSECURE_SSL:
        # 인증서 검증을 끄는 것은 중간자 공격 위험이 있으므로 기본값으로 사용하지 않습니다.
        # 로컬에서 원인 확인을 위한 일시적 테스트에만 사용하세요.
        session.verify = False
        warnings.simplefilter("ignore", InsecureRequestWarning)

    return session


def ssl_mode_description() -> str:
    if CUSTOM_CA_BUNDLE:
        return f"사용자 지정 CA 번들: {CUSTOM_CA_BUNDLE}"
    if ALLOW_INSECURE_SSL:
        return "검증 비활성화(테스트 전용, 보안상 권장하지 않음)"
    if TRUSTSTORE_ENABLED:
        return "운영체제 시스템 인증서 저장소(truststore)"
    return "requests 기본 CA 묶음(certifi)"


def ssl_error_message(url: str, exc: Exception) -> str:
    install_command = f'"{sys.executable}" -m pip install -U truststore requests urllib3 certifi'
    lines = [
        f"HTTPS 인증서 검증에 실패했습니다: {url}",
        "회사망/VPN/백신의 HTTPS 검사 인증서가 Python의 CA 목록에 없을 때 "
        "주로 발생합니다.",
        f"현재 SSL 방식: {ssl_mode_description()}",
    ]

    if not TRUSTSTORE_ENABLED and not CUSTOM_CA_BUNDLE and not ALLOW_INSECURE_SSL:
        lines.extend(
            [
                "현재 실행 중인 Python에 truststore를 설치한 뒤 다시 실행하세요:",
                f"  {install_command}",
            ]
        )
    else:
        lines.extend(
            [
                "그래도 실패하면 회사/기관의 루트 CA 인증서를 PEM 형식으로 받은 뒤",
                "MOYOPLAN_CA_BUNDLE 환경 변수에 그 파일 경로를 지정하세요.",
            ]
        )

    lines.append(f"원래 오류: {exc}")
    return "\n".join(lines)


def fetch_soup(
    session: requests.Session,
    url: str,
    params: Optional[Dict[str, Any]] = None,
) -> BeautifulSoup:
    try:
        response = session.get(url, params=params, timeout=TIMEOUT)
    except requests.exceptions.SSLError as exc:
        raise RuntimeError(ssl_error_message(url, exc)) from exc

    response.raise_for_status()

    # 차단 또는 비정상 응답을 조기에 발견합니다.
    content_type = response.headers.get("Content-Type", "")
    if "text/html" not in content_type.lower():
        raise RuntimeError(f"HTML이 아닌 응답입니다: {content_type} / {response.url}")

    return BeautifulSoup(response.text, "html.parser")


def canonical_plan_url(href: str) -> Optional[str]:
    absolute = urljoin(BASE_URL, href)
    parsed = urlparse(absolute)
    path = parsed.path.rstrip("/")

    if parsed.netloc not in {"moyoplan.com", "www.moyoplan.com"}:
        return None
    if not re.fullmatch(r"/plans/\d+", path):
        return None

    return BASE_URL + path


def extract_plan_cards(soup: BeautifulSoup) -> List[Dict[str, str]]:
    """목록 페이지의 상세 링크와 카드 표시 문구를 보이는 순서대로 수집합니다.

    동일한 상세 URL이나 동일한 카드 문구가 여러 번 나오더라도 제거하지 않습니다.
    즉, 상단 추천 영역과 일반 결과 목록에 같은 요금제가 반복 노출되면 각각 별도
    행으로 저장됩니다.
    """

    cards: List[Dict[str, str]] = []

    for anchor in soup.select("a[href]"):
        if not isinstance(anchor, Tag):
            continue

        detail_url = canonical_plan_url(anchor.get("href", ""))
        if not detail_url:
            continue

        card_text = normalize_spaces(anchor.get_text(" ", strip=True))
        if not card_text:
            continue

        # 요금제 카드가 아닌 단순 링크를 제외합니다.
        if not NETWORK_RE.search(card_text):
            continue
        if "원" not in card_text or "통화" not in card_text or "문자" not in card_text:
            continue

        cards.append(
            {
                "detail_url": detail_url,
                "card_text": card_text,
            }
        )

    return cards

def parse_card_text(card_text: str) -> Dict[str, Any]:
    """카드 표시 문구에서 제공량과 가격 정보를 파싱합니다."""

    text = normalize_spaces(card_text)

    network_matches = list(NETWORK_RE.finditer(text))
    if not network_matches:
        raise ValueError(f"통신망/LTE·5G 구간을 찾지 못했습니다: {text}")

    # 요금제명 안에 '통화' 또는 '문자'가 들어갈 수 있으므로 뒤에서부터 찾습니다.
    network_match = network_matches[-1]
    before_network = text[: network_match.start()].rstrip()
    tail = text[network_match.end() :].strip()

    sms_pos = before_network.rfind(SMS_TOKEN)
    call_pos = before_network.rfind(CALL_TOKEN, 0, sms_pos)
    if sms_pos < 0 or call_pos < 0:
        raise ValueError(f"통화/문자 구간을 찾지 못했습니다: {text}")

    prefix = before_network[:call_pos].strip()
    call_raw = before_network[call_pos + len(CALL_TOKEN) : sms_pos].strip()
    sms_raw = before_network[sms_pos + len(SMS_TOKEN) :].strip()

    data_match = DATA_SEGMENT_RE.search(prefix)
    if data_match:
        data_amount = normalize_spaces(data_match.group("data"))
        plan_part = prefix[: data_match.start()].strip()
    else:
        data_amount = ""
        plan_part = prefix

    # 카드 맨 앞의 평점(예: 4.6)을 제거합니다. 상세 페이지 제목으로 다시 보정됩니다.
    fallback_plan_name = re.sub(
        r"^\d+(?:\.\d+)?\s*", "", plan_part, count=1
    ).strip()

    monthly_match = MONTHLY_PRICE_RE.search(tail)
    if not monthly_match:
        raise ValueError(f"월 요금을 찾지 못했습니다: {text}")

    after_match = AFTER_PRICE_RE.search(tail)

    return {
        "fallback_plan_name": fallback_plan_name,
        "통화제공량": normalize_call(call_raw),
        "문자제공량": normalize_sms(sms_raw),
        "데이터 제공량": data_amount,
        "망정보": normalize_spaces(network_match.group(1)),
        "LTE/5G 구분": network_match.group(2).upper(),
        "월 요금": number_only(monthly_match.group(1)),
        "할인 기간": int(after_match.group(1)) if after_match else None,
        "기간 이후 요금": number_only(after_match.group(2)) if after_match else None,
    }


def parse_detail_summary(
    soup: BeautifulSoup, plan_name: str
) -> Dict[str, str]:
    """상세 본문 앞부분에서 데이터/통화/문자/LTE·5G 값을 보조 추출합니다."""

    body = soup.body
    if body is None or not plan_name:
        return {}

    body_text = normalize_spaces(body.get_text(" ", strip=True))
    result: Dict[str, str] = {}

    # 본문에서 요금제명이 여러 번 나올 수 있으므로 각 위치를 시험합니다.
    for occurrence in re.finditer(re.escape(plan_name), body_text):
        candidate = body_text[occurrence.end() : occurrence.end() + 700].strip()
        network_match = NETWORK_RE.search(candidate)
        if not network_match:
            continue

        before_network = candidate[: network_match.start()].rstrip()
        sms_pos = before_network.rfind(SMS_TOKEN)
        call_pos = before_network.rfind(CALL_TOKEN, 0, sms_pos)
        if sms_pos < 0 or call_pos < 0:
            continue

        data_text = before_network[:call_pos].strip()
        call_raw = before_network[call_pos + len(CALL_TOKEN) : sms_pos].strip()
        sms_raw = before_network[sms_pos + len(SMS_TOKEN) :].strip()

        if not DATA_FULL_RE.fullmatch(data_text):
            continue

        result["데이터 제공량"] = data_text
        result["통화제공량"] = normalize_call(call_raw)
        result["문자제공량"] = normalize_sms(sms_raw)
        result["망정보"] = normalize_spaces(network_match.group(1))
        result["LTE/5G 구분"] = network_match.group(2).upper()
        return result

    # 현재 상세 화면에서는 데이터 표기가 h1인 경우가 있어 추가로 보완합니다.
    for heading in soup.find_all(["h1", "h2", "h3"]):
        heading_text = normalize_spaces(heading.get_text(" ", strip=True))
        if DATA_FULL_RE.fullmatch(heading_text):
            result["데이터 제공량"] = heading_text
            break

    return result


def parse_detail_page(soup: BeautifulSoup) -> Dict[str, str]:
    title = normalize_spaces(soup.title.get_text(" ", strip=True)) if soup.title else ""
    title_match = DETAIL_TITLE_RE.search(title)

    provider = ""
    plan_name = ""
    if title_match:
        provider = normalize_spaces(title_match.group("provider"))
        plan_name = normalize_spaces(title_match.group("plan"))

    # 제목 패턴이 달라질 때를 위한 사업자명 보조 추출입니다.
    if not provider:
        for heading in soup.find_all(["h1", "h2", "h3"]):
            heading_text = normalize_spaces(heading.get_text(" ", strip=True))
            if heading_text.endswith(" 후기"):
                provider = re.sub(r"\s*후기$", "", heading_text).strip()
                break

    detail = {
        "사업자": provider,
        "요금제": plan_name,
    }
    detail.update(parse_detail_summary(soup, plan_name))
    return detail


def detect_total_count(soup: BeautifulSoup) -> Optional[int]:
    """목록 페이지에 표시된 전체 요금제 수를 가능한 여러 형식에서 찾습니다.

    모요가 A/B 테스트나 렌더링 방식을 변경하면 ``2,196개의 결과``가
    ``2,196 개의 결과``처럼 분리되거나, 화면 텍스트 대신 JSON 데이터에만
    들어갈 수 있습니다. 전체 건수를 못 찾더라도 호출부에서 자동 페이지 탐색으로
    계속 수집하므로 이 함수의 실패는 치명적인 오류가 아닙니다.
    """

    visible_text = normalize_spaces(soup.get_text(" ", strip=True))
    visible_patterns = [
        r"([\d,]+)\s*개의\s*(?:결과|요금제)",
        r"(?:전체|총)\s*([\d,]+)\s*(?:개의?\s*)?(?:결과|요금제|건)",
        r"(?:결과|요금제)\s*(?:전체|총)?\s*([\d,]+)\s*(?:개|건)",
    ]

    for pattern in visible_patterns:
        match = re.search(pattern, visible_text, flags=re.I)
        if match:
            return number_only(match.group(1))

    # Next.js/React의 직렬화 데이터에만 전체 건수가 들어 있는 경우를 보완합니다.
    raw_html = str(soup)
    json_patterns = [
        r'(?:totalCount|totalElements|totalPlanCount)(?:\\?["\'])?\s*:\s*(?:\\?["\'])?([\d,]+)',
        r'(?:totalCount|totalElements|totalPlanCount)\s*=\s*(?:\\?["\'])?([\d,]+)',
    ]
    for pattern in json_patterns:
        match = re.search(pattern, raw_html, flags=re.I)
        if match:
            return number_only(match.group(1))

    return None


def polite_sleep() -> None:
    time.sleep(random.uniform(*REQUEST_DELAY_RANGE))


def empty_result_frame(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    columns = [
        "사업자",
        "요금제",
        "통화제공량",
        "문자제공량",
        "데이터 제공량",
        "망정보",
        "LTE/5G 구분",
        "월 요금",
        "할인 기간",
        "기간 이후 요금",
    ]

    df = pd.DataFrame(rows, columns=columns)
    for column in ["월 요금", "할인 기간", "기간 이후 요금"]:
        df[column] = pd.array(df[column], dtype="Int64")
    return df


def format_excel_sheet(worksheet, dataframe: pd.DataFrame) -> None:
    """필터, 틀 고정, 열 너비와 숫자 형식을 적용합니다."""

    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.sheet_view.showGridLines = False

    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    preferred_widths = {
        "사업자": 18,
        "요금제": 38,
        "통화제공량": 14,
        "문자제공량": 14,
        "데이터 제공량": 30,
        "망정보": 12,
        "LTE/5G 구분": 13,
        "월 요금": 14,
        "할인 기간": 12,
        "기간 이후 요금": 16,
    }

    for column_index, column_name in enumerate(dataframe.columns, start=1):
        letter = get_column_letter(column_index)
        worksheet.column_dimensions[letter].width = preferred_widths.get(
            column_name, 15
        )

        for cell in worksheet[letter][1:]:
            cell.alignment = Alignment(vertical="center")
            if column_name in {"월 요금", "할인 기간", "기간 이후 요금"}:
                cell.number_format = "#,##0"


DEFAULT_REFERENCE_ROWS = [
    ("월 125GB + 5Mbps", "RS"),
    ("월 95GB + 3Mbps", "RS"),
    ("월 150GB + 5Mbps", "RS"),
    ("월 100GB + 5Mbps", "RS"),
    ("월 11GB + 매일 2GB + 3Mbps", "RS"),
    ("매일 5GB + 5Mbps", "RS"),
    ("월 15GB + 3Mbps", "RS"),
    ("월 180GB + 10Mbps", "RS"),
    ("월 31GB + 1Mbps", "RS"),
    ("월 50GB + 1Mbps", "RS"),
    ("월 80GB + 1Mbps", "RS"),
    ("월 7GB + 1Mbps", "RS"),
    ("월 7GB + 3Mbps", "RS"),
]


def _ensure_reference_sheet(workbook) -> None:
    """기존 참조 시트를 보존하고, 비어 있을 때만 최초 RS 기준값을 복원합니다."""
    if "참조" not in workbook.sheetnames:
        ws = workbook.create_sheet("참조")
    else:
        ws = workbook["참조"]

    # 사용자가 입력해 둔 B:C 데이터가 하나라도 있으면 절대 덮어쓰지 않습니다.
    has_user_data = any(
        ws.cell(r, 2).value not in (None, "") or ws.cell(r, 3).value not in (None, "")
        for r in range(2, ws.max_row + 1)
    )
    if has_user_data:
        return

    # 참조 시트가 비어 있는 경우에만 원래 기준 데이터를 복원합니다.
    ws.cell(1, 2).value = "모태"
    ws.cell(1, 3).value = "구분"
    for r, (data_amount, fee_type) in enumerate(DEFAULT_REFERENCE_ROWS, start=2):
        ws.cell(r, 2).value = data_amount
        ws.cell(r, 3).value = fee_type


def _reference_map(workbook) -> Dict[Tuple[str, str, str], str]:
    """참조 시트 B:E를 읽어 (데이터, 음성, 문자) -> 요금구분 매핑을 만듭니다."""
    if "참조" not in workbook.sheetnames:
        return {}
    ws = workbook["참조"]
    result: Dict[Tuple[str, str, str], str] = {}
    for row in range(2, ws.max_row + 1):
        data_amount = ws.cell(row, 2).value   # B: 모태/데이터
        fee_type = ws.cell(row, 3).value      # C: 구분
        voice = ws.cell(row, 4).value         # D: 음성
        sms = ws.cell(row, 5).value           # E: 문자
        if (
            data_amount not in (None, "")
            and fee_type not in (None, "")
            and voice not in (None, "")
            and sms not in (None, "")
        ):
            key = (
                str(data_amount).strip(),
                str(voice).strip(),
                str(sms).strip(),
            )
            result[key] = str(fee_type).strip()
    return result


def _sheet_rows_as_dicts(worksheet) -> List[Dict[str, Any]]:
    """전일 비교용으로 기존 Sheet2 값을 읽습니다."""
    if worksheet is None or worksheet.max_row < 2:
        return []
    headers = [worksheet.cell(1, c).value for c in range(1, worksheet.max_column + 1)]
    rows: List[Dict[str, Any]] = []
    for r in range(2, worksheet.max_row + 1):
        values = [worksheet.cell(r, c).value for c in range(1, worksheet.max_column + 1)]
        if not any(v not in (None, "") for v in values[:10]):
            continue
        rows.append({headers[i]: values[i] for i in range(len(headers)) if headers[i]})
    return rows


def _clear_sheet_values(worksheet) -> None:
    for row in worksheet.iter_rows():
        for cell in row:
            cell.value = None


def _write_raw_sheet(worksheet, dataframe: pd.DataFrame) -> None:
    """마스터 Sheet1/2의 양식/수식은 유지하고 A:J 크롤링 데이터만 갱신합니다.

    K열은 마스터의 K2 일반 수식을 행별로 번역해서 복사합니다.
    ArrayFormula로 변환하지 않으므로 Excel의 XLOOKUP 등 최신 수식을 훼손하지 않습니다.
    """
    from copy import copy
    from openpyxl.formula.translate import Translator

    # 마스터 K2 수식이 source of truth
    k2_formula = worksheet["K2"].value
    if not (isinstance(k2_formula, str) and k2_formula.startswith("=")):
        raise RuntimeError(
            f"[{worksheet.title}] K2가 일반 Excel 수식이 아닙니다. "
            "GitHub의 마스터 Excel K2 수식을 확인하세요."
        )

    # 2행의 서식을 신규 데이터 행에 복사
    style_template = {}
    if worksheet.max_row >= 2:
        for c in range(1, 12):
            cell = worksheet.cell(2, c)
            style_template[c] = {
                "font": copy(cell.font),
                "fill": copy(cell.fill),
                "border": copy(cell.border),
                "alignment": copy(cell.alignment),
                "number_format": cell.number_format,
                "protection": copy(cell.protection),
            }

    # 기존 데이터/수식 값만 삭제. 시트, 열너비, 조건부서식 등은 유지.
    old_last_row = max(worksheet.max_row, 2)
    for r in range(2, old_last_row + 1):
        for c in range(1, 12):
            worksheet.cell(r, c).value = None

    # 오늘 크롤링 A:J 작성 + K 수식 복사
    for r_idx, row in enumerate(dataframe.itertuples(index=False, name=None), start=2):
        for c_idx, value in enumerate(row, start=1):
            if pd.isna(value):
                value = None
            worksheet.cell(r_idx, c_idx).value = value

        worksheet.cell(r_idx, 11).value = Translator(
            k2_formula, origin="K2"
        ).translate_formula(f"K{r_idx}")

        if style_template:
            for c in range(1, 12):
                target = worksheet.cell(r_idx, c)
                st = style_template[c]
                target.font = copy(st["font"])
                target.fill = copy(st["fill"])
                target.border = copy(st["border"])
                target.alignment = copy(st["alignment"])
                target.number_format = st["number_format"]
                target.protection = copy(st["protection"])

    last_row = max(2, len(dataframe) + 1)
    worksheet.auto_filter.ref = f"A1:K{last_row}"
    if not worksheet.freeze_panes:
        worksheet.freeze_panes = "A2"



def _classify_rows(
    rows: List[Dict[str, Any]],
    ref_map: Dict[Tuple[str, str, str], str],
) -> List[Dict[str, Any]]:
    """엑셀 K열과 동일하게 데이터+음성+문자 3개 조건으로 요금구분을 판정합니다."""
    classified = []
    for row in rows:
        item = dict(row)
        key = (
            str(item.get("데이터 제공량") or "").strip(),
            str(item.get("통화제공량") or "").strip(),
            str(item.get("문자제공량") or "").strip(),
        )
        item["요금구분"] = ref_map.get(key, "")
        classified.append(item)
    return classified


def _write_guide_sheet(worksheet, current_rows: List[Dict[str, Any]], ref_map: Dict[Tuple[str, str, str], str]) -> Tuple[int, int]:
    """I3/I19에 적어둔 기준으로 RS/RM 가이드 위반을 자동 작성합니다."""
    from copy import copy
    from openpyxl.styles import Alignment, Font, PatternFill

    classified = _classify_rows(current_rows, ref_map)

    def n(value):
        try:
            return int(value) if value not in (None, "") else None
        except (TypeError, ValueError):
            return None

    rs = []
    rm = []
    for row in classified:
        if str(row.get("망정보") or "").strip() != "LG U+":
            continue
        fee_type = str(row.get("요금구분") or "").strip().upper()
        price = n(row.get("월 요금"))
        months = n(row.get("할인 기간"))
        if fee_type == "RS" and (price == 0 or (months is not None and months <= 6)):
            rs.append(row)
        elif fee_type == "RM" and (price == 0 or (months is not None and months <= 5)):
            rm.append(row)
        # 빈칸은 RS/RM 어느 쪽으로도 추정하지 않습니다.

    # 사용자가 작성한 I3/I19 설명은 보존합니다.
    i3 = worksheet["I3"].value
    i19 = worksheet["I19"].value
    _clear_sheet_values(worksheet)

    title_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    header_fill = PatternFill(fill_type="solid", fgColor="E2F0D9")

    def section(start_row: int, title: str, items: List[Dict[str, Any]], note: Optional[str]):
        worksheet.cell(start_row, 2).value = title
        worksheet.cell(start_row, 2).font = Font(bold=True, size=12)
        worksheet.cell(start_row, 2).fill = title_fill
        headers = ["순번", "사업자", "요금제 명", "월 요금", "할인 기간", "기간 이후 요금"]
        for c, h in enumerate(headers, 2):
            cell = worksheet.cell(start_row + 1, c)
            cell.value = h
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        if note:
            worksheet.cell(start_row + 1, 9).value = note
        for idx, item in enumerate(items, 1):
            rr = start_row + 1 + idx
            values = [idx, item.get("사업자"), item.get("요금제"), item.get("월 요금"), item.get("할인 기간"), item.get("기간 이후 요금")]
            for c, value in enumerate(values, 2):
                worksheet.cell(rr, c).value = value
            for c in (5, 6, 7):
                worksheet.cell(rr, c).number_format = "#,##0"
        return start_row + 2 + len(items)

    next_row = section(2, "■RS요금제 가이드 위반", rs, i3)
    rm_start = max(next_row + 2, 17)
    section(rm_start, "■RM요금제 가이드 위반", rm, i19)

    for col, width in {"B":8, "C":18, "D":48, "E":14, "F":12, "G":16, "I":80}.items():
        worksheet.column_dimensions[col].width = width
    worksheet.sheet_view.showGridLines = False
    return len(rs), len(rm)


def _identity(row: Dict[str, Any]) -> Tuple[str, str, str, str]:
    return (
        str(row.get("사업자") or "").strip(),
        str(row.get("요금제") or "").strip(),
        str(row.get("망정보") or "").strip(),
        str(row.get("LTE/5G 구분") or "").strip(),
    )


def _write_change_sheet(worksheet, previous_rows: List[Dict[str, Any]], current_rows: List[Dict[str, Any]]) -> Dict[str, int]:
    """페이백 미포함 기준 전일 대비 신규/삭제/가격·기간 변동을 작성합니다."""
    from collections import defaultdict
    from openpyxl.styles import Alignment, Font, PatternFill

    prev_map: Dict[Tuple[str, str, str, str], List[Dict[str, Any]]] = defaultdict(list)
    curr_map: Dict[Tuple[str, str, str, str], List[Dict[str, Any]]] = defaultdict(list)
    for row in previous_rows:
        prev_map[_identity(row)].append(row)
    for row in current_rows:
        curr_map[_identity(row)].append(row)

    changes: List[List[Any]] = []
    stats = {"신규": 0, "삭제": 0, "월요금 인상": 0, "월요금 인하": 0, "기타 변경": 0}
    all_keys = set(prev_map) | set(curr_map)

    def signature(row):
        return (row.get("월 요금"), row.get("할인 기간"), row.get("기간 이후 요금"))

    for key in sorted(all_keys):
        old_list = sorted(prev_map.get(key, []), key=lambda x: str(signature(x)))
        new_list = sorted(curr_map.get(key, []), key=lambda x: str(signature(x)))
        common = min(len(old_list), len(new_list))
        for i in range(common):
            old, new = old_list[i], new_list[i]
            if signature(old) == signature(new):
                continue
            old_price, new_price = old.get("월 요금"), new.get("월 요금")
            if isinstance(old_price, (int, float)) and isinstance(new_price, (int, float)) and new_price > old_price:
                change_type = "월요금 인상"
            elif isinstance(old_price, (int, float)) and isinstance(new_price, (int, float)) and new_price < old_price:
                change_type = "월요금 인하"
            else:
                change_type = "기타 변경"
            stats[change_type] += 1
            changes.append([change_type, *key, old.get("월 요금"), new.get("월 요금"), old.get("할인 기간"), new.get("할인 기간"), old.get("기간 이후 요금"), new.get("기간 이후 요금")])
        for new in new_list[common:]:
            stats["신규"] += 1
            changes.append(["신규", *key, None, new.get("월 요금"), None, new.get("할인 기간"), None, new.get("기간 이후 요금")])
        for old in old_list[common:]:
            stats["삭제"] += 1
            changes.append(["삭제", *key, old.get("월 요금"), None, old.get("할인 기간"), None, old.get("기간 이후 요금"), None])

    _clear_sheet_values(worksheet)
    headers = ["변동구분", "사업자", "요금제", "망정보", "LTE/5G", "전일 월요금", "금일 월요금", "전일 할인기간", "금일 할인기간", "전일 이후요금", "금일 이후요금"]
    for c, h in enumerate(headers, 1):
        worksheet.cell(1, c).value = h
    for r, values in enumerate(changes, 2):
        for c, value in enumerate(values, 1):
            worksheet.cell(r, c).value = value
    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:K{max(worksheet.max_row,1)}"
    worksheet.sheet_view.showGridLines = False
    widths = [14,18,48,12,10,14,14,14,14,16,16]
    from openpyxl.utils import get_column_letter
    for i,w in enumerate(widths,1): worksheet.column_dimensions[get_column_letter(i)].width=w
    for row in worksheet.iter_rows(min_row=2):
        for c in (6,7,8,9,10,11): row[c-1].number_format = "#,##0"
    return stats


def _write_summary_sheet(worksheet, current_rows: List[Dict[str, Any]], ref_map: Dict[Tuple[str, str, str], str], rs_count: int, rm_count: int, change_stats: Dict[str, int]) -> None:
    from collections import Counter
    from datetime import datetime
    from openpyxl.styles import Font, PatternFill

    classified = _classify_rows(current_rows, ref_map)
    network = Counter(str(r.get("망정보") or "미확인") for r in classified)
    fee = Counter(str(r.get("요금구분") or "미분류") for r in classified)
    generation = Counter(str(r.get("LTE/5G 구분") or "미확인") for r in classified)

    _clear_sheet_values(worksheet)
    rows = [
        ["모요 요금제 요약 및 통계", None],
        ["생성시각", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["기준", "페이백 미포함"],
        [None, None],
        ["전체 저장 행", len(classified)],
        ["RS", fee.get("RS",0)],
        ["RM", fee.get("RM",0)],
        ["미분류", fee.get("미분류",0)],
        ["RS 가이드 위반", rs_count],
        ["RM 가이드 위반", rm_count],
        [None, None],
        ["LG U+망", network.get("LG U+",0)],
        ["KT망", network.get("KT",0)],
        ["SKT망", network.get("SKT",0)],
        ["망정보 미확인", network.get("미확인",0)],
        ["LTE", generation.get("LTE",0)],
        ["5G", generation.get("5G",0)],
        [None, None],
        ["전일 대비 신규", change_stats.get("신규",0)],
        ["전일 대비 삭제", change_stats.get("삭제",0)],
        ["월요금 인상", change_stats.get("월요금 인상",0)],
        ["월요금 인하", change_stats.get("월요금 인하",0)],
        ["기타 변경", change_stats.get("기타 변경",0)],
    ]
    for r, values in enumerate(rows,1):
        for c,value in enumerate(values,1): worksheet.cell(r,c).value=value
    worksheet["A1"].font = Font(bold=True, size=14)
    worksheet["A1"].fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    for r in range(5, worksheet.max_row+1):
        worksheet.cell(r,1).font = Font(bold=True)
    worksheet.column_dimensions["A"].width=24
    worksheet.column_dimensions["B"].width=24
    worksheet.sheet_view.showGridLines=False


def _mail_html_table(title: str, headers: List[str], rows: List[List[Any]]) -> str:
    def esc(v: Any) -> str:
        if v is None:
            return ""
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            if isinstance(v, float) and v.is_integer():
                v = int(v)
            return html.escape(f"{v:,}")
        return html.escape(str(v))

    parts = [
        f'<h3 style="margin:24px 0 8px;color:#1F4E78;">{html.escape(title)}</h3>',
        '<table style="border-collapse:collapse;width:100%;font-family:Arial,Malgun Gothic,sans-serif;font-size:13px;">',
        '<tr>',
    ]
    for h in headers:
        parts.append(f'<th style="border:1px solid #cfd8e3;background:#D9EAF7;padding:7px;text-align:center;">{html.escape(h)}</th>')
    parts.append('</tr>')
    if not rows:
        parts.append(f'<tr><td colspan="{len(headers)}" style="border:1px solid #cfd8e3;padding:8px;text-align:center;color:#666;">해당 없음</td></tr>')
    else:
        for row in rows:
            parts.append('<tr>')
            for v in row:
                align = 'right' if isinstance(v, (int, float)) and not isinstance(v, bool) else 'left'
                parts.append(f'<td style="border:1px solid #cfd8e3;padding:7px;text-align:{align};vertical-align:top;">{esc(v)}</td>')
            parts.append('</tr>')
    parts.append('</table>')
    return ''.join(parts)


def _build_moyo_price_mail_rows(
    workbook,
    current_rows: List[Dict[str, Any]],
    ref_map: Dict[Tuple[str, str, str], str],
) -> Tuple[List[str], List[List[Any]]]:
    """모요 판가 시트 수식과 같은 조건으로 당일 메일용 판가표를 계산합니다."""
    if "모요 판가" not in workbook.sheetnames or "참조" not in workbook.sheetnames:
        return [], []

    price_ws = workbook["모요 판가"]
    ref_ws = workbook["참조"]

    category_map: Dict[str, Tuple[str, str, str]] = {}
    for r in range(2, ref_ws.max_row + 1):
        key = ref_ws.cell(r, 6).value
        data = ref_ws.cell(r, 2).value
        voice = ref_ws.cell(r, 4).value
        sms = ref_ws.cell(r, 5).value
        if key not in (None, "") and data not in (None, ""):
            category_map[str(key).strip()] = (
                str(data).strip(),
                str(voice or "").strip(),
                str(sms or "").strip(),
            )

    business_map: Dict[str, str] = {}
    for r in range(2, ref_ws.max_row + 1):
        canonical = ref_ws.cell(r, 11).value
        moyo = ref_ws.cell(r, 12).value
        fallback = ref_ws.cell(r, 13).value
        if canonical not in (None, ""):
            mapped = moyo if moyo not in (None, "") else fallback
            if mapped not in (None, ""):
                business_map[str(canonical).strip()] = str(mapped).strip()

    classified = _classify_rows(current_rows, ref_map)
    category_cols = [4, 7, 10, 13, 16, 19, 22, 25]
    headers = ["사업자", "망"]
    categories = []

    for col in category_cols:
        title = price_ws.cell(3, col).value
        if title in (None, ""):
            continue
        title_text = str(title).strip()
        criteria = category_map.get(title_text.replace("GB", "G"))
        if not criteria:
            continue
        categories.append((title_text, criteria))
        headers.extend([
            f"{title_text} 월 요금",
            f"{title_text} 할인 기간",
            f"{title_text} 기간 이후 요금",
        ])

    result_rows: List[List[Any]] = []
    previous_business = ""

    for r in range(5, price_ws.max_row + 1):
        raw_business = price_ws.cell(r, 2).value
        network = price_ws.cell(r, 3).value
        if network in (None, ""):
            continue

        if isinstance(raw_business, str) and raw_business.startswith("="):
            business_label = previous_business
        elif raw_business not in (None, ""):
            business_label = str(raw_business).strip()
            previous_business = business_label
        else:
            business_label = previous_business

        if not business_label:
            continue

        target_business = business_map.get(business_label, business_label)
        network_text = str(network).strip()
        target_network = "SKT" if network_text in {"SK", "SKT"} else network_text
        output: List[Any] = [business_label, network_text]

        for _, (data_amount, voice, sms) in categories:
            candidates = [
                item for item in classified
                if str(item.get("사업자") or "").strip() == target_business
                and str(item.get("통화제공량") or "").strip() == voice
                and str(item.get("문자제공량") or "").strip() == sms
                and str(item.get("데이터 제공량") or "").strip() == data_amount
                and str(item.get("망정보") or "").strip() == target_network
                and str(item.get("요금구분") or "").strip().upper() == "RS"
            ]

            if candidates:
                def sort_key(item: Dict[str, Any]):
                    price = item.get("월 요금")
                    months = item.get("할인 기간")
                    after = item.get("기간 이후 요금")
                    price_key = price if isinstance(price, (int, float)) else 10**18
                    months_key = months if isinstance(months, (int, float)) else 10**18
                    after_key = after if isinstance(after, (int, float)) else 0
                    return (price_key, -months_key, after_key)

                best = sorted(candidates, key=sort_key)[0]
                output.extend([
                    best.get("월 요금"),
                    best.get("할인 기간"),
                    best.get("기간 이후 요금"),
                ])
            else:
                output.extend(["", "", ""])

        result_rows.append(output)

    return headers, result_rows


def generate_mail_body(
    workbook,
    current_rows: List[Dict[str, Any]],
    ref_map: Dict[Tuple[str, str, str], str],
    rs_count: int,
    rm_count: int,
    change_stats: Dict[str, int],
) -> None:
    """메일: 크롤링 요약 → 모요 판가 → 가이드 위반 순서."""
    from collections import Counter
    from datetime import datetime

    classified = _classify_rows(current_rows, ref_map)
    network = Counter(str(r.get("망정보") or "미확인") for r in classified)
    fee = Counter(str(r.get("요금구분") or "미분류") for r in classified)
    generation = Counter(str(r.get("LTE/5G 구분") or "미확인") for r in classified)

    def n(value: Any) -> Optional[int]:
        try:
            return int(value) if value not in (None, "") else None
        except (TypeError, ValueError):
            return None

    rs_items, rm_items = [], []
    for row in classified:
        if str(row.get("망정보") or "").strip() != "LG U+":
            continue
        fee_type = str(row.get("요금구분") or "").strip().upper()
        price = n(row.get("월 요금"))
        months = n(row.get("할인 기간"))
        item = [row.get("사업자"), row.get("요금제"), row.get("월 요금"),
                row.get("할인 기간"), row.get("기간 이후 요금")]
        if fee_type == "RS" and (price == 0 or (months is not None and months <= 6)):
            rs_items.append(item)
        elif fee_type == "RM" and (price == 0 or (months is not None and months <= 5)):
            rm_items.append(item)

    summary_rows = [
        ["생성시각", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["기준", "페이백 미포함"],
        ["전체 저장 행", len(classified)],
        ["RS", fee.get("RS", 0)],
        ["RM", fee.get("RM", 0)],
        ["미분류", fee.get("미분류", 0)],
        ["RS 가이드 위반", rs_count],
        ["RM 가이드 위반", rm_count],
        ["LG U+망", network.get("LG U+", 0)],
        ["KT망", network.get("KT", 0)],
        ["SKT망", network.get("SKT", 0)],
        ["망정보 미확인", network.get("미확인", 0)],
        ["LTE", generation.get("LTE", 0)],
        ["5G", generation.get("5G", 0)],
        ["전일 대비 신규", change_stats.get("신규", 0)],
        ["전일 대비 삭제", change_stats.get("삭제", 0)],
        ["월요금 인상", change_stats.get("월요금 인상", 0)],
        ["월요금 인하", change_stats.get("월요금 인하", 0)],
        ["기타 변경", change_stats.get("기타 변경", 0)],
    ]

    moyo_headers, moyo_rows = _build_moyo_price_mail_rows(
        workbook, current_rows, ref_map
    )
    guide_headers = ["사업자", "요금제 명", "월 요금", "할인 기간", "기간 이후 요금"]

    body = [
        '<!doctype html><html><body style="font-family:Arial,Malgun Gothic,sans-serif;color:#222;line-height:1.5;">',
        '<div style="max-width:1400px;margin:0 auto;">',
        '<h2 style="color:#17365D;margin-bottom:4px;">모요 요금제 Daily Report</h2>',
        '<p style="color:#666;margin-top:0;">자동 크롤링 및 가이드 점검 결과입니다.</p>',
        _mail_html_table("크롤링 요약", ["항목", "값"], summary_rows),
        _mail_html_table("모요 판가", moyo_headers, moyo_rows)
            if moyo_headers else _mail_html_table("모요 판가", ["결과"], []),
        _mail_html_table("RS 요금제 가이드 위반", guide_headers, rs_items),
        _mail_html_table("RM 요금제 가이드 위반", guide_headers, rm_items),
        '<p style="margin-top:28px;color:#666;">상세 내용은 첨부된 <b>moyoplan_parsed_plans.xlsx</b> 파일을 확인해 주세요.</p>',
        '<p style="color:#888;font-size:12px;">이 메일은 자동으로 생성되었습니다.</p>',
        '</div></body></html>',
    ]
    Path("mail_body.html").write_text("".join(body), encoding="utf-8")
    print("메일 본문 HTML: mail_body.html")



def save_outputs(
    results: List[Tuple[Dict[str, Any], List[Dict[str, Any]]]],
    errors: List[Dict[str, Any]],
) -> None:
    """기존 엑셀의 참조/가이드 양식을 보존하면서 6개 시트를 자동 갱신합니다."""
    from openpyxl import Workbook, load_workbook

    dataframes: List[Tuple[Dict[str, Any], pd.DataFrame]] = []
    result_by_sheet: Dict[str, List[Dict[str, Any]]] = {}
    for config, rows in results:
        df = empty_result_frame(rows)
        dataframes.append((config, df))
        result_by_sheet[config["sheet_name"]] = rows
        csv_path = f"{OUTPUT_PREFIX}_{config['csv_suffix']}.csv"
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        print(f"CSV [{config['label']}]: {csv_path}")

    output_path = Path(OUTPUT_XLSX)
    previous_rows: List[Dict[str, Any]] = []

    if not output_path.exists():
        raise FileNotFoundError(
            f"마스터 Excel이 없습니다: {OUTPUT_XLSX}. "
            "GitHub 저장소 루트에 마스터 파일을 먼저 올려주세요."
        )

    workbook = load_workbook(output_path)
    required_sheets = [
        "페이백 포함", "페이백 미포함", "가이드 위반",
        "전일 대비 변동", "요약 및 통계", "모요 판가", "참조",
    ]
    missing = [s for s in required_sheets if s not in workbook.sheetnames]
    if missing:
        raise RuntimeError("마스터 Excel 누락 시트: " + ", ".join(missing))

    # 오늘 데이터로 덮기 전에 Sheet2를 전일 데이터로 확보
    previous_rows = _sheet_rows_as_dicts(workbook["페이백 미포함"])

    # 참조 시트는 사용자가 관리합니다. Python은 읽기만 합니다.
    ref_map = _reference_map(workbook)

    for config, df in dataframes:
        _write_raw_sheet(workbook[config["sheet_name"]], df)

    current_excluded = result_by_sheet.get("페이백 미포함", [])

    # 모요 판가 시트는 값/수식/서식을 전혀 수정하지 않습니다.
    rs_count, rm_count = _write_guide_sheet(workbook["가이드 위반"], current_excluded, ref_map)
    change_stats = _write_change_sheet(workbook["전일 대비 변동"], previous_rows, current_excluded)
    _write_summary_sheet(workbook["요약 및 통계"], current_excluded, ref_map, rs_count, rm_count, change_stats)

    # 마스터의 기존 시트 순서/수식/서식을 유지한 채 저장합니다.
    workbook.save(output_path)

    if errors:
        pd.DataFrame(errors).to_csv(
            f"{OUTPUT_PREFIX}_errors.csv", index=False, encoding="utf-8-sig"
        )

    print(f"\nXLSX: {OUTPUT_XLSX}")
    for index, (config, df) in enumerate(dataframes, start=1):
        print(f"Sheet{index} [{config['sheet_name']}]: {len(df):,}건")
    print(f"가이드 위반: RS {rs_count:,}건 / RM {rm_count:,}건")
    print(f"전일 대비 변동: {sum(change_stats.values()):,}건")
    print("참조 시트: 기존 내용 보존")
    if errors:
        print(f"파싱/요청 오류: {len(errors):,}건 - {OUTPUT_PREFIX}_errors.csv 확인")


def make_page_params(config: Dict[str, Any], page_index: int) -> Dict[str, Any]:
    params = dict(config.get("params", {}))
    params["page.page"] = page_index
    params["page.size"] = PAGE_SIZE
    return params


def crawl_dataset(
    session: requests.Session,
    config: Dict[str, Any],
    detail_cache: Dict[str, Dict[str, str]],
    all_errors: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    label = config["label"]
    rows: List[Dict[str, Any]] = []

    print("\n" + "=" * 70)
    print(f"수집 시작: {label}")
    print("=" * 70)

    first_soup = fetch_soup(
        session,
        BASE_URL + LIST_PATH,
        params=make_page_params(config, 0),
    )

    total_count: Optional[int] = None
    page_count: Optional[int] = None

    if MAX_PAGES is not None:
        page_count = MAX_PAGES
        print(f"지정 모드: {page_count}페이지 수집")
    else:
        total_count = detect_total_count(first_soup)
        if total_count is not None:
            page_count = math.ceil(total_count / PAGE_SIZE)
            print(f"전체 {total_count:,}건 / 예상 {page_count:,}페이지")
        else:
            print(
                "[안내] 전체 건수 문구를 찾지 못했습니다. "
                "빈 페이지 또는 동일 페이지의 반복을 감지할 때까지 "
                "다음 페이지를 자동 탐색합니다."
            )

    page_index = 0
    consecutive_empty_pages = 0
    consecutive_identical_pages = 0
    previous_page_signature: Optional[Tuple[Tuple[str, str], ...]] = None

    while True:
        if page_count is not None and page_index >= page_count:
            break

        if page_index >= AUTO_PAGE_LIMIT:
            print(
                f"[안전 중단] 자동 탐색 상한 {AUTO_PAGE_LIMIT:,}페이지에 "
                "도달했습니다. 사이트 페이지 구조를 확인하세요."
            )
            break

        if page_index == 0:
            list_soup = first_soup
        else:
            # 목록 페이지 사이에도 짧게 대기하여 과도한 연속 요청을 피합니다.
            polite_sleep()
            list_soup = fetch_soup(
                session,
                BASE_URL + LIST_PATH,
                params=make_page_params(config, page_index),
            )

        cards = extract_plan_cards(list_soup)

        progress = (
            f"{page_index + 1}/{page_count}"
            if page_count is not None
            else f"{page_index + 1}/자동"
        )
        print(
            f"[{progress}] page.page={page_index}: "
            f"카드 {len(cards)}개"
        )

        if not cards and page_index == 0:
            raise RuntimeError(
                f"[{label}] 첫 페이지에서 요금제 카드를 찾지 못했습니다. "
                "사이트 HTML 구조, 로그인/차단 화면 또는 네트워크 응답을 확인하세요."
            )

        if not cards:
            consecutive_empty_pages += 1
            print(
                f"  빈 페이지 "
                f"({consecutive_empty_pages}/{MAX_CONSECUTIVE_EMPTY_PAGES})"
            )

            if consecutive_empty_pages >= MAX_CONSECUTIVE_EMPTY_PAGES:
                print(
                    "  연속으로 빈 페이지가 반환되어 마지막 페이지로 판단하고 "
                    "이 조건의 수집을 종료합니다."
                )
                break

            page_index += 1
            continue

        consecutive_empty_pages = 0

        # 전체 건수를 읽지 못한 경우에만 동일 페이지 반복 여부를 종료 안전장치로 확인합니다.
        # 아래 비교는 저장 행을 제거하지 않습니다. 해당 페이지의 모든 카드를 먼저 저장한 뒤,
        # 완전히 같은 페이지가 계속 반환될 때 무한 탐색만 중단합니다.
        if page_count is None:
            current_signature = tuple(
                (card["detail_url"], card["card_text"]) for card in cards
            )
            if current_signature == previous_page_signature:
                consecutive_identical_pages += 1
            else:
                consecutive_identical_pages = 0
            previous_page_signature = current_signature

        # 중복 여부와 관계없이 현재 페이지에서 발견된 모든 카드를 순서대로 저장합니다.
        for card in cards:
            detail_url = card["detail_url"]

            card_info: Dict[str, Any] = {}
            detail_info: Dict[str, str] = {}

            try:
                card_info = parse_card_text(card["card_text"])
            except Exception as exc:
                all_errors.append(
                    {
                        "구분": label,
                        "단계": "목록 카드 파싱",
                        "URL": detail_url,
                        "오류": repr(exc),
                        "원문": card["card_text"],
                    }
                )

            # 상세 페이지 캐시는 중복 행을 제거하는 기능이 아닙니다.
            # 같은 URL을 다시 요청하지 않고 사업자/요금제 정보를 재사용할 뿐이며,
            # 현재 카드에 대응하는 행은 매번 별도로 rows에 추가됩니다.
            if detail_url in detail_cache:
                detail_info = detail_cache[detail_url]
            else:
                try:
                    polite_sleep()
                    detail_soup = fetch_soup(session, detail_url)
                    detail_info = parse_detail_page(detail_soup)
                    detail_cache[detail_url] = detail_info
                except Exception as exc:
                    all_errors.append(
                        {
                            "구분": label,
                            "단계": "상세 페이지 요청/파싱",
                            "URL": detail_url,
                            "오류": repr(exc),
                            "원문": "",
                        }
                    )

            row = {
                "사업자": detail_info.get("사업자", ""),
                "요금제": detail_info.get("요금제")
                or card_info.get("fallback_plan_name", ""),
                "통화제공량": detail_info.get("통화제공량")
                or card_info.get("통화제공량", ""),
                "문자제공량": detail_info.get("문자제공량")
                or card_info.get("문자제공량", ""),
                "데이터 제공량": detail_info.get("데이터 제공량")
                or card_info.get("데이터 제공량", ""),
                "망정보": detail_info.get("망정보")
                or card_info.get("망정보", ""),
                "LTE/5G 구분": detail_info.get("LTE/5G 구분")
                or card_info.get("LTE/5G 구분", ""),
                "월 요금": card_info.get("월 요금"),
                "할인 기간": card_info.get("할인 기간"),
                "기간 이후 요금": card_info.get("기간 이후 요금"),
            }
            rows.append(row)

        print(f"  [{label}] 누적 저장 행: {len(rows):,}건 (중복 포함)")
        page_index += 1

        if (
            page_count is None
            and consecutive_identical_pages >= MAX_CONSECUTIVE_IDENTICAL_PAGES
        ):
            print(
                "  완전히 동일한 페이지가 연속 반환되어 무한 반복 방지를 위해 "
                "자동 탐색을 종료합니다. 반복 페이지의 행도 저장되어 있습니다."
            )
            break

    if total_count is not None and len(rows) != total_count:
        print(
            f"[안내] 사이트 표시 결과는 {total_count:,}건이고 실제 저장 행은 "
            f"{len(rows):,}건입니다. 중복을 제거하지 않으므로 상단 추천 영역이나 "
            "페이지 내 반복 노출이 모두 포함되어 건수가 다를 수 있습니다."
        )

    return rows

def crawl() -> None:
    session = build_session()
    print(f"SSL 인증서 검증: {ssl_mode_description()}")
    if (
        TRUSTSTORE_IMPORT_ERROR is not None
        and not CUSTOM_CA_BUNDLE
        and not ALLOW_INSECURE_SSL
    ):
        print(
            "[안내] truststore가 설치되어 있지 않습니다. SSL 오류가 발생하면 "
            "다음 명령을 실행하세요:"
        )
        print(f'  "{sys.executable}" -m pip install -U truststore')

    all_results: List[Tuple[Dict[str, Any], List[Dict[str, Any]]]] = []
    all_errors: List[Dict[str, Any]] = []
    detail_cache: Dict[str, Dict[str, str]] = {}

    for config in SHEET_CONFIGS:
        rows = crawl_dataset(
            session=session,
            config=config,
            detail_cache=detail_cache,
            all_errors=all_errors,
        )
        all_results.append((config, rows))

    save_outputs(all_results, all_errors)


if __name__ == "__main__":
    crawl()
