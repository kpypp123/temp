from __future__ import annotations

import csv
import html
import json
import io
import math
import re
import urllib.parse
import urllib.request
import uuid
from datetime import date, datetime, time, timedelta
from typing import Any
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st


APP_TITLE = "현장 폭염 조치 기록"
APP_VERSION = "Professional UI v3.11 · 2026-08-10"
WORKSHEET_DEFAULT = "records"
SPREADSHEET_URL_FALLBACK = (
    "https://docs.google.com/spreadsheets/d/"
    "18c-qnfPmGG25qyAM497R7czDw3F7J7WRKmdLX3IGtY0"
)
KST = ZoneInfo("Asia/Seoul")
KMA_POINT_API_URL = (
    "https://apihub.kma.go.kr/api/typ01/cgi-bin/url/"
    "nph-sfc_obs_nc_pt_api"
)
KMA_ULTRA_NOWCAST_API_URL = (
    "https://apihub.kma.go.kr/api/typ02/openApi/"
    "VilageFcstInfoService_2.0/getUltraSrtNcst"
)
KAKAO_PLACE_API_URL = "https://dapi.kakao.com/v2/local/search/keyword.json"

COLUMNS = [
    "id",
    "작업날짜",
    "현장명",
    "팀",
    "근무시작",
    "근무종료",
    "작성자",
    "작업인원",
    "폭염시작",
    "폭염종료",
    "체감온도",
    "휴게시작",
    "휴게종료",
    "휴게시간",
    "조치사항",
    "특이사항",
    "등록시간",
    "수정시간",
]

TEAM_OPTIONS = ["중계팀", "영상팀"]

MEASURE_OPTIONS = [
    "1시간 이내 10분 이상 휴식",
    "2시간 이내 20분 이상 휴식",
    "냉방기 가동 25도 이하",
    "제작팀 협의 조정",
]

TIME_FIELD_COLUMNS = {
    "work_start": "근무시작",
    "work_end": "근무종료",
    "heat_start": "폭염시작",
    "heat_end": "폭염종료",
}


st.set_page_config(
    page_title=APP_TITLE,
    page_icon=":material/health_and_safety:",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    :root {
        --bg: #f3f5f7;
        --surface: #ffffff;
        --surface-subtle: #f8f9fb;
        --ink: #111827;
        --muted: #667085;
        --line: #d8dee7;
        --line-strong: #c7ced8;
        --navy: #172b4d;
        --navy-hover: #0f213d;
        --danger: #b42318;
        --danger-soft: #fef3f2;
        --warning: #b54708;
        --warning-soft: #fff7ed;
        --success: #067647;
        --success-soft: #ecfdf3;
    }

    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI",
                     "Noto Sans KR", sans-serif;
    }

    .stApp {
        background: var(--bg);
        color: var(--ink);
    }

    header[data-testid="stHeader"] {
        background: transparent;
    }

    #MainMenu, footer {
        visibility: hidden;
    }

    .block-container {
        max-width: 860px;
        padding-top: 1.35rem;
        padding-bottom: 5rem;
    }

    .app-header {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 1.25rem;
        padding: 0.2rem 0 1.15rem;
        border-bottom: 1px solid var(--line);
        margin-bottom: 1rem;
    }

    .app-kicker {
        color: var(--muted);
        font-size: 0.74rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        margin-bottom: 0.28rem;
    }

    .app-title {
        color: var(--ink);
        font-size: clamp(1.72rem, 5vw, 2.2rem);
        font-weight: 760;
        letter-spacing: -0.035em;
        line-height: 1.2;
        margin: 0;
    }

    .app-subtitle {
        color: var(--muted);
        font-size: 0.9rem;
        line-height: 1.55;
        margin: 0.42rem 0 0;
    }

    .app-version {
        flex: 0 0 auto;
        color: var(--muted);
        background: var(--surface-subtle);
        border: 1px solid var(--line);
        border-radius: 6px;
        padding: 0.38rem 0.55rem;
        font-size: 0.7rem;
        font-weight: 650;
        white-space: nowrap;
    }

    .section-heading {
        display: flex;
        align-items: flex-start;
        gap: 0.72rem;
        margin: 0 0 0.9rem;
    }

    .section-number {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 1.75rem;
        height: 1.75rem;
        border-radius: 5px;
        background: var(--navy);
        color: #ffffff;
        font-size: 0.72rem;
        font-weight: 750;
        line-height: 1;
    }

    .section-heading b {
        display: block;
        color: var(--ink);
        font-size: 1rem;
        font-weight: 730;
        line-height: 1.3;
    }

    .section-heading small {
        display: block;
        color: var(--muted);
        font-size: 0.78rem;
        line-height: 1.45;
        margin-top: 0.12rem;
    }

    div[data-testid="stVerticalBlockBorderWrapper"] {
        background: var(--surface);
        border: 1px solid var(--line) !important;
        border-radius: 10px !important;
        box-shadow: 0 1px 2px rgba(16, 24, 40, 0.03);
    }

    div[data-testid="stVerticalBlockBorderWrapper"] > div {
        padding: 1.05rem 1.05rem 1.15rem;
    }

    div.stButton > button,
    div[data-testid="stDownloadButton"] > button,
    div[data-testid="stLinkButton"] a {
        min-height: 3.35rem;
        border-radius: 9px;
        border: 1px solid var(--line-strong);
        background: var(--surface);
        color: var(--ink);
        font-size: 1rem;
        font-weight: 720;
        box-shadow: none;
    }

    .field-label {
        color: #344054;
        font-size: 0.9rem;
        font-weight: 700;
        margin: 0.15rem 0 0.45rem;
    }

    div.stButton > button[kind="primary"] {
        background: var(--navy);
        border-color: var(--navy);
        color: #ffffff;
    }

    div.stButton > button[kind="primary"]:hover {
        background: var(--navy-hover);
        border-color: var(--navy-hover);
        color: #ffffff;
    }

    div[data-baseweb="input"] > div,
    div[data-baseweb="select"] > div,
    div[data-baseweb="textarea"] > div,
    textarea {
        border-radius: 7px !important;
        border-color: var(--line-strong) !important;
        background: #ffffff !important;
    }

    label[data-testid="stWidgetLabel"] p {
        color: #344054;
        font-size: 0.84rem;
        font-weight: 650;
    }

    .quick-time-summary,
    .metric-grid,
    .record-details {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 0.62rem;
    }

    .quick-time-summary {
        margin: 0.72rem 0 0.7rem;
    }

    .metric-grid {
        margin: 0.65rem 0 1rem;
    }

    .record-details {
        gap: 0.48rem;
        margin: 0.8rem 0 0.75rem;
    }

    .quick-time-card,
    .metric-card,
    .record-detail {
        background: var(--surface-subtle);
        border: 1px solid var(--line);
        border-radius: 8px;
        padding: 0.78rem 0.82rem;
    }

    .metric-card {
        background: var(--surface);
        padding: 0.9rem 0.95rem;
    }

    .quick-time-card span,
    .metric-card span,
    .record-detail span {
        display: block;
        color: var(--muted);
        font-size: 0.72rem;
        font-weight: 650;
        margin-bottom: 0.2rem;
    }

    .quick-time-card strong {
        display: block;
        color: var(--ink);
        font-size: 0.97rem;
        font-weight: 720;
        white-space: nowrap;
    }

    .metric-card b {
        display: block;
        color: var(--ink);
        font-size: 1.45rem;
        font-weight: 740;
        line-height: 1.15;
    }

    .record-detail b {
        display: block;
        color: var(--ink);
        font-size: 0.84rem;
        font-weight: 680;
        line-height: 1.35;
    }

    .status-pill {
        display: inline-flex;
        align-items: center;
        min-height: 1.7rem;
        padding: 0.18rem 0.5rem;
        border-radius: 5px;
        border: 1px solid var(--line);
        background: var(--surface-subtle);
        color: #475467;
        font-size: 0.72rem;
        font-weight: 720;
        margin-bottom: 0.42rem;
    }

    .status-pill.status-caution {
        background: var(--warning-soft);
        border-color: #fed7aa;
        color: var(--warning);
    }

    .status-pill.status-danger {
        background: var(--danger-soft);
        border-color: #fecdca;
        color: var(--danger);
    }

    .status-pill.status-normal {
        background: var(--success-soft);
        border-color: #abefc6;
        color: var(--success);
    }

    .setup-box {
        background: #fffcf5;
        border: 1px solid #fedf89;
        border-radius: 8px;
        color: #7a2e0e;
        padding: 0.9rem 1rem;
        margin-bottom: 1rem;
    }

    @media (max-width: 600px) {
        .block-container {
            padding-left: 0.78rem;
            padding-right: 0.78rem;
            padding-top: 0.85rem;
        }

        .app-header {
            display: block;
        }

        .app-version {
            display: inline-block;
            margin-top: 0.7rem;
        }

        .quick-time-summary,
        .metric-grid,
        .record-details {
            grid-template-columns: 1fr;
        }

        div[data-testid="stVerticalBlockBorderWrapper"] > div {
            padding: 0.9rem;
        }

        div.stButton > button,
        div[data-testid="stDownloadButton"] > button,
        div[data-testid="stLinkButton"] a {
            min-height: 3.6rem;
            font-size: 1.03rem;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def empty_dataframe() -> pd.DataFrame:
    return pd.DataFrame(columns=COLUMNS)


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def parse_int(value: Any, default: int = 0) -> int:
    try:
        return int(float(clean_text(value)))
    except (TypeError, ValueError):
        return default


def parse_float(value: Any) -> float | None:
    try:
        text = clean_text(value)
        return float(text) if text else None
    except (TypeError, ValueError):
        return None


def format_number(value: float) -> str:
    if value.is_integer():
        return str(int(value))
    return f"{value:.1f}".rstrip("0").rstrip(".")


def normalize_temperature_text(value: Any) -> tuple[str, str | None]:
    """
    체감온도 입력값을 검증하고 저장용 문자열로 정리합니다.

    허용 예:
    - 34
    - 33~35
    - 33.5~35.2
    - 33-35
    - 33～35℃
    """
    text = clean_text(value)
    if not text:
        return "", None

    text = (
        text.replace("°C", "")
        .replace("°c", "")
        .replace("℃", "")
        .replace("도", "")
        .replace("～", "~")
        .replace("〜", "~")
        .replace("–", "~")
        .replace("—", "~")
    )

    # 양수 두 값 사이의 일반 하이픈은 범위 기호로 처리합니다.
    text = re.sub(r"(?<=\d)\s*-\s*(?=\d)", "~", text)
    text = re.sub(r"\s+", "", text)

    match = re.fullmatch(
        r"(-?\d+(?:\.\d+)?)(?:~(-?\d+(?:\.\d+)?))?",
        text,
    )
    if not match:
        return "", "체감온도는 34 또는 33~35 형식으로 입력해 주세요."

    first = float(match.group(1))
    second_text = match.group(2)
    second = float(second_text) if second_text is not None else None

    values = [first] if second is None else [first, second]
    if any(number < -20 or number > 60 for number in values):
        return "", "체감온도는 -20℃에서 60℃ 사이로 입력해 주세요."

    if second is not None and first > second:
        return "", "체감온도 범위는 낮은 값부터 입력해 주세요. 예: 33~35"

    normalized = format_number(first)
    if second is not None:
        normalized += f"~{format_number(second)}"

    return normalized, None


def temperature_numbers(value: Any) -> list[float]:
    normalized, error = normalize_temperature_text(value)
    if error or not normalized:
        return []

    numbers: list[float] = []
    for part in normalized.split("~"):
        try:
            numbers.append(float(part))
        except ValueError:
            continue
    return numbers


def max_temperature(value: Any) -> float | None:
    numbers = temperature_numbers(value)
    return max(numbers) if numbers else None


def temperature_display(value: Any) -> str:
    normalized, error = normalize_temperature_text(value)
    if error:
        return clean_text(value)
    return normalized


def parse_time_value(value: Any, default: time | None = None) -> time | None:
    text = clean_text(value)
    if not text:
        return default

    normalized = text.replace("시", ":00")
    if re.fullmatch(r"\d{1,2}", normalized):
        normalized = f"{normalized}:00"

    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(normalized, fmt).time().replace(
                second=0,
                microsecond=0,
            )
        except ValueError:
            continue
    return default


def format_time_value(value: time | None) -> str:
    return value.strftime("%H:%M") if value is not None else ""


def markdown_escape(value: Any) -> str:
    return re.sub(
        r"([\\`*_{}\[\]()#+\-.!|>])",
        r"\\\1",
        clean_text(value),
    )


def get_secret(path: tuple[str, ...], default: str = "") -> str:
    try:
        current: Any = st.secrets
        for key in path:
            current = current[key]
        return clean_text(current)
    except (KeyError, TypeError):
        return default


def time_state_key(field: str, nonce: int) -> str:
    return f"time_{field}_{nonce}"


def initialize_time_state(
    editing_record: dict[str, Any],
    nonce: int,
) -> None:
    for field, column in TIME_FIELD_COLUMNS.items():
        key = time_state_key(field, nonce)
        if key not in st.session_state:
            st.session_state[key] = parse_time_value(
                editing_record.get(column)
            )


def set_time_now(field: str, nonce: int) -> None:
    st.session_state[time_state_key(field, nonce)] = (
        datetime.now(KST).time().replace(second=0, microsecond=0)
    )


def weather_notice_key(nonce: int) -> str:
    return f"weather_notice_{nonce}"


def get_site_coordinates(site_name: str) -> tuple[float, float] | None:
    """Streamlit Secrets에서 현장명의 위도·경도를 찾습니다."""
    try:
        sites = st.secrets["weather"]["sites"]
    except (KeyError, TypeError):
        return None

    target = clean_text(site_name).casefold()

    for configured_name, configured_value in sites.items():
        if clean_text(configured_name).casefold() != target:
            continue

        try:
            if isinstance(configured_value, str):
                parts = [
                    part.strip()
                    for part in configured_value.split(",")
                ]
                latitude, longitude = map(float, parts[:2])
            elif hasattr(configured_value, "get"):
                latitude = float(configured_value.get("lat"))
                longitude = float(configured_value.get("lon"))
            else:
                latitude = float(configured_value[0])
                longitude = float(configured_value[1])
        except (IndexError, TypeError, ValueError):
            return None

        if not (33 <= latitude <= 39 and 124 <= longitude <= 132):
            return None

        return latitude, longitude

    return None


def search_kakao_site_coordinates(
    site_name: str,
    rest_api_key: str,
) -> tuple[float, float, str]:
    """카카오 장소 검색으로 현장명의 위도·경도를 찾습니다."""
    params = urllib.parse.urlencode({"query": site_name, "size": "5"})
    request = urllib.request.Request(
        f"{KAKAO_PLACE_API_URL}?{params}",
        headers={
            "Authorization": f"KakaoAK {rest_api_key}",
            "User-Agent": "checktemp-streamlit/1.0",
        },
    )

    with urllib.request.urlopen(request, timeout=8) as response:
        result = json.loads(response.read().decode("utf-8"))

    documents = result.get("documents") or []
    if not documents:
        raise ValueError("카카오 장소 검색 결과가 없습니다")

    place = documents[0]
    latitude = float(place["y"])
    longitude = float(place["x"])
    matched_name = clean_text(place.get("place_name")) or site_name

    if not (33 <= latitude <= 39 and 124 <= longitude <= 132):
        raise ValueError("검색된 장소가 대한민국 범위를 벗어났습니다")

    return latitude, longitude, matched_name


def parse_kma_temperature_response(response_text: str) -> tuple[str, str]:
    """기상청 특정지점 ASCII 응답에서 최신 체감온도를 추출합니다."""
    observations: list[tuple[str, float]] = []

    for raw_line in response_text.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        # 응답 버전에 따라 공백/쉼표 구분 및 초(14자리) 포함 여부가
        # 달라질 수 있으므로 행 전체에서 시각과 값을 찾습니다.
        timestamp_match = re.search(r"(?<!\d)(\d{12}|\d{14})(?!\d)", line)
        if timestamp_match:
            timestamp = timestamp_match.group(1)[:12]
            value_text = line[timestamp_match.end():]
        else:
            separated_time = re.search(
                r"(\d{4})[-/.]?(\d{2})[-/.]?(\d{2})"
                r"[ T]?(\d{2}):?(\d{2})",
                line,
            )
            if not separated_time:
                continue
            timestamp = "".join(separated_time.groups())
            value_text = line[separated_time.end():]

        numeric_values = re.findall(r"[-+]?\d+(?:\.\d+)?", value_text)
        if not numeric_values:
            continue

        # 특정지점 단일요소 응답의 관측값은 행의 마지막 숫자입니다.
        value = float(numeric_values[-1])
        if -50 <= value <= 80 and value not in (-99.0, -999.0, -9999.0):
            observations.append((timestamp, value))

    if not observations:
        raise ValueError("기상청 응답에서 체감온도 값을 찾지 못했습니다.")

    timestamp, value = observations[-1]
    observed_at = datetime.strptime(timestamp, "%Y%m%d%H%M").strftime(
        "%H:%M"
    )
    return format_number(value), observed_at


def fetch_kma_apparent_temperature(
    latitude: float,
    longitude: float,
    auth_key: str,
) -> tuple[str, str]:
    """기상청 500m 격자 특정지점의 최신 체감온도를 조회합니다."""
    current_time = datetime.now(KST).replace(second=0, microsecond=0)
    query_end = current_time - timedelta(
        minutes=(current_time.minute % 5) + 10
    )
    query_start = query_end - timedelta(minutes=55)
    params = urllib.parse.urlencode(
        {
            "obs": "ta_chi",
            "tm1": query_start.strftime("%Y%m%d%H%M"),
            "tm2": query_end.strftime("%Y%m%d%H%M"),
            "itv": "5",
            "lon": f"{longitude:.6f}",
            "lat": f"{latitude:.6f}",
            "authKey": auth_key,
        }
    )
    request = urllib.request.Request(
        f"{KMA_POINT_API_URL}?{params}",
        headers={"User-Agent": "checktemp-streamlit/1.0"},
    )

    with urllib.request.urlopen(request, timeout=8) as response:
        payload = response.read()

    for encoding in ("utf-8", "euc-kr"):
        try:
            response_text = payload.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        response_text = payload.decode("utf-8", errors="replace")

    return parse_kma_temperature_response(response_text)


def latitude_longitude_to_grid(
    latitude: float,
    longitude: float,
) -> tuple[int, int]:
    """위·경도를 기상청 동네예보 Lambert 격자 좌표로 변환합니다."""
    earth_radius = 6371.00877
    grid_size = 5.0
    standard_parallel_1 = 30.0
    standard_parallel_2 = 60.0
    origin_longitude = 126.0
    origin_latitude = 38.0
    origin_x = 43.0
    origin_y = 136.0
    radians = math.pi / 180.0

    radius = earth_radius / grid_size
    parallel_1 = standard_parallel_1 * radians
    parallel_2 = standard_parallel_2 * radians
    origin_lon = origin_longitude * radians
    origin_lat = origin_latitude * radians

    sn = math.tan(math.pi * 0.25 + parallel_2 * 0.5) / math.tan(
        math.pi * 0.25 + parallel_1 * 0.5
    )
    sn = math.log(math.cos(parallel_1) / math.cos(parallel_2)) / math.log(sn)
    sf = math.tan(math.pi * 0.25 + parallel_1 * 0.5) ** sn
    sf = sf * math.cos(parallel_1) / sn
    ro = math.tan(math.pi * 0.25 + origin_lat * 0.5) ** sn
    ro = radius * sf / ro

    ra = math.tan(math.pi * 0.25 + latitude * radians * 0.5) ** sn
    ra = radius * sf / ra
    theta = longitude * radians - origin_lon
    if theta > math.pi:
        theta -= 2.0 * math.pi
    if theta < -math.pi:
        theta += 2.0 * math.pi
    theta *= sn

    grid_x = int(ra * math.sin(theta) + origin_x + 0.5)
    grid_y = int(ro - ra * math.cos(theta) + origin_y + 0.5)
    return grid_x, grid_y


def calculate_summer_apparent_temperature(
    air_temperature: float,
    relative_humidity: float,
) -> float:
    """기상청 여름철 산식으로 기온·습도 기반 체감온도를 계산합니다."""
    wet_bulb = (
        air_temperature
        * math.atan(0.151977 * math.sqrt(relative_humidity + 8.313659))
        + math.atan(air_temperature + relative_humidity)
        - math.atan(relative_humidity - 1.67633)
        + 0.00391838
        * relative_humidity ** 1.5
        * math.atan(0.023101 * relative_humidity)
        - 4.686035
    )
    apparent = (
        -0.2442
        + 0.55399 * wet_bulb
        + 0.45535 * air_temperature
        - 0.0022 * wet_bulb**2
        + 0.00278 * wet_bulb * air_temperature
        + 3.0
    )
    return round(apparent, 1)


def fetch_kma_regional_apparent_temperature(
    latitude: float,
    longitude: float,
    auth_key: str,
) -> tuple[str, str, str, str]:
    """동네예보 초단기실황의 기온·습도로 지역 체감온도를 계산합니다."""
    grid_x, grid_y = latitude_longitude_to_grid(latitude, longitude)
    current_time = datetime.now(KST).replace(second=0, microsecond=0)
    latest_candidate = current_time.replace(minute=0)
    if current_time.minute < 15:
        latest_candidate -= timedelta(hours=1)

    last_error = "조회 가능한 초단기실황이 없습니다."
    for hours_back in range(4):
        base_time = latest_candidate - timedelta(hours=hours_back)
        params = urllib.parse.urlencode(
            {
                "pageNo": "1",
                "numOfRows": "1000",
                "dataType": "JSON",
                "base_date": base_time.strftime("%Y%m%d"),
                "base_time": base_time.strftime("%H%M"),
                "nx": str(grid_x),
                "ny": str(grid_y),
                "authKey": auth_key,
            }
        )
        request = urllib.request.Request(
            f"{KMA_ULTRA_NOWCAST_API_URL}?{params}",
            headers={"User-Agent": "checktemp-streamlit/1.0"},
        )

        try:
            with urllib.request.urlopen(request, timeout=8) as response:
                result = json.loads(response.read().decode("utf-8"))
            response_body = result.get("response") or {}
            header = response_body.get("header") or {}
            if clean_text(header.get("resultCode")) != "00":
                last_error = clean_text(header.get("resultMsg")) or last_error
                continue

            raw_items = (
                response_body.get("body", {}).get("items", {}).get("item", [])
            )
            values = {
                clean_text(item.get("category")): parse_float(item.get("obsrValue"))
                for item in raw_items
            }
            air_temperature = values.get("T1H")
            relative_humidity = values.get("REH")
            if air_temperature is None or relative_humidity is None:
                last_error = "초단기실황에 기온 또는 습도 값이 없습니다."
                continue
            if not (-50 <= air_temperature <= 60 and 0 <= relative_humidity <= 100):
                last_error = "초단기실황 기온·습도 값이 정상 범위를 벗어났습니다."
                continue

            apparent = calculate_summer_apparent_temperature(
                air_temperature,
                relative_humidity,
            )
            return (
                format_number(apparent),
                base_time.strftime("%H:%M"),
                format_number(air_temperature),
                format_number(relative_humidity),
            )
        except Exception as error:  # noqa: BLE001
            last_error = str(error)

    raise ValueError(last_error)


def record_heat_start_with_weather(
    nonce: int,
    record_start_time: bool = True,
) -> None:
    """폭염 시작시간을 기록하고 설정 시 체감온도도 자동 입력합니다."""
    if record_start_time:
        set_time_now("heat_start", nonce)
    notice_key = weather_notice_key(nonce)
    site_name = clean_text(st.session_state.get(f"site_{nonce}"))

    if not site_name:
        st.session_state[notice_key] = (
            "warning",
            "현장명을 먼저 입력하면 체감온도를 자동 조회할 수 있습니다.",
        )
        return

    auth_key = get_secret(("weather", "kma_auth_key"))
    if not auth_key:
        st.session_state[notice_key] = (
            "info",
            "기상청 인증키가 아직 설정되지 않아 시간만 기록했습니다.",
        )
        return

    coordinates = get_site_coordinates(site_name)
    matched_name = site_name
    if coordinates is None:
        kakao_key = get_secret(("location", "kakao_rest_api_key"))
        if not kakao_key:
            st.session_state[notice_key] = (
                "warning",
                "현장 좌표가 등록되지 않았고 카카오 REST API 키도 없습니다. "
                "폭염 시작시간만 기록했습니다.",
            )
            return

        try:
            latitude, longitude, matched_name = (
                search_kakao_site_coordinates(site_name, kakao_key)
            )
            coordinates = (latitude, longitude)
        except Exception as error:  # noqa: BLE001
            st.session_state[notice_key] = (
                "warning",
                f"'{site_name}' 장소 검색에 실패해 시작시간만 기록했습니다. "
                f"현장명을 더 정확하게 입력해 주세요. ({error})",
            )
            return

    grid_temperature: str | None = None
    grid_observed_at = ""
    regional_temperature: str | None = None
    regional_observed_at = ""
    regional_air_temperature = ""
    regional_humidity = ""
    lookup_errors: list[str] = []

    try:
        grid_temperature, grid_observed_at = fetch_kma_apparent_temperature(
            coordinates[0],
            coordinates[1],
            auth_key,
        )
    except Exception as error:  # noqa: BLE001
        lookup_errors.append(f"500m 격자: {error}")

    try:
        (
            regional_temperature,
            regional_observed_at,
            regional_air_temperature,
            regional_humidity,
        ) = fetch_kma_regional_apparent_temperature(
            coordinates[0],
            coordinates[1],
            auth_key,
        )
    except Exception as error:  # noqa: BLE001
        lookup_errors.append(f"지역 실황: {error}")

    available_temperatures = [
        value
        for value in (grid_temperature, regional_temperature)
        if value is not None
    ]
    if not available_temperatures:
        st.session_state[notice_key] = (
            "warning",
            "기상청 체감온도 조회에 실패해 시간만 기록했습니다. "
            f"직접 입력해 주세요. ({' / '.join(lookup_errors)})",
        )
        return

    applied_temperature = max(
        available_temperatures,
        key=lambda value: float(value),
    )
    st.session_state[f"temperature_{nonce}"] = applied_temperature

    detail_parts: list[str] = []
    if grid_temperature is not None:
        detail_parts.append(
            f"500m 격자 {grid_observed_at} {grid_temperature}℃"
        )
    if regional_temperature is not None:
        detail_parts.append(
            f"지역 실황 {regional_observed_at} {regional_temperature}℃"
            f"(기온 {regional_air_temperature}℃·습도 {regional_humidity}%)"
        )
    if lookup_errors:
        detail_parts.append(f"일부 조회 실패: {' / '.join(lookup_errors)}")

    st.session_state[notice_key] = (
        "success",
        f"{matched_name} · {' / '.join(detail_parts)} · "
        f"안전을 위해 높은 값 {applied_temperature}℃를 자동 적용했습니다.",
    )


def clear_time(field: str, nonce: int) -> None:
    st.session_state[time_state_key(field, nonce)] = None


def time_state_text(field: str, nonce: int) -> str:
    return format_time_value(
        st.session_state.get(time_state_key(field, nonce))
    )


def team_state_key(nonce: int) -> str:
    return f"selected_team_{nonce}"


def initialize_team_state(
    editing_record: dict[str, Any],
    nonce: int,
) -> None:
    key = team_state_key(nonce)
    if key not in st.session_state:
        current_team = clean_text(editing_record.get("팀"))
        st.session_state[key] = (
            current_team
            if current_team in TEAM_OPTIONS
            else TEAM_OPTIONS[0]
        )


def set_team(team: str, nonce: int) -> None:
    if team in TEAM_OPTIONS:
        st.session_state[team_state_key(nonce)] = team


def measures_state_key(nonce: int) -> str:
    return f"selected_measures_{nonce}"


def initialize_measures_state(
    editing_record: dict[str, Any],
    nonce: int,
) -> None:
    key = measures_state_key(nonce)
    if key not in st.session_state:
        st.session_state[key] = selected_measures(
            editing_record.get("조치사항")
        )


def toggle_measure(measure: str, nonce: int) -> None:
    key = measures_state_key(nonce)
    current = list(st.session_state.get(key, []))

    if measure in current:
        current.remove(measure)
    else:
        # 10분/20분 휴식 기준은 동시에 선택하지 않도록 합니다.
        if measure in MEASURE_OPTIONS[:2]:
            current = [
                option
                for option in current
                if option not in MEASURE_OPTIONS[:2]
            ]
        current.append(measure)

    # 화면/저장 순서는 MEASURE_OPTIONS 순서로 고정
    st.session_state[key] = [
        option for option in MEASURE_OPTIONS if option in current
    ]


def rest_minutes_key(nonce: int) -> str:
    return f"rest_minutes_{nonce}"


def initialize_rest_minutes(
    editing_record: dict[str, Any],
    nonce: int,
) -> None:
    key = rest_minutes_key(nonce)
    if key not in st.session_state:
        st.session_state[key] = max(
            0,
            parse_int(editing_record.get("휴게시간"), 0),
        )


def add_rest_minutes(minutes: int, nonce: int) -> None:
    key = rest_minutes_key(nonce)
    current = parse_int(st.session_state.get(key), 0)
    st.session_state[key] = max(0, current + minutes)


def clear_rest_minutes(nonce: int) -> None:
    st.session_state[rest_minutes_key(nonce)] = 0


def calculate_minutes(start: str, end: str) -> int:
    if not start or not end:
        return 0

    start_hour, start_minute = map(int, start.split(":"))
    end_hour, end_minute = map(int, end.split(":"))

    minutes = (
        end_hour * 60
        + end_minute
        - start_hour * 60
        - start_minute
    )
    return minutes + 1440 if minutes < 0 else minutes


def render_time_summary(nonce: int) -> None:
    work_start = time_state_text("work_start", nonce)
    work_end = time_state_text("work_end", nonce)
    heat_start = time_state_text("heat_start", nonce)
    heat_end = time_state_text("heat_end", nonce)
    rest_minutes = parse_int(
        st.session_state.get(rest_minutes_key(nonce)),
        0,
    )

    work_text = f"{work_start or '-'} ~ {work_end or '-'}"
    heat_text = f"{heat_start or '-'} ~ {heat_end or '-'}"
    rest_text = f"{rest_minutes}분" if rest_minutes > 0 else "-"

    st.markdown(
        f"""
        <div class="quick-time-summary">
            <div class="quick-time-card">
                <span>근무 시간</span>
                <strong>{html.escape(work_text)}</strong>
            </div>
            <div class="quick-time-card">
                <span>폭염 노출</span>
                <strong>{html.escape(heat_text)}</strong>
            </div>
            <div class="quick-time-card">
                <span>누적 휴게시간</span>
                <strong>{html.escape(rest_text)}</strong>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_resource(show_spinner=False)
def get_worksheet() -> gspread.Worksheet:
    service_account = dict(st.secrets["google_service_account"])

    if "private_key" in service_account:
        service_account["private_key"] = str(
            service_account["private_key"]
        ).replace("\\n", "\n")

    spreadsheet_url = get_secret(("app", "spreadsheet_url"))
    worksheet_name = get_secret(
        ("app", "worksheet"),
        WORKSHEET_DEFAULT,
    )

    if not spreadsheet_url:
        raise ValueError(
            "Streamlit Secrets의 app.spreadsheet_url 값이 비어 있습니다."
        )

    client = gspread.service_account_from_dict(service_account)
    spreadsheet = client.open_by_url(spreadsheet_url)
    return spreadsheet.worksheet(worksheet_name)


def ensure_headers(worksheet: gspread.Worksheet) -> list[str]:
    values = worksheet.get_all_values()

    if not values:
        worksheet.update(
            range_name="A1:R1",
            values=[COLUMNS],
        )
        return COLUMNS

    headers = [clean_text(value) for value in values[0]]

    if headers[: len(COLUMNS)] != COLUMNS:
        missing = [
            column
            for column in COLUMNS
            if column not in headers
        ]
        missing_text = (
            ", ".join(missing)
            if missing
            else "열 순서 불일치"
        )
        raise ValueError(
            "records 시트의 첫 행 제목이 앱 형식과 다릅니다. "
            f"확인할 항목: {missing_text}"
        )

    return headers


def load_records() -> pd.DataFrame:
    worksheet = get_worksheet()
    ensure_headers(worksheet)
    values = worksheet.get_all_values()

    if len(values) <= 1:
        return empty_dataframe()

    rows: list[dict[str, str]] = []

    for raw_row in values[1:]:
        padded = raw_row + [""] * (
            len(COLUMNS) - len(raw_row)
        )
        record = {
            column: clean_text(padded[index])
            for index, column in enumerate(COLUMNS)
        }

        if any(record.values()):
            rows.append(record)

    if not rows:
        return empty_dataframe()

    return pd.DataFrame(
        rows,
        columns=COLUMNS,
    ).fillna("")


def record_values(record: dict[str, Any]) -> list[str]:
    return [
        clean_text(record.get(column, ""))
        for column in COLUMNS
    ]


def append_record(record: dict[str, Any]) -> None:
    worksheet = get_worksheet()
    ensure_headers(worksheet)
    worksheet.append_row(
        record_values(record),
        value_input_option="USER_ENTERED",
        insert_data_option="INSERT_ROWS",
    )


def update_record(
    record_id: str,
    record: dict[str, Any],
) -> None:
    worksheet = get_worksheet()
    ensure_headers(worksheet)
    cell = worksheet.find(record_id, in_column=1)

    if cell is None:
        raise ValueError(
            "수정할 기록을 찾지 못했습니다. "
            "목록을 새로고침해 주세요."
        )

    worksheet.update(
        range_name=f"A{cell.row}:R{cell.row}",
        values=[record_values(record)],
        value_input_option="USER_ENTERED",
    )


def delete_record(record_id: str) -> None:
    worksheet = get_worksheet()
    ensure_headers(worksheet)
    cell = worksheet.find(record_id, in_column=1)

    if cell is None:
        raise ValueError(
            "삭제할 기록을 찾지 못했습니다. "
            "목록을 새로고침해 주세요."
        )

    worksheet.delete_rows(cell.row)


def heat_level(value: Any) -> str:
    temperature = max_temperature(value)

    if temperature is None:
        return "온도 미입력"
    if temperature >= 38:
        return "매우 위험"
    if temperature >= 35:
        return "위험"
    if temperature >= 33:
        return "주의"
    if temperature >= 31:
        return "관심"
    return "일반"


def heat_level_class(value: Any) -> str:
    temperature = max_temperature(value)

    if temperature is None:
        return ""
    if temperature >= 35:
        return "status-danger"
    if temperature >= 31:
        return "status-caution"
    return "status-normal"


def option_index(
    options: list[str],
    value: Any,
    fallback: int = 0,
) -> int:
    text = clean_text(value)

    try:
        return options.index(text)
    except ValueError:
        return fallback


def selected_measures(value: Any) -> list[str]:
    text = clean_text(value)

    if not text:
        return []

    parts = [
        part.strip()
        for part in re.split(r"\s*\|\s*|,\s*", text)
        if part.strip()
    ]

    return [
        part
        for part in parts
        if part in MEASURE_OPTIONS
    ]


def make_csv_bytes(dataframe: pd.DataFrame) -> bytes:
    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\n")
    writer.writerow(COLUMNS)

    for _, row in dataframe[COLUMNS].iterrows():
        writer.writerow(
            [clean_text(row[column]) for column in COLUMNS]
        )

    return output.getvalue().encode("utf-8-sig")


def init_state() -> None:
    defaults = {
        "page": "form",
        "editing_id": None,
        "pending_delete": None,
        "form_nonce": 0,
        "flash": "",
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def reset_form(*, go_to_form: bool = True) -> None:
    st.session_state.editing_id = None
    st.session_state.pending_delete = None
    st.session_state.form_nonce += 1

    if go_to_form:
        st.session_state.page = "form"


def show_flash() -> None:
    message = clean_text(
        st.session_state.get("flash", "")
    )

    if message:
        st.success(message)
        st.session_state.flash = ""


def render_section_heading(
    number: str,
    title: str,
    subtitle: str,
) -> None:
    st.markdown(
        f"""
        <div class="section-heading">
            <span class="section-number">
                {html.escape(number)}
            </span>
            <div>
                <b>{html.escape(title)}</b>
                <small>{html.escape(subtitle)}</small>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    st.markdown(
        f"""
        <div class="app-header">
            <div>
                <div class="app-kicker">현장 안전관리</div>
                <h1 class="app-title">
                    현장 폭염 조치 기록
                </h1>
                <p class="app-subtitle">
                    현장별 근무·폭염 노출·휴게 조치 내역을
                    기록하고 공동 관리합니다.
                </p>
            </div>
            <div class="app-version">
                {html.escape(APP_VERSION)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_navigation() -> None:
    left, right = st.columns(2)

    with left:
        if st.button(
            "새 기록",
            use_container_width=True,
            type=(
                "primary"
                if st.session_state.page == "form"
                else "secondary"
            ),
            key="nav_form",
        ):
            reset_form(go_to_form=True)
            st.rerun()

    with right:
        if st.button(
            "기록 조회",
            use_container_width=True,
            type=(
                "primary"
                if st.session_state.page == "records"
                else "secondary"
            ),
            key="nav_records",
        ):
            st.session_state.page = "records"
            st.session_state.pending_delete = None
            st.rerun()


def render_setup_error(error: Exception) -> None:
    st.markdown(
        """
        <div class="setup-box">
        <b>Google Sheets 연결 설정이 아직 완료되지 않았습니다.</b>
        <br>
        화면은 미리 볼 수 있지만, 연결 전에는
        저장·수정·삭제가 비활성화됩니다.
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.expander("설정 오류 자세히 보기"):
        st.code(str(error))
        st.caption(
            "Streamlit App settings → Secrets에 "
            "스프레드시트 주소와 서비스 계정 정보를 "
            "등록해야 합니다."
        )


def render_form(
    records: pd.DataFrame,
    store_error: Exception | None,
) -> None:
    editing_id = clean_text(
        st.session_state.editing_id
    )
    editing_record: dict[str, Any] = {}

    if editing_id and not records.empty:
        matches = records[
            records["id"].astype(str) == editing_id
        ]

        if not matches.empty:
            editing_record = matches.iloc[0].to_dict()
        else:
            st.warning(
                "수정하려던 기록을 찾지 못해 "
                "새 기록 화면으로 전환했습니다."
            )
            reset_form(go_to_form=True)
            st.rerun()

    title = "기록 수정" if editing_id else "새 기록 작성"
    st.markdown(f"## {title}")

    nonce = st.session_state.form_nonce
    initialize_time_state(editing_record, nonce)
    initialize_rest_minutes(editing_record, nonce)
    initialize_team_state(editing_record, nonce)
    initialize_measures_state(editing_record, nonce)

    default_date = datetime.now(KST).date()
    date_text = clean_text(
        editing_record.get("작업날짜")
    )

    if date_text:
        try:
            default_date = date.fromisoformat(date_text)
        except ValueError:
            pass

    temperature_default = clean_text(
        editing_record.get("체감온도")
    )

    with st.container(border=True):
        render_section_heading(
            "01",
            "기본 정보",
            "현장과 담당 정보를 입력합니다.",
        )

        work_date = st.date_input(
            "작업 날짜 *",
            value=default_date,
            key=f"date_{nonce}",
        )

        site = st.text_input(
            "현장명 *",
            value=clean_text(
                editing_record.get("현장명")
            ),
            placeholder="예: ○○골프장, ○○야구장",
            key=f"site_{nonce}",
        )

        st.markdown(
            '<div class="field-label">팀 선택 *</div>',
            unsafe_allow_html=True,
        )

        selected_team = clean_text(
            st.session_state.get(team_state_key(nonce))
        )
        team_left, team_right = st.columns(2)

        with team_left:
            st.button(
                "중계",
                key=f"team_relay_{nonce}",
                use_container_width=True,
                type=(
                    "primary"
                    if selected_team == "중계팀"
                    else "secondary"
                ),
                on_click=set_team,
                args=("중계팀", nonce),
            )

        with team_right:
            st.button(
                "영상",
                key=f"team_video_{nonce}",
                use_container_width=True,
                type=(
                    "primary"
                    if selected_team == "영상팀"
                    else "secondary"
                ),
                on_click=set_team,
                args=("영상팀", nonce),
            )

        team = clean_text(
            st.session_state.get(team_state_key(nonce))
        )

        author = st.text_input(
            "작성자",
            value=clean_text(
                editing_record.get("작성자")
            ),
            placeholder="예: 홍길동",
            key=f"author_{nonce}",
        )

        st.divider()

        render_section_heading(
            "02",
            "시간 기록",
            "근무·폭염 시간을 직접 입력합니다.",
        )

        st.caption(
            "24시간 형식으로 입력하세요. 예: 오전 9시는 09:00, "
            "오후 6시는 18:00"
        )

        work_left, work_right = st.columns(2)

        with work_left:
            work_start_input = st.text_input(
                "근무 시작 *",
                value=clean_text(editing_record.get("근무시작")),
                placeholder="예: 09:00",
                key=f"manual_work_start_{nonce}",
            )

        with work_right:
            work_end_input = st.text_input(
                "근무 종료 *",
                value=clean_text(editing_record.get("근무종료")),
                placeholder="예: 18:00",
                key=f"manual_work_end_{nonce}",
            )

        heat_left, heat_right = st.columns(2)

        with heat_left:
            heat_start_input = st.text_input(
                "폭염 시작",
                value=clean_text(editing_record.get("폭염시작")),
                placeholder="예: 13:00",
                key=f"manual_heat_start_{nonce}",
            )

        with heat_right:
            heat_end_input = st.text_input(
                "폭염 종료",
                value=clean_text(editing_record.get("폭염종료")),
                placeholder="예: 17:00",
                key=f"manual_heat_end_{nonce}",
            )

        work_start = parse_time_value(work_start_input)
        work_end = parse_time_value(work_end_input)
        heat_start = parse_time_value(heat_start_input)
        heat_end = parse_time_value(heat_end_input)

        st.button(
            "현장 체감온도 자동 조회",
            key=f"lookup_weather_{nonce}",
            use_container_width=True,
            on_click=record_heat_start_with_weather,
            args=(nonce, False),
        )

        weather_notice = st.session_state.get(
            weather_notice_key(nonce)
        )
        if weather_notice:
            notice_type, notice_message = weather_notice
            if notice_type == "success":
                st.success(notice_message)
            elif notice_type == "warning":
                st.warning(notice_message)
            else:
                st.info(notice_message)

        st.caption(
            "휴게시간은 아래 조치사항의 휴식 버튼을 선택해 기록합니다."
        )

        st.divider()

        render_section_heading(
            "03",
            "폭염 정보",
            "현장에서 확인한 체감온도를 "
            "직접 입력합니다.",
        )

        temperature = st.text_input(
            "체감온도 (℃)",
            value=temperature_default,
            placeholder="예: 33~35 또는 34",
            help=(
                "단일값은 34, 범위는 33~35처럼 "
                "입력합니다. ℃는 입력하지 않아도 됩니다."
            ),
            key=f"temperature_{nonce}",
        )

        st.caption(
            "자동 조회 시 현장 좌표 기준 기상청 500m 격자값과 지역 "
            "초단기실황 체감온도를 비교해 높은 값을 적용합니다. 현장 "
            "측정값이 더 높다면 직접 수정하세요. 범위로 입력한 경우에는 "
            "높은 온도를 기준으로 폭염 단계와 통계를 계산합니다."
        )

        st.divider()

        render_section_heading(
            "04",
            "조치 사항",
            "시행한 조치와 특이사항을 남깁니다.",
        )

        st.markdown(
            '<div class="field-label">시행한 조치</div>',
            unsafe_allow_html=True,
        )
        st.caption("복수 선택 가능합니다. 다시 누르면 선택이 해제됩니다.")

        selected_measure_list = list(
            st.session_state.get(
                measures_state_key(nonce),
                [],
            )
        )

        measure_labels = {
            "1시간 이내 10분 이상 휴식": "1시간 이내 10분 이상 휴식",
            "2시간 이내 20분 이상 휴식": "2시간 이내 20분 이상 휴식",
            "냉방기 가동 25도 이하": "냉방기 가동 25℃ 이하",
            "제작팀 협의 조정": "제작팀 협의 조정",
        }

        measure_row1_left, measure_row1_right = st.columns(2)
        with measure_row1_left:
            st.button(
                measure_labels[MEASURE_OPTIONS[0]],
                key=f"measure_0_{nonce}",
                use_container_width=True,
                type=(
                    "primary"
                    if MEASURE_OPTIONS[0] in selected_measure_list
                    else "secondary"
                ),
                on_click=toggle_measure,
                args=(MEASURE_OPTIONS[0], nonce),
            )

        with measure_row1_right:
            st.button(
                measure_labels[MEASURE_OPTIONS[1]],
                key=f"measure_1_{nonce}",
                use_container_width=True,
                type=(
                    "primary"
                    if MEASURE_OPTIONS[1] in selected_measure_list
                    else "secondary"
                ),
                on_click=toggle_measure,
                args=(MEASURE_OPTIONS[1], nonce),
            )

        measure_row2_left, measure_row2_right = st.columns(2)
        with measure_row2_left:
            st.button(
                measure_labels[MEASURE_OPTIONS[2]],
                key=f"measure_2_{nonce}",
                use_container_width=True,
                type=(
                    "primary"
                    if MEASURE_OPTIONS[2] in selected_measure_list
                    else "secondary"
                ),
                on_click=toggle_measure,
                args=(MEASURE_OPTIONS[2], nonce),
            )

        with measure_row2_right:
            st.button(
                measure_labels[MEASURE_OPTIONS[3]],
                key=f"measure_3_{nonce}",
                use_container_width=True,
                type=(
                    "primary"
                    if MEASURE_OPTIONS[3] in selected_measure_list
                    else "secondary"
                ),
                on_click=toggle_measure,
                args=(MEASURE_OPTIONS[3], nonce),
            )

        measures = list(
            st.session_state.get(
                measures_state_key(nonce),
                [],
            )
        )

        notes = st.text_area(
            "상세 조치 및 특이사항",
            value=clean_text(
                editing_record.get("특이사항")
            ),
            placeholder=(
                "예: 설치·철수 시간 조정, "
                "제작팀 협의사항 기재"
            ),
            height=120,
            key=f"notes_{nonce}",
        )

        save_col, reset_col = st.columns(2)

        with save_col:
            save_clicked = st.button(
                (
                    "수정 내용 저장"
                    if editing_id
                    else "기록 저장"
                ),
                key=f"save_record_{nonce}",
                use_container_width=True,
                type="primary",
                disabled=store_error is not None,
            )

        with reset_col:
            reset_clicked = st.button(
                "입력 초기화",
                key=f"reset_record_{nonce}",
                use_container_width=True,
            )

    if reset_clicked:
        reset_form(go_to_form=True)
        st.rerun()

    if not save_clicked:
        return

    work_start_text = format_time_value(work_start)
    work_end_text = format_time_value(work_end)
    heat_start_text = format_time_value(heat_start)
    heat_end_text = format_time_value(heat_end)
    # 선택한 휴식 조치가 Google Sheets의 휴게시간 열에도 저장됩니다.
    if MEASURE_OPTIONS[0] in measures:
        rest_minutes = 10
    elif MEASURE_OPTIONS[1] in measures:
        rest_minutes = 20
    else:
        rest_minutes = 0

    validation_errors: list[str] = []

    if not site.strip():
        validation_errors.append(
            "현장명을 입력해 주세요."
        )

    if not work_start_text or not work_end_text:
        validation_errors.append(
            "근무 시작과 종료 시간을 기록해 주세요."
        )

    if bool(heat_start_text) != bool(heat_end_text):
        validation_errors.append(
            "폭염 노출 시간은 시작과 종료를 "
            "모두 입력하거나 모두 비워 주세요."
        )

    normalized_temperature, temperature_error = (
        normalize_temperature_text(temperature)
    )

    if temperature_error:
        validation_errors.append(temperature_error)

    if validation_errors:
        for message in validation_errors:
            st.error(message)
        return

    now_text = datetime.now(KST).strftime(
        "%Y-%m-%d %H:%M:%S"
    )
    created_at = (
        clean_text(
            editing_record.get("등록시간")
        )
        or now_text
    )
    record_id = editing_id or str(uuid.uuid4())

    record = {
        "id": record_id,
        "작업날짜": work_date.isoformat(),
        "현장명": site.strip(),
        "팀": team,
        "근무시작": work_start_text,
        "근무종료": work_end_text,
        "작성자": author.strip(),
        "작업인원": "",
        "폭염시작": heat_start_text,
        "폭염종료": heat_end_text,
        "체감온도": normalized_temperature,
        "휴게시작": (
            clean_text(editing_record.get("휴게시작"))
            if editing_id
            else ""
        ),
        "휴게종료": (
            clean_text(editing_record.get("휴게종료"))
            if editing_id
            else ""
        ),
        "휴게시간": str(rest_minutes),
        "조치사항": " | ".join(measures),
        "특이사항": notes.strip(),
        "등록시간": created_at,
        "수정시간": now_text,
    }

    try:
        if editing_id:
            update_record(editing_id, record)
            st.session_state.flash = (
                "기록을 수정했습니다."
            )
        else:
            append_record(record)
            st.session_state.flash = (
                "기록을 저장했습니다."
            )
    except Exception as exc:  # noqa: BLE001
        st.error(f"저장에 실패했습니다: {exc}")
        return

    reset_form(go_to_form=False)
    st.session_state.page = "records"
    st.rerun()


def render_metrics(records: pd.DataFrame) -> None:
    hot_count = 0
    rest_action_count = 0

    if not records.empty:
        hot_count = sum(
            (
                max_temperature(value)
                if max_temperature(value) is not None
                else -999
            )
            >= 33
            for value in records["체감온도"]
        )

        rest_action_count = sum(
            "휴식" in clean_text(value)
            for value in records["조치사항"]
        )

    st.markdown(
        f"""
        <div class="metric-grid">
            <div class="metric-card">
                <span>전체 기록</span>
                <b>{len(records)}</b>
            </div>
            <div class="metric-card">
                <span>체감온도 33℃ 이상</span>
                <b>{hot_count}</b>
            </div>
            <div class="metric-card">
                <span>휴식 조치 기록</span>
                <b>{rest_action_count}</b>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_records(
    records: pd.DataFrame,
    store_error: Exception | None,
) -> None:
    st.markdown("## 기록 조회")

    spreadsheet_url = get_secret(
        ("app", "spreadsheet_url"),
        SPREADSHEET_URL_FALLBACK,
    )

    st.link_button(
        "Google Sheets 열기",
        spreadsheet_url,
        use_container_width=True,
    )

    if store_error is not None:
        st.info(
            "Google Sheets 연결이 완료되면 "
            "공동 기록이 여기에 표시됩니다."
        )
        return

    render_metrics(records)

    with st.expander(
        "검색 및 필터",
        expanded=False,
    ):
        search_text = st.text_input(
            "현장명·작성자 검색",
            placeholder="검색어 입력",
            key="record_search",
        )

        team_filter = st.selectbox(
            "팀",
            ["전체"] + TEAM_OPTIONS,
            key="team_filter",
        )

    filtered = records.copy()

    if search_text.strip() and not filtered.empty:
        keyword = search_text.strip().lower()
        mask = (
            filtered["현장명"]
            .astype(str)
            .str.lower()
            .str.contains(
                keyword,
                na=False,
                regex=False,
            )
            |
            filtered["작성자"]
            .astype(str)
            .str.lower()
            .str.contains(
                keyword,
                na=False,
                regex=False,
            )
        )
        filtered = filtered[mask]

    if (
        team_filter != "전체"
        and not filtered.empty
    ):
        filtered = filtered[
            filtered["팀"] == team_filter
        ]

    if not filtered.empty:
        filtered = filtered.assign(
            _sort_key=(
                filtered["작업날짜"].astype(str)
                + " "
                + filtered["등록시간"].astype(str)
            )
        ).sort_values(
            "_sort_key",
            ascending=False,
        )

    refresh_col, download_col = st.columns(2)

    with refresh_col:
        if st.button(
            "새로고침",
            use_container_width=True,
            key="refresh_records",
        ):
            st.rerun()

    with download_col:
        st.download_button(
            "현재 목록 CSV",
            data=make_csv_bytes(
                filtered.drop(
                    columns=["_sort_key"],
                    errors="ignore",
                )
            ),
            file_name=(
                "현장_폭염_조치_기록_"
                f"{datetime.now(KST).strftime('%Y%m%d')}"
                ".csv"
            ),
            mime="text/csv",
            use_container_width=True,
            disabled=filtered.empty,
        )

    if filtered.empty:
        st.info("조건에 맞는 기록이 없습니다.")
        return

    admin_pin = get_secret(
        ("security", "admin_pin")
    )

    for _, row in filtered.iterrows():
        record = row.to_dict()
        record_id = clean_text(record.get("id"))
        site = (
            markdown_escape(record.get("현장명"))
            or "현장명 미입력"
        )
        work_date = markdown_escape(
            record.get("작업날짜")
        )
        temperature_text = temperature_display(
            record.get("체감온도")
        )
        temperature_label = (
            f" · {temperature_text}℃"
            if temperature_text
            else ""
        )

        with st.container(border=True):
            status_class = heat_level_class(
                temperature_text
            )
            team_text = (
                clean_text(record.get("팀"))
                or "팀 미입력"
            )
            author_text = (
                clean_text(record.get("작성자"))
                or "-"
            )
            work_time_text = (
                f"{clean_text(record.get('근무시작')) or '-'}"
                " ~ "
                f"{clean_text(record.get('근무종료')) or '-'}"
            )
            heat_time_text = (
                f"{clean_text(record.get('폭염시작')) or '-'}"
                " ~ "
                f"{clean_text(record.get('폭염종료')) or '-'}"
            )
            measure_text = clean_text(record.get("조치사항"))
            has_10_minute_rest = MEASURE_OPTIONS[0] in measure_text
            has_20_minute_rest = MEASURE_OPTIONS[1] in measure_text
            if has_10_minute_rest and has_20_minute_rest:
                rest_action_text = "10분·20분 휴식조치"
            elif has_10_minute_rest:
                rest_action_text = "10분 휴식조치"
            elif has_20_minute_rest:
                rest_action_text = "20분 휴식조치"
            else:
                rest_action_text = "-"

            st.markdown(f"### {site}")

            pill_text = (
                f"{heat_level(temperature_text)}"
                f"{temperature_label}"
            )
            st.markdown(
                (
                    f'<span class="status-pill '
                    f'{status_class}">'
                    f'{html.escape(pill_text)}'
                    "</span>"
                ),
                unsafe_allow_html=True,
            )

            st.caption(
                f"{work_date} · {team_text} · "
                f"작성자 {author_text}"
            )

            st.markdown(
                f"""
                <div class="record-details">
                    <div class="record-detail">
                        <span>근무 시간</span>
                        <b>{html.escape(work_time_text)}</b>
                    </div>
                    <div class="record-detail">
                        <span>폭염 노출</span>
                        <b>{html.escape(heat_time_text)}</b>
                    </div>
                    <div class="record-detail">
                        <span>휴식 조치</span>
                        <b>{html.escape(rest_action_text)}</b>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            st.write(
                "**조치사항**  "
                f"{clean_text(record.get('조치사항')) or '-'}"
            )

            if clean_text(record.get("특이사항")):
                st.write(
                    "**특이사항**  "
                    f"{clean_text(record.get('특이사항'))}"
                )

            st.caption(
                "등록 "
                f"{clean_text(record.get('등록시간'))}"
                " · 수정 "
                f"{clean_text(record.get('수정시간'))}"
            )

            edit_col, delete_col = st.columns(2)

            with edit_col:
                if st.button(
                    "수정",
                    use_container_width=True,
                    key=f"edit_{record_id}",
                ):
                    st.session_state.editing_id = (
                        record_id
                    )
                    st.session_state.form_nonce += 1
                    st.session_state.page = "form"
                    st.session_state.pending_delete = (
                        None
                    )
                    st.rerun()

            with delete_col:
                if st.button(
                    "삭제",
                    use_container_width=True,
                    key=f"delete_{record_id}",
                ):
                    st.session_state.pending_delete = (
                        record_id
                    )
                    st.rerun()

            if (
                st.session_state.pending_delete
                == record_id
            ):
                st.warning(
                    "이 기록을 삭제할까요? "
                    "삭제 후 복구는 Google Sheets "
                    "변경 기록에서만 가능합니다."
                )

                entered_pin = ""

                if admin_pin:
                    entered_pin = st.text_input(
                        "관리자 삭제 PIN",
                        type="password",
                        key=f"delete_pin_{record_id}",
                    )

                confirm_col, cancel_col = st.columns(2)

                with confirm_col:
                    if st.button(
                        "삭제 확인",
                        type="primary",
                        use_container_width=True,
                        key=f"confirm_{record_id}",
                    ):
                        if (
                            admin_pin
                            and entered_pin != admin_pin
                        ):
                            st.error(
                                "관리자 PIN이 "
                                "올바르지 않습니다."
                            )
                        else:
                            try:
                                delete_record(record_id)
                                st.session_state.pending_delete = (
                                    None
                                )
                                st.session_state.flash = (
                                    "기록을 삭제했습니다."
                                )
                                st.rerun()
                            except Exception as exc:  # noqa: BLE001
                                st.error(
                                    f"삭제에 실패했습니다: {exc}"
                                )

                with cancel_col:
                    if st.button(
                        "취소",
                        use_container_width=True,
                        key=f"cancel_{record_id}",
                    ):
                        st.session_state.pending_delete = (
                            None
                        )
                        st.rerun()


init_state()
render_header()
render_navigation()
show_flash()

records_df = empty_dataframe()
connection_error: Exception | None = None

try:
    records_df = load_records()
except Exception as exc:  # noqa: BLE001
    connection_error = exc
    render_setup_error(exc)

if st.session_state.page == "form":
    render_form(
        records_df,
        connection_error,
    )
else:
    render_records(
        records_df,
        connection_error,
    )
