from __future__ import annotations

import csv
import base64
import html
import json
import io
import logging
import math
import re
import urllib.parse
import urllib.request
import urllib.error
import uuid
import time as time_module
import zipfile
from pathlib import Path
from datetime import date, datetime, time, timedelta
from typing import Any
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


APP_TITLE = "현장 폭염 조치 기록"
APP_VERSION = "Professional UI v3.31 · 2026-08-12"
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
KMA_VILLAGE_FORECAST_API_URL = (
    "https://apihub.kma.go.kr/api/typ02/openApi/"
    "VilageFcstInfoService_2.0/getVilageFcst"
)
KAKAO_PLACE_API_URL = "https://dapi.kakao.com/v2/local/search/keyword.json"
FORECAST_HEAT_THRESHOLD = 31.0

COLUMNS = [
    "id",
    "작업날짜",
    "종목",
    "현장명",
    "팀",
    "근무시작",
    "근무종료",
    "작성자",
    "폭염시작",
    "폭염종료",
    "체감온도",
    "휴게시간",
    "공통 조치사항",
    "조치사항",
    "특이사항",
    "등록시간",
    "수정시간",
]

# v3.16 시트 구조: 공통 조치사항이 조치사항에 합쳐져 있던 상태
PRE_COMMON_COLUMNS = [
    "id",
    "작업날짜",
    "종목",
    "현장명",
    "팀",
    "근무시작",
    "근무종료",
    "작성자",
    "폭염시작",
    "폭염종료",
    "체감온도",
    "휴게시간",
    "조치사항",
    "특이사항",
    "등록시간",
    "수정시간",
]

LEGACY_COLUMNS = [
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

SPORT_OPTIONS = ["야구", "남자골프", "여자골프", "기타스포츠(실내)"]

SPORT_COMMON_MEASURES = {
    "남자골프": (
        "- 중계차, 중계룸, 카메라룸, W/L룸, 몽골 텐트 냉방 가동\n"
        "  (냉방 공간 체감온도 27℃ 이하 유지)\n"
        "- 개인별 아이스박스·우산 지급 및 생수·얼음물 비치\n"
        "- 식염포도당·폭염질환 응급키트 위치 공유 및 사용 안내\n"
        "- 폭염 시간대 불필요한 외부 활동 최소화\n"
        "- 외부 근무자 1시간 이내 10분 이상 휴식 부여"
    ),
    "여자골프": (
        "- 중계차, 중계룸, 카메라룸, W/L룸, 몽골 텐트 냉방 가동\n"
        "  (냉방 공간 체감온도 27℃ 이하 유지)\n"
        "- 개인별 아이스박스·우산 지급 및 생수·얼음물 비치\n"
        "- 식염포도당·폭염질환 응급키트 위치 공유 및 사용 안내\n"
        "- 폭염 시간대 불필요한 외부 활동 최소화\n"
        "- 외부 근무자 1시간 이내 10분 이상 휴식 부여"
    ),
    "야구": (
        "- 중계차·장비차·중계석·중계스태프실 냉방 가동\n"
        "  · 냉방 공간 체감온도 27℃ 이하 유지\n"
        "- 생수·냉음료·식염포도당·폭염질환 응급키트 비치 및 위치 공유\n"
        "- 폭염시간대 불필요한 야외활동 최소화\n"
        "- 이상 증상 발생 시 10~15분간 냉방 공간에서 휴식하도록\n"
        "  제작팀과 사전 협의"
    ),
    "기타스포츠(실내)": (
        "- 중계차, 휴게실, 체육관 냉방 가동\n"
        "  (냉방 공간 체감온도 27℃ 이하 유지)\n"
        "- 식염포도당, 폭염질환 응급키트 위치 공유 및 생수, 음료 지급\n"
        "- 폭염 시간대 불필요한 외부 활동 최소화\n"
        "- 중계차, 체육관 냉방 가동 공간에서 근무,\n"
        "  폭염 작업 해당 없음\n"
        "- 케이블 설치 등 외부 작업 완료\n"
        "  1시간 이내 10분 이상 휴게시간 부여"
    ),
}

# v3.16에서 조치사항에 함께 저장했던 기존 공통 문구.
# 마이그레이션 시 이 항목만 제거하고 실제 추가 조치는 보존합니다.
LEGACY_COMMON_MEASURES = {
    "야구": (
        "생수·냉수 비치 | 그늘·냉방 휴게공간 확인 | "
        "폭염시간대 업무강도 조정"
    ),
    "골프": (
        "생수·냉수 비치 | 그늘·냉방 휴게공간 확인 | "
        "코스 이동 동선 및 업무강도 조정"
    ),
    "남자골프": (
        "생수·냉수 비치 | 그늘·냉방 휴게공간 확인 | "
        "코스 이동 동선 및 업무강도 조정"
    ),
    "여자골프": (
        "생수·냉수 비치 | 그늘·냉방 휴게공간 확인 | "
        "코스 이동 동선 및 업무강도 조정"
    ),
    "기타스포츠": (
        "생수·냉수 비치 | 그늘·냉방 휴게공간 확인 | "
        "폭염시간대 업무강도 조정"
    ),
    "기타스포츠(실내)": (
        "생수·냉수 비치 | 그늘·냉방 휴게공간 확인 | "
        "폭염시간대 업무강도 조정"
    ),
}

MEASURE_OPTIONS = [
    "1시간 이내 10분 이상 휴식",
    "2시간 이내 20분 이상 휴식",
]

TIME_FIELD_COLUMNS = {
    "work_start": "근무시작",
    "work_end": "근무종료",
    "heat_start": "폭염시작",
    "heat_end": "폭염종료",
}


WEATHER_LOGGER = logging.getLogger("checktemp.weather")
if not WEATHER_LOGGER.handlers:
    _weather_handler = logging.StreamHandler()
    _weather_handler.setFormatter(
        logging.Formatter(
            "%(asctime)s [CHECKTEMP-WEATHER] %(levelname)s %(message)s"
        )
    )
    WEATHER_LOGGER.addHandler(_weather_handler)
WEATHER_LOGGER.setLevel(logging.INFO)
WEATHER_LOGGER.propagate = False


def weather_debug_log_key(nonce: int) -> str:
    return f"weather_debug_log_{nonce}"


def clear_weather_debug_log(nonce: int) -> None:
    st.session_state[weather_debug_log_key(nonce)] = []


def add_weather_debug_log(
    nonce: int | None,
    message: str,
    *,
    level: str = "info",
) -> None:
    """기상조회 진단 로그를 Streamlit 로그와 화면용 세션에 함께 남깁니다."""
    now_text = datetime.now(KST).strftime("%H:%M:%S")
    clean_message = clean_text(message).replace("\n", " ")
    rendered = f"[{now_text}] {clean_message}"

    if level == "error":
        WEATHER_LOGGER.error(clean_message)
    elif level == "warning":
        WEATHER_LOGGER.warning(clean_message)
    else:
        WEATHER_LOGGER.info(clean_message)

    if nonce is None:
        return

    key = weather_debug_log_key(nonce)
    existing = list(st.session_state.get(key, []))
    existing.append(rendered)
    st.session_state[key] = existing[-80:]


def safe_error_body(error: urllib.error.HTTPError) -> str:
    """HTTP 오류 응답을 인증정보 없이 짧게 기록합니다."""
    try:
        payload = error.read(1200)
    except Exception:
        return ""

    if not payload:
        return ""

    for encoding in ("utf-8", "euc-kr"):
        try:
            return payload.decode(encoding, errors="replace")[:1000]
        except Exception:
            continue
    return repr(payload[:500])


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
        color-scheme: light;
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
    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input,
    div[data-testid="stDateInput"] input,
    div[data-testid="stTimeInput"] input,
    div[data-testid="stTextArea"] textarea,
    textarea {
        border-radius: 7px !important;
        border-color: var(--line-strong) !important;
        background-color: #ffffff !important;
        color: #111827 !important;
        -webkit-text-fill-color: #111827 !important;
    }

    div[data-baseweb="input"] input,
    div[data-baseweb="textarea"] textarea,
    div[data-baseweb="select"] *,
    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input,
    div[data-testid="stDateInput"] input,
    div[data-testid="stTimeInput"] input,
    div[data-testid="stTextArea"] textarea {
        color: #111827 !important;
        -webkit-text-fill-color: #111827 !important;
        caret-color: #111827 !important;
    }

    input::placeholder,
    textarea::placeholder {
        color: #98a2b3 !important;
        -webkit-text-fill-color: #98a2b3 !important;
        opacity: 1 !important;
    }

    div[data-baseweb="popover"],
    div[data-baseweb="popover"] > div,
    ul[role="listbox"],
    li[role="option"] {
        background-color: #ffffff !important;
        color: #111827 !important;
        -webkit-text-fill-color: #111827 !important;
    }

    li[role="option"] *,
    div[data-baseweb="menu"] *,
    div[data-baseweb="calendar"] * {
        color: #111827 !important;
        -webkit-text-fill-color: #111827 !important;
    }

    div[data-baseweb="calendar"],
    div[data-baseweb="calendar"] > div,
    div[data-baseweb="calendar"] button {
        background-color: #ffffff !important;
    }

    div[data-baseweb="calendar"] button[aria-selected="true"] {
        background-color: var(--navy) !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }

    div[data-baseweb="calendar"] button[aria-selected="true"] * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
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
    """시간을 HH:MM 형식으로 해석합니다. 930→09:30, 18→18:00도 허용합니다."""
    text = clean_text(value)
    if not text:
        return default

    normalized = text.replace("시", ":00").replace(" ", "")

    if re.fullmatch(r"\d{1,2}", normalized):
        normalized = f"{normalized}:00"
    elif re.fullmatch(r"\d{3}", normalized):
        normalized = f"0{normalized[0]}:{normalized[1:]}"
    elif re.fullmatch(r"\d{4}", normalized):
        normalized = f"{normalized[:2]}:{normalized[2:]}"

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


def parse_kma_temperature_observations(
    response_text: str,
) -> list[tuple[str, float]]:
    """기상청 특정지점 ASCII 응답에서 체감온도 관측값을 추출합니다."""
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

    observations.sort(key=lambda item: item[0])
    return observations


def parse_kma_temperature_response(response_text: str) -> tuple[str, str]:
    """기상청 특정지점 ASCII 응답에서 최신 체감온도를 추출합니다."""
    observations = parse_kma_temperature_observations(response_text)

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


def fetch_kma_apparent_temperature_range(
    latitude: float,
    longitude: float,
    auth_key: str,
    range_start: datetime,
    range_end: datetime,
    *,
    debug_nonce: int | None = None,
) -> dict[str, Any]:
    """500m 격자 체감온도를 55분 이하 구간으로 독립 조회합니다.

    일부 구간이 실패해도 성공한 관측값은 보존합니다.
    연속 네트워크 타임아웃이 2회 발생하면 500m 조회를 중단하고
    초단기실황 조회로 넘어갈 수 있도록 부분 결과를 반환합니다.
    """
    now = datetime.now(KST).replace(second=0, microsecond=0)

    if range_start > now:
        raise ValueError("조회 시작시간이 현재보다 이후입니다.")

    actual_end = min(range_end, now)
    if actual_end < range_start:
        raise ValueError("조회할 시간대가 없습니다.")

    add_weather_debug_log(
        debug_nonce,
        (
            "500m 분할조회 시작 | "
            f"range={range_start.strftime('%Y-%m-%d %H:%M')}"
            f"~{actual_end.strftime('%Y-%m-%d %H:%M')} | "
            f"lat={latitude:.6f}, lon={longitude:.6f}"
        ),
    )

    observations_by_time: dict[str, float] = {}
    failed_chunks: list[str] = []
    requested_chunks = 0
    successful_chunks = 0
    consecutive_timeouts = 0

    chunk_start = range_start

    while chunk_start <= actual_end:
        # 공식 최대 조회기간보다 여유를 둔 55분 단위.
        chunk_end = min(
            chunk_start + timedelta(minutes=55),
            actual_end,
        )
        requested_chunks += 1
        chunk_label = (
            f"{chunk_start.strftime('%H:%M')}~"
            f"{chunk_end.strftime('%H:%M')}"
        )

        params = urllib.parse.urlencode(
            {
                "obs": "ta_chi",
                "tm1": chunk_start.strftime("%Y%m%d%H%M"),
                "tm2": chunk_end.strftime("%Y%m%d%H%M"),
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

        add_weather_debug_log(
            debug_nonce,
            (
                f"500m 구간 요청 {requested_chunks} | {chunk_label} | "
                "timeout=8s"
            ),
        )

        started = time_module.perf_counter()
        payload: bytes | None = None
        network_timeout = False

        try:
            with urllib.request.urlopen(request, timeout=8) as response:
                payload = response.read()
                elapsed = time_module.perf_counter() - started
                status = getattr(response, "status", None) or response.getcode()

            add_weather_debug_log(
                debug_nonce,
                (
                    f"500m 구간 성공 | {chunk_label} | status={status} | "
                    f"elapsed={elapsed:.2f}s | bytes={len(payload)}"
                ),
            )
            consecutive_timeouts = 0

        except urllib.error.HTTPError as error:
            elapsed = time_module.perf_counter() - started
            body = safe_error_body(error)
            failed_chunks.append(chunk_label)
            consecutive_timeouts = 0
            add_weather_debug_log(
                debug_nonce,
                (
                    f"500m 구간 HTTPError | {chunk_label} | "
                    f"code={error.code} | elapsed={elapsed:.2f}s"
                    + (f" | body={body}" if body else "")
                ),
                level="error",
            )

        except (urllib.error.URLError, TimeoutError) as error:
            elapsed = time_module.perf_counter() - started
            failed_chunks.append(chunk_label)
            consecutive_timeouts += 1
            network_timeout = True
            reason = clean_text(getattr(error, "reason", error))
            add_weather_debug_log(
                debug_nonce,
                (
                    f"500m 구간 타임아웃 | {chunk_label} | "
                    f"type={type(error).__name__} | reason={reason} | "
                    f"elapsed={elapsed:.2f}s | "
                    f"consecutive={consecutive_timeouts}"
                ),
                level="error",
            )

        except Exception as error:
            elapsed = time_module.perf_counter() - started
            failed_chunks.append(chunk_label)
            consecutive_timeouts = 0
            add_weather_debug_log(
                debug_nonce,
                (
                    f"500m 구간 예외 | {chunk_label} | "
                    f"type={type(error).__name__} | "
                    f"elapsed={elapsed:.2f}s | message={error}"
                ),
                level="error",
            )

        if payload is not None:
            response_text = ""
            for encoding in ("utf-8", "euc-kr"):
                try:
                    response_text = payload.decode(encoding)
                    break
                except UnicodeDecodeError:
                    continue

            if not response_text:
                response_text = payload.decode(
                    "utf-8",
                    errors="replace",
                )

            try:
                chunk_observations = parse_kma_temperature_observations(
                    response_text
                )
                for timestamp, value in chunk_observations:
                    observations_by_time[timestamp] = value

                successful_chunks += 1
                add_weather_debug_log(
                    debug_nonce,
                    (
                        f"500m 구간 파싱 성공 | {chunk_label} | "
                        f"observations={len(chunk_observations)}"
                    ),
                )
            except Exception as error:
                failed_chunks.append(chunk_label)
                preview = re.sub(r"\s+", " ", response_text[:600])
                add_weather_debug_log(
                    debug_nonce,
                    (
                        f"500m 구간 파싱 실패 | {chunk_label} | "
                        f"type={type(error).__name__} | "
                        f"message={error} | preview={preview}"
                    ),
                    level="error",
                )

        # 서버가 연속으로 응답하지 않으면 이후 구간까지 오래 기다리지 않습니다.
        if network_timeout and consecutive_timeouts >= 2:
            add_weather_debug_log(
                debug_nonce,
                (
                    "500m 회로차단 | 연속 2개 구간 타임아웃으로 "
                    "나머지 500m 조회를 중단하고 지역실황으로 진행"
                ),
                level="warning",
            )
            break

        if chunk_end >= actual_end:
            break

        # itv=5에 맞춰 다음 관측 시점부터 시작.
        chunk_start = chunk_end + timedelta(minutes=5)

    values = list(observations_by_time.values())

    if not values:
        add_weather_debug_log(
            debug_nonce,
            (
                f"500m 분할조회 결과 없음 | requested={requested_chunks} | "
                f"failed={len(failed_chunks)}"
            ),
            level="warning",
        )
        return {
            "available": False,
            "source": "500m 격자",
            "minimum": None,
            "maximum": None,
            "actual_end": actual_end.strftime("%H:%M"),
            "observations": 0,
            "requested": requested_chunks,
            "successful": successful_chunks,
            "failed": failed_chunks,
        }

    minimum = min(values)
    maximum = max(values)

    add_weather_debug_log(
        debug_nonce,
        (
            f"500m 분할조회 완료 | min={minimum:.1f} | max={maximum:.1f} | "
            f"observations={len(values)} | "
            f"chunks={successful_chunks}/{requested_chunks}"
        ),
    )

    return {
        "available": True,
        "source": "500m 격자",
        "minimum": minimum,
        "maximum": maximum,
        "actual_end": actual_end.strftime("%H:%M"),
        "observations": len(values),
        "requested": requested_chunks,
        "successful": successful_chunks,
        "failed": failed_chunks,
    }


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


def latest_village_forecast_base(now: datetime) -> datetime:
    """현재 시각에 이미 발표된 최신 단기예보 기준시각을 구합니다."""
    base_hours = (2, 5, 8, 11, 14, 17, 20, 23)
    # 단기예보는 기준시각 약 10분 뒤 제공되므로 15분의 여유를 둡니다.
    available_before = now - timedelta(minutes=15)
    candidates = [
        available_before.replace(hour=hour, minute=0, second=0, microsecond=0)
        for hour in base_hours
        if hour <= available_before.hour
    ]
    if candidates:
        return candidates[-1]
    previous_day = available_before - timedelta(days=1)
    return previous_day.replace(hour=23, minute=0, second=0, microsecond=0)


def fetch_kma_forecast_apparent_temperature_range(
    latitude: float,
    longitude: float,
    auth_key: str,
    range_start: datetime,
    range_end: datetime,
) -> tuple[str, str, str | None, str | None, int]:
    """단기예보로 근무시간의 예상 체감온도와 예상 폭염시간을 구합니다."""
    grid_x, grid_y = latitude_longitude_to_grid(latitude, longitude)
    base_time = latest_village_forecast_base(datetime.now(KST))
    params = urllib.parse.urlencode(
        {
            "pageNo": "1",
            "numOfRows": "2000",
            "dataType": "JSON",
            "base_date": base_time.strftime("%Y%m%d"),
            "base_time": base_time.strftime("%H%M"),
            "nx": str(grid_x),
            "ny": str(grid_y),
            "authKey": auth_key,
        }
    )
    request = urllib.request.Request(
        f"{KMA_VILLAGE_FORECAST_API_URL}?{params}",
        headers={"User-Agent": "checktemp-streamlit/1.0"},
    )

    with urllib.request.urlopen(request, timeout=20) as response:
        result = json.loads(response.read().decode("utf-8"))

    response_body = result.get("response") or {}
    header = response_body.get("header") or {}
    if clean_text(header.get("resultCode")) != "00":
        raise ValueError(
            clean_text(header.get("resultMsg")) or "단기예보 응답 오류"
        )

    raw_items = (
        response_body.get("body", {}).get("items", {}).get("item", [])
    )
    hourly_values: dict[datetime, dict[str, float]] = {}
    for item in raw_items:
        category = clean_text(item.get("category"))
        if category not in {"TMP", "REH"}:
            continue
        forecast_date = clean_text(item.get("fcstDate"))
        forecast_time = clean_text(item.get("fcstTime")).zfill(4)
        try:
            forecast_at = datetime.strptime(
                f"{forecast_date}{forecast_time}", "%Y%m%d%H%M"
            ).replace(tzinfo=KST)
        except ValueError:
            continue
        if not (range_start <= forecast_at <= range_end):
            continue
        value = parse_float(item.get("fcstValue"))
        if value is not None:
            hourly_values.setdefault(forecast_at, {})[category] = value

    forecasts: list[tuple[datetime, float]] = []
    for forecast_at, values in sorted(hourly_values.items()):
        air_temperature = values.get("TMP")
        relative_humidity = values.get("REH")
        if air_temperature is None or relative_humidity is None:
            continue
        if not (-50 <= air_temperature <= 60 and 0 <= relative_humidity <= 100):
            continue
        forecasts.append(
            (
                forecast_at,
                calculate_summer_apparent_temperature(
                    air_temperature,
                    relative_humidity,
                ),
            )
        )

    if not forecasts:
        raise ValueError(
            "선택한 날짜·근무시간은 현재 제공되는 단기예보 범위에 없습니다."
        )

    temperatures = [temperature for _, temperature in forecasts]
    heat_forecasts = [
        forecast_at
        for forecast_at, temperature in forecasts
        if temperature >= FORECAST_HEAT_THRESHOLD
    ]
    expected_start: str | None = None
    expected_end: str | None = None
    if heat_forecasts:
        heat_start = heat_forecasts[0]
        heat_end = min(heat_forecasts[-1] + timedelta(hours=1), range_end)
        expected_start = heat_start.strftime("%H:%M")
        expected_end = heat_end.strftime("%H:%M")

    return (
        format_number(min(temperatures)),
        format_number(max(temperatures)),
        expected_start,
        expected_end,
        len(forecasts),
    )


def fetch_kma_regional_apparent_temperature_at(
    latitude: float,
    longitude: float,
    auth_key: str,
    base_time: datetime,
    *,
    debug_nonce: int | None = None,
) -> tuple[float, float, float]:
    """지정 정시 초단기실황(T1H·REH)으로 체감온도를 계산합니다."""
    grid_x, grid_y = latitude_longitude_to_grid(latitude, longitude)
    base_time = base_time.astimezone(KST).replace(
        minute=0,
        second=0,
        microsecond=0,
    )

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

    started = time_module.perf_counter()

    try:
        with urllib.request.urlopen(request, timeout=6) as response:
            result = json.loads(response.read().decode("utf-8"))
            elapsed = time_module.perf_counter() - started
            status = getattr(response, "status", None) or response.getcode()

        add_weather_debug_log(
            debug_nonce,
            (
                f"지역실황 HTTP 성공 | base={base_time.strftime('%Y%m%d %H:%M')} | "
                f"status={status} | elapsed={elapsed:.2f}s"
            ),
        )

    except Exception as error:
        elapsed = time_module.perf_counter() - started
        add_weather_debug_log(
            debug_nonce,
            (
                f"지역실황 HTTP 실패 | base={base_time.strftime('%Y%m%d %H:%M')} | "
                f"type={type(error).__name__} | elapsed={elapsed:.2f}s | "
                f"message={error}"
            ),
            level="error",
        )
        raise

    response_body = result.get("response") or {}
    header = response_body.get("header") or {}

    if clean_text(header.get("resultCode")) != "00":
        result_message = (
            clean_text(header.get("resultMsg"))
            or "초단기실황 응답 오류"
        )
        raise ValueError(result_message)

    raw_items = (
        response_body.get("body", {})
        .get("items", {})
        .get("item", [])
    )

    values = {
        clean_text(item.get("category")): parse_float(item.get("obsrValue"))
        for item in raw_items
    }

    air_temperature = values.get("T1H")
    relative_humidity = values.get("REH")

    if air_temperature is None or relative_humidity is None:
        raise ValueError("초단기실황에 기온 또는 습도 값이 없습니다.")

    if not (-50 <= air_temperature <= 60):
        raise ValueError("초단기실황 기온 값이 정상 범위를 벗어났습니다.")

    if not (0 <= relative_humidity <= 100):
        raise ValueError("초단기실황 습도 값이 정상 범위를 벗어났습니다.")

    apparent = calculate_summer_apparent_temperature(
        air_temperature,
        relative_humidity,
    )

    add_weather_debug_log(
        debug_nonce,
        (
            f"지역실황 값 | {base_time.strftime('%H:%M')} | "
            f"T1H={air_temperature:.1f} | REH={relative_humidity:.1f} | "
            f"apparent={apparent:.1f}"
        ),
    )

    return apparent, air_temperature, relative_humidity


def fetch_kma_regional_apparent_temperature_range(
    latitude: float,
    longitude: float,
    auth_key: str,
    range_start: datetime,
    range_end: datetime,
    *,
    debug_nonce: int | None = None,
) -> dict[str, Any]:
    """근무시간 안의 정시별 지역실황을 독립 조회해 체감온도 범위를 계산합니다."""
    now = datetime.now(KST).replace(second=0, microsecond=0)

    if range_start > now:
        raise ValueError("조회 시작시간이 현재보다 이후입니다.")

    actual_end = min(range_end, now)
    if actual_end < range_start:
        raise ValueError("조회할 시간대가 없습니다.")

    latest_available = now.replace(
        minute=0,
        second=0,
        microsecond=0,
    )
    if now.minute < 15:
        latest_available -= timedelta(hours=1)

    # 시작 이후 첫 정시부터 사용합니다.
    first_hour = range_start.replace(
        minute=0,
        second=0,
        microsecond=0,
    )
    if range_start.minute > 0:
        first_hour += timedelta(hours=1)

    last_hour = min(
        actual_end.replace(
            minute=0,
            second=0,
            microsecond=0,
        ),
        latest_available,
    )

    # 1시간보다 짧고 범위 안에 정시가 하나도 없으면
    # 시작 직전 정시를 대표값으로 1건 사용합니다.
    if first_hour > last_hour:
        fallback_hour = min(
            range_start.replace(
                minute=0,
                second=0,
                microsecond=0,
            ),
            latest_available,
        )
        sample_times = [fallback_hour]
    else:
        sample_times = []
        cursor = first_hour
        while cursor <= last_hour:
            sample_times.append(cursor)
            cursor += timedelta(hours=1)

    add_weather_debug_log(
        debug_nonce,
        (
            "지역실황 범위조회 시작 | "
            f"samples={len(sample_times)} | "
            f"from={sample_times[0].strftime('%Y-%m-%d %H:%M')} | "
            f"to={sample_times[-1].strftime('%Y-%m-%d %H:%M')}"
        ),
    )

    apparent_values: list[float] = []
    failed_hours: list[str] = []
    consecutive_failures = 0

    for sample_time in sample_times:
        label = sample_time.strftime("%H:%M")
        try:
            apparent, _, _ = fetch_kma_regional_apparent_temperature_at(
                latitude,
                longitude,
                auth_key,
                sample_time,
                debug_nonce=debug_nonce,
            )
            apparent_values.append(apparent)
            consecutive_failures = 0

        except Exception:
            failed_hours.append(label)
            consecutive_failures += 1

            # 지역실황 서버 자체가 연속으로 응답하지 않을 때도
            # 지나치게 오래 대기하지 않습니다.
            if consecutive_failures >= 2 and not apparent_values:
                add_weather_debug_log(
                    debug_nonce,
                    (
                        "지역실황 회로차단 | 첫 2개 시간 연속 실패로 "
                        "나머지 시간 조회 중단"
                    ),
                    level="warning",
                )
                break

    if not apparent_values:
        add_weather_debug_log(
            debug_nonce,
            (
                f"지역실황 범위조회 결과 없음 | "
                f"requested={len(sample_times)} | failed={len(failed_hours)}"
            ),
            level="warning",
        )
        return {
            "available": False,
            "source": "초단기실황",
            "minimum": None,
            "maximum": None,
            "actual_end": actual_end.strftime("%H:%M"),
            "observations": 0,
            "requested": len(sample_times),
            "successful": 0,
            "failed": failed_hours,
        }

    minimum = min(apparent_values)
    maximum = max(apparent_values)

    add_weather_debug_log(
        debug_nonce,
        (
            f"지역실황 범위조회 완료 | min={minimum:.1f} | "
            f"max={maximum:.1f} | "
            f"hours={len(apparent_values)}/{len(sample_times)}"
        ),
    )

    return {
        "available": True,
        "source": "초단기실황",
        "minimum": minimum,
        "maximum": maximum,
        "actual_end": actual_end.strftime("%H:%M"),
        "observations": len(apparent_values),
        "requested": len(sample_times),
        "successful": len(apparent_values),
        "failed": failed_hours,
    }


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
    clear_weather_debug_log(nonce)
    add_weather_debug_log(
        nonce,
        f"조회 트리거 | record_start_time={record_start_time}",
    )
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

    heat_start_text = clean_text(
        st.session_state.get(f"manual_heat_start_{nonce}")
    )
    heat_end_text = clean_text(
        st.session_state.get(f"manual_heat_end_{nonce}")
    )
    work_start_text = clean_text(
        st.session_state.get(f"manual_work_start_{nonce}")
    )
    work_end_text = clean_text(
        st.session_state.get(f"manual_work_end_{nonce}")
    )

    selected_date = st.session_state.get(f"date_{nonce}")
    if not isinstance(selected_date, date):
        selected_date = datetime.now(KST).date()

    add_weather_debug_log(
        nonce,
        (
            f"입력값 | date={selected_date.isoformat()} | site={site_name} | "
            f"matched_site={matched_name} | "
            f"work={work_start_text or '-'}~{work_end_text or '-'} | "
            f"heat={heat_start_text or '-'}~{heat_end_text or '-'} | "
            f"coords={coordinates[0]:.6f},{coordinates[1]:.6f}"
        ),
    )

    if selected_date > datetime.now(KST).date():
        work_start_value = parse_time_value(work_start_text)
        work_end_value = parse_time_value(work_end_text)
        if work_start_value is None or work_end_value is None:
            st.session_state[notice_key] = (
                "warning",
                "미래 날짜의 예상 폭염시간을 조회하려면 근무 시작과 종료를 "
                "HH:MM 형식으로 모두 입력해 주세요.",
            )
            return

        forecast_start = datetime.combine(
            selected_date,
            work_start_value,
            tzinfo=KST,
        )
        forecast_end = datetime.combine(
            selected_date,
            work_end_value,
            tzinfo=KST,
        )
        if forecast_end < forecast_start:
            forecast_end += timedelta(days=1)

        try:
            (
                minimum,
                maximum,
                expected_start,
                expected_end,
                forecast_count,
            ) = fetch_kma_forecast_apparent_temperature_range(
                coordinates[0],
                coordinates[1],
                auth_key,
                forecast_start,
                forecast_end,
            )
        except Exception as error:  # noqa: BLE001
            if isinstance(error, urllib.error.HTTPError) and error.code == 403:
                error_detail = (
                    "현재 인증키에 '4.3 단기예보조회' 권한이 없습니다. "
                    "기상청 API허브에서 해당 API를 활용신청한 뒤 다시 조회해 "
                    "주세요. 기존에 승인된 2.2 초단기예보는 약 6시간 범위라 "
                    "미래 날짜 전체 조회에는 사용할 수 없습니다."
                )
            else:
                error_detail = (
                    "날짜가 단기예보 범위인지 확인해 주세요. "
                    f"({error})"
                )
            st.session_state[notice_key] = (
                "warning",
                "선택한 미래 날짜의 예상 체감온도 조회에 실패했습니다. "
                f"{error_detail}",
            )
            return

        temperature_range = (
            minimum if minimum == maximum else f"{minimum}~{maximum}"
        )
        st.session_state[f"temperature_{nonce}"] = temperature_range
        st.session_state[f"manual_heat_start_{nonce}"] = expected_start or ""
        st.session_state[f"manual_heat_end_{nonce}"] = expected_end or ""

        if expected_start and expected_end:
            heat_message = (
                f"체감온도 {format_number(FORECAST_HEAT_THRESHOLD)}℃ 이상 예상시간 "
                f"{expected_start}~{expected_end}를 예상 폭염시간으로 입력했습니다."
            )
        else:
            heat_message = (
                f"근무시간 중 체감온도 {format_number(FORECAST_HEAT_THRESHOLD)}℃ "
                "이상으로 예상되는 시간이 없어 폭염시간을 비워 두었습니다."
            )
        st.session_state[notice_key] = (
            "success",
            f"{matched_name} · {selected_date.strftime('%Y-%m-%d')} 단기예보 "
            f"체감온도 최저 {minimum}℃ · 최고 {maximum}℃를 "
            f"{forecast_count}개 시간자료로 확인했습니다. {heat_message} "
            "예보는 변경될 수 있으므로 당일 또는 작업 후 다시 조회해 주세요.",
        )
        return

    range_label = ""
    range_start_text = ""
    range_end_text = ""
    if heat_start_text and heat_end_text:
        range_label = "폭염시간"
        range_start_text = heat_start_text
        range_end_text = heat_end_text
    elif heat_start_text or heat_end_text:
        st.session_state[notice_key] = (
            "warning",
            "폭염시간을 사용하려면 폭염 시작과 종료를 모두 입력해 주세요.",
        )
        return
    elif work_start_text and work_end_text:
        range_label = "근무시간"
        range_start_text = work_start_text
        range_end_text = work_end_text
    elif work_start_text or work_end_text:
        st.session_state[notice_key] = (
            "warning",
            "폭염시간이 비어 있으면 근무시간을 사용합니다. 근무 시작과 "
            "종료를 모두 입력해 주세요.",
        )
        return

    if range_label:
        range_start_value = parse_time_value(range_start_text)
        range_end_value = parse_time_value(range_end_text)
        if range_start_value is None or range_end_value is None:
            st.session_state[notice_key] = (
                "warning",
                f"{range_label} 시작과 종료를 HH:MM 형식으로 입력해 주세요.",
            )
            return

        range_start = datetime.combine(
            selected_date,
            range_start_value,
            tzinfo=KST,
        )
        range_end = datetime.combine(
            selected_date,
            range_end_value,
            tzinfo=KST,
        )
        if range_end < range_start:
            range_end += timedelta(days=1)
            add_weather_debug_log(
                nonce,
                "자정 경과 인식 | 종료시간을 익일로 처리",
            )

        add_weather_debug_log(
            nonce,
            (
                f"{range_label} 해석 완료 | "
                f"start={range_start.strftime('%Y-%m-%d %H:%M')} | "
                f"end={range_end.strftime('%Y-%m-%d %H:%M')}"
            ),
        )

        grid_result: dict[str, Any] | None = None
        regional_result: dict[str, Any] | None = None
        range_errors: list[str] = []

        try:
            grid_result = fetch_kma_apparent_temperature_range(
                coordinates[0],
                coordinates[1],
                auth_key,
                range_start,
                range_end,
                debug_nonce=nonce,
            )
        except Exception as error:  # noqa: BLE001
            range_errors.append(f"500m 격자: {error}")
            add_weather_debug_log(
                nonce,
                (
                    f"500m 소스 예외 종료 | type={type(error).__name__} | "
                    f"message={error}"
                ),
                level="error",
            )

        try:
            regional_result = fetch_kma_regional_apparent_temperature_range(
                coordinates[0],
                coordinates[1],
                auth_key,
                range_start,
                range_end,
                debug_nonce=nonce,
            )
        except Exception as error:  # noqa: BLE001
            range_errors.append(f"초단기실황: {error}")
            add_weather_debug_log(
                nonce,
                (
                    f"지역실황 소스 예외 종료 | type={type(error).__name__} | "
                    f"message={error}"
                ),
                level="error",
            )

        available_results = [
            result
            for result in (grid_result, regional_result)
            if result and result.get("available")
        ]

        if not available_results:
            error_detail = (
                " / ".join(range_errors)
                if range_errors
                else "두 기상자료 모두 유효한 값을 받지 못했습니다."
            )
            st.session_state[notice_key] = (
                "warning",
                f"입력한 {range_label}의 체감온도 조회에 실패했습니다. "
                f"{error_detail} 기상조회 진단 로그를 확인해 주세요.",
            )
            return

        # 안전 우선: 최고 체감온도가 더 높은 소스의 전체 범위를 적용합니다.
        # 최고값이 같으면 최저값이 더 높은 결과를 우선합니다.
        selected_result = max(
            available_results,
            key=lambda result: (
                float(result["maximum"]),
                float(result["minimum"]),
            ),
        )

        selected_minimum = float(selected_result["minimum"])
        selected_maximum = float(selected_result["maximum"])

        minimum_text = format_number(selected_minimum)
        maximum_text = format_number(selected_maximum)
        temperature_range = (
            minimum_text
            if minimum_text == maximum_text
            else f"{minimum_text}~{maximum_text}"
        )

        st.session_state[f"temperature_{nonce}"] = temperature_range

        requested_end_text = range_end.strftime("%H:%M")
        actual_end_text = clean_text(selected_result.get("actual_end"))
        end_note = (
            ""
            if actual_end_text == requested_end_text
            else f" (현재 조회 가능한 {actual_end_text}까지)"
        )

        source_details: list[str] = []

        if grid_result:
            if grid_result.get("available"):
                grid_min = format_number(float(grid_result["minimum"]))
                grid_max = format_number(float(grid_result["maximum"]))
                grid_range = (
                    grid_min
                    if grid_min == grid_max
                    else f"{grid_min}~{grid_max}"
                )
                source_details.append(
                    "500m 격자 "
                    f"{grid_range}℃ "
                    f"({grid_result['successful']}/{grid_result['requested']}구간, "
                    f"{grid_result['observations']}개 관측)"
                )
            else:
                source_details.append("500m 격자 조회 실패")

        if regional_result:
            if regional_result.get("available"):
                regional_min = format_number(
                    float(regional_result["minimum"])
                )
                regional_max = format_number(
                    float(regional_result["maximum"])
                )
                regional_range = (
                    regional_min
                    if regional_min == regional_max
                    else f"{regional_min}~{regional_max}"
                )
                source_details.append(
                    "초단기실황 "
                    f"{regional_range}℃ "
                    f"({regional_result['successful']}/"
                    f"{regional_result['requested']}시간)"
                )
            else:
                source_details.append("초단기실황 조회 실패")

        if range_errors:
            source_details.append(
                f"일부 소스 오류: {' / '.join(range_errors)}"
            )

        selected_source = clean_text(selected_result.get("source"))

        add_weather_debug_log(
            nonce,
            (
                f"안전 우선 최종선택 | source={selected_source} | "
                f"range={temperature_range} | "
                f"max={maximum_text}"
            ),
        )

        st.session_state[notice_key] = (
            "success",
            f"{matched_name} · {range_label} "
            f"{range_start.strftime('%H:%M')}~"
            f"{requested_end_text}{end_note} · "
            f"{' / '.join(source_details)} · "
            f"안전을 위해 최고 체감온도가 더 높은 "
            f"{selected_source} 결과 {temperature_range}℃를 자동 적용했습니다.",
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


def auto_lookup_future_weather(nonce: int) -> None:
    selected_date = st.session_state.get(f"date_{nonce}")
    site_name = clean_text(st.session_state.get(f"site_{nonce}"))
    work_start = clean_text(
        st.session_state.get(f"manual_work_start_{nonce}")
    )
    work_end = clean_text(
        st.session_state.get(f"manual_work_end_{nonce}")
    )

    if (
        not isinstance(selected_date, date)
        or selected_date <= datetime.now(KST).date()
        or not site_name
        or parse_time_value(work_start) is None
        or parse_time_value(work_end) is None
    ):
        return

    request_signature = (
        selected_date.isoformat(),
        site_name,
        work_start,
        work_end,
    )
    signature_key = f"auto_forecast_signature_{nonce}"
    if st.session_state.get(signature_key) == request_signature:
        return

    st.session_state[signature_key] = request_signature
    record_heat_start_with_weather(nonce, False)


def normalize_manual_time_field(
    field: str,
    nonce: int,
    *,
    trigger_weather: bool = False,
) -> None:
    """직접 입력한 시간을 HH:MM으로 정규화하고 필요 시 예보 조회를 실행합니다.

    예: 2256 -> 22:56, 930 -> 09:30, 18 -> 18:00
    """
    key = f"manual_{field}_{nonce}"
    raw_value = clean_text(st.session_state.get(key))

    if not raw_value:
        if trigger_weather:
            auto_lookup_future_weather(nonce)
        return

    parsed = parse_time_value(raw_value)
    if parsed is None:
        # 잘못된 값은 사용자가 수정할 수 있도록 그대로 둡니다.
        return

    # callback 단계에서 widget state를 먼저 정규화하므로
    # Enter 후 화면에도 즉시 HH:MM 형식으로 표시됩니다.
    st.session_state[key] = format_time_value(parsed)

    if trigger_weather:
        auto_lookup_future_weather(nonce)


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


def column_letter(index: int) -> str:
    result = ""
    value = index
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(65 + remainder) + result
    return result


def infer_sport(site_name: Any) -> str:
    site = clean_text(site_name)
    if any(keyword in site for keyword in ("야구", "구장", "베이스볼")):
        return "야구"
    if any(keyword in site.upper() for keyword in ("KLPGA", "여자골프")):
        return "여자골프"
    if (
        "KPGA" in site.upper()
        or any(keyword in site for keyword in ("골프", "CC", "cc", "컨트리클럽"))
    ):
        return "남자골프"
    return "기타스포츠(실내)"


def normalize_sport(value: Any, site_name: Any = "") -> str:
    sport = clean_text(value)
    if sport in ("기타", "기타스포츠"):
        return "기타스포츠(실내)"
    if sport == "골프":
        return infer_sport(site_name) if site_name else "남자골프"
    if sport in SPORT_OPTIONS:
        return sport
    return infer_sport(site_name)


def common_measures_for_sport(sport: Any) -> str:
    normalized = normalize_sport(sport)
    return SPORT_COMMON_MEASURES.get(
        normalized,
        SPORT_COMMON_MEASURES["기타스포츠(실내)"],
    )


def strip_legacy_common_measures(sport: Any, value: Any) -> str:
    """기존 조치사항에서 v3.16 공통조치 문구만 제거합니다."""
    text = clean_text(value)
    if not text:
        return ""

    normalized_sport = normalize_sport(sport)
    legacy = LEGACY_COMMON_MEASURES.get(normalized_sport, "")
    legacy_parts = {
        part.strip()
        for part in legacy.split("|")
        if part.strip()
    }
    parts = [
        part.strip()
        for part in text.split("|")
        if part.strip()
    ]
    return " | ".join(
        part for part in parts
        if part not in legacy_parts
    )


def migrate_sheet_to_current(
    worksheet: gspread.Worksheet,
    values: list[list[str]],
) -> None:
    backup_name = (
        f"{worksheet.title}_backup_"
        f"{datetime.now(KST).strftime('%Y%m%d_%H%M%S')}"
    )
    worksheet.duplicate(new_sheet_name=backup_name)

    old_headers = [clean_text(value) for value in values[0]]
    migrated_rows: list[list[str]] = [COLUMNS]

    for raw_row in values[1:]:
        padded = raw_row + [""] * max(
            0,
            len(old_headers) - len(raw_row),
        )
        old_record = {
            header: clean_text(padded[index])
            for index, header in enumerate(old_headers)
            if header
        }
        if not any(old_record.values()):
            continue

        sport = normalize_sport(
            old_record.get("종목"),
            old_record.get("현장명"),
        )
        new_record = {
            column: old_record.get(column, "")
            for column in COLUMNS
        }
        new_record["종목"] = sport
        new_record["공통 조치사항"] = (
            old_record.get("공통 조치사항")
            or common_measures_for_sport(sport)
        )
        new_record["조치사항"] = strip_legacy_common_measures(
            sport,
            old_record.get("조치사항"),
        )
        migrated_rows.append(record_values(new_record))

    worksheet.clear()
    worksheet.update(
        range_name="A1",
        values=migrated_rows,
        value_input_option="USER_ENTERED",
    )


def ensure_headers(worksheet: gspread.Worksheet) -> list[str]:
    headers = [
        clean_text(value)
        for value in worksheet.row_values(1)
    ]

    if not headers:
        worksheet.update(
            range_name=f"A1:{column_letter(len(COLUMNS))}1",
            values=[COLUMNS],
        )
        return COLUMNS

    if headers[: len(COLUMNS)] == COLUMNS:
        return headers

    if (
        headers[: len(PRE_COMMON_COLUMNS)] == PRE_COMMON_COLUMNS
        or headers[: len(LEGACY_COLUMNS)] == LEGACY_COLUMNS
    ):
        values = worksheet.get_all_values()
        migrate_sheet_to_current(worksheet, values)
        return COLUMNS

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


def _sheet_grid_borders(
    style: str = "SOLID",
) -> dict[str, Any]:
    """Google Sheets 표 전체 셀에 적용할 네 방향 테두리입니다."""
    color = {"red": 0.76, "green": 0.79, "blue": 0.84}
    return {
        edge: {
            "style": style,
            "color": color,
        }
        for edge in ("top", "bottom", "left", "right")
    }


def _data_row_format(row_number: int) -> dict[str, Any]:
    """데이터 행 구분용 교차 음영·전체 테두리·줄바꿈 서식입니다."""
    background = (
        {"red": 1.0, "green": 1.0, "blue": 1.0}
        if row_number % 2 == 0
        else {"red": 0.965, "green": 0.973, "blue": 0.984}
    )
    return {
        "backgroundColor": background,
        "textFormat": {
            "bold": False,
            "foregroundColor": {"red": 0.08, "green": 0.10, "blue": 0.14},
        },
        "verticalAlignment": "MIDDLE",
        "wrapStrategy": "WRAP",
        "borders": _sheet_grid_borders(),
    }


def _special_notes_format(row_number: int, has_notes: bool) -> dict[str, Any]:
    if has_notes:
        return {
            "backgroundColor": {"red": 1.0, "green": 0.93, "blue": 0.80},
            "textFormat": {
                "bold": True,
                "foregroundColor": {"red": 0.55, "green": 0.20, "blue": 0.05},
            },
            "verticalAlignment": "MIDDLE",
            "wrapStrategy": "WRAP",
            "borders": _sheet_grid_borders(),
        }

    base = _data_row_format(row_number)
    return {
        "backgroundColor": base["backgroundColor"],
        "textFormat": base["textFormat"],
        "verticalAlignment": "MIDDLE",
        "wrapStrategy": "WRAP",
        "borders": _sheet_grid_borders(),
    }


def format_special_notes_cell(
    worksheet: gspread.Worksheet,
    row_number: int,
    notes: Any,
) -> None:
    notes_column = column_letter(COLUMNS.index("특이사항") + 1)
    end_column = column_letter(len(COLUMNS))
    has_notes = bool(clean_text(notes))

    # 기록 행 전체에 교차 음영과 하단 테두리를 적용해 행 구분을 명확히 합니다.
    worksheet.format(
        f"A{row_number}:{end_column}{row_number}",
        _data_row_format(row_number),
    )

    # 특이사항이 있으면 해당 셀만 기존 주황색 강조를 유지합니다.
    worksheet.format(
        f"{notes_column}{row_number}",
        _special_notes_format(row_number, has_notes),
    )


def normalize_existing_special_notes_format(
    worksheet: gspread.Worksheet,
    values: list[list[str]],
) -> None:
    """기존 데이터 전체 서식 함수.

    앱 시작 시 호출하지 않습니다. 전체 행 서식은 느릴 수 있으므로
    신규/수정 행은 format_special_notes_cell()로 개별 적용합니다.
    """
    if not values:
        return

    session_key = (
        f"sheet_visual_v320_{worksheet.id}_{len(values)}_{len(COLUMNS)}"
    )
    if st.session_state.get(session_key):
        return

    headers = [clean_text(value) for value in values[0]]
    if "특이사항" not in headers:
        return

    notes_index = headers.index("특이사항")
    notes_column = column_letter(notes_index + 1)
    end_column = column_letter(len(COLUMNS))

    formats: list[dict[str, Any]] = [
        {
            "range": f"A1:{end_column}1",
            "format": {
                "backgroundColor": {"red": 0.09, "green": 0.17, "blue": 0.30},
                "textFormat": {
                    "bold": True,
                    "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "wrapStrategy": "WRAP",
                "borders": {
                    edge: {
                        "style": "SOLID_MEDIUM",
                        "color": {"red": 0.09, "green": 0.17, "blue": 0.30},
                    }
                    for edge in ("top", "bottom", "left", "right")
                },
            },
        }
    ]

    for row_number, raw_row in enumerate(values[1:], start=2):
        formats.append(
            {
                "range": f"A{row_number}:{end_column}{row_number}",
                "format": _data_row_format(row_number),
            }
        )

        notes = raw_row[notes_index] if notes_index < len(raw_row) else ""
        if clean_text(notes):
            formats.append(
                {
                    "range": f"{notes_column}{row_number}",
                    "format": _special_notes_format(row_number, True),
                }
            )

    worksheet.batch_format(formats)

    try:
        worksheet.freeze(rows=1)
    except Exception:  # noqa: BLE001
        pass

    st.session_state[session_key] = True

@st.cache_data(ttl=20, show_spinner=False)
def load_records() -> pd.DataFrame:
    """records 시트를 1회만 읽어 앱 로딩을 가볍게 유지합니다.

    기존 전체 행 서식 재적용은 앱 시작 경로에서 제거했습니다.
    새 기록/수정 행은 저장 시 해당 행만 서식 적용합니다.
    """
    worksheet = get_worksheet()
    values = worksheet.get_all_values()

    if not values:
        worksheet.update(
            range_name=f"A1:{column_letter(len(COLUMNS))}1",
            values=[COLUMNS],
        )
        return empty_dataframe()

    headers = [clean_text(value) for value in values[0]]

    if headers[: len(COLUMNS)] == COLUMNS:
        pass
    elif (
        headers[: len(PRE_COMMON_COLUMNS)] == PRE_COMMON_COLUMNS
        or headers[: len(LEGACY_COLUMNS)] == LEGACY_COLUMNS
    ):
        migrate_sheet_to_current(worksheet, values)
        # 마이그레이션은 예외적인 1회 작업이므로 이후에만 다시 읽습니다.
        values = worksheet.get_all_values()
    else:
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
    append_result = worksheet.append_row(
        record_values(record),
        value_input_option="USER_ENTERED",
        insert_data_option="INSERT_ROWS",
    )
    updated_range = clean_text(
        (append_result.get("updates") or {}).get("updatedRange")
        if isinstance(append_result, dict)
        else ""
    )
    row_match = re.search(r"(\d+)(?::[A-Z]+\d+)?$", updated_range)
    row_number = int(row_match.group(1)) if row_match else 0
    if row_number <= 0:
        appended_cell = worksheet.find(
            clean_text(record.get("id")),
            in_column=1,
        )
        row_number = appended_cell.row if appended_cell else 0
    try:
        if row_number > 0:
            format_special_notes_cell(
                worksheet,
                row_number,
                record.get("특이사항"),
            )
    except Exception:  # noqa: BLE001
        # 기록 저장은 성공했으므로 서식 오류 때문에 중복 저장되지 않게 합니다.
        pass

    load_records.clear()


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
        range_name=(
            f"A{cell.row}:"
            f"{column_letter(len(COLUMNS))}{cell.row}"
        ),
        values=[record_values(record)],
        value_input_option="USER_ENTERED",
    )
    try:
        format_special_notes_cell(
            worksheet,
            cell.row,
            record.get("특이사항"),
        )
    except Exception:  # noqa: BLE001
        pass

    load_records.clear()


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
    load_records.clear()


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
    text = (
        normalize_sport(value)
        if options == SPORT_OPTIONS
        else clean_text(value)
    )

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


def unique_texts(values: list[Any]) -> list[str]:
    results: list[str] = []
    for value in values:
        text = clean_text(value)
        if text and text not in results:
            results.append(text)
    return results


def report_team_summary(records: pd.DataFrame, team: str) -> str:
    team_rows = records[records["팀"] == team]
    if team_rows.empty:
        return "- 기록 없음"

    measures = unique_texts(team_rows["조치사항"].tolist())
    rest_total = sum(
        parse_int(value)
        for value in team_rows["휴게시간"].tolist()
    )
    parts = [item.replace(" | ", " / ") for item in measures]
    if rest_total > 0:
        parts.append(f"누적 휴게시간 {rest_total}분")
    return "- " + (" / ".join(parts) if parts else "조치사항 기록 없음")


def report_team_work_time(records: pd.DataFrame, team: str) -> str:
    team_rows = records[records["팀"] == team]
    if team_rows.empty:
        return f"{team} -"

    starts = sorted(unique_texts(team_rows["근무시작"].tolist()))
    ends = sorted(unique_texts(team_rows["근무종료"].tolist()))
    if not starts or not ends:
        return f"{team} -"
    return f"{team} {starts[0]}~{ends[-1]}"


def replace_report_text(xml_text: str, old: str, new: str) -> str:
    return xml_text.replace(old, html.escape(new, quote=False), 1)


def make_heat_report_bytes(records: pd.DataFrame) -> bytes:
    """같은 날짜·현장의 기록을 원본 HWPX 양식에 채웁니다."""
    if records.empty:
        raise ValueError("보고서로 만들 기록이 없습니다.")

    template_path = Path(__file__).with_name("report_template.b64")
    template_bytes = base64.b64decode(
        template_path.read_text(encoding="ascii").strip()
    )

    first = records.iloc[0]
    work_date = clean_text(first.get("작업날짜"))
    try:
        work_date_text = datetime.strptime(
            work_date,
            "%Y-%m-%d",
        ).strftime("%Y.%m.%d")
    except ValueError:
        work_date_text = work_date.replace("-", ".")

    site = clean_text(first.get("현장명")) or "현장명 미입력"
    sport = normalize_sport(first.get("종목"), site)
    site_text = f"{sport} / {site}"

    heat_starts = sorted(unique_texts(records["폭염시작"].tolist()))
    heat_ends = sorted(unique_texts(records["폭염종료"].tolist()))
    heat_time = (
        f"{heat_starts[0]}~{heat_ends[-1]}"
        if heat_starts and heat_ends
        else "해당 없음"
    )

    temperatures: list[float] = []
    for value in records["체감온도"].tolist():
        temperatures.extend(temperature_numbers(value))
    if temperatures:
        low = format_number(min(temperatures))
        high = format_number(max(temperatures))
        temperature_text = (
            f"{low}℃" if low == high else f"{low}℃~{high}℃"
        )
    else:
        temperature_text = "미기록"

    common_values = unique_texts(records["공통 조치사항"].tolist())
    common_text = common_values[0] if common_values else common_measures_for_sport(sport)
    common_lines = [
        line.strip().lstrip("- ").strip()
        for line in common_text.splitlines()
        if line.strip()
    ]
    while len(common_lines) < 3:
        common_lines.append("")

    notes = " / ".join(unique_texts(records["특이사항"].tolist()))
    notes = notes or "특이사항 없음"

    replacements = [
        ("종목 / 장소명", site_text),
        ("2026.08.10", work_date_text),
        ("중계팀 08:00~17:00", report_team_work_time(records, "중계팀")),
        ("영상팀 08:00~17:00", report_team_work_time(records, "영상팀")),
        ("08:00~15:00", heat_time),
        ("31℃~33℃", temperature_text),
        (
            "- 중계차, 중계석, 휴게실 냉방 가동(체감온도 OO℃ 이하 유지)",
            f"- {common_lines[0]}",
        ),
        (
            "식염포도당, 폭염질환 응급키트 위치 공유 및 생수 및 이온 음료",
            common_lines[1],
        ),
        (" 지급", ""),
        ("  - 폭염 시간대 불필요한 외부 활동 최소화", f"  - {common_lines[2]}"),
        ("영상팀 필드카메라 중요 선수 외 이글 또는 버디", notes),
        ("상황까지", ""),
        ("만", ""),
        ("                                     촬영하고 그 외 추가 휴식시간 부여", ""),
        ("  - 예시 2) 일정 조정 : 기존 출근시간 대비 1시간 조기 출근하여 실외 작업을 조기", ""),
        ("                       진행하고, 폭염시간대에는 1시간 이내 20분 이상 휴게시간을", ""),
        ("                       확보하여 운영", ""),
        ("- 중계차 및 중계석 등 냉방공간 근무 → ", report_team_summary(records, "중계팀")),
        ("폭염작업 해당 없음", ""),
        ("- 1시간 이내 10분 이상 휴게시간 부여", report_team_summary(records, "영상팀")),
        (" · 필드카메라 : 중요 선수 외 이글·버디 워킹 촬영 후 휴식", ""),
        (" · W/L 카메라 : 혹서기 지원인력을 활용한 교대근무 실시", ""),
    ]

    source = io.BytesIO(template_bytes)
    output = io.BytesIO()
    with zipfile.ZipFile(source, "r") as input_zip:
        with zipfile.ZipFile(output, "w") as output_zip:
            for info in input_zip.infolist():
                payload = input_zip.read(info.filename)
                if info.filename == "Contents/section0.xml":
                    section_xml = payload.decode("utf-8")
                    for old, new in replacements:
                        section_xml = replace_report_text(section_xml, old, new)
                    payload = section_xml.encode("utf-8")
                output_zip.writestr(info, payload)

    return output.getvalue()


def make_excel_bytes(dataframe: pd.DataFrame) -> bytes:
    """전체 기록을 관리용 XLSX 파일로 생성합니다."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "전체기록"

    header_fill = PatternFill("solid", fgColor="172B4D")
    header_font = Font(color="FFFFFF", bold=True)
    zebra_fill = PatternFill("solid", fgColor="F6F8FB")
    notes_fill = PatternFill("solid", fgColor="FFECCD")
    notes_font = Font(color="8C330D", bold=True)

    thin_side = Side(style="thin", color="C4CAD4")
    grid_border = Border(
        left=thin_side,
        right=thin_side,
        top=thin_side,
        bottom=thin_side,
    )

    for column_index, column_name in enumerate(COLUMNS, start=1):
        cell = worksheet.cell(row=1, column=column_index, value=column_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = grid_border
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    export_df = dataframe.reindex(columns=COLUMNS).fillna("")

    for row_index, (_, row) in enumerate(export_df.iterrows(), start=2):
        for column_index, column_name in enumerate(COLUMNS, start=1):
            value = clean_text(row.get(column_name, ""))
            cell = worksheet.cell(
                row=row_index,
                column=column_index,
                value=value,
            )
            cell.border = grid_border
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True,
            )
            if row_index % 2 == 1:
                cell.fill = zebra_fill

        notes_cell = worksheet.cell(
            row=row_index,
            column=COLUMNS.index("특이사항") + 1,
        )
        if clean_text(notes_cell.value):
            notes_cell.fill = notes_fill
            notes_cell.font = notes_font

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = (
        f"A1:{get_column_letter(len(COLUMNS))}{max(1, worksheet.max_row)}"
    )

    preferred_widths = {
        "id": 18,
        "작업날짜": 12,
        "종목": 12,
        "현장명": 24,
        "팀": 10,
        "근무시작": 10,
        "근무종료": 10,
        "작성자": 12,
        "폭염시작": 10,
        "폭염종료": 10,
        "체감온도": 12,
        "휴게시간": 10,
        "공통 조치사항": 48,
        "조치사항": 34,
        "특이사항": 36,
        "등록시간": 20,
        "수정시간": 20,
    }

    for index, column_name in enumerate(COLUMNS, start=1):
        worksheet.column_dimensions[
            get_column_letter(index)
        ].width = preferred_widths.get(column_name, 16)

    worksheet.row_dimensions[1].height = 26

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


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
            on_change=auto_lookup_future_weather,
            args=(nonce,),
        )

        sport = st.selectbox(
            "종목 *",
            SPORT_OPTIONS,
            index=option_index(
                SPORT_OPTIONS,
                normalize_sport(
                    editing_record.get("종목"),
                    editing_record.get("현장명"),
                ),
            ),
            key=f"sport_{nonce}",
        )

        site = st.text_input(
            "현장명 *",
            value=clean_text(
                editing_record.get("현장명")
            ),
            placeholder="예: ○○골프장, ○○야구장",
            key=f"site_{nonce}",
            on_change=auto_lookup_future_weather,
            args=(nonce,),
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
                placeholder="예: 2256 → 22:56",
                key=f"manual_work_start_{nonce}",
                on_change=normalize_manual_time_field,
                args=("work_start", nonce),
                kwargs={"trigger_weather": True},
            )

        with work_right:
            work_end_input = st.text_input(
                "근무 종료 *",
                value=clean_text(editing_record.get("근무종료")),
                placeholder="예: 1830 → 18:30",
                key=f"manual_work_end_{nonce}",
                on_change=normalize_manual_time_field,
                args=("work_end", nonce),
                kwargs={"trigger_weather": True},
            )

        st.caption(
            "숫자만 입력해도 됩니다. Enter를 누르면 "
            "2256→22:56, 930→09:30, 18→18:00으로 자동 변환됩니다."
        )

        heat_left, heat_right = st.columns(2)

        with heat_left:
            heat_start_input = st.text_input(
                "폭염 시작",
                value=clean_text(editing_record.get("폭염시작")),
                placeholder="예: 1330 → 13:30",
                key=f"manual_heat_start_{nonce}",
                on_change=normalize_manual_time_field,
                args=("heat_start", nonce),
            )

        with heat_right:
            heat_end_input = st.text_input(
                "폭염 종료",
                value=clean_text(editing_record.get("폭염종료")),
                placeholder="예: 1700 → 17:00",
                key=f"manual_heat_end_{nonce}",
                on_change=normalize_manual_time_field,
                args=("heat_end", nonce),
            )

        work_start = parse_time_value(work_start_input)
        work_end = parse_time_value(work_end_input)
        heat_start = parse_time_value(heat_start_input)
        heat_end = parse_time_value(heat_end_input)

        st.button(
            "시간대 체감온도 자동 조회",
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

        weather_debug_lines = st.session_state.get(
            weather_debug_log_key(nonce),
            [],
        )
        if weather_debug_lines:
            with st.expander("기상조회 진단 로그", expanded=False):
                st.caption(
                    "오류 분석용 로그입니다. 인증키는 기록되지 않습니다. "
                    "조회 실패 시 아래 내용을 복사해 전달해 주세요."
                )
                st.code(
                    "\n".join(weather_debug_lines),
                    language="text",
                )

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

        is_forecast_entry = work_date > datetime.now(KST).date()
        if is_forecast_entry:
            st.info(
                "미래 날짜 사전입력입니다. 현장명과 근무시간을 모두 "
                "입력하면 단기예보 체감온도와 예상 폭염시간이 자동 "
                "적용됩니다. 작업 후 기록 조회의 수정 버튼으로 실제값을 "
                "반영할 수 있습니다."
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
            "미래 날짜는 근무시간의 단기예보를 이용해 예상 체감온도와 "
            "31℃ 이상 예상 폭염시간을 자동 입력합니다. 오늘 또는 지난 "
            "날짜는 기존처럼 폭염시간을 우선 조회하고, 비어 있으면 "
            "근무시간의 기상청 500m 격자값을 조회합니다."
        )

        st.divider()

        render_section_heading(
            "04",
            "조치 사항",
            "시행한 조치와 특이사항을 남깁니다.",
        )

        common_measures = common_measures_for_sport(sport)

        st.markdown(
            '<div class="field-label">공통 조치사항</div>',
            unsafe_allow_html=True,
        )
        st.info(
            f"**{sport} 공통 조치사항**\n\n"
            f"{common_measures}"
        )

        st.markdown(
            '<div class="field-label">추가 시행 조치</div>',
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
        }

        measure_left, measure_right = st.columns(2)
        with measure_left:
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

        with measure_right:
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

        measures = list(
            st.session_state.get(
                measures_state_key(nonce),
                [],
            )
        )

        notes = st.text_area(
            "특이사항",
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
        "종목": sport,
        "현장명": site.strip(),
        "팀": team,
        "근무시작": work_start_text,
        "근무종료": work_end_text,
        "작성자": author.strip(),
        "폭염시작": heat_start_text,
        "폭염종료": heat_end_text,
        "체감온도": normalized_temperature,
        "휴게시간": str(rest_minutes),
        "공통 조치사항": common_measures,
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
                "사전 예보 기록을 저장했습니다. 작업 후 실제값으로 수정해 주세요."
                if work_date > datetime.now(KST).date()
                else "기록을 저장했습니다."
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
            "종목·현장명·작성자 검색",
            placeholder="검색어 입력",
            key="record_search",
        )

        team_filter = st.selectbox(
            "팀",
            ["전체"] + TEAM_OPTIONS,
            key="team_filter",
        )

        sport_filter = st.selectbox(
            "종목",
            ["전체"] + SPORT_OPTIONS,
            key="sport_filter",
        )

    filtered = records.copy()

    if search_text.strip() and not filtered.empty:
        keyword = search_text.strip().lower()
        mask = (
            filtered["종목"]
            .astype(str)
            .str.lower()
            .str.contains(
                keyword,
                na=False,
                regex=False,
            )
            |
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

    if (
        sport_filter != "전체"
        and not filtered.empty
    ):
        normalized_sports = filtered.apply(
            lambda row: normalize_sport(
                row.get("종목"),
                row.get("현장명"),
            ),
            axis=1,
        )
        filtered = filtered[
            normalized_sports == sport_filter
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
            load_records.clear()
            st.rerun()

    with download_col:
        st.download_button(
            "엑셀로 추출하기",
            data=make_excel_bytes(records),
            file_name=(
                "현장_폭염_조치_전체기록_"
                f"{datetime.now(KST).strftime('%Y%m%d_%H%M')}"
                ".xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
            disabled=records.empty,
        )

    st.caption(
        "엑셀 파일은 필터와 관계없이 전체 기록을 저장하며, "
        "헤더·테두리·줄바꿈·특이사항 강조가 포함됩니다."
    )

    if filtered.empty:
        st.info("조건에 맞는 기록이 없습니다.")
        return

    st.markdown("### 폭염 보고서 다운로드")
    st.caption(
        "같은 작업날짜와 현장명의 중계팀·영상팀 기록을 "
        "한글 HWPX 보고서 한 장으로 묶습니다."
    )

    report_candidates = (
        filtered[["작업날짜", "현장명"]]
        .drop_duplicates()
        .to_dict("records")
    )
    report_labels = [
        f"{clean_text(item['작업날짜'])} · {clean_text(item['현장명'])}"
        for item in report_candidates
    ]
    selected_report_label = st.selectbox(
        "보고서 대상",
        report_labels,
        key="report_target",
    )
    selected_report_index = report_labels.index(selected_report_label)
    selected_report = report_candidates[selected_report_index]
    report_rows = filtered[
        (filtered["작업날짜"] == selected_report["작업날짜"])
        & (filtered["현장명"] == selected_report["현장명"])
    ]

    try:
        report_bytes = make_heat_report_bytes(report_rows)
        safe_site_name = re.sub(
            r"[^0-9A-Za-z가-힣_-]+",
            "_",
            clean_text(selected_report["현장명"]),
        ).strip("_") or "현장"
        report_date_name = clean_text(
            selected_report["작업날짜"]
        ).replace("-", "")
        st.download_button(
            "한글 보고서(HWPX) 다운로드",
            data=report_bytes,
            file_name=(
                f"폭염작업_조치_결과_보고서_"
                f"{report_date_name}_{safe_site_name}.hwpx"
            ),
            mime=(
                "application/vnd.hancom.hwpx"
            ),
            use_container_width=True,
            key="download_heat_report",
        )
    except Exception as exc:  # noqa: BLE001
        st.warning(f"보고서를 만들 수 없습니다: {exc}")

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
            sport_text = (
                clean_text(record.get("종목"))
                or "종목 미입력"
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
                f"{work_date} · {sport_text} · {team_text} · "
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

            common_text = clean_text(
                record.get("공통 조치사항")
            )
            if common_text:
                st.markdown("**공통 조치사항**")
                st.markdown(common_text)

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

# 새 기록 첫 화면은 Google Sheets 응답을 기다리지 않고 즉시 표시합니다.
# 기록 목록이 필요하거나 기존 기록을 수정할 때만 캐시된 1회 읽기를 실행합니다.
needs_records = (
    st.session_state.page == "records"
    or bool(clean_text(st.session_state.editing_id))
)

if needs_records:
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
