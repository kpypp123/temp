from __future__ import annotations

import csv
import io
import re
import uuid
from datetime import date, datetime, time, timedelta
from typing import Any
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st


APP_TITLE = "현장 폭염 조치 기록"
APP_VERSION = "빠른 시간 기록 v2 · 2026-08-07"
WORKSHEET_DEFAULT = "records"
SPREADSHEET_URL_FALLBACK = "https://docs.google.com/spreadsheets/d/18c-qnfPmGG25qyAM497R7czDw3F7J7WRKmdLX3IGtY0"
KST = ZoneInfo("Asia/Seoul")

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
    "rest_start": "휴게시작",
    "rest_end": "휴게종료",
}


st.set_page_config(
    page_title=APP_TITLE,
    page_icon="☀️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    :root {
        --ink: #17231d;
        --muted: #66736b;
        --line: #dce4df;
        --green: #176b45;
        --green-soft: #e8f3ec;
        --amber-soft: #fff0dc;
    }
    .stApp {
        background: linear-gradient(135deg, #edf5ee 0%, #f7f5ee 50%, #eef3f0 100%);
        color: var(--ink);
    }
    .block-container {
        max-width: 760px;
        padding-top: 1.1rem;
        padding-bottom: 5rem;
    }
    .app-eyebrow {
        color: var(--green);
        font-size: 0.78rem;
        font-weight: 800;
        letter-spacing: 0.09em;
        margin-bottom: 0.25rem;
    }
    .app-title {
        font-size: clamp(1.75rem, 7vw, 2.5rem);
        font-weight: 850;
        letter-spacing: -0.04em;
        line-height: 1.15;
        margin: 0;
    }
    .app-subtitle {
        color: var(--muted);
        font-size: 0.9rem;
        margin: 0.45rem 0 1.15rem;
    }
    div[data-testid="stForm"], div[data-testid="stVerticalBlockBorderWrapper"] {
        background: rgba(255, 255, 255, 0.94);
        border-color: rgba(220, 228, 223, 0.95);
        border-radius: 18px;
    }
    div.stButton > button, div[data-testid="stFormSubmitButton"] > button,
    div[data-testid="stDownloadButton"] > button {
        min-height: 3rem;
        border-radius: 12px;
        font-weight: 750;
    }
    div[data-baseweb="input"] > div, div[data-baseweb="select"] > div,
    textarea {
        border-radius: 11px !important;
    }
    .time-help {
        color: var(--muted);
        font-size: 0.78rem;
        margin: -0.2rem 0 0.65rem;
    }

    .version-badge {
        display: inline-block;
        margin: 0.15rem 0 0.9rem;
        padding: 0.28rem 0.58rem;
        border-radius: 999px;
        background: var(--green-soft);
        color: var(--green);
        font-size: 0.75rem;
        font-weight: 800;
    }
    .quick-time-summary {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 0.5rem;
        margin: 0.55rem 0 0.8rem;
    }
    .quick-time-card {
        background: #f6faf7;
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 0.65rem 0.72rem;
        font-size: 0.8rem;
        line-height: 1.45;
    }
    .quick-time-card b {
        display: block;
        color: var(--green);
        font-size: 0.74rem;
        margin-bottom: 0.15rem;
    }
    .metric-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 0.55rem;
        margin: 0.5rem 0 1rem;
    }
    .metric-card {
        background: rgba(255, 255, 255, 0.92);
        border: 1px solid var(--line);
        border-radius: 14px;
        padding: 0.8rem;
    }
    .metric-card b {
        display: block;
        font-size: 1.35rem;
        line-height: 1.2;
    }
    .metric-card span {
        color: var(--muted);
        font-size: 0.76rem;
    }
    .status-pill {
        display: inline-block;
        padding: 0.22rem 0.55rem;
        border-radius: 999px;
        background: var(--amber-soft);
        color: #9a5012;
        font-size: 0.76rem;
        font-weight: 800;
        margin-bottom: 0.35rem;
    }
    .setup-box {
        background: #fff8e8;
        border: 1px solid #efd9aa;
        border-radius: 14px;
        padding: 0.9rem 1rem;
        margin-bottom: 1rem;
    }
    @media (max-width: 520px) {
        .block-container {
            padding-left: 0.78rem;
            padding-right: 0.78rem;
            padding-top: 0.8rem;
        }
        .metric-grid, .quick-time-summary {
            grid-template-columns: repeat(2, minmax(0, 1fr));
        }
        .metric-grid .metric-card:last-child,
        .quick-time-summary .quick-time-card:last-child {
            grid-column: 1 / -1;
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


def parse_time_value(value: Any, default: time | None = None) -> time | None:
    text = clean_text(value)
    if not text:
        return default

    normalized = text.replace("시", ":00")
    if re.fullmatch(r"\d{1,2}", normalized):
        normalized = f"{normalized}:00"

    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(normalized, fmt).time().replace(second=0, microsecond=0)
        except ValueError:
            continue
    return default


def format_time_value(value: time | None) -> str:
    return value.strftime("%H:%M") if value is not None else ""


def markdown_escape(value: Any) -> str:
    return re.sub(r"([\\`*_{}\[\]()#+\-.!|>])", r"\\\1", clean_text(value))


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


def initialize_time_state(editing_record: dict[str, Any], nonce: int) -> None:
    for field, column in TIME_FIELD_COLUMNS.items():
        key = time_state_key(field, nonce)
        if key not in st.session_state:
            st.session_state[key] = parse_time_value(editing_record.get(column))


def set_time_now(field: str, nonce: int) -> None:
    st.session_state[time_state_key(field, nonce)] = (
        datetime.now(KST).time().replace(second=0, microsecond=0)
    )


def clear_time(field: str, nonce: int) -> None:
    st.session_state[time_state_key(field, nonce)] = None


def time_state_text(field: str, nonce: int) -> str:
    return format_time_value(st.session_state.get(time_state_key(field, nonce)))


def render_time_summary(nonce: int) -> None:
    work_start = time_state_text("work_start", nonce)
    work_end = time_state_text("work_end", nonce)
    heat_start = time_state_text("heat_start", nonce)
    heat_end = time_state_text("heat_end", nonce)
    rest_start = time_state_text("rest_start", nonce)
    rest_end = time_state_text("rest_end", nonce)

    work_text = f"{work_start or '-'} ~ {work_end or '-'}"
    heat_text = f"{heat_start or '-'} ~ {heat_end or '-'}"
    rest_text = f"{rest_start or '-'} ~ {rest_end or '-'}"
    if rest_start and rest_end:
        rest_text += f" · {calculate_minutes(rest_start, rest_end)}분"

    st.markdown(
        f"""
        <div class="quick-time-summary">
            <div class="quick-time-card"><b>근무 시간</b>{work_text}</div>
            <div class="quick-time-card"><b>폭염 노출</b>{heat_text}</div>
            <div class="quick-time-card"><b>휴게 시간</b>{rest_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_resource(show_spinner=False)
def get_worksheet() -> gspread.Worksheet:
    service_account = dict(st.secrets["google_service_account"])
    if "private_key" in service_account:
        service_account["private_key"] = str(service_account["private_key"]).replace("\\n", "\n")

    spreadsheet_url = get_secret(("app", "spreadsheet_url"))
    worksheet_name = get_secret(("app", "worksheet"), WORKSHEET_DEFAULT)
    if not spreadsheet_url:
        raise ValueError("Streamlit Secrets의 app.spreadsheet_url 값이 비어 있습니다.")

    client = gspread.service_account_from_dict(service_account)
    spreadsheet = client.open_by_url(spreadsheet_url)
    return spreadsheet.worksheet(worksheet_name)


def ensure_headers(worksheet: gspread.Worksheet) -> list[str]:
    values = worksheet.get_all_values()
    if not values:
        worksheet.update(range_name="A1:R1", values=[COLUMNS])
        return COLUMNS

    headers = [clean_text(value) for value in values[0]]
    if headers[: len(COLUMNS)] != COLUMNS:
        missing = [column for column in COLUMNS if column not in headers]
        missing_text = ", ".join(missing) if missing else "열 순서 불일치"
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
        padded = raw_row + [""] * (len(COLUMNS) - len(raw_row))
        record = {column: clean_text(padded[index]) for index, column in enumerate(COLUMNS)}
        if any(record.values()):
            rows.append(record)

    if not rows:
        return empty_dataframe()
    return pd.DataFrame(rows, columns=COLUMNS).fillna("")


def record_values(record: dict[str, Any]) -> list[str]:
    return [clean_text(record.get(column, "")) for column in COLUMNS]


def append_record(record: dict[str, Any]) -> None:
    worksheet = get_worksheet()
    ensure_headers(worksheet)
    worksheet.append_row(
        record_values(record),
        value_input_option="USER_ENTERED",
        insert_data_option="INSERT_ROWS",
    )


def update_record(record_id: str, record: dict[str, Any]) -> None:
    worksheet = get_worksheet()
    ensure_headers(worksheet)
    cell = worksheet.find(record_id, in_column=1)

    if cell is None:
        raise ValueError("수정할 기록을 찾지 못했습니다. 목록을 새로고침해 주세요.")

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
        raise ValueError("삭제할 기록을 찾지 못했습니다. 목록을 새로고침해 주세요.")

    worksheet.delete_rows(cell.row)


def calculate_minutes(start: str, end: str) -> int:
    if not start or not end:
        return 0
    start_hour, start_minute = map(int, start.split(":"))
    end_hour, end_minute = map(int, end.split(":"))
    minutes = (end_hour * 60 + end_minute) - (start_hour * 60 + start_minute)
    return minutes + 1440 if minutes < 0 else minutes


def heat_level(value: Any) -> str:
    temperature = parse_float(value)
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


def option_index(options: list[str], value: Any, fallback: int = 0) -> int:
    text = clean_text(value)
    try:
        return options.index(text)
    except ValueError:
        return fallback


def selected_measures(value: Any) -> list[str]:
    text = clean_text(value)
    if not text:
        return []
    parts = [part.strip() for part in re.split(r"\s*\|\s*|,\s*", text) if part.strip()]
    return [part for part in parts if part in MEASURE_OPTIONS]


def make_csv_bytes(dataframe: pd.DataFrame) -> bytes:
    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\n")
    writer.writerow(COLUMNS)
    for _, row in dataframe[COLUMNS].iterrows():
        writer.writerow([clean_text(row[column]) for column in COLUMNS])
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
    message = clean_text(st.session_state.get("flash", ""))
    if message:
        st.success(message)
        st.session_state.flash = ""


def render_header() -> None:
    st.markdown(
        f"""
        <div class="app-eyebrow">HEAT SAFETY LOG</div>
        <h1 class="app-title">현장 폭염 조치 기록</h1>
        <p class="app-subtitle">모바일에서 현장 조치와 휴게 내역을 기록하고 Google Sheets에 공동 저장합니다.</p>
        <div class="version-badge">{APP_VERSION}</div>
        """,
        unsafe_allow_html=True,
    )


def render_navigation() -> None:
    left, right = st.columns(2)
    with left:
        if st.button(
            "✍️ 새 기록",
            use_container_width=True,
            type="primary" if st.session_state.page == "form" else "secondary",
            key="nav_form",
        ):
            reset_form(go_to_form=True)
            st.rerun()
    with right:
        if st.button(
            "📋 기록 조회",
            use_container_width=True,
            type="primary" if st.session_state.page == "records" else "secondary",
            key="nav_records",
        ):
            st.session_state.page = "records"
            st.session_state.pending_delete = None
            st.rerun()


def render_setup_error(error: Exception) -> None:
    st.markdown(
        """
        <div class="setup-box">
        <b>Google Sheets 연결 설정이 아직 완료되지 않았습니다.</b><br>
        화면은 미리 볼 수 있지만, 연결 전에는 저장·수정·삭제가 비활성화됩니다.
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.expander("설정 오류 자세히 보기"):
        st.code(str(error))
        st.caption("Streamlit App settings → Secrets에 스프레드시트 주소와 서비스 계정 정보를 등록해야 합니다.")


def render_form(records: pd.DataFrame, store_error: Exception | None) -> None:
    editing_id = clean_text(st.session_state.editing_id)
    editing_record: dict[str, Any] = {}
    if editing_id and not records.empty:
        matches = records[records["id"].astype(str) == editing_id]
        if not matches.empty:
            editing_record = matches.iloc[0].to_dict()
        else:
            st.warning("수정하려던 기록을 찾지 못해 새 기록 화면으로 전환했습니다.")
            reset_form(go_to_form=True)
            st.rerun()

    title = "기록 수정" if editing_id else "새 기록 작성"
    st.subheader(title)
    nonce = st.session_state.form_nonce
    initialize_time_state(editing_record, nonce)

    default_date = datetime.now(KST).date()
    date_text = clean_text(editing_record.get("작업날짜"))
    if date_text:
        try:
            default_date = date.fromisoformat(date_text)
        except ValueError:
            pass

    team_options = TEAM_OPTIONS.copy()
    current_team = clean_text(editing_record.get("팀"))
    if current_team and current_team not in team_options:
        team_options.append(current_team)

    temperature_default = parse_float(editing_record.get("체감온도"))

    with st.container(border=True):
        st.markdown("#### 기본 정보")
        work_date = st.date_input(
            "작업 날짜 *",
            value=default_date,
            key=f"date_{nonce}",
        )
        site = st.text_input(
            "현장명 *",
            value=clean_text(editing_record.get("현장명")),
            placeholder="예: ○○골프장, ○○야구장",
            key=f"site_{nonce}",
        )
        team = st.selectbox(
            "팀 선택 *",
            options=team_options,
            index=option_index(team_options, current_team, 0),
            key=f"team_{nonce}",
        )
        author = st.text_input(
            "작성자",
            value=clean_text(editing_record.get("작성자")),
            placeholder="예: 홍길동",
            key=f"author_{nonce}",
        )

        st.divider()
        st.markdown("#### 빠른 시간 기록")
        st.caption("장비를 들고 있어도 해당 버튼을 한 번 누르면 현재 한국 시간이 기록됩니다.")

        work_left, work_right = st.columns(2)
        with work_left:
            st.button(
                "▶ 근무 시작",
                key=f"quick_work_start_{nonce}",
                use_container_width=True,
                on_click=set_time_now,
                args=("work_start", nonce),
            )
        with work_right:
            st.button(
                "■ 근무 종료",
                key=f"quick_work_end_{nonce}",
                use_container_width=True,
                on_click=set_time_now,
                args=("work_end", nonce),
            )

        heat_left, heat_right = st.columns(2)
        with heat_left:
            st.button(
                "☀️ 폭염 시작",
                key=f"quick_heat_start_{nonce}",
                use_container_width=True,
                on_click=set_time_now,
                args=("heat_start", nonce),
            )
        with heat_right:
            st.button(
                "🌤️ 폭염 종료",
                key=f"quick_heat_end_{nonce}",
                use_container_width=True,
                on_click=set_time_now,
                args=("heat_end", nonce),
            )

        rest_left, rest_right = st.columns(2)
        with rest_left:
            st.button(
                "☕ 휴게 시작",
                key=f"quick_rest_start_{nonce}",
                use_container_width=True,
                on_click=set_time_now,
                args=("rest_start", nonce),
            )
        with rest_right:
            st.button(
                "✅ 휴게 종료",
                key=f"quick_rest_end_{nonce}",
                use_container_width=True,
                on_click=set_time_now,
                args=("rest_end", nonce),
            )

        render_time_summary(nonce)
        st.caption("버튼으로 기록한 시간은 아래 '기록 저장'을 눌러야 Google Sheets에 저장됩니다.")

        with st.expander("시간 직접 수정", expanded=False):
            st.caption("빠른 기록이 잘못됐을 때만 수정하세요.")

            direct_work_left, direct_work_right = st.columns(2)
            with direct_work_left:
                work_start = st.time_input(
                    "근무 시작",
                    key=time_state_key("work_start", nonce),
                    step=timedelta(minutes=1),
                )
            with direct_work_right:
                work_end = st.time_input(
                    "근무 종료",
                    key=time_state_key("work_end", nonce),
                    step=timedelta(minutes=1),
                )

            direct_heat_left, direct_heat_right = st.columns(2)
            with direct_heat_left:
                heat_start = st.time_input(
                    "폭염 시작",
                    key=time_state_key("heat_start", nonce),
                    step=timedelta(minutes=1),
                )
                st.button(
                    "폭염 시작 지우기",
                    key=f"clear_heat_start_{nonce}",
                    use_container_width=True,
                    on_click=clear_time,
                    args=("heat_start", nonce),
                )
            with direct_heat_right:
                heat_end = st.time_input(
                    "폭염 종료",
                    key=time_state_key("heat_end", nonce),
                    step=timedelta(minutes=1),
                )
                st.button(
                    "폭염 종료 지우기",
                    key=f"clear_heat_end_{nonce}",
                    use_container_width=True,
                    on_click=clear_time,
                    args=("heat_end", nonce),
                )

            direct_rest_left, direct_rest_right = st.columns(2)
            with direct_rest_left:
                rest_start = st.time_input(
                    "휴게 시작",
                    key=time_state_key("rest_start", nonce),
                    step=timedelta(minutes=1),
                )
                st.button(
                    "휴게 시작 지우기",
                    key=f"clear_rest_start_{nonce}",
                    use_container_width=True,
                    on_click=clear_time,
                    args=("rest_start", nonce),
                )
            with direct_rest_right:
                rest_end = st.time_input(
                    "휴게 종료",
                    key=time_state_key("rest_end", nonce),
                    step=timedelta(minutes=1),
                )
                st.button(
                    "휴게 종료 지우기",
                    key=f"clear_rest_end_{nonce}",
                    use_container_width=True,
                    on_click=clear_time,
                    args=("rest_end", nonce),
                )

        st.divider()
        st.markdown("#### 폭염 정보")
        temperature = st.number_input(
            "체감온도 (℃)",
            min_value=-20.0,
            max_value=60.0,
            step=0.1,
            value=temperature_default,
            placeholder="예: 33.5",
            key=f"temperature_{nonce}",
        )

        st.divider()
        st.markdown("#### 조치 사항")
        measures = st.multiselect(
            "시행한 조치",
            options=MEASURE_OPTIONS,
            default=selected_measures(editing_record.get("조치사항")),
            key=f"measures_{nonce}",
        )
        notes = st.text_area(
            "상세 조치 및 특이사항",
            value=clean_text(editing_record.get("특이사항")),
            placeholder="예: 설치·철수 시간 조정, 제작팀 협의사항 등",
            height=120,
            key=f"notes_{nonce}",
        )

        save_col, reset_col = st.columns(2)
        with save_col:
            save_clicked = st.button(
                "수정 내용 저장" if editing_id else "기록 저장",
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
    rest_start_text = format_time_value(rest_start)
    rest_end_text = format_time_value(rest_end)

    validation_errors: list[str] = []
    if not site.strip():
        validation_errors.append("현장명을 입력해 주세요.")
    if not work_start_text or not work_end_text:
        validation_errors.append("근무 시작과 종료 시간을 기록해 주세요.")
    if bool(heat_start_text) != bool(heat_end_text):
        validation_errors.append("폭염 노출 시간은 시작과 종료를 모두 입력하거나 모두 비워 주세요.")
    if bool(rest_start_text) != bool(rest_end_text):
        validation_errors.append("휴게 시간은 시작과 종료를 모두 입력하거나 모두 비워 주세요.")

    if validation_errors:
        for message in validation_errors:
            st.error(message)
        return

    calculated_rest_minutes = calculate_minutes(rest_start_text, rest_end_text)

    now_text = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")
    created_at = clean_text(editing_record.get("등록시간")) or now_text
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
        "체감온도": "" if temperature is None else f"{float(temperature):.1f}",
        "휴게시작": rest_start_text,
        "휴게종료": rest_end_text,
        "휴게시간": str(calculated_rest_minutes),
        "조치사항": " | ".join(measures),
        "특이사항": notes.strip(),
        "등록시간": created_at,
        "수정시간": now_text,
    }

    try:
        if editing_id:
            update_record(editing_id, record)
            st.session_state.flash = "기록을 수정했습니다."
        else:
            append_record(record)
            st.session_state.flash = "기록을 저장했습니다."
    except Exception as exc:  # noqa: BLE001 - user-facing boundary
        st.error(f"저장에 실패했습니다: {exc}")
        return

    reset_form(go_to_form=False)
    st.session_state.page = "records"
    st.rerun()


def render_metrics(records: pd.DataFrame) -> None:
    hot_count = 0
    total_rest = 0
    if not records.empty:
        hot_count = sum((parse_float(value) or -999) >= 33 for value in records["체감온도"])
        total_rest = sum(parse_int(value) for value in records["휴게시간"])

    st.markdown(
        f"""
        <div class="metric-grid">
            <div class="metric-card"><b>{len(records)}</b><span>전체 기록</span></div>
            <div class="metric-card"><b>{hot_count}</b><span>33℃ 이상</span></div>
            <div class="metric-card"><b>{total_rest}분</b><span>누적 휴게</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_records(records: pd.DataFrame, store_error: Exception | None) -> None:
    st.subheader("저장된 기록")
    spreadsheet_url = get_secret(("app", "spreadsheet_url"), SPREADSHEET_URL_FALLBACK)
    st.link_button(
        "📊 Google Sheets에서 기록 열기",
        spreadsheet_url,
        use_container_width=True,
    )

    if store_error is not None:
        st.info("Google Sheets 연결이 완료되면 공동 기록이 여기에 표시됩니다.")
        return

    render_metrics(records)

    with st.expander("검색 및 필터", expanded=False):
        search_text = st.text_input("현장명·작성자 검색", placeholder="검색어 입력", key="record_search")
        team_filter = st.selectbox("팀", ["전체"] + TEAM_OPTIONS, key="team_filter")

    filtered = records.copy()
    if search_text.strip() and not filtered.empty:
        keyword = search_text.strip().lower()
        mask = (
            filtered["현장명"].astype(str).str.lower().str.contains(keyword, na=False, regex=False)
            | filtered["작성자"].astype(str).str.lower().str.contains(keyword, na=False, regex=False)
        )
        filtered = filtered[mask]
    if team_filter != "전체" and not filtered.empty:
        filtered = filtered[filtered["팀"] == team_filter]

    if not filtered.empty:
        filtered = filtered.assign(
            _sort_key=filtered["작업날짜"].astype(str) + " " + filtered["등록시간"].astype(str)
        ).sort_values("_sort_key", ascending=False)

    refresh_col, download_col = st.columns(2)
    with refresh_col:
        if st.button("새로고침", use_container_width=True, key="refresh_records"):
            st.rerun()
    with download_col:
        st.download_button(
            "현재 목록 CSV",
            data=make_csv_bytes(filtered.drop(columns=["_sort_key"], errors="ignore")),
            file_name=f"현장_폭염_조치_기록_{datetime.now(KST).strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
            disabled=filtered.empty,
        )

    if filtered.empty:
        st.info("조건에 맞는 기록이 없습니다.")
        return

    admin_pin = get_secret(("security", "admin_pin"))

    for _, row in filtered.iterrows():
        record = row.to_dict()
        record_id = clean_text(record.get("id"))
        site = markdown_escape(record.get("현장명")) or "현장명 미입력"
        work_date = markdown_escape(record.get("작업날짜"))
        temperature_text = clean_text(record.get("체감온도"))
        temperature_label = f" · {temperature_text}℃" if temperature_text else ""

        with st.container(border=True):
            st.markdown(f"### {site}")
            st.markdown(
                f'<span class="status-pill">{heat_level(temperature_text)}{temperature_label}</span>',
                unsafe_allow_html=True,
            )
            st.caption(
                f"{work_date} · {clean_text(record.get('팀')) or '팀 미입력'} · "
                f"작성자 {clean_text(record.get('작성자')) or '-'}"
            )
            st.write(f"근무: {clean_text(record.get('근무시작'))} ~ {clean_text(record.get('근무종료'))}")
            st.write(
                "폭염 노출: "
                f"{clean_text(record.get('폭염시작')) or '-'} ~ {clean_text(record.get('폭염종료')) or '-'}"
            )
            st.write(
                "휴게: "
                f"{clean_text(record.get('휴게시작')) or '-'} ~ {clean_text(record.get('휴게종료')) or '-'} "
                f"({parse_int(record.get('휴게시간'), 0)}분)"
            )
            st.write(f"조치사항: {clean_text(record.get('조치사항')) or '-'}")
            if clean_text(record.get("특이사항")):
                st.write(f"특이사항: {clean_text(record.get('특이사항'))}")
            st.caption(f"등록 {clean_text(record.get('등록시간'))} · 수정 {clean_text(record.get('수정시간'))}")

            edit_col, delete_col = st.columns(2)
            with edit_col:
                if st.button("수정", use_container_width=True, key=f"edit_{record_id}"):
                    st.session_state.editing_id = record_id
                    st.session_state.form_nonce += 1
                    st.session_state.page = "form"
                    st.session_state.pending_delete = None
                    st.rerun()
            with delete_col:
                if st.button("삭제", use_container_width=True, key=f"delete_{record_id}"):
                    st.session_state.pending_delete = record_id
                    st.rerun()

            if st.session_state.pending_delete == record_id:
                st.warning("이 기록을 삭제할까요? 삭제 후 복구는 Google Sheets 변경 기록에서만 가능합니다.")
                entered_pin = ""
                if admin_pin:
                    entered_pin = st.text_input(
                        "관리자 삭제 PIN",
                        type="password",
                        key=f"delete_pin_{record_id}",
                    )
                confirm_col, cancel_col = st.columns(2)
                with confirm_col:
                    if st.button("삭제 확인", type="primary", use_container_width=True, key=f"confirm_{record_id}"):
                        if admin_pin and entered_pin != admin_pin:
                            st.error("관리자 PIN이 올바르지 않습니다.")
                        else:
                            try:
                                delete_record(record_id)
                                st.session_state.pending_delete = None
                                st.session_state.flash = "기록을 삭제했습니다."
                                st.rerun()
                            except Exception as exc:  # noqa: BLE001
                                st.error(f"삭제에 실패했습니다: {exc}")
                with cancel_col:
                    if st.button("취소", use_container_width=True, key=f"cancel_{record_id}"):
                        st.session_state.pending_delete = None
                        st.rerun()


init_state()
render_header()
render_navigation()
show_flash()

records_df = empty_dataframe()
connection_error: Exception | None = None
try:
    records_df = load_records()
except Exception as exc:  # noqa: BLE001 - setup errors are displayed in the UI
    connection_error = exc
    render_setup_error(exc)

if st.session_state.page == "form":
    render_form(records_df, connection_error)
else:
    render_records(records_df, connection_error)
