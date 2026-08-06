from __future__ import annotations

import csv
import io
import re
import uuid
from datetime import date, datetime
from typing import Any
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st


APP_TITLE = "현장 폭염 조치 기록"
WORKSHEET_DEFAULT = "records"
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
TIME_OPTIONS = ["선택"] + [f"{hour:02d}:{minute:02d}" for hour in range(24) for minute in range(0, 60, 5)]
HEAT_HOUR_OPTIONS = ["선택"] + [f"{hour:02d}시" for hour in range(24)]


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
        .metric-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr));
        }
        .metric-grid .metric-card:last-child {
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
    if start == "선택" or end == "선택" or not start or not end:
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
        """
        <div class="app-eyebrow">HEAT SAFETY LOG</div>
        <h1 class="app-title">현장 폭염 조치 기록</h1>
        <p class="app-subtitle">모바일에서 현장 조치와 휴게 내역을 기록하고 Google Sheets에 공동 저장합니다.</p>
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
            editing_id = ""

    title = "기록 수정" if editing_id else "새 기록 작성"
    st.subheader(title)
    nonce = st.session_state.form_nonce

    default_date = date.today()
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
    work_count_default = max(parse_int(editing_record.get("작업인원"), 1), 1)
    rest_minutes_default = max(parse_int(editing_record.get("휴게시간"), 0), 0)

    with st.form(f"record_form_{nonce}", border=True):
        st.markdown("#### 기본 정보")
        work_date = st.date_input("작업 날짜 *", value=default_date, key=f"date_{nonce}")
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
        work_start = st.selectbox(
            "근무 시작 *",
            options=TIME_OPTIONS,
            index=option_index(TIME_OPTIONS, editing_record.get("근무시작"), option_index(TIME_OPTIONS, "08:00")),
            key=f"work_start_{nonce}",
        )
        work_end = st.selectbox(
            "근무 종료 *",
            options=TIME_OPTIONS,
            index=option_index(TIME_OPTIONS, editing_record.get("근무종료"), option_index(TIME_OPTIONS, "17:00")),
            key=f"work_end_{nonce}",
        )
        author = st.text_input(
            "작성자",
            value=clean_text(editing_record.get("작성자")),
            placeholder="예: 홍길동",
            key=f"author_{nonce}",
        )
        work_count = st.number_input(
            "작업 인원",
            min_value=1,
            step=1,
            value=work_count_default,
            key=f"work_count_{nonce}",
        )

        st.divider()
        st.markdown("#### 폭염 및 휴게 정보")
        heat_start = st.selectbox(
            "폭염 노출 시작",
            options=HEAT_HOUR_OPTIONS,
            index=option_index(HEAT_HOUR_OPTIONS, editing_record.get("폭염시작"), 0),
            key=f"heat_start_{nonce}",
        )
        heat_end = st.selectbox(
            "폭염 노출 종료",
            options=HEAT_HOUR_OPTIONS,
            index=option_index(HEAT_HOUR_OPTIONS, editing_record.get("폭염종료"), 0),
            key=f"heat_end_{nonce}",
        )
        temperature = st.number_input(
            "체감온도 (℃)",
            min_value=-20.0,
            max_value=60.0,
            step=0.1,
            value=temperature_default,
            placeholder="예: 33.5",
            key=f"temperature_{nonce}",
        )
        rest_start = st.selectbox(
            "휴게 시작",
            options=TIME_OPTIONS,
            index=option_index(TIME_OPTIONS, editing_record.get("휴게시작"), 0),
            key=f"rest_start_{nonce}",
        )
        rest_end = st.selectbox(
            "휴게 종료",
            options=TIME_OPTIONS,
            index=option_index(TIME_OPTIONS, editing_record.get("휴게종료"), 0),
            key=f"rest_end_{nonce}",
        )
        rest_minutes = st.number_input(
            "총 휴게시간 (분)",
            min_value=0,
            step=1,
            value=rest_minutes_default,
            help="0으로 두면 휴게 시작·종료 시간을 기준으로 자동 계산합니다.",
            key=f"rest_minutes_{nonce}",
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

        save_clicked = st.form_submit_button(
            "수정 내용 저장" if editing_id else "기록 저장",
            use_container_width=True,
            type="primary",
            disabled=store_error is not None,
        )
        reset_clicked = st.form_submit_button("입력 초기화", use_container_width=True)

    if reset_clicked:
        reset_form(go_to_form=True)
        st.rerun()

    if not save_clicked:
        return

    validation_errors: list[str] = []
    if not site.strip():
        validation_errors.append("현장명을 입력해 주세요.")
    if work_start == "선택" or work_end == "선택":
        validation_errors.append("근무 시작과 종료 시간을 선택해 주세요.")

    if validation_errors:
        for message in validation_errors:
            st.error(message)
        return

    calculated_rest_minutes = int(rest_minutes)
    if calculated_rest_minutes == 0:
        calculated_rest_minutes = calculate_minutes(rest_start, rest_end)

    now_text = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")
    created_at = clean_text(editing_record.get("등록시간")) or now_text
    record_id = editing_id or str(uuid.uuid4())

    record = {
        "id": record_id,
        "작업날짜": work_date.isoformat(),
        "현장명": site.strip(),
        "팀": team,
        "근무시작": work_start,
        "근무종료": work_end,
        "작성자": author.strip(),
        "작업인원": str(int(work_count)),
        "폭염시작": "" if heat_start == "선택" else heat_start,
        "폭염종료": "" if heat_end == "선택" else heat_end,
        "체감온도": "" if temperature is None else f"{float(temperature):.1f}",
        "휴게시작": "" if rest_start == "선택" else rest_start,
        "휴게종료": "" if rest_end == "선택" else rest_end,
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
                f"작성자 {clean_text(record.get('작성자')) or '-'} · {parse_int(record.get('작업인원'), 0)}명"
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
