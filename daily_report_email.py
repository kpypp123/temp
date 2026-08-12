from __future__ import annotations

import ast
import os
import re
import smtplib
import tomllib
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path
from types import ModuleType
from zoneinfo import ZoneInfo

import gspread
import pandas as pd


KST = ZoneInfo("Asia/Seoul")
APP_SOURCE = Path(
    os.environ.get("CHECKTEMP_APP_SOURCE", Path(__file__).with_name("streamlit_app.py"))
)
SPREADSHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "18c-qnfPmGG25qyAM497R7czDw3F7J7WRKmdLX3IGtY0"
)
WORKSHEET_NAME = "records"


def required_env(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"필수 환경변수 {name}가 비어 있습니다.")
    return value


def load_report_module() -> ModuleType:
    """UI를 실행하지 않고 앱의 기존 Word 보고서 생성 로직만 불러옵니다."""
    source = APP_SOURCE.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(APP_SOURCE))
    allowed = []
    for node in tree.body:
        if isinstance(node, (ast.Import, ast.ImportFrom, ast.FunctionDef, ast.ClassDef)):
            allowed.append(node)
        elif isinstance(node, (ast.Assign, ast.AnnAssign)) and node.lineno < 1000:
            allowed.append(node)
    library_tree = ast.Module(body=allowed, type_ignores=[])
    ast.fix_missing_locations(library_tree)
    module = ModuleType("checktemp_report_library")
    module.__file__ = str(APP_SOURCE)
    exec(compile(library_tree, str(APP_SOURCE), "exec"), module.__dict__)
    return module


def service_account_info(raw_toml: str) -> dict:
    parsed = tomllib.loads(raw_toml)
    info = parsed.get("google_service_account")
    if not isinstance(info, dict):
        raise RuntimeError(
            "GOOGLE_SERVICE_ACCOUNT_TOML에 [google_service_account] 항목이 없습니다."
        )
    normalized = dict(info)
    normalized["private_key"] = str(normalized.get("private_key", "")).replace(
        "\\n", "\n"
    )
    return normalized


def target_report_date() -> date:
    manual = os.environ.get("REPORT_DATE", "").strip()
    if manual:
        return datetime.strptime(manual, "%Y-%m-%d").date()
    return datetime.now(KST).date() - timedelta(days=1)


def load_records(module: ModuleType) -> pd.DataFrame:
    credentials = service_account_info(required_env("GOOGLE_SERVICE_ACCOUNT_TOML"))
    client = gspread.service_account_from_dict(credentials)
    values = client.open_by_url(SPREADSHEET_URL).worksheet(WORKSHEET_NAME).get_all_values()
    if not values:
        return pd.DataFrame(columns=module.COLUMNS)
    headers = values[0]
    rows = [row + [""] * (len(headers) - len(row)) for row in values[1:]]
    frame = pd.DataFrame([row[: len(headers)] for row in rows], columns=headers)
    for column in module.COLUMNS:
        if column not in frame.columns:
            frame[column] = ""
    return frame[module.COLUMNS].fillna("")


def safe_site_name(value: object) -> str:
    cleaned = re.sub(r"[^0-9A-Za-z가-힣_-]+", "_", str(value).strip()).strip("_")
    return cleaned or "근무장소"


def build_reports(module: ModuleType, records: pd.DataFrame, report_date: date):
    date_text = report_date.isoformat()
    daily = records[records["작업날짜"].astype(str).str.strip() == date_text]
    reports: list[tuple[str, bytes]] = []
    for site in daily["현장명"].astype(str).str.strip().drop_duplicates():
        site_rows = daily[daily["현장명"].astype(str).str.strip() == site]
        filename = (
            f"폭염작업_조치_결과_보고서_{report_date:%Y%m%d}_"
            f"{safe_site_name(site)}.docx"
        )
        reports.append((filename, module.make_heat_report_docx_bytes(site_rows)))
    return reports


def recipients_from_env() -> list[str]:
    raw = required_env("MAIL_RECIPIENTS")
    recipients = [item.strip() for item in re.split(r"[,;\n]+", raw) if item.strip()]
    if not recipients:
        raise RuntimeError("MAIL_RECIPIENTS에 수신자 주소가 없습니다.")
    return recipients


def send_email(report_date: date, reports: list[tuple[str, bytes]]) -> None:
    sender = required_env("MAIL_SENDER")
    password = required_env("MAIL_APP_PASSWORD").replace(" ", "")
    recipients = recipients_from_env()

    message = EmailMessage()
    message["Subject"] = f"[CheckTemp] {report_date:%Y-%m-%d} 폭염작업 조치 결과 보고서"
    message["From"] = sender
    message["To"] = sender
    message["Bcc"] = ", ".join(recipients)
    names = "\n".join(f"- {name}" for name, _ in reports)
    message.set_content(
        f"{report_date:%Y-%m-%d} 근무 기록으로 자동 작성한 보고서입니다.\n\n"
        f"첨부 보고서: {len(reports)}개\n{names}\n\n"
        "이 메일은 CheckTemp에서 자동 발송되었습니다."
    )
    for filename, payload in reports:
        message.add_attachment(
            payload,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=filename,
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
        smtp.login(sender, password)
        smtp.send_message(message)


def main() -> None:
    report_date = target_report_date()
    module = load_report_module()
    records = load_records(module)
    reports = build_reports(module, records, report_date)
    if not reports:
        print(f"{report_date:%Y-%m-%d} 기록이 없어 메일 발송을 생략했습니다.")
        return
    send_email(report_date, reports)
    print(f"{report_date:%Y-%m-%d} 보고서 {len(reports)}개를 발송했습니다.")


if __name__ == "__main__":
    main()
