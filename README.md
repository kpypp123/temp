# 현장 폭염 조치 기록 — Streamlit 버전

기존 `index.html`의 현장 폭염 기록 기능을 모바일 우선 Streamlit 앱으로 옮긴 버전입니다. 기록은 브라우저가 아니라 Google Sheets의 `records` 시트에 공동 저장됩니다.

## PC에 Python이 없어도 됩니다

이 저장소를 Streamlit Community Cloud에 배포하면 서버가 `requirements.txt`를 읽어 Python과 필요한 라이브러리를 자동 설치합니다. 로컬 PC에 Python을 설치할 필요가 없습니다.

## 주요 기능

- 모바일 기준 1열 입력 UI
- 새 기록 저장
- Google Sheets 공동 조회
- 기록 검색 및 팀 필터
- 기록 수정·삭제
- 휴게시간 자동 계산
- CSV 다운로드
- 선택형 삭제 관리자 PIN

## Google Sheet 첫 행

시트 탭 이름은 `records`로 만들고 A1부터 아래 제목을 순서대로 입력합니다.

```text
id | 작업날짜 | 현장명 | 팀 | 근무시작 | 근무종료 | 작성자 | 작업인원 | 폭염시작 | 폭염종료 | 체감온도 | 휴게시작 | 휴게종료 | 휴게시간 | 조치사항 | 특이사항 | 등록시간 | 수정시간
```

빈 시트인 경우 앱이 제목 행을 자동으로 만들 수 있지만, `records` 탭 자체는 미리 존재해야 합니다.

## Streamlit 배포

1. 이 변경사항을 `main` 브랜치에 병합합니다.
2. Streamlit Community Cloud에서 새 앱을 만듭니다.
3. Repository: `kpypp123/temp`
4. Branch: `main`
5. Main file path: `streamlit_app.py`
6. App settings → Secrets에 아래 형식으로 인증정보를 등록합니다.

## Secrets 형식

`.streamlit/secrets.toml.example`을 참고합니다. 다운로드한 서비스 계정 JSON에서 같은 이름의 값을 복사합니다.

```toml
[app]
spreadsheet_url = "Google 스프레드시트 전체 주소"
worksheet = "records"

[google_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = """-----BEGIN PRIVATE KEY-----
...
-----END PRIVATE KEY-----
"""
client_email = "..."
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
```

서비스 계정 JSON 파일이나 `private_key`를 GitHub에 올리면 안 됩니다.

삭제 PIN이 필요하면 아래 항목을 추가합니다. 로그인 기능은 아니며 삭제 동작만 보호합니다.

```toml
[security]
admin_pin = "원하는 숫자"
```

## Google 권한 확인

- Google Sheets API가 활성화되어 있어야 합니다.
- 만든 스프레드시트를 서비스 계정 `client_email`에 **편집자**로 공유해야 합니다.
- 연결 오류가 계속되면 같은 Google Cloud 프로젝트에서 Google Drive API도 활성화합니다.
