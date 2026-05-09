"""
gdrive_fetch.py — Google Drive 폴더에서 매출 XLS 파일 다운로드
================================================================

GitHub Actions에서 실행되는 스크립트.

인증:
  - 환경변수 GDRIVE_CREDENTIALS (서비스 계정 JSON 문자열) 필요
  - 또는 GDRIVE_CREDENTIALS_FILE (파일 경로)

대상 폴더:
  - 환경변수 GDRIVE_FOLDER_ID 또는 --folder-id 인자

기본 동작:
  어제 날짜(KST) 매출 파일을 Drive에서 찾아 ./POS_DOWNLOADS/{date}/로 다운로드
  파일명 패턴: *{YYYY-MM-DD}*.xls 또는 .xlsx
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
except ImportError:
    print("[!] 'google-api-python-client', 'google-auth' 가 필요합니다.", file=sys.stderr)
    print("    pip install google-api-python-client google-auth", file=sys.stderr)
    sys.exit(1)

KST = timezone(timedelta(hours=9))
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
DEFAULT_FOLDER_ID = "1UZ4IXk-0p_kryyPRmtEp0ZelWgRDVZ8d"


def get_drive_service():
    creds_json = os.environ.get("GDRIVE_CREDENTIALS")
    creds_file = os.environ.get("GDRIVE_CREDENTIALS_FILE")

    if creds_json:
        try:
            info = json.loads(creds_json)
        except json.JSONDecodeError as e:
            print(f"[!] GDRIVE_CREDENTIALS JSON 파싱 실패: {e}", file=sys.stderr)
            sys.exit(2)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    elif creds_file:
        creds = service_account.Credentials.from_service_account_file(creds_file, scopes=SCOPES)
    else:
        print("[!] GDRIVE_CREDENTIALS 또는 GDRIVE_CREDENTIALS_FILE 환경변수가 필요합니다.", file=sys.stderr)
        sys.exit(2)

    return build("drive", "v3", credentials=creds, cache_discovery=False)


def list_files_for_date(service, folder_id: str, date: str):
    query = f"'{folder_id}' in parents and trashed = false"
    fields = "nextPageToken, files(id, name, modifiedTime, size, mimeType)"
    files = []
    page_token = None
    while True:
        resp = service.files().list(
            q=query, fields=fields, pageSize=1000, pageToken=page_token,
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
        files.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    matched = [
        f for f in files
        if date in f["name"]
        and (f["name"].lower().endswith(".xls") or f["name"].lower().endswith(".xlsx"))
    ]
    return matched, files


def download_file(service, file_id: str, dest_path: Path):
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    with open(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request, chunksize=1024 * 1024)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return dest_path.stat().st_size


def yesterday_kst() -> str:
    return (datetime.now(KST) - timedelta(days=1)).strftime("%Y-%m-%d")


def main() -> int:
    ap = argparse.ArgumentParser(description="Drive 매출 폴더에서 XLS 다운로드")
    ap.add_argument("--date", default=yesterday_kst(),
                    help="대상 날짜 YYYY-MM-DD (기본: 어제 KST)")
    ap.add_argument("--folder-id",
                    default=os.environ.get("GDRIVE_FOLDER_ID", DEFAULT_FOLDER_ID),
                    help="Drive 폴더 ID")
    ap.add_argument("--output", default="./POS_DOWNLOADS",
                    help="저장 루트 (실제: {output}/{date}/)")
    args = ap.parse_args()

    out_dir = Path(args.output) / args.date
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"[gdrive_fetch] 대상 날짜  : {args.date}")
    print(f"[gdrive_fetch] Drive 폴더 : {args.folder_id}")
    print(f"[gdrive_fetch] 저장 위치  : {out_dir}")

    service = get_drive_service()
    matched, all_files = list_files_for_date(service, args.folder_id, args.date)

    if not matched:
        print(f"[gdrive_fetch] {args.date} 날짜의 파일이 Drive에 없습니다.")
        print(f"[gdrive_fetch] (참고) 폴더 안 전체 파일 {len(all_files)}개:")
        for f in all_files[:20]:
            print(f"  - {f['name']}")
        if len(all_files) > 20:
            print(f"  ... 외 {len(all_files) - 20}개")
        return 1

    print(f"[gdrive_fetch] 다운로드 대상 {len(matched)}개:")
    for f in matched:
        print(f"  · {f['name']} ({f.get('size', '?')} bytes)")

    total_bytes = 0
    for f in matched:
        dest = out_dir / f["name"]
        size = download_file(service, f["id"], dest)
        total_bytes += size
        print(f"  ✓ {f['name']} → {dest} ({size:,} bytes)")

    print(f"[gdrive_fetch] 완료: {len(matched)}개, 총 {total_bytes:,} bytes")
    return 0


if __name__ == "__main__":
    sys.exit(main())
