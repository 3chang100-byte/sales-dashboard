"""
gdrive_fetch.py — Google Drive 폴더에서 매출 XLS 파일 다운로드
================================================================

지원 폴더 구조 (자동 감지):
  A) 평면 구조: 부모 폴더 안에 직접 *_{date}.xls 파일들
  B) 날짜 하위 폴더: 부모 폴더 안에 {date}/ 폴더 → 그 안에 파일들

인증:
  - 환경변수 GDRIVE_CREDENTIALS (서비스 계정 JSON 문자열)
  - 또는 GDRIVE_CREDENTIALS_FILE (파일 경로)
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
    sys.exit(1)

KST = timezone(timedelta(hours=9))
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
DEFAULT_FOLDER_ID = "1UZ4IXk-0p_kryyPRmtEp0ZelWgRDVZ8d"
FOLDER_MIME = "application/vnd.google-apps.folder"


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


def list_folder_children(service, folder_id: str):
    """폴더의 직계 자식들 (파일+폴더) 모두 반환."""
    query = f"'{folder_id}' in parents and trashed = false"
    fields = "nextPageToken, files(id, name, mimeType, size, modifiedTime)"
    items = []
    page_token = None
    while True:
        resp = service.files().list(
            q=query, fields=fields, pageSize=1000, pageToken=page_token,
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
        items.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return items


def find_xls_files(service, parent_folder_id: str, date: str):
    """
    부모 폴더에서 date 날짜의 .xls/.xlsx 파일들을 찾음.
    - 우선: 부모 안에 '{date}' 이름의 하위폴더가 있으면 그 안의 모든 .xls/.xlsx 반환
    - 차선: 부모 안에 직접 *_{date}.xls(x) 파일이 있으면 그것 반환
    """
    children = list_folder_children(service, parent_folder_id)

    # A) 날짜 이름의 하위폴더 찾기
    date_folder = next(
        (c for c in children if c["mimeType"] == FOLDER_MIME and c["name"] == date),
        None
    )
    if date_folder:
        print(f"[gdrive_fetch] 날짜 하위폴더 발견: '{date_folder['name']}' (id={date_folder['id'][:12]}...)")
        sub_items = list_folder_children(service, date_folder["id"])
        files = [
            f for f in sub_items
            if f["mimeType"] != FOLDER_MIME
            and (f["name"].lower().endswith(".xls") or f["name"].lower().endswith(".xlsx"))
        ]
        return files, "subfolder", children

    # B) 평면 구조 fallback
    flat_files = [
        f for f in children
        if f["mimeType"] != FOLDER_MIME
        and date in f["name"]
        and (f["name"].lower().endswith(".xls") or f["name"].lower().endswith(".xlsx"))
    ]
    return flat_files, "flat", children


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
    ap.add_argument("--date", default=yesterday_kst())
    ap.add_argument("--folder-id",
                    default=os.environ.get("GDRIVE_FOLDER_ID", DEFAULT_FOLDER_ID))
    ap.add_argument("--output", default="./POS_DOWNLOADS")
    args = ap.parse_args()

    out_dir = Path(args.output) / args.date
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"[gdrive_fetch] 대상 날짜  : {args.date}")
    print(f"[gdrive_fetch] Drive 폴더 : {args.folder_id}")
    print(f"[gdrive_fetch] 저장 위치  : {out_dir}")

    service = get_drive_service()
    files, mode, all_children = find_xls_files(service, args.folder_id, args.date)

    if not files:
        print(f"[gdrive_fetch] {args.date} 날짜의 .xls/.xlsx 파일을 찾지 못함 (탐색 모드: {mode}).")
        print(f"[gdrive_fetch] 부모 폴더 안 항목 {len(all_children)}개:")
        for c in all_children[:30]:
            kind = "DIR " if c["mimeType"] == FOLDER_MIME else "FILE"
            print(f"  [{kind}] {c['name']}")
        if len(all_children) > 30:
            print(f"  ... 외 {len(all_children) - 30}개")
        return 1

    print(f"[gdrive_fetch] 탐색 모드: {mode}")
    print(f"[gdrive_fetch] 다운로드 대상 {len(files)}개:")
    for f in files:
        print(f"  · {f['name']} ({f.get('size', '?')} bytes)")

    total_bytes = 0
    for f in files:
        dest = out_dir / f["name"]
        size = download_file(service, f["id"], dest)
        total_bytes += size
        print(f"  ✓ {f['name']} → {dest.name} ({size:,} bytes)")

    print(f"[gdrive_fetch] 완료: {len(files)}개, 총 {total_bytes:,} bytes")
    return 0


if __name__ == "__main__":
    sys.exit(main())
