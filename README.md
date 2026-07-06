# sales-dashboard 자동화

장어만세 5개 지점 매출 대시보드의 자동 갱신 인프라.

## 구조

- `index.html` — 대시보드 본체 (GitHub Pages가 서빙)
- `gdrive_fetch.py` — Google Drive 매출 폴더에서 XLS 파일 다운로드
- `daily_report_generator.py` — XLS 파싱 → `sales_history.json` 갱신
- `build_dashboard.py` — `sales_history.json` → ALL_DATA JS 객체 빌드
- `update_index.py` — index.html의 ALL_DATA 줄 교체
- `sales_history.json` — 누적 매출 데이터 (Actions가 자동 commit)
- `.github/workflows/daily.yml` — 평일 12시 KST cron 트리거

## 흐름

```
황주임 → Google Drive 폴더 업로드
   ↓ (평일 12시 KST cron)
GitHub Actions
   ├─ gdrive_fetch.py    : Drive → POS_DOWNLOADS/{어제}/
   ├─ daily_report_generator.py : XLS → sales_history.json
   ├─ update_index.py    : sales_history.json → index.html
   └─ git commit + push  → GitHub Pages 배포
```

## 수동 실행

Actions 탭 → "Daily Dashboard Update" → "Run workflow" → 날짜 입력(또는 비워두면 어제) → Run

## 필요한 Secrets

- `GDRIVE_CREDENTIALS` — 서비스 계정 JSON 전체 내용
- `GDRIVE_FOLDER_ID` — Drive 폴더 ID (선택, 기본값은 코드에 박혀있음)

<!-- redeploy 2026-07-05: 7월 최신분 반영 트리거 -->
