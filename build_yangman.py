# -*- coding: utf-8 -*-
"""
build_yangman.py — 삼창양만 급이일지 구글시트를 읽어 월별 양만 손익(yangman-pl.json) 생성.
대시보드 '양만/그룹 손익' 탭 소스.

시트: 급이일지 v8.0 - 삼창양만 (YANGMAN_SHEET_ID). netlify 앱이 동기화하는 라이브 시트.
탭 이름에 의존하지 않고 헤더로 자동 식별(장어출하·일일급이·사료입고·설정).
인증: 환경변수 GDRIVE_CREDENTIALS (서비스 계정 JSON). 시트를 SA에 공유해야 읽힘.

산식(앱 월별 손익 분석과 동일):
  매출   = 장어출하 단가>0 행의 kg×단가 (호지이동 '->'·선별·단가0 제외). 내부/외부 분리.
  사료비 = 일일급이 (오전+오후)kg × 사료단가(사료입고 최신단가)
  전기료/인건비/기타운영비/입식비/이자비용 = 설정의 '<항목>_YYYY-MM' 키
  순이익 = 매출 − 사료비 − 전기 − 인건 − 기타 − 입식 − 이자
"""
from __future__ import annotations
import json, os, re, sys
from pathlib import Path
from collections import defaultdict

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
except ImportError:
    print("[!] google-api-python-client / google-auth 필요", file=sys.stderr); sys.exit(1)

SCOPES = ["https://www.googleapis.com/auth/drive.readonly",
          "https://www.googleapis.com/auth/spreadsheets.readonly"]
YANGMAN_SHEET_ID = os.environ.get("YANGMAN_SHEET_ID", "1nfBMAXZ4ZbUmzPxm0UQLM79faZC52spykZ8oH_CBEGk")
INTERNAL_KEYS = ("삼창", "조합", "반듯한", "평택본점")  # 그룹 내부거래 거래처
OUT = Path(__file__).parent / "yangman-pl.json"


def get_sheets():
    cj = os.environ.get("GDRIVE_CREDENTIALS"); cf = os.environ.get("GDRIVE_CREDENTIALS_FILE")
    if cj:
        creds = service_account.Credentials.from_service_account_info(json.loads(cj), scopes=SCOPES)
    elif cf:
        creds = service_account.Credentials.from_service_account_file(cf, scopes=SCOPES)
    else:
        print("[!] GDRIVE_CREDENTIALS 환경변수 필요", file=sys.stderr); sys.exit(2)
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def num(x):
    if x is None: return 0.0
    s = str(x).replace(",", "").replace("원", "").strip()
    if s in ("", "-"): return 0.0
    try: return float(s)
    except ValueError: return 0.0


def col_idx(header, *names):
    for i, h in enumerate(header):
        hh = str(h).replace(" ", "")
        for n in names:
            if n in hh: return i
    return -1


def read_all_tabs(svc):
    meta = svc.spreadsheets().get(spreadsheetId=YANGMAN_SHEET_ID,
                                  fields="sheets.properties.title").execute()
    titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
    tabs = {}
    for t in titles:
        try:
            r = svc.spreadsheets().values().get(
                spreadsheetId=YANGMAN_SHEET_ID, range=t,
                valueRenderOption="UNFORMATTED_VALUE").execute()
            tabs[t] = r.get("values", [])
        except Exception as e:
            print(f"[!] {t} 읽기 실패: {e}", file=sys.stderr)
    return tabs


def is_internal(party):
    return any(k in str(party) for k in INTERNAL_KEYS)


def is_move(party, price):
    p = str(party)
    return p.startswith("->") or p.startswith("→") or "선별" in p or num(price) <= 0


def build_pl(tabs):
    months = defaultdict(lambda: {"매출_내부": 0.0, "매출_외부": 0.0, "출하kg": 0.0,
                                  "사료비": 0.0, "전기료": 0.0, "인건비": 0.0,
                                  "기타운영비": 0.0, "입식비": 0.0, "이자비용": 0.0})
    feed_price = {}   # 사료명 -> 최신 단가
    chulha = giip = None
    feed_rows = []

    # 1) 탭 분류
    for title, rows in tabs.items():
        if not rows: continue
        # 헤더 후보(데이터 있는 첫 행)
        header = rows[0]
        hjoin = "".join(str(c) for c in header)
        if ("출하" in hjoin and "단가" in hjoin and "거래처" in hjoin):
            chulha = rows
        elif ("오전" in hjoin and "오후" in hjoin):
            giip = rows
        elif ("입고" in hjoin and "사료" in hjoin and "단가" in hjoin):
            feed_rows = rows
        # 설정류: 어디 있든 '항목_YYYY-MM' 패턴 스캔
        for r in rows:
            if len(r) >= 2 and isinstance(r[0], str):
                m = re.match(r"^(전기료|인건비|기타운영비|입식비|이자비용)_(\d{4}-\d{2})$", r[0].strip())
                if m:
                    months[m.group(2)][m.group(1)] += num(r[1])

    # 2) 사료 단가 맵 (최신 입고단가)
    if feed_rows:
        h = feed_rows[0]
        ci_date = col_idx(h, "입고일", "날짜"); ci_feed = col_idx(h, "사료")
        ci_price = col_idx(h, "단가")
        seen_date = {}
        for r in feed_rows[1:]:
            if ci_feed < 0 or ci_feed >= len(r): continue
            name = str(r[ci_feed]).strip()
            if not name: continue
            d = str(r[ci_date]) if 0 <= ci_date < len(r) else ""
            p = num(r[ci_price]) if 0 <= ci_price < len(r) else 0
            if p <= 0: continue
            if name not in seen_date or d >= seen_date[name]:
                seen_date[name] = d; feed_price[name] = p

    def price_for(feed):
        feed = str(feed).strip()
        if feed in feed_price: return feed_price[feed]
        for k, v in feed_price.items():       # 부분일치 (한방흑자5호 ~ 한방흑자)
            if k and (k in feed or feed in k): return v
        return 0.0

    # 3) 매출 (장어출하)
    if chulha:
        h = chulha[0]
        ci_date = col_idx(h, "출하일", "날짜"); ci_party = col_idx(h, "거래처")
        ci_kg = col_idx(h, "출하kg", "출하량", "kg"); ci_price = col_idx(h, "단가")
        for r in chulha[1:]:
            if ci_date < 0 or ci_date >= len(r): continue
            d = str(r[ci_date]).strip()
            m = re.match(r"(\d{4}-\d{2})", d)
            if not m: continue
            party = r[ci_party] if 0 <= ci_party < len(r) else ""
            price = r[ci_price] if 0 <= ci_price < len(r) else 0
            kg = num(r[ci_kg]) if 0 <= ci_kg < len(r) else 0
            if is_move(party, price): continue
            amt = kg * num(price); ym = m.group(1)
            months[ym]["출하kg"] += kg
            months[ym]["매출_내부" if is_internal(party) else "매출_외부"] += amt

    # 4) 사료비 (일일급이)
    if giip:
        h = giip[0]
        ci_date = col_idx(h, "날짜"); ci_feed = col_idx(h, "사료")
        ci_am = col_idx(h, "오전"); ci_pm = col_idx(h, "오후")
        for r in giip[1:]:
            if ci_date < 0 or ci_date >= len(r): continue
            d = str(r[ci_date]).strip()
            m = re.match(r"(\d{4}-\d{2})", d)
            if not m: continue
            feed = r[ci_feed] if 0 <= ci_feed < len(r) else ""
            kg = (num(r[ci_am]) if 0 <= ci_am < len(r) else 0) + \
                 (num(r[ci_pm]) if 0 <= ci_pm < len(r) else 0)
            months[m.group(1)]["사료비"] += kg * price_for(feed)

    # 5) 정리 + 순이익
    out = {}
    for ym in sorted(months):
        v = months[ym]
        rev = v["매출_내부"] + v["매출_외부"]
        cost = v["사료비"] + v["전기료"] + v["인건비"] + v["기타운영비"] + v["입식비"] + v["이자비용"]
        rec = {k: round(v[k]) for k in v}
        rec["매출"] = round(rev)
        rec["순이익"] = round(rev - cost)
        rec["마진율"] = round((rev - cost) / rev * 100, 1) if rev else None
        out[ym] = rec
    return out


def main():
    svc = get_sheets()
    tabs = read_all_tabs(svc)
    pl = build_pl(tabs)
    OUT.write_text(json.dumps(pl, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"[OK] yangman-pl.json — {len(pl)}개월")
    for ym in list(pl)[-3:]:
        r = pl[ym]; print(f"  {ym}: 매출 {r['매출']:,} 순이익 {r['순이익']:,} (마진 {r['마진율']}%)")


if __name__ == "__main__":
    main()
