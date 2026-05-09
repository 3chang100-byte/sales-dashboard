"""
매출대시보드.html 생성기
========================
이 스크립트를 실행해야만 reports/매출대시보드.html 이 업데이트됩니다.

파일 역할 분리:
  unified_dashboard.py -> reports/월간_대시보드.html  (월간 뷰, 별도)
  build_dashboard.py   -> reports/매출대시보드.html   (캘린더+드릴다운 메인 대시보드)

실행 방법:
  python build_dashboard.py
"""
import json, re
from datetime import datetime, timedelta
from pathlib import Path

BASE = Path(__file__).parent.resolve()
HISTORY_FILE = BASE / "sales_history.json"
TEMPLATE_FILE = BASE / "reports" / "매출대시보드.html"
OUTPUT_FILE   = BASE / "reports" / "매출대시보드.html"

STORES = ["평택점", "오산점", "안성점", "복대점", "율량점"]
STORE_COLORS = {
    "평택점": "#3b82f6", "오산점": "#06b6d4", "안성점": "#a78bfa",
    "복대점": "#f97316", "율량점": "#34d399",
}
ALCOHOL_KEYWORDS = ["소주", "맥주", "백세주", "복분자", "막걸리"]

def build_all_data(history):
    dates = sorted(history.keys())
    all_data = {}

    for target_date in dates:
        rec = history.get(target_date, {})
        prev_date = (datetime.fromisoformat(target_date) - timedelta(days=1)).date().isoformat()
        prev_rec = history.get(prev_date, {})

        trend7_dates = [
            (datetime.fromisoformat(target_date) - timedelta(days=i)).date().isoformat()
            for i in range(6, -1, -1)
        ]

        store_data = {}
        for st in STORES:
            r = rec.get(st, {})
            prev = prev_rec.get(st, {})
            l  = r.get("점심매출", 0); d  = r.get("저녁매출", 0)
            lc = r.get("점심객수", 0); dc = r.get("저녁객수", 0)
            t  = l + d;               c  = lc + dc

            trend = []
            for dd in trend7_dates:
                dr = history.get(dd, {}).get(st, {})
                lv = dr.get("점심매출", 0); dv = dr.get("저녁매출", 0)
                trend.append({"date": dd,
                              "label": datetime.fromisoformat(dd).strftime("%m/%d"),
                              "total": lv+dv, "lunch": lv, "dinner": dv})

            raw_menu = r.get("메뉴", {})
            menu_top = sorted(raw_menu.items(), key=lambda x: -x[1].get("sales", 0))[:10]

            # 주류 지표 계산
            alcohol_qty   = sum(v.get("qty", 0)   for k, v in raw_menu.items()
                                if any(kw in k for kw in ALCOHOL_KEYWORDS))
            alcohol_sales = sum(v.get("sales", 0) for k, v in raw_menu.items()
                                if any(kw in k for kw in ALCOHOL_KEYWORDS))
            tables = r.get("회전율", {}).get("tables", 0)
            alcohol_per_table = round(alcohol_qty / tables, 2) if tables else 0
            alcohol_sales_pct = round(alcohol_sales / t * 100, 1) if t else 0

            store_data[st] = {
                "color": STORE_COLORS.get(st, "#999"),
                "total": t, "lunch": l, "dinner": d,
                "guests": c, "lunch_guests": lc, "dinner_guests": dc,
                "ticket": round(t/c) if c else 0,
                "lunch_ticket": round(l/lc) if lc else 0,
                "dinner_ticket": round(d/dc) if dc else 0,
                "lunch_pct": round(l/t*100, 1) if t else 0,
                "prev_total": prev.get("점심매출",0)+prev.get("저녁매출",0),
                "week_total": 0,
                "trend": trend,
                "missing": not bool(r),
                "menu": [{"name":k,"qty":v.get("qty",0),"sales":v.get("sales",0),
                          "lunch":v.get("lunch",0),"dinner":v.get("dinner",0)}
                         for k,v in menu_top],
                "hourly": r.get("시간별", {}),
                "turnover": r.get("회전율", {}),
                "alcohol_qty": alcohol_qty,
                "alcohol_sales": alcohol_sales,
                "alcohol_per_table": alcohol_per_table,
                "alcohol_sales_pct": alcohol_sales_pct,
            }

        all_total  = sum(v["total"]  for v in store_data.values())
        all_lunch  = sum(v["lunch"]  for v in store_data.values())
        all_dinner = sum(v["dinner"] for v in store_data.values())
        all_guests = sum(v["guests"] for v in store_data.values())

        hourly_all = {}
        for st in STORES:
            for h, amt in (store_data[st].get("hourly") or {}).items():
                hourly_all[h] = hourly_all.get(h, 0) + amt

        trend7_all = []
        for dd in trend7_dates:
            dr = history.get(dd, {})
            tv = sum(dr.get(s,{}).get("점심매출",0)+dr.get(s,{}).get("저녁매출",0) for s in STORES)
            lv = sum(dr.get(s,{}).get("점심매출",0) for s in STORES)
            dv = sum(dr.get(s,{}).get("저녁매출",0) for s in STORES)
            trend7_all.append({"date":dd,"label":datetime.fromisoformat(dd).strftime("%m/%d"),
                               "total":tv,"lunch":lv,"dinner":dv})

        all_data[target_date] = {
            "date": target_date,
            "dateDisp": datetime.fromisoformat(target_date).strftime("%Y년 %m월 %d일"),
            "stores": STORES,
            "storeData": store_data,
            "summary": {
                "total": all_total, "lunch": all_lunch, "dinner": all_dinner,
                "guests": all_guests,
                "ticket": round(all_total/all_guests) if all_guests else 0,
                "lunchPct": round(all_lunch/all_total*100,1) if all_total else 0,
                "prevTotal": sum(v["prev_total"] for v in store_data.values()),
                "weekTotal": 0,
                "hourly": hourly_all,
            },
            "trend7": trend7_all,
        }

    return all_data


def main():
    with open(HISTORY_FILE, encoding="utf-8") as f:
        history = json.load(f)

    all_data = build_all_data(history)
    print(f"날짜 {len(all_data)}개 처리: {', '.join(sorted(all_data.keys()))}")

    template = TEMPLATE_FILE.read_text(encoding="utf-8")
    new_data_js = "const ALL_DATA = " + json.dumps(all_data, ensure_ascii=False) + ";"

    lines = template.split("\n")
    replaced = False
    for i, line in enumerate(lines):
        if line.startswith("const ALL_DATA"):
            lines[i] = new_data_js
            replaced = True
            break
    if not replaced:
        print("오류: ALL_DATA 줄을 찾지 못했습니다.")
        return
    new_html = "\n".join(lines)

    import os, stat
    try:
        os.chmod(OUTPUT_FILE, stat.S_IRUSR|stat.S_IWUSR|stat.S_IRGRP|stat.S_IROTH)
    except Exception:
        pass
    OUTPUT_FILE.write_text(new_html, encoding="utf-8")
    try:
        os.chmod(OUTPUT_FILE, stat.S_IRUSR|stat.S_IRGRP|stat.S_IROTH)
    except Exception:
        pass
    print(f"완료: {OUTPUT_FILE}  ({OUTPUT_FILE.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
