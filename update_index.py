"""
update_index.py — sales_history.json을 index.html의 ALL_DATA에 반영

GitHub Actions 환경에서 실행:
  - sales_history.json 로드
  - build_dashboard.build_all_data() 호출
  - index.html의 'const ALL_DATA = ...;' 줄 교체
"""
from __future__ import annotations
import json, sys
from pathlib import Path

import build_dashboard as bd

BASE = Path(__file__).parent.resolve()
HISTORY = BASE / "sales_history.json"
INDEX = BASE / "index.html"


def main() -> int:
    if not HISTORY.exists():
        print(f"[!] {HISTORY} 없음", file=sys.stderr)
        return 1
    if not INDEX.exists():
        print(f"[!] {INDEX} 없음", file=sys.stderr)
        return 1

    with open(HISTORY, encoding="utf-8") as f:
        history = json.load(f)

    all_data = bd.build_all_data(history)
    new_js = "const ALL_DATA = " + json.dumps(all_data, ensure_ascii=False) + ";"

    text = INDEX.read_text(encoding="utf-8")
    lines = text.split("\n")
    replaced = False
    for i, line in enumerate(lines):
        stripped = line.lstrip()
        if stripped.startswith("const ALL_DATA"):
            indent = line[:len(line) - len(stripped)]
            lines[i] = indent + new_js
            replaced = True
            break
    if not replaced:
        print("[!] index.html에서 'const ALL_DATA' 줄을 찾지 못했습니다.", file=sys.stderr)
        return 2

    INDEX.write_text("\n".join(lines), encoding="utf-8")
    print(f"[update_index] 완료: {len(history)} 날짜, {len(all_data)} 처리")
    return 0


if __name__ == "__main__":
    sys.exit(main())
