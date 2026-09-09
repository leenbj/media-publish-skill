#!/usr/bin/env python3
"""最终清理：pending 去重 + 删除 b-1 失败旧条目 + links.csv 清掉串行错误记录"""
import csv
import json
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent

pending = json.loads((BASE / "pending-links.json").read_text())
# 1) 只留海宝源那条有效条目（b-1 的旧条目已确认被平台拦截，删）
valid = [p for p in pending if p["title"].startswith("烟台海烟水产")]
seen = set()
dedup = []
for p in valid:
    key = p["title"]
    if key in seen:
        continue
    seen.add(key)
    dedup.append(p)
(BASE / "pending-links.json").write_text(json.dumps(dedup, ensure_ascii=False, indent=2))
print("pending 最终:", [(p["code"], p["title"][:20]) for p in dedup])

# 2) links.csv 清掉高沃错误记录
rows = list(csv.DictReader(open(BASE / "links.csv", encoding="utf-8-sig")))
keep = [r for r in rows if "1038440412" not in r["link"]]
with open(BASE / "links.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=["code", "media", "account", "title", "status",
                                      "link", "published_at", "collected_at"])
    w.writeheader()
    w.writerows(keep)
print(f"links.csv: {len(rows)} → {len(keep)} 条")
