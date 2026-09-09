#!/usr/bin/env python3
"""回查发布结果：逐个平台检查 pending-links.json 里的文章状态，抓正式链接存 links.csv。

审核中的跳过（下轮再查）；已发布的抓链接；CSDN 直接拼链接并验证 HTTP 状态。

用法：
  python3 scripts/collect_links.py            # 回查所有 pending
  python3 scripts/collect_links.py --code b-1 # 只查指定编码
"""
import argparse
import csv
import json
import sys
from datetime import datetime
from pathlib import Path

import requests
from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
import common


def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def check_toutiao(pg, item: dict) -> dict:
    """头条：文章管理列表按标题找，取状态+正式链接。
    注意：列表页"审核中/未通过"可能查不到被风控静默拦截的文章——
    补查前台主页（www.toutiao.com/c/user/<id>/#log）确认是否真实在线。"""
    pg.goto("https://mp.toutiao.com/profile_v4/graphic/articles")
    pg.wait_for_timeout(6000)
    r = pg.evaluate("""((title) => {
      // 标题行 a：文本完全匹配或前缀匹配
      const a = [...document.querySelectorAll('a')]
        .find(a => {
          const t = (a.textContent || '').trim()
          return t === title || (title.length > 10 && t && title.startsWith(t))
        })
      if (!a) return {found: false}
      // 状态从 .content-list-main 级容器抓（避免匹配到筛选 tab 文本）
      let row = a
      for (let i = 0; i < 8 && row; i++) {
        row = row.parentElement
        if (row && row.className && row.className.toString().includes('content-list')) break
      }
      const scope = row || a.closest('div')
      const statusM = scope ? (scope.textContent.match(/已发布|审核中|审核未通过/) || []) : []
      return {found: true, status: statusM[0] || '', href: a.href}
    })""", item["title"])
    if r["found"]:
        link = r["href"] if r["status"] == "已发布" else ""
        return {"status": r["status"], "link": link}
    # 后台列表没找到 → 查前台主页是否真实在线（含搜索关键词前8字）
    key = item["title"][:8]
    user_id = common.media_param("toutiao", "user_id")
    pg.goto(f"https://www.toutiao.com/c/user/{user_id}/#log")
    pg.wait_for_timeout(6000)
    online = pg.evaluate("""((key) => {
      const a = [...document.querySelectorAll('a')].find(x => (x.textContent || '').includes(key))
      return a ? {found: true, href: a.href} : {found: false}
    })""", key)
    if online["found"]:
        return {"status": "已发布（前台在线）", "link": online["href"]}
    return {"status": "未找到（可能被平台拦截或审核超时）", "link": ""}


def check_sohu(pg, item: dict) -> dict:
    """搜狐：内容管理列表找标题行。
    实测规则：
    - 刚过审的新文章：列表页标题 a 指向 articlepreview（后台预览），正式链接未渲染
    - 发布较久的文章：列表页标题 a 直接是 www.sohu.com/a/ 正式链接
    - 坑：列表前几篇可能是其他文章，a.textContent.trim()===title 必须全等匹配，
      且匹配后要校验行容器内确有该标题文本（否则会串行拿到别的文章的链接）"""
    pg.goto("https://mp.sohu.com/mpfe/v4/contentManagement/first/page")
    pg.wait_for_timeout(7000)
    r = pg.evaluate("""((title) => {
      const norm = s => (s || '').replace(/\\s+/g, '')
      const target = norm(title)
      const a = [...document.querySelectorAll('a')]
        .find(x => norm(x.textContent) === target ||
                   (x.href || '').includes('www.sohu.com/a/') && norm(x.textContent).includes(target.slice(0, 12)))
      if (!a) return {found: false, matched: false}
      let row = a
      for (let i = 0; i < 8 && row; i++) {
        row = row.parentElement
        if (row && row.className && row.className.toString().includes('content-list-main')) break
      }
      const scope = row || a.parentElement
      // 校验：scope 内必须真的有这篇标题（防止匹配到别的文章行）
      if (scope && !norm(scope.textContent).includes(target.slice(0, 12)))
        return {found: false, matched: 'row-mismatch'}
      const statusM = scope ? (scope.textContent.match(/审核中|已发布|未通过/) || []) : []
      // 只有 a.href 本身是正式链接才采用；articlepreview 一律不算
      const link = a.href && a.href.includes('www.sohu.com/a/') ? a.href : ''
      return {found: true, status: statusM[0] || '', link}
    })""", item["title"])
    if not r["found"]:
        # 没匹配到或串行 → 当作"链接未渲染/未找到"，留 pending
        return {"status": "已发布（正式链接未渲染）", "link": ""}
    return {"status": r["status"], "link": r["link"]}


def check_csdn(pg, item: dict) -> dict:
    """CSDN：success 页 URL 已含 articleId（发布时记录），直接拼正式链接验 HTTP 状态。"""
    # 从 pending 里的 success_url 或重新进列表拿 articleId
    art_id = None
    if item.get("success_url"):
        art_id = item["success_url"].rstrip("/").split("/")[-1]
    else:
        # 创作管理列表查
        pg.goto("https://mp.csdn.net/mp_blog/manage/article?NotNeedLoginNoLeftPanel=true")
        pg.wait_for_timeout(6000)
        art_id = pg.evaluate("""((title) => {
          const a = [...document.querySelectorAll('a')].find(a => (a.textContent || '').includes(title.slice(0, 15)))
          if (!a) return null
          const m = (a.href || '').match(/article\\/details\\/(\\d+)/)
          return m ? m[1] : null
        })""", item["title"])
    if not art_id:
        return {"status": "未找到", "link": ""}
    user = common.media_param("csdn", "user")
    link = f"https://blog.csdn.net/{user}/article/details/{art_id}"
    try:
        resp = requests.get(link, timeout=15, allow_redirects=True,
                            headers={"User-Agent": "Mozilla/5.0"})
        if resp.status_code == 200 and ("blog" in resp.url or "文章" in resp.text):
            return {"status": "已发布（在线）", "link": link}
        return {"status": f"链接待生效(HTTP {resp.status_code})", "link": ""}
    except Exception as e:
        return {"status": f"验证失败({type(e).__name__})", "link": ""}


CHECKERS = {"toutiao": check_toutiao, "sohu": check_sohu, "csdn": check_csdn}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--code", help="只查指定编码")
    ap.add_argument("--pending", default="", help="待回查文件（默认 $MEDIA_PENDING）")
    ap.add_argument("--links", default="", help="结果 CSV（默认 $MEDIA_LINKS）")
    a = ap.parse_args()

    PENDING = Path(a.pending) if a.pending else common.pending_path()
    LINKS_CSV = Path(a.links) if a.links else common.links_path()
    if not PENDING.exists():
        print("没有待回查条目")
        return 0
    pending = json.loads(PENDING.read_text(encoding="utf-8"))
    targets = [x for x in pending if (not a.code or x["code"] == a.code)]
    if not targets:
        print("没有待回查条目")
        return 0

    # 已有 csv 结果读进来做合并去重
    existing = {}
    if LINKS_CSV.exists():
        with open(LINKS_CSV, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                existing[(row["code"], row["title"])] = row

    new_rows = []
    still_pending = []
    media_groups = {}
    for x in targets:
        media_groups.setdefault(x["media"], []).append(x)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        for media, items in media_groups.items():
            # 每个媒体用对应账号登录态（同媒体多账号时逐条切换）
            for item in items:
                state_file = common.state_for(media, item["account"])
                if not state_file.exists():
                    log(f"[{item['code']}] 缺登录态文件 {state_file.name}，跳过")
                    still_pending.append(item)
                    continue
                ctx = browser.new_context(storage_state=str(state_file))
                pg = ctx.new_page()
                try:
                    r = CHECKERS[media](pg, item)
                except Exception as e:
                    log(f"[{item['code']}] 回查异常：{type(e).__name__} {e}")
                    r = {"status": f"回查异常({type(e).__name__})", "link": ""}
                ctx.close()
                log(f"[{item['code']}/{media}] {r['status']} {r['link']}")
                if r["link"]:
                    new_rows.append({
                        "code": item["code"], "media": media, "account": item["account"],
                        "title": item["title"], "status": r["status"], "link": r["link"],
                        "published_at": item["published_at"],
                        "collected_at": datetime.now().isoformat(timespec="seconds"),
                    })
                elif r["status"] in ("未找到",) or "未通过" in r["status"]:
                    new_rows.append({
                        "code": item["code"], "media": media, "account": item["account"],
                        "title": item["title"], "status": r["status"], "link": "",
                        "published_at": item["published_at"],
                        "collected_at": datetime.now().isoformat(timespec="seconds"),
                    })
                else:
                    item["status"] = r["status"]
                    still_pending.append(item)
        browser.close()

    # 写 csv（合并旧记录）
    all_rows = list(existing.values()) + new_rows
    seen = set()
    dedup = []
    for row in all_rows:
        key = (row["code"], row["title"])
        if key in seen:
            continue
        seen.add(key)
        dedup.append(row)
    with open(LINKS_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["code", "media", "account", "title", "status",
                                          "link", "published_at", "collected_at"])
        w.writeheader()
        w.writerows(dedup)

    # 更新 pending：只留还没拿到链接的
    got_keys = {(r["code"], r["title"]) for r in new_rows if r["link"]}
    PENDING.write_text(json.dumps(
        [x for x in pending if (x["code"], x["title"]) not in got_keys],
        ensure_ascii=False, indent=2))

    print(f"\n回查完成：{len(new_rows)} 条更新，{len(still_pending)} 条仍待下轮")
    print(f"结果文件：{LINKS_CSV}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
