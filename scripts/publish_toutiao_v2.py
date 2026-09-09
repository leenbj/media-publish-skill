#!/usr/bin/env python3
"""头条号（b-1）自动发布：填稿 → 免费图库随机选图 → 单图封面 → 预览并发布。

用法：
  dry-run（填稿+配图+截图，不发布）：
    python3 scripts/publish_toutiao_v2.py --title "标题" --body-file 正文.txt
  真发：
    python3 scripts/publish_toutiao_v2.py --title "标题" --body-file 正文.txt --go
  真发并登记待回查链接：
    python3 scripts/publish_toutiao_v2.py --title "标题" --body-file 正文.txt --go --code b-1
"""
import argparse
import json
import random
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
import common
PENDING = common.pending_path()
SHOTS = Path("/tmp/media-publish-dryrun")
SHOTS.mkdir(parents=True, exist_ok=True)

PUBLISH_URL = "https://mp.toutiao.com/profile_v4/graphic/publish"


def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def drawer_open(pg) -> bool:
    return pg.evaluate("""(() => {
      const panel = document.querySelector('.upload-image-panel')
      if (!panel) return false
      const drawer = panel.closest('.byte-drawer-wrapper')
      return !!drawer && drawer.offsetWidth > 0
    })()""")


def close_ai_drawer(pg) -> None:
    """关 AI 助手抽屉（它会挡工具栏点击）；确认关闭，关不掉打日志继续。"""
    for _ in range(3):
        gone = pg.evaluate("""(() => {
          const d = document.querySelector('.ai-assistant-drawer')
          if (!d || !d.offsetParent) return true
          const b = d.querySelector('svg.close-btn')
          if (b) b.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true}))
          return false
        })()""")
        pg.wait_for_timeout(1200)
        if gone:
            return
    log("AI 抽屉未能确认关闭，继续（可能挡住工具栏）")


def fill_article(pg, title: str, paras: list[str]) -> None:
    pg.goto(PUBLISH_URL)
    pg.wait_for_timeout(5000)
    close_ai_drawer(pg)
    pg.get_by_placeholder("请输入文章标题").fill(title)
    pg.wait_for_timeout(500)
    pg.evaluate("document.querySelector('.ProseMirror[contenteditable=true]').focus()")
    pg.wait_for_timeout(500)
    for i, p in enumerate(paras):
        pg.keyboard.type(p)
        if i < len(paras) - 1:
            pg.keyboard.press("Enter")
            pg.wait_for_timeout(200)
    pg.wait_for_timeout(2000)
    log(f"填稿完成：标题 {len(title)} 字，正文 {len(paras)} 段")


def open_gallery(pg) -> None:
    """点工具栏图片按钮打开图库抽屉（实测 = 可见工具栏按钮 index 11）。
    偶发打不开（AI 抽屉遮挡/工具栏未稳定），最多重试 3 次。"""
    for attempt in (1, 2, 3):
        close_ai_drawer(pg)  # 重试前确保遮挡已除
        pg.evaluate("window.scrollTo(0, 0)")
        pg.wait_for_timeout(800)
        info = pg.evaluate("""(() => {
          const btns = [...document.querySelectorAll('button.syl-toolbar-button')].filter(b => b.offsetParent)
          if (btns.length <= 11) return {n: btns.length}
          const r = btns[11].getBoundingClientRect()
          return {n: btns.length, x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)}
        })()""")
        if "x" not in info:
            log(f"图库按钮缺失（仅 {info['n']} 个可见按钮），重试 {attempt}/3")
            pg.wait_for_timeout(2000)
            continue
        pg.mouse.click(info["x"], info["y"])
        pg.wait_for_timeout(2500)
        if drawer_open(pg):
            log(f"图库抽屉已打开（{attempt}/3 次）")
            return
        log(f"第 {attempt}/3 次点击无反应，重试…")
        pg.wait_for_timeout(1500)
    pg.screenshot(path=str(SHOTS / "toutiao-gallery-fail.png"))
    raise RuntimeError("图库抽屉 3 次未打开，截图见 toutiao-gallery-fail.png（可能页面改版，检查 NOTES-toutiao.md）")


def ensure_drawer(pg) -> None:
    """抽屉意外关闭（页面重渲染偶发）时重开，保证后续步骤有抽屉可用。"""
    if not drawer_open(pg):
        log("图库抽屉意外关闭，重新打开…")
        open_gallery(pg)


def search_gallery(pg, keyword: str) -> None:
    """切到免费正版图片 tab 并搜索。组合词可能空结果，逐个降级。"""
    ensure_drawer(pg)
    pg.evaluate("""(() => {
      const t = [...document.querySelectorAll('.byte-tabs-header-title')]
        .find(e => e.textContent.trim() === '免费正版图片')
      if (t) t.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true}))
    })()""")
    pg.wait_for_timeout(2000)
    for kw in keyword.split() or [keyword]:
        pg.evaluate("""((kw) => {
          const drawer = document.querySelector('.upload-image-panel').closest('.byte-drawer-wrapper')
          const pane = [...drawer.querySelectorAll('.byte-tabs-content-item')].find(p => p.className.includes('active'))
          const input = [...pane.querySelectorAll('input')].find(i => i.placeholder && i.placeholder.includes('关键词'))
          input.focus()
          const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set
          setter.call(input, '')
          input.dispatchEvent(new Event('input', {bubbles: true}))
          setter.call(input, kw)
          input.dispatchEvent(new Event('input', {bubbles: true}))
        })""", kw)
        pg.keyboard.press("Enter")
        pg.wait_for_timeout(5000)
        n = pg.evaluate("""(() => {
          const drawer = document.querySelector('.upload-image-panel').closest('.byte-drawer-wrapper')
          const pane = [...drawer.querySelectorAll('.byte-tabs-content-item')].find(p => p.className.includes('active'))
          return pane.querySelectorAll('li.item').length
        })()""")
        log(f"关键词「{kw}」命中 {n} 张")
        if n > 0:
            return
    raise RuntimeError(f"所有关键词都无图：{keyword}")


def pick_random_image(pg) -> None:
    """随机点一张卡片（只选视口内的，视口外的点不中），等选中标记出现。"""
    ensure_drawer(pg)
    n, xy = pg.evaluate("""(() => {
      const drawer = document.querySelector('.upload-image-panel').closest('.byte-drawer-wrapper')
      const pane = [...drawer.querySelectorAll('.byte-tabs-content-item')].find(p => p.className.includes('active'))
      const items = [...pane.querySelectorAll('li.item')]
      const ok = []
      items.forEach((li, i) => {
        const r = li.getBoundingClientRect()
        if (r.width > 50 && r.y >= 0 && r.y + r.height < 800 && r.x >= 280)
          ok.push([i, Math.round(r.x + r.width / 2), Math.round(r.y + r.height / 2)])
      })
      return [ok.length, ok[Math.floor(Math.random() * ok.length)]]
    })()""")
    idx, cx, cy = xy[0], xy[1], xy[2]
    pg.mouse.click(cx, cy)
    pg.wait_for_timeout(2000)
    picked = pg.evaluate("""((idx) => {
      const drawer = document.querySelector('.upload-image-panel').closest('.byte-drawer-wrapper')
      const pane = [...drawer.querySelectorAll('.byte-tabs-content-item')].find(p => p.className.includes('active'))
      return !!(pane.querySelectorAll('li.item')[idx] && pane.querySelectorAll('li.item')[idx].querySelector('.pop-number'))
    })""", idx)
    if not picked:
        raise RuntimeError(f"第 {idx} 张未选中（无 pop-number）")
    log(f"随机选中第 {idx + 1}/{n} 张")


def confirm_insert(pg) -> None:
    btn = pg.get_by_role("button", name="确定", exact=True)
    btn.click()
    pg.wait_for_timeout(5000)
    gone = pg.evaluate("!document.querySelector('.upload-image-panel')")
    if not gone:
        raise RuntimeError("插入后抽屉未关闭")
    log("图片已插入正文，抽屉已关闭")


def set_cover_single(pg) -> None:
    pg.evaluate("""(() => {
      const radios = [...document.querySelectorAll('.article-cover .byte-radio')]
      const target = radios.find(r => r.textContent.trim() === '单图')
      if (!target) return 'nf'
      target.scrollIntoView({block: 'center'})
      const input = target.querySelector('input')
      if (input && !input.checked) input.click()
      return 'ok'
    })()""")
    pg.wait_for_timeout(1000)
    checked = pg.evaluate("""(() => {
      const radios = [...document.querySelectorAll('.article-cover .byte-radio')]
      const t = radios.find(r => r.textContent.trim() === '单图')
      return !!(t && t.querySelector('input') && t.querySelector('input').checked)
    })()""")
    log(f"封面单图：{'已选' if checked else '选中失败'}")


def do_publish(pg) -> str:
    pg.get_by_role("button", name="预览并发布").click()
    pg.wait_for_timeout(4000)
    pg.get_by_role("button", name="确认发布").click()
    pg.wait_for_timeout(6000)
    url = pg.url
    toast = pg.evaluate("""(() => {
      const els = [...document.querySelectorAll('*')]
        .filter(e => e.children.length === 0 && /审核中|发布成功|失败/.test(e.textContent || ''))
      return els.length ? els[0].textContent.trim() : ''
    })()""")
    log(f"发布提交：URL={url} toast={toast}")
    return url


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--title", required=True)
    ap.add_argument("--body-file", required=True)
    ap.add_argument("--keyword", default="", help="图库检索词，默认取标题前2-4字的核心词")
    ap.add_argument("--code", default="b-1")
    ap.add_argument("--go", action="store_true")
    ap.add_argument("--headless", default="false", choices=["true", "false"],
                    help="头条风控会静默拦截 headless 发布，默认有头（false）。仅调试时用 true")
    a = ap.parse_args()

    text = Path(a.body_file).read_text(encoding="utf-8")
    paras = [p.strip() for p in text.split("\n") if p.strip()]
    keyword = a.keyword or "互联网"
    media, account, state_file = common.state_for_code(a.code)
    warn = common.title_warning(a.title, media)
    if warn:
        print(f"⚠ {warn}")

    with sync_playwright() as p:
        # 反检测：禁用自动化标志（实测 navigator.webdriver=None）+ 默认有头
        browser = p.chromium.launch(
            headless=a.headless == "true",
            args=["--disable-blink-features=AutomationControlled", "--no-first-run",
                  "--no-default-browser-check"],
        )
        ctx = browser.new_context(storage_state=str(state_file))
        ctx.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        pg = ctx.new_page()
        try:
            fill_article(pg, a.title, paras)
            open_gallery(pg)
            search_gallery(pg, keyword)
            pick_random_image(pg)
            confirm_insert(pg)
            set_cover_single(pg)
            pg.screenshot(path=str(SHOTS / "toutiao-v2-filled.png"))
            if not a.go:
                print("DRYRUN 完成，未发布。截图：", SHOTS / "toutiao-v2-filled.png")
                return 0
            do_publish(pg)
        finally:
            browser.close()

    if a.go:
        pending = json.loads(PENDING.read_text()) if PENDING.exists() else []
        pending.append({
            "code": a.code, "media": media, "account": account,
            "title": a.title, "published_at": datetime.now().isoformat(timespec="seconds"),
            "status": "审核中",
        })
        PENDING.write_text(json.dumps(pending, ensure_ascii=False, indent=2))
        print(f"已登记待回查：{PENDING}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
