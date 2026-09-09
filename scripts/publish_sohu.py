#!/usr/bin/env python3
"""搜狐号（a-1）自动发布：填稿 → 必选声明 → 发布。

用法：
  dry-run（填稿+截图，不发布）：
    python3 scripts/publish_sohu.py --title "标题" --body-file 正文.txt
  真发：
    python3 scripts/publish_sohu.py --title "标题" --body-file 正文.txt --go --code a-1

搜狐号无免费图库，封面可跳过（实测有效）。每日限发 5 篇。
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent.parent
STATES = BASE / "states"
PENDING = BASE / "pending-links.json"
SHOTS = Path("/tmp/media-publish-dryrun")
SHOTS.mkdir(parents=True, exist_ok=True)

HOME_URL = "https://mp.sohu.com/mpfe/v4/contentManagement/first/page"


def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def open_editor(pg) -> None:
    pg.goto(HOME_URL)
    pg.wait_for_timeout(6000)
    # 顶部"发布内容"按钮
    pg.evaluate("""(() => {
      const btn = document.querySelector('button.mt-button.publish-btn')
      if (btn) btn.click()
    })()""")
    pg.wait_for_timeout(6000)
    if "addarticle" not in pg.url:
        raise RuntimeError(f"未进入编辑页：{pg.url}")
    # 关"排版新升级"提示
    pg.evaluate("""(() => {
      const b = [...document.querySelectorAll('button, a, [role=button]')]
        .find(b => (b.textContent || '').trim() === '我知道了' && b.offsetParent)
      if (b) b.click()
    })()""")
    pg.wait_for_timeout(1000)
    log("编辑页已打开")


def fill_article(pg, title: str, paras: list[str]) -> None:
    pg.get_by_placeholder("请输入标题（5-72字）").fill(title)
    pg.wait_for_timeout(500)
    pg.evaluate("document.querySelector('.ql-editor[contenteditable=true]').focus()")
    pg.wait_for_timeout(500)
    for i, para in enumerate(paras):
        pg.keyboard.type(para)
        if i < len(paras) - 1:
            pg.keyboard.press("Enter")
            pg.wait_for_timeout(150)
    pg.wait_for_timeout(2000)
    log(f"填稿完成：标题 {len(title)} 字，正文 {len(paras)} 段")
    # 顶部截图（标题+正文）供 dry-run 核验
    pg.evaluate("window.scrollTo(0, 0)")
    pg.wait_for_timeout(800)
    pg.screenshot(path=str(SHOTS / "sohu-filled-top.png"))


def set_declaration(pg) -> None:
    """必选创作声明，默认选「无需声明」。"""
    # 滚到声明区并展开
    pg.evaluate("""(() => {
      const p = document.querySelector('.statement-panel')
      if (p) p.scrollIntoView({block: 'center'})
    })()""")
    pg.wait_for_timeout(800)
    ok = pg.evaluate("""(() => {
      const label = [...document.querySelectorAll('label.el-radio')]
        .find(l => (l.textContent || '').trim() === '无需声明')
      if (!label) return false
      const input = label.querySelector('input.el-radio__original')
      if (!input) return false
      if (input.checked) return true
      // Vue ElementUI 必须原生事件流
      ;['mousedown', 'mouseup', 'click', 'change'].forEach(t => {
        input.dispatchEvent(new MouseEvent(t, {bubbles: true, cancelable: true}))
      })
      return input.checked
    })()""")
    if not ok:
        # 再查一次（事件循环渲染延迟）
        pg.wait_for_timeout(800)
        ok = pg.evaluate("""(() => {
          const label = [...document.querySelectorAll('label.el-radio')]
            .find(l => (l.textContent || '').trim() === '无需声明')
          return label && label.className.includes('is-checked')
        })()""")
    if not ok:
        raise RuntimeError("创作声明「无需声明」未选中")
    log("创作声明：无需声明 ✓")


def do_publish(pg) -> None:
    pg.evaluate("""(() => {
      const li = document.querySelector('li.publish-report-btn.active')
      if (li) li.click()
    })()""")
    pg.wait_for_timeout(4000)
    # 成功标志：跳回 first/page + toast"审核中"
    ok = pg.evaluate("""(() => ({
      back: location.pathname.includes('first/page') || location.pathname.includes('contentManagement'),
      toast: [...document.querySelectorAll('*')]
        .filter(e => e.children.length === 0 && /审核中|发布成功|失败|必填|请添加/.test(e.textContent || '') && e.offsetParent)
        .map(e => e.textContent.trim()).find(t => t.length < 20) || ''
    }))""")
    if not ok["back"]:
        raise RuntimeError(f"发布后未跳转，toast={ok['toast']}")
    log(f"发布提交成功，toast={ok['toast'] or '(无)'}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--title", required=True)
    ap.add_argument("--body-file", required=True)
    ap.add_argument("--code", default="a-1")
    ap.add_argument("--go", action="store_true")
    a = ap.parse_args()

    text = Path(a.body_file).read_text(encoding="utf-8")
    paras = [p.strip() for p in text.split("\n") if p.strip()]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(storage_state=str(STATES / "sohu-默认账号.json"))
        pg = ctx.new_page()
        try:
            open_editor(pg)
            fill_article(pg, a.title, paras)
            set_declaration(pg)
            pg.screenshot(path=str(SHOTS / "sohu-filled.png"))
            if not a.go:
                print("DRYRUN 完成，未发布。截图：", SHOTS / "sohu-filled.png")
                return 0
            do_publish(pg)
        finally:
            browser.close()

    pending = json.loads(PENDING.read_text()) if PENDING.exists() else []
    pending.append({
        "code": a.code, "media": "sohu", "account": "默认账号",
        "title": a.title, "published_at": datetime.now().isoformat(timespec="seconds"),
        "status": "审核中",
    })
    PENDING.write_text(json.dumps(pending, ensure_ascii=False, indent=2))
    print(f"已登记待回查：{PENDING}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
