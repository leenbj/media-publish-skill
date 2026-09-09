#!/usr/bin/env python3
"""CSDN（c-1）自动发布：填标题+正文 → 文章标签 → 发布博客。

用法：
  dry-run：
    python3 scripts/publish_csdn.py --title "标题" --body-file 正文.txt
  真发：
    python3 scripts/publish_csdn.py --title "标题" --body-file 正文.txt --go --code c-1

说明：
- 正文填入 CKEditor iframe（富文本模式）
- 文章标签必填（默认搜"域名"选"域名与商标"，可用 --tag 改）
- 创作活动/话题当前无进行中活动，可跳过（实测不阻塞发布）
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
import common
PENDING = common.pending_path()
SHOTS = Path("/tmp/media-publish-dryrun")
SHOTS.mkdir(parents=True, exist_ok=True)

EDITOR_URL = "https://mp.csdn.net/mp_blog/creation/editor"
DEFAULT_TAG_SEARCH = "域名"
DEFAULT_TAG = "域名与商标"


def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def html_escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def fill_article(pg, title: str, paras: list[str]) -> None:
    pg.goto(EDITOR_URL)
    pg.wait_for_timeout(8000)
    if pg.locator('iframe[src*="passport.csdn.net"]').count():
        raise RuntimeError("CSDN 登录态失效，请重新跑 login.py csdn")
    # 标题（textarea）
    pg.evaluate("""((title) => {
      const t = [...document.querySelectorAll('textarea')]
        .find(x => x.placeholder && x.placeholder.includes('文章标题'))
      const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set
      t.focus()
      setter.call(t, title)
      t.dispatchEvent(new Event('input', {bubbles: true}))
    })""", title)
    pg.wait_for_timeout(800)
    # 正文：CKEditor 4 必须真实键入（直接改 innerHTML 不触发编辑器同步，字数始终为0）
    frame_el = pg.query_selector("iframe.cke_wysiwyg_frame")
    box = frame_el.bounding_box()
    pg.mouse.click(box["x"] + 100, box["y"] + 20)
    pg.wait_for_timeout(800)
    for i, para in enumerate(paras):
        pg.keyboard.type(para, delay=15)
        if i < len(paras) - 1:
            pg.keyboard.press("Enter")
            pg.wait_for_timeout(200)
    pg.wait_for_timeout(2000)
    # 校验编辑器字数（'共 N 字'必须 > 0，否则 CKEditor 未接收内容）
    counter = pg.evaluate("""(() => {
      const c = [...document.querySelectorAll('*')]
        .filter(e => e.offsetParent && /共 \\d+ 字/.test(e.textContent || '') && e.children.length === 0)
      return c.length ? c[0].textContent.match(/共 (\\d+) 字/)[1] : '0'
    })""")
    if int(counter) == 0:
        raise RuntimeError("CKEditor 未接收正文内容（共 0 字）")
    log(f"填稿完成：标题 {len(title)} 字，正文编辑器计数 {counter} 字")


def set_tag(pg, search: str, tag: str) -> None:
    pg.evaluate("""(() => {
      const b = [...document.querySelectorAll('button.tag__btn-tag')]
        .find(x => (x.textContent || '').trim() === '添加文章标签')
      if (b) { b.scrollIntoView({block: 'center'}); b.click() }
    })()""")
    pg.wait_for_timeout(2000)
    # 搜索框：真实键入（search-recommend-tag 接口由 Vue watch 触发，需真实输入）
    # 实测搜索接口对'域名'无结果 → 用 Enter 直接创建自定义标签（官方支持：Enter键入可添加自定义标签）
    loc = pg.locator("input[placeholder*='自定义标签']").first
    loc.click()
    loc.type(tag, delay=120)
    pg.keyboard.press("Enter")
    pg.wait_for_timeout(2500)
    picked = pg.evaluate("""((tag) => {
      const sel = [...document.querySelectorAll('[class*=tag-box] *')]
        .filter(e => e.offsetParent && e.children.length === 0 && e.textContent.trim() === tag)
      return sel.length > 0
    })""", tag)
    if not picked:
        raise RuntimeError(f"自定义标签「{tag}」创建失败")
    pg.wait_for_timeout(800)
    pg.mouse.click(200, 300)  # 关弹层
    pg.wait_for_timeout(500)
    log(f"文章标签（自定义创建）：{tag} ✓")


def do_publish(pg) -> str:
    pg.evaluate("""(() => {
      const btn = [...document.querySelectorAll('button')]
        .find(b => b.offsetParent && /发布博客/.test((b.textContent || '').trim()))
      if (btn) { btn.scrollIntoView({block: 'center'}); btn.click() }
    })()""")
    pg.wait_for_timeout(5000)
    url = pg.url
    if "creation/success" not in url:
        toast = pg.evaluate("""(() => {
          const t = [...document.querySelectorAll('*')]
            .filter(e => e.children.length === 0 && /发布成功|失败|请|必须/.test(e.textContent || '') && e.offsetParent)
          return t.length ? t[0].textContent.trim() : ''
        })()""")
        raise RuntimeError(f"发布后未跳转成功页，toast={toast}")
    log(f"发布提交成功 → {url}")
    return url


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--title", required=True)
    ap.add_argument("--body-file", required=True)
    ap.add_argument("--tag-search", default=DEFAULT_TAG_SEARCH)
    ap.add_argument("--tag", default=DEFAULT_TAG)
    ap.add_argument("--code", default="c-1")
    ap.add_argument("--go", action="store_true")
    a = ap.parse_args()

    text = Path(a.body_file).read_text(encoding="utf-8")
    paras = [p.strip() for p in text.split("\n") if p.strip()]
    media, account, state_file = common.state_for_code(a.code)
    warn = common.title_warning(a.title, media)
    if warn:
        print(f"⚠ {warn}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(storage_state=str(state_file))
        pg = ctx.new_page()
        try:
            fill_article(pg, a.title, paras)
            set_tag(pg, a.tag_search, a.tag)
            pg.screenshot(path=str(SHOTS / "csdn-filled.png"))
            if not a.go:
                print("DRYRUN 完成，未发布。截图：", SHOTS / "csdn-filled.png")
                return 0
            do_publish(pg)
        finally:
            browser.close()

    pending = json.loads(PENDING.read_text()) if PENDING.exists() else []
    pending.append({
        "code": a.code, "media": media, "account": account,
        "title": a.title, "published_at": datetime.now().isoformat(timespec="seconds"),
        "status": "审核中",
        "success_url": None,  # success 页含文章 id，回查时从列表页拿正式链接
    })
    PENDING.write_text(json.dumps(pending, ensure_ascii=False, indent=2))
    print(f"已登记待回查：{PENDING}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
