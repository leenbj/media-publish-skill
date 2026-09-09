#!/usr/bin/env python3
"""媒体账号登录态管理。

一次扫码保存 storage_state（cookie+localStorage），之后发布任务自动加载对应账号，
各账号存储完全隔离、互不串号。

用法：
  python3 scripts/login.py sohu --account 账号A        # 打开浏览器扫码登录并保存
  python3 scripts/login.py sohu --account 账号A --check # 校验该账号登录态是否有效
  python3 scripts/login.py --list                       # 列出已保存的账号
"""
import argparse
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
import common
STATES = BASE / "states"
STATES.mkdir(exist_ok=True)

MEDIA = {
    "sohu": {
        "label": "搜狐号",
        "login_url": "https://mp.sohu.com/mpfe/v4/login",
        "check_url": "https://mp.sohu.com/mpfe/v4/index",
    },
    "toutiao": {
        "label": "头条号",
        "login_url": "https://mp.toutiao.com/profile_v4/index",
        "check_url": "https://mp.toutiao.com/profile_v4/index",
    },
    "csdn": {
        "label": "CSDN博客",
        "login_url": "https://passport.csdn.net/login?code=applets",
        "check_url": "https://mp.csdn.net/edit",
    },
}


def state_path(media: str, account: str) -> Path:
    return common.state_for(media, account)


def do_login(media: str, account: str) -> int:
    cfg = MEDIA[media]
    path = state_path(media, account)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        ctx = browser.new_context()
        page = ctx.new_page()
        page.goto(cfg["login_url"])
        print(f"[{cfg['label']}/{account}] 已打开登录页，请在弹出的浏览器窗口中完成登录（扫码/账密）。")
        try:
            input("登录完成后按回车保存登录态…")
        except EOFError:
            pass
        ctx.storage_state(path=str(path))
        print(f"已保存：{path}（{path.stat().st_size} 字节）")
        # 就地校验
        page.goto(cfg["check_url"])
        page.wait_for_timeout(3000)
        print(f"校验：当前 URL = {page.url}")
        print(f"校验：标题 = {page.title()}")
        browser.close()
    return 0


def do_check(media: str, account: str) -> int:
    cfg = MEDIA[media]
    path = state_path(media, account)
    if not path.exists():
        print(f"[{cfg['label']}/{account}] 未找到登录态文件：{path}，请先跑 login 保存。")
        return 1
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(storage_state=str(path))
        page = ctx.new_page()
        page.goto(cfg["check_url"])
        page.wait_for_timeout(4000)
        url, title = page.url, page.title()
        browser.close()
    print(f"[{cfg['label']}/{account}] check_url 跳转后 URL = {url}")
    print(f"[{cfg['label']}/{account}] 标题 = {title}")
    if "passport.csdn.net" in url or "mpfe/v4/login" in url or "sso.toutiao.com" in url:
        print("→ 登录态已失效，需要重新扫码。")
        return 2
    print("→ 登录态有效。")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("media", nargs="?", choices=list(MEDIA) + ["--list"])
    ap.add_argument("--account", default="默认账号")
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--list", action="store_true")
    a = ap.parse_args()
    if a.list or a.media == "--list":
        files = sorted(STATES.glob("*.json"))
        print("已保存账号：" if files else "暂无已保存账号。")
        for f in files:
            print(f"  - {f.stem}（{f.stat().st_size} 字节）")
        return 0
    if not a.media:
        ap.error("需要指定媒体：sohu / toutiao / csdn")
    if a.check:
        return do_check(a.media, a.account)
    return do_login(a.media, a.account)


if __name__ == "__main__":
    sys.exit(main())
