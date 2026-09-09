#!/usr/bin/env python3
"""首次使用引导：检查各账号登录态，缺失/失效的逐个引导录入。

用法：python3 scripts/setup_accounts.py [--media sohu|toutiao|csdn]
"""
import json
import subprocess
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
import common
STATES = BASE / "states"

ACCOUNTS = common.default_accounts()

CHECK_URL = {
    "sohu": "https://mp.sohu.com/mpfe/v4/index",
    "toutiao": "https://mp.toutiao.com/profile_v4/index",
    "csdn": "https://mp.csdn.net/mp_blog/creation/editor",
}
FAIL_MARK = {
    # sohu：登录失效时首页 iframe 里有"安全验证"；正常时页面有"发布内容"菜单
    "sohu": lambda url, html: "安全验证" in html or "请登录" in html[:5000],
    "toutiao": lambda url, html: "sso.toutiao" in url or "passport" in url,
    "csdn": lambda url, html: "passport.csdn.net" in url,
}


def state_summary(path: Path) -> str:
    d = json.loads(path.read_text())
    nc = len(d.get("cookies", []))
    nl = sum(len(o.get("localStorage", [])) for o in d.get("origins", []))
    return f"cookie {nc} 条 + localStorage {nl} 键"


def check_valid(media: str, state_file: Path) -> tuple[bool, str]:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(storage_state=str(state_file))
        pg = ctx.new_page()
        pg.goto(CHECK_URL[media])
        pg.wait_for_timeout(5000)
        url, html = pg.url, pg.content()
        browser.close()
    if FAIL_MARK[media](url, html):
        return False, f"已失效（跳转 {url[:60]}）"
    return True, "有效"


def guide_one(media: str, account: str, label: str) -> None:
    print(f"\n═══ {label}（{media}/{account}）登录录入 ═══")
    print("即将弹出真实浏览器窗口 → 请在窗口中登录该平台账号（扫码/账密均可）")
    print("登录完成后回到终端按回车，系统自动保存登录态。")
    input("按回车打开浏览器…")
    subprocess.run([sys.executable, str(BASE / "scripts" / "login.py"), media, "--account", account], check=True)


def main() -> int:
    args = sys.argv[1:]
    only = args[args.index("--media") + 1] if "--media" in args else None
    STATES.mkdir(exist_ok=True)

    print("╔══════════════════════════════════════════╗")
    print("║   媒体多账号发布系统 · 账号登录态检查     ║")
    print("╚══════════════════════════════════════════╝")
    print("保存方式说明：每个账号保存一份 storage_state 文件")
    print("（= cookie + localStorage 快照，比单纯导 cookie 更可靠，")
    print("  因为很多平台的登录凭证存在 localStorage 里）\n")

    missing = []
    for acc in ACCOUNTS:
        if only and acc["media"] != only:
            continue
        path = common.state_for(acc["media"], acc["account"])
        if not path.exists():
            print(f"[{acc['label']}] ✗ 无登录态文件")
            missing.append(acc)
            continue
        ok, msg = check_valid(acc["media"], path)
        mark = "✓" if ok else "✗"
        print(f"[{acc['label']}] {mark} {msg}（{state_summary(path)}）")
        if not ok:
            missing.append(acc)

    if not missing:
        print("\n全部账号登录态有效，可以直接发布。")
        return 0

    print(f"\n共 {len(missing)} 个账号需要录入登录态，开始逐个引导：")
    for acc in missing:
        guide_one(acc["media"], acc["account"], acc["label"])
    print("\n全部录入完成。用 --check 校验：python3 scripts/login.py --list")
    return 0


if __name__ == "__main__":
    sys.exit(main())
