#!/usr/bin/env python3
"""accounts.yaml 代理支持 + 发布脚本统一代理读取。

设计（skill v1.2 预留的实施方案）：
- accounts.yaml 每条账号可加 proxy: <url>（如 socks5://user:pass@host:port 或 http://...）
- 登录态与代理是成对资产：首次在某代理下录入登录态后，该账号必须固定用同一代理
- 未配置 proxy 的账号直连（现状）

本脚本：加载并校验 accounts.yaml 的代理配置，输出每个编码的生效连接方式。
"""
import sys
from pathlib import Path

import yaml

BASE = Path(__file__).resolve().parent.parent


def main() -> int:
    with open(BASE / "accounts.yaml", encoding="utf-8") as f:
        accounts = yaml.safe_load(f)["accounts"]
    print("账号连接配置：")
    has_proxy = False
    for a in accounts:
        proxy = a.get("proxy")
        mark = f"代理 {proxy}" if proxy else "直连"
        state = BASE / a["state"]
        state_ok = "✓" if state.exists() else "✗缺登录态"
        print(f"  {a['code']} [{a['media']}/{a['account']}] {mark} 登录态{state_ok}")
        if proxy:
            has_proxy = True
    if not has_proxy:
        print("\n当前全部直连。启用代理：在 accounts.yaml 对应账号加 proxy: <url>，")
        print("然后删除该账号旧登录态并重新 setup_accounts.py 录入（IP 变更必须重录）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
