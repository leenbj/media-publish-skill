#!/usr/bin/env python3
"""共享配置：所有脚本从 accounts.yaml 读编码/账号/登录态/平台参数。

accounts.yaml 是唯一真实来源；环境变量可覆盖：
  ANYSEARCH_CMD  企业资料检索命令（默认自动探测，见 resolve_anysearch）
  MEDIA_PENDING  待回查文件路径（默认 skill 根目录 pending-links.json）
  MEDIA_LINKS    回查结果 CSV 路径（默认 skill 根目录 links.csv）
  MEDIA_LLM_CMD  新闻稿生成 LLM 命令（为空则用内置模板，见 pipeline.py）
"""
from __future__ import annotations

import os
import shlex
import shutil
from pathlib import Path

import yaml

BASE = Path(__file__).resolve().parent.parent
CONFIG = BASE / "accounts.yaml"

# 内置兜底：本机历史路径，不存在则忽略（靠 PATH 探测）
_LEGACY_ANYSEARCH = "/Users/ethan/.hermes/hermes-agent/venv/bin/python3 /Users/ethan/.hermes/skills/anysearch/scripts/anysearch_cli.py"

_cfg = None


def config() -> dict:
    """解析 accounts.yaml（带缓存）。"""
    global _cfg
    if _cfg is None:
        _cfg = yaml.safe_load(CONFIG.read_text(encoding="utf-8")) or {}
    return _cfg


def codes() -> dict:
    """编码 → 条目（含绝对路径 state）。"""
    out = {}
    for acc in config().get("accounts", []):
        code = str(acc.get("code", "")).strip()
        if not code:
            continue
        entry = dict(acc)
        entry["state"] = str(BASE / acc.get("state", f"states/{acc.get('media')}-{acc.get('account')}.json"))
        out[code] = entry
    return out


def media_of(code: str) -> str | None:
    e = codes().get(code)
    return e["media"] if e else None


def state_for(media: str, account: str) -> Path:
    """按 (media, account) 找 state；accounts.yaml 未登记时回退默认命名。"""
    for e in codes().values():
        if e.get("media") == media and e.get("account") == account:
            return Path(e["state"])
    return BASE / f"states/{media}-{account}.json"


def state_for_code(code: str) -> tuple[str, str, Path]:
    """按编码返回 (media, account, state)。未知编码抛 KeyError（附已知列表）。"""
    e = codes().get(code)
    if not e:
        raise KeyError(f"未知媒体编码 {code!r}，已知：{sorted(codes())}")
    return e["media"], e["account"], Path(e["state"])


def media_param(media: str, key: str, default: str = "") -> str:
    """读 accounts.yaml 顶层 media_params.<media>.<key>。"""
    return str(config().get("media_params", {}).get(media, {}).get(key, default))


def default_accounts() -> list[dict]:
    """setup_accounts/login 用的账号清单：{media, account, label}。"""
    return [{"media": a["media"], "account": a["account"],
             "label": a.get("note") or a["media"]} for a in config().get("accounts", [])]


def resolve_anysearch() -> list[str]:
    """检索命令 argv 前缀：$ANYSEARCH_CMD > accounts.yaml > PATH > 历史路径。"""
    raw = (os.environ.get("ANYSEARCH_CMD", "")
           or str(config().get("anysearch_cmd", "")))
    if raw:
        return shlex.split(raw)
    found = shutil.which("anysearch_cli.py")
    if found:
        py = shutil.which("python3") or "python3"
        return [py, found]
    if _LEGACY_ANYSEARCH.split()[-1] and Path(_LEGACY_ANYSEARCH.split()[-1]).exists():
        return shlex.split(_LEGACY_ANYSEARCH)
    raise FileNotFoundError(
        "找不到企业检索命令：设置 ANYSEARCH_CMD 环境变量，或把 anysearch_cli.py 放入 PATH")


def pending_path() -> Path:
    return Path(os.environ.get("MEDIA_PENDING", str(BASE / "pending-links.json")))


def links_path() -> Path:
    return Path(os.environ.get("MEDIA_LINKS", str(BASE / "links.csv")))
