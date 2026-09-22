#!/usr/bin/env python3
"""共享配置：所有脚本从 accounts.yaml 读编码/账号/登录态/平台参数。

accounts.yaml 是唯一真实来源；环境变量可覆盖：
  ANYSEARCH_CMD  企业资料检索命令（默认自动探测，见 resolve_anysearch）
  MEDIA_PENDING  待回查文件路径（默认 skill 根目录 pending-links.json）
  MEDIA_LINKS    回查结果 CSV 路径（默认 skill 根目录 links.csv）
  MEDIA_LLM_CMD  兼容的双阶段 LLM 命令（为空则用内置模板，见 pipeline.py）
  MEDIA_HUMAN_WRITING_CMD  human-writing 第一阶段命令
  MEDIA_HUMANIZER_ZH_CMD   humanizer-zh 第二阶段命令
"""
from __future__ import annotations

import os
import re
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
    """按编码返回 (media, account, state)。未知编码抛 KeyError（附已知列表）。

    登录态文件缺失时抛 FileNotFoundError 并提示补录命令，调用方（pipeline/
    publish_*.py）直接把这段话展示给用户，不抛无头 traceback。"""
    e = codes().get(code)
    if not e:
        raise KeyError(f"未知媒体编码 {code!r}，已知：{sorted(codes())}")
    media, account, path = e["media"], e["account"], Path(e["state"])
    if not path.exists():
        raise FileNotFoundError(
            f"[{code}] 无登录态文件：{path}。请先补录登录："
            f"python3 scripts/login.py {media} --account {account} "
            f"（或跑 python3 scripts/setup_accounts.py 逐个检查）")
    return media, account, path


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


# 各平台标题字数上限（实测：头条输入框限 30 字、搜狐 placeholder 5-72 字、CSDN 5~100 字）
TITLE_LIMIT = {"toutiao": 30, "sohu": 72, "csdn": 100}
DEFAULT_TITLE_LIMIT = 30  # 多媒体同发时取各目标中最严的一个

# 公司名缩写：超长时剥离这些尾缀（保留地域+品牌，如 …集团有限公司→…集团）
# 注意顺序：“有限公司”必须排在“集团有限公司”之前，否则后者先命中会多剥掉“集团”
COMPANY_TAILS = ("有限责任公司", "股份有限公司", "有限公司",
                  "集团有限公司", "实业有限公司", "集团公司")

# 三级缩写都装不下时的截断锚点：切到“地域+品牌+集团/公司”为止
ABBR_BOUNDARIES = ("集团", "控股", "实业", "公司", "中心")

# 企业名里的通用字样：不能单独当作“标题已出现企业名称”的依据
GENERIC_NAME_FRAGMENTS = ("有限", "公司", "集团", "股份", "责任", "控股", "实业", "中心")


def short_company(name: str) -> str:
    for t in COMPANY_TAILS:
        if name.endswith(t) and len(name) - len(t) >= 2:
            return name[:len(name) - len(t)]
    return name


def abbr_company(name: str) -> str:
    """激进缩写：先去尾缀，再按 集团/控股/公司 等边界截断，保留地域+品牌。
    如：中国长江三峡集团有限公司新能源发展有限责任公司 → 中国长江三峡集团。"""
    s = short_company(name)
    for kw in ABBR_BOUNDARIES:
        i = s.find(kw)
        if i >= 2:
            return s[:i + len(kw)]
    return s


def title_name_variants(company: str, *aliases: str) -> list[str]:
    """标题可用名称：全称、调用方给的简称、去尾缀简称、激进缩写（去重保序）。"""
    candidates = [company, *aliases, short_company(company), abbr_company(company)]
    return [c.strip() for c in dict.fromkeys(candidates) if c and c.strip()]


def title_name_hit(title: str, names: list[str], width: int = 3) -> str:
    """返回标题中用到的企业名称写法，空串表示没认出来。

    先认完整变体（全称/简称/缩写），再认企业名里 ≥width 字的连续片段，这样
    “烟台海烟水产食品有限公司”写成“海烟水产……”这类自然简称也算出现企业名；
    “有限/公司”这类通用字样不算。"""
    for name in sorted(names, key=len, reverse=True):
        if name and name in title:
            return name
    for name in names:
        for i in range(max(0, len(name) - width + 1)):
            fragment = name[i:i + width]
            if any(generic in fragment for generic in GENERIC_NAME_FRAGMENTS):
                continue
            if fragment in title:
                return fragment
    return ""


def title_missing_core(title: str, names: list[str], limit: int = 0) -> str:
    """返回标题缺失的硬要素，空串表示合规。

    企业新闻标题的硬要求只有一条：标题里出现企业名称（全称/简称/缩写，或企业名里
    的自然简称片段）。其余句式、用词和角度不限；只多一条防呆——标题不能只是企业
    名称本身，否则读起来不是标题。"""
    if not title:
        return "标题为空"
    if limit and len(title) > limit:
        return f"超过{limit}字（当前{len(title)}字）"
    if not names:
        return ""
    hit = title_name_hit(title, names)
    if not hit:
        return "缺少企业名称"
    if not title.replace(hit, "", 1).strip("，,：:。 、-—"):
        return "只有企业名称，缺少新闻角度"
    return ""


# 标题完全自由后可能带路径字符，直接当文件名会抛 OSError
ILLEGAL_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def safe_filename(name: str, max_len: int = 80) -> str:
    """把标题转成可用文件名：替换路径非法字符并截断。"""
    return ILLEGAL_FILENAME_CHARS.sub("_", (name or "").strip())[:max_len].strip() or "未命名"


def title_warning(title: str, media: str) -> str:
    """标题超该平台上限时返回提示，否则返回空串。"""
    limit = TITLE_LIMIT.get(media, DEFAULT_TITLE_LIMIT)
    if len(title) > limit:
        return f"标题 {len(title)} 字超{media}上限 {limit} 字，可能被平台截断/拒收"
    return ""


def pending_path() -> Path:
    return Path(os.environ.get("MEDIA_PENDING", str(BASE / "pending-links.json")))


def links_path() -> Path:
    return Path(os.environ.get("MEDIA_LINKS", str(BASE / "links.csv")))
