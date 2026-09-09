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


# 各平台标题字数上限（实测：头条输入框限 30 字、搜狐 placeholder 5-72 字、CSDN 5~100 字）
TITLE_LIMIT = {"toutiao": 30, "sohu": 72, "csdn": 100}
DEFAULT_TITLE_LIMIT = 30  # 多媒体同发时取各目标中最严的一个

# 公司名缩写：超长时剥离这些尾缀（保留地域+品牌，如 …集团有限公司→…集团）
# 注意顺序：“有限公司”必须排在“集团有限公司”之前，否则后者先命中会多剥掉“集团”
COMPANY_TAILS = ("有限责任公司", "股份有限公司", "有限公司",
                  "集团有限公司", "实业有限公司", "集团公司")

# 三级缩写都装不下时的截断锚点：切到“地域+品牌+集团/公司”为止
ABBR_BOUNDARIES = ("集团", "控股", "实业", "公司", "中心")

# 常见行业词（公司名里命中即用，最长匹配优先，如 新能源/医疗器械）
INDUSTRY_WORDS = (
    "医疗器械", "新能源", "新材料", "半导体", "机器人", "无人机",
    "水产", "食品", "医药", "医疗", "科技", "软件", "电子", "机械",
    "化工", "建材", "纺织", "服装", "汽车", "物流", "旅游", "餐饮",
    "教育", "金融", "农业", "种业", "出版", "传媒", "建筑", "地产",
    "物业", "家具", "家电", "照明", "模具", "轴承", "阀门", "电缆",
    "光伏", "锂电", "芯片", "环保", "节能", "养殖", "种植", "茶叶",
    "白酒", "乳业", "饮料", "海鲜", "水果", "蔬菜", "花卉", "苗木",
    "宠物", "母婴", "玩具", "文具", "体育", "健身", "生物",
)

# 无行业词时的中性兜底（事实陈述，无广告色彩）
NEUTRAL_SUFFIXES = ("新入口", "官网直达", "中文直达", "品牌直达")


def industry_of(company: str) -> str:
    for w in sorted(INDUSTRY_WORDS, key=len, reverse=True):
        if w in company:
            return w
    return ""


def title_suffixes(company: str, pick: int = 0) -> list[str]:
    """标题描述位：行业相关优先（{行业}新入口/官网直达…），无行业回退中性词。
    按行号轮换保证同批多样性。"""
    pool: list[str] = []
    ind = industry_of(company)
    if ind:
        pool += [f"{ind}新入口", f"{ind}官网直达", f"{ind}品牌直达", f"{ind}中文直达"]
    pool += list(NEUTRAL_SUFFIXES)
    pool = list(dict.fromkeys(pool))
    return [pool[(pick + i) % len(pool)] for i in range(len(pool))]


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


def build_title(company: str, domain: str, limit: int = DEFAULT_TITLE_LIMIT,
                pick: int = 0, short_name: str = "") -> str:
    """组装标题：名称 + 官网启用 + 域名 + 短描述，保证 len <= limit。
    名称优先级：xlsx 简称列 > 全称 > 去尾缀 > 地域品牌截断；
    描述只是装饰，空间不够时先换短词、再直接省略，保名称+官网启用+域名完整。"""
    suffixes = title_suffixes(company, pick)
    bases = list(dict.fromkeys(
        [b for b in (short_name.strip(), company,
                     short_company(company), abbr_company(company)) if b]))
    for base in bases:
        stem = f"{base}官网启用{domain}"
        for sfx in suffixes:
            t = f"{stem}，{sfx}"
            if len(t) <= limit:
                return t
    for base in bases:  # 描述位让路：名称+官网启用+域名优先
        stem = f"{base}官网启用{domain}"
        if len(stem) <= limit:
            return stem
    # 极端超长：截公司名，保“官网启用+域名”完整
    core = f"官网启用{domain}"
    keep = max(2, limit - len(core))
    return bases[-1][:keep] + core


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
