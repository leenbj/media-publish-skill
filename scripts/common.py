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

# 标题角度规则：只在正文实际出现相应事实时提供可读的新闻标题短语。
# 候选会按正文命中次数排序，并由行号轮换；它们不是固定标题模板。
TITLE_TOPIC_RULES = (
    (("启用", "上线", "入口", "地址换成", "官网地址"),
     ("官网入口更新", "新入口上线")),
    (("浏览器", "输入", "访问", "直达"),
     ("可直接输入访问", "访问方式更直观")),
    (("海产品", "水产品", "海参", "调味品", "食品"),
     ("主营海产品", "主营海产品与调味品")),
    (("主营业务", "主营", "研发", "生产", "供应", "加工", "制造"),
     ("主营业务梳理", "业务资料更新")),
    (("数字化", "转型", "线上", "互联网", "信息化"),
     ("线上服务完善", "数字化进展")),
    (("品牌资产", "防伪", "保护", "仿冒", "钓鱼"),
     ("线上品牌保护", "品牌保护再添入口")),
    (("客户", "用户", "体验"),
     ("用户访问更清晰", "访问体验说明")),
)


def title_topics_from_content(content: str, domain: str = "", pick: int = 0,
                              company: str = "") -> list[str]:
    """从正文事实提取标题角度，不命中正文就不生成该角度。

    域名只有在正文同时提到访问/入口/启用等事实时才进入候选，避免把域名
    无条件塞进每个标题。返回值按命中强度排序并按行号轮换，便于同批标题
    有变化而不改变标题的事实边界。
    """
    text = content or ""
    domain_present = bool(domain and domain in text)
    # 企业名/域名本身可能含“品牌、服务、入口”等字样，不能仅凭身份字段
    # 触发标题角度；角度应来自正文对事实或动作的描述。
    if company:
        for name in dict.fromkeys((company, short_company(company), abbr_company(company))):
            if name:
                text = text.replace(name, "")
    if domain:
        text = text.replace(domain, "")
    ranked: list[tuple[int, int, tuple[str, ...]]] = []
    for index, (keywords, labels) in enumerate(TITLE_TOPIC_RULES):
        hits = sum(text.count(keyword) for keyword in keywords)
        if hits:
            ranked.append((hits, index, labels))
    ranked.sort(key=lambda item: (-item[0], item[1]))

    # 业务、产品、数字化、品牌等实质角度优先于“入口/访问”这类通用事实；
    # 只有正文没有实质角度时，才用入口或访问方式作标题补充。
    specific = [item for item in ranked if item[1] >= 2]
    general = [item for item in ranked if item[1] < 2]
    topics: list[str] = []
    access_terms = ("启用", "上线", "入口", "地址", "访问", "输入", "直达")
    if specific:
        ranked_for_title = specific + general
    else:
        if domain_present and any(term in text for term in access_terms):
            topics.append(f"{domain}入口")
        ranked_for_title = general
    for _, _, labels in ranked_for_title:
        topics.extend(labels)

    unique = list(dict.fromkeys(topic for topic in topics if topic))
    if unique:
        offset = pick % len(unique)
        unique = unique[offset:] + unique[:offset]
    return unique


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
                pick: int = 0, short_name: str = "", content: str = "") -> str:
    """组装可阅读的企业新闻标题，保证 len <= limit。

    标题只要求出现企业全称或简称，并补充正文支持的动作或事实角度；域名的
    完整写法留在正文首段，不再强制塞进标题。这样既能保留企业识别，又能让
    标题读起来像正常新闻标题，而不是字段拼接。
    """
    company = company.strip()
    short_name = short_name.strip()
    bases = list(dict.fromkeys(
        [b for b in (company, short_name, short_company(company), abbr_company(company)) if b]))
    topics = title_topics_from_content(content, domain, pick, company)
    access_terms = ("启用", "上线", "入口", "地址", "访问", "输入", "直达")
    # 事件句只保留新闻动作，不携带完整域名；完整域名在正文首段核实。
    event_labels = ("官网入口更新", "官网地址更新", "官网上线")
    event_label = (event_labels[0]
                   if domain and domain in (content or "")
                   and any(term in (content or "") for term in access_terms)
                   else "")

    for base in bases:
        if len(base) > limit:
            continue
        if event_label:
            candidate = f"{base}{event_label}"
            if len(candidate) <= limit:
                return candidate
            for topic in topics:
                candidate = f"{base}，{topic}"
                if len(candidate) <= limit:
                    return candidate
            # 当前名称放不下动作句时，继续尝试 xlsx 简称或品牌缩写，避免标题
            # 退化成只有企业名称的字段。
            continue
        for topic in topics:
            candidate = f"{base}，{topic}"
            if len(candidate) <= limit:
                return candidate
        return base

    # 极端长企业名：保留最短可用名称；这是平台硬上限下的最后兜底。
    fallback = next((base for base in reversed(bases) if base), "企业")
    action = "，入口更新"
    if limit > len(action) + 3:
        return f"{fallback[:limit - len(action)]}{action}"
    return fallback[:max(0, limit)]


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
