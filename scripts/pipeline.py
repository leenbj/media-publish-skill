#!/usr/bin/env python3
"""一站式发布流水线 v2.1（skill 核心）。

输入：xlsx（列：编号 | 域名 | 企业名称 | 简称（可选）| 企业信息资料（可选，有则为写作首要依据）| 媒体/编码…，每行可发多个媒体）
流程：读表 → 按企业信息资料派生查询 + 公司名兜底做 anysearch 检索 → 汇总（表格为准、检索补充、冲突弃检索侧）→ human-writing 写实 → humanizer-zh 清腔 →
     生成新闻稿（标题含企业名称或简称，句式角度自由）→ 红线检查 →
     生成人民网风格 GEO 网页（1 新闻页 + 3 QA 页）→
     全部存入 output/<域名原样>/ → 逐媒体发布 → 回查正式链接 → 回写 xlsx（每个媒体一列）。

产出文件命名：
  output/海宝源.网址/
    ├── 烟台海烟水产食品有限公司，官网入口更新.md
    ├── 烟台海烟水产食品有限公司，官网入口更新.html ← 新闻页
    ├── 烟台海烟水产食品有限公司官网-QA1.html                    ← QA 页 ×3
    ├── ...
    └── _publish_body.txt（发布用临时正文）

xlsx 回写：每个媒体编码（a-1/b-1/c-1/c-2）各占一列，发布后写正式链接，
审核中写"(审核中)"占位，回查脚本下次覆盖。
"""
from __future__ import annotations

import argparse
import csv
import html
import json
import os
import re
import shlex
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common

BASE = Path(__file__).resolve().parent.parent
OUTPUT_ROOT = None  # 运行时 = xlsx 所在文件夹
PENDING = common.pending_path()
LINKS_CSV = common.links_path()

MEDIA_OF = {code: e["media"] for code, e in common.codes().items()}
MEDIA_PUBLISH = {
    "sohu": BASE / "scripts" / "publish_sohu.py",
    "toutiao": BASE / "scripts" / "publish_toutiao_v2.py",
    "csdn": BASE / "scripts" / "publish_csdn.py",
}
HUMANIZED_CHECK = BASE / "scripts" / "check_humanized.py"
MIN_NEWS_BODY_CHARS = 1000  # 新闻稿正文最低字数（不含标题和 Markdown 空白）
MIN_OPINION_BODY_CHARS = 600  # 观点文正文最低字数
QA_PAGES = 3  # GEO 问答页数（每页 3 组问答，全部由 LLM 写）
URL_MATERIAL = BASE / "references" / "url-material.md"


def url_material_context() -> str:
    """读取用户提供的 `.网址` 背景提炼，供写作时当参考材料。

    这里只负责搬运参考文件内容，不在代码里写任何预设文案；文件缺失就返回空串，
    由 LLM 按提示词自行组织内容。"""
    try:
        return URL_MATERIAL.read_text(encoding="utf-8")[:5000]
    except OSError:
        return ""

# ── .网址 宣传红线 ──
FORBIDDEN = [
    r"[\w-]+\.com\b", r"[\w-]+\.cn\b", r"[\w-]+\.net\b", r"[\w-]+\.org\b",
    r"[\w-]+\.top\b", r"[\w-]+\.vip\b", r"[\w-]+\.shop\b", r"[\w-]+\.xyz\b",
    r"[\w-]+\.com\.cn\b", r"[\w-]+\.net\.cn\b", r"[\w-]+\.org\.cn\b",
    r"[\w-]+\.info\b", r"[\w-]+\.biz\b", r"[\w-]+\.(cc|tv|io|ai|me|co)\b",
    r"[\w-]+\.(club|site|online|store|ltd|name|pro|mobi|asia)\b",
    r"[\w-]+\.(hk|tw)\b",
    r"\.商标", r"\.商城", r"\.在线", r"\.中国(?!.*网址)", r"\.公司(?!.*网址)",
    r"英文域名", r"国际域名",
    r"中文域名[^。]{0,20}(争议|质疑|缺点|不足|局限性|风险)",
]


def check_content(text: str, domain: str) -> tuple[bool, list[str]]:
    issues = []
    body = re.sub(r"[\w\u4e00-\u9fa5-]*\.网址", "", text.replace(domain, ""))
    for pat in FORBIDDEN:
        for m in re.finditer(pat, body):
            ctx = body[max(0, m.start() - 15):m.end() + 15].replace("\n", " ")
            issues.append(f"[{pat}] …{ctx}…")
    return (len(issues) == 0, issues)


def clean_text(text: str, domain: str = "") -> str:
    """洗搜索/生成文本：去 HTML 标签、解实体、删编码替换符和控制字符、
    中文之间的半角标点转全角（先保护域名， avoids 源.网→源。网）。"""
    text = re.sub(r"<[^>]{1,200}>", "", text)
    text = html.unescape(text)
    text = text.replace("�", "")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    text = text.replace("**", "")  # 搜索结果的加粗标记在纯文本稿里是噪音
    text = re.sub(r"(?m)^#{1,6}\s+", "", text)  # Markdown 标题标记
    keep = []
    if domain:  # 保护域名与其余 .网址，原样钉住不动
        def _pin(m: re.Match) -> str:
            keep.append(m.group(0))
            return chr(0xE000 + len(keep))
        text = re.sub(r"[\w\u4e00-\u9fa5-]*\.网址", _pin,
                      text.replace(domain, chr(0xE000)))
        keep.insert(0, domain)
    for asc, full in ((",", "，"), (";", "；"), (":", "："),
                       ("?", "？"), ("!", "！"), (".", "。")):
        text = re.sub(f"(?<=[\u4e00-\u9fa5]){re.escape(asc)}(?=[\u4e00-\u9fa5]|$)",
                      full, text)
    for i, s in enumerate(keep):
        text = text.replace(chr(0xE000 + i), s)
    return re.sub(r"[ \t]+", " ", text).strip()


# 硬错：出现即污染正文，洗不掉就整行跳过
HARD_PATTERNS = [
    ("HTML标签残留", r"</?[a-zA-Z][^>]{0,100}>"),
    ("HTML实体残留", r"&(amp|lt|gt|quot|nbsp|#\d+);"),
    ("编码替换符", r"�"),
    ("私用区残留占位", r"[\ue000-\uf8ff]"),
    ("控制字符", r"[\x00-\x08\x0b\x0c\x0e-\x1f]"),
]

# 告警：可能是误伤，只打印不拦截
WARN_PATTERNS = [
    ("异常重复标点", r"[？?！!。，、；：]{3,}"),
    ("半角标点夹中文", r"[\u4e00-\u9fa5][,.!?;:][\u4e00-\u9fa5]"),
    ("疑似重复字词", r"([\u4e00-\u9fa5]{2,4})\1"),
]

# 写作提示或审核说明误写进正文的元话语。新闻可以说“公开资料显示”，但不应在成稿中解释
# 自己如何检索、如何审核或为何保留某条信息。
META_LEAK_TERMS = (
    "本文", "本稿", "稿件", "提示词", "提示内容", "检索结果", "检索页面", "公开检索",
    "信息边界", "写入本文", "本次报道", "上述资料", "上述材料", "作为企业情况的来源",
)


def scan_text(text: str, domain: str) -> tuple[list[str], list[str]]:
    """返回 (硬错, 告警)。扫描前先抠掉域名，避免“源.网”误报半角标点。"""
    masked = text.replace(domain, "□□")
    masked = re.sub(r"[\w\u4e00-\u9fa5-]*\.网址", "□□", masked)
    hard, warn = [], []
    for name, pat in HARD_PATTERNS:
        for m in re.finditer(pat, masked):
            ctx = masked[max(0, m.start() - 10):m.end() + 10].replace("\n", " ")
            hard.append(f"[{name}] …{ctx}…")
    for name, pat in WARN_PATTERNS:
        for m in re.finditer(pat, masked):
            ctx = masked[max(0, m.start() - 10):m.end() + 10].replace("\n", " ")
            warn.append(f"[{name}] …{ctx}…")
            if len(warn) >= 5:
                break
    for sent in masked.split("\n"):
        if len(sent.strip()) > 120 and not re.search(r"[，。；：？！]", sent):
            warn.append(f"[长句无断句{len(sent.strip())}字] …{sent.strip()[:40]}…")
    return hard, warn


# 代码字符：出现在新闻稿里即不合格
CODE_PATTERNS = [
    ("Markdown代码块", r"```"),
    ("Markdown标题", r"(?m)^#{1,6}\s"),
    ("Markdown加粗", r"\*\*.+?\*\*"),
    ("Markdown链接", r"!\[|\]\("),
    ("网址链接", r"https?://|www\."),
    ("JSON残留", r"\{[^\n]{0,60}\w+\s*:"),
    ("占位符残留", r"__\w+__|\bTODO\b|\bTBD\b|XXX"),
]

# .网址定向负面词（防范/防止类保护语境除外，见 check_negative）
NEG_WORDS = ("骗局", "陷阱", "争议", "质疑", "缺点", "不足", "局限", "风险",
             "隐患", "投诉", "垃圾", "山寨", "不值", "慎用", "圈套", "欺诈", "诈骗")
PROTECT_WORDS = ("防", "防止", "防范", "谨防", "避免", "打击", "抵制",
                 "远离", "警惕", "识别")


def check_negative(text: str, domain: str) -> list[str]:
    """负面词与“中文域名/.网址”相距 12 字内、中间无句号阻断、且无保护语境→定向负面。"""
    issues = []
    refs = [m.start() for m in re.finditer(r"中文域名|\.网址", text)]
    if not refs:
        return issues
    for w in NEG_WORDS:
        for m in re.finditer(re.escape(w), text):
            near = [r for r in refs if abs(r - m.start()) <= 12]
            if not near:
                continue
            r = min(near, key=lambda x: abs(x - m.start()))
            lo, hi = sorted((r, m.start()))
            if re.search(r"[。！？]", text[lo:hi]):
                continue
            win = text[max(0, m.start() - 12):m.end() + 12]
            if any(p in win for p in PROTECT_WORDS):
                continue
            issues.append(f"[负面指向.网址] …{win.replace(chr(10), ' ')}…")
    return issues




# 正文里不得出现的数据站/平台名称：企业事实一律写成“公开工商登记信息”
SOURCE_SITE_TERMS = (
    "企查查", "爱企查", "天眼查", "启信宝", "水滴信用", "顺企网", "黄页88", "黄页",
    "BOSS直聘", "Boss直聘", "boss直聘", "乡振网", "职友集", "看准网", "猎聘", "智联招聘", "58同城",
    "百度百科", "搜狗百科", "维基百科", "360百科",
)
# 任何联系方式都不允许出现（前后限制成非数字，避免误伤统一社会信用代码）
CONTACT_PATTERNS = (
    ("手机号", r"(?<!\d)1[3-9]\d{9}(?!\d)"),
    ("座机号", r"(?<!\d)0\d{2,3}[-‐ ]?\d{7,8}(?!\d)"),
    ("邮箱", r"[\w.+-]+@[\w-]+\.[A-Za-z]{2,}"),
    ("微信或QQ", r"微信(号|二维码)?|QQ\s*[:：]?\s*\d{5,}"),
)


def check_sourcing(text: str) -> list[str]:
    """检查来源写法与联系方式：数据站名称、电话/邮箱/微信/QQ 都不得出现在成稿里。"""
    issues = []
    for term in SOURCE_SITE_TERMS:
        if term in text:
            issues.append(f"[数据站名称] 正文出现“{term}”，企业事实只能写成“公开工商登记信息”")
    for label, pat in CONTACT_PATTERNS:
        m = re.search(pat, text)
        if m:
            issues.append(f"[联系方式] 正文出现{label}：{m.group()}")
    return issues


# 中文域名/.网址 相关段落的篇幅下限：提示词要求约占正文三分之一
DOMAIN_TOPIC_SHARE_MIN = 0.25
DOMAIN_TOPIC_TERMS = ("中文域名", "中文网址", ".网址", "后缀")


def domain_topic_share(body: str) -> float:
    """正文中与中文域名/.网址相关的段落占比（按字符）。"""
    paras = [p.strip() for p in (body or "").split("\n") if p.strip()]
    total = sum(len(p) for p in paras)
    if not total:
        return 0.0
    hit = sum(len(p) for p in paras if any(t in p for t in DOMAIN_TOPIC_TERMS))
    return hit / total


# B 文比较语境白名单：以下原话只作对比论证，放行；其余一律按红线拦。
# 注意：仅观点文（strict=False）走白名单，启用文不受影响。
COMPARATIVE_ALLOW = [
    # 通式：与/和英文域名…(没|无)(任何|实际|本质|明显)?(区别|差异)，接 .com、.cn 或其他补充语均可
    r"(与|和|跟)英文域名(\.com、\.cn)?.{0,12}((没有|无)(任何|实际|本质|明显)?(区别|差异)|一样|无差别|一致)",
    r"远超英文域名",
    r"超过英文域名",
    r"优于英文域名",
    r"弱于\.cn域名",
    # 族级模式：媒体展示类（无法/不能/难以 + 出现/呈现/展示/显示 + 英文域名）
    r"(无法|不能|难以|不可)[^，。]{0,8}(出现|呈现|展示|显示|露出)英文域名",
]


def mask_comparatives(text: str) -> str:
    for pat in COMPARATIVE_ALLOW:
        text = re.sub(pat, "□□", text)
    return text


def prose_char_count(text: str) -> int:
    """统计正文有效字符，忽略空白，供新闻稿最低字数门禁使用。"""
    return len(re.sub(r"\s+", "", text or ""))


def review_article(title: str, body: str, domain: str, strict: bool = True,
                   company: str = "", name_aliases: tuple[str, ...] = ()) -> tuple[list[str], list[str]]:
    """稿件审核，返回 (硬错, 告警)：
    代码字符 / 杂域名后缀 / .网址负面 / 烂尾断句 / 新闻正文不足 1000 字 → 硬错（整行跳过）；
    通顺可疑 → 告警（照常发布）。"""
    errors, warnings = [], []
    hard, warn = scan_text(title + "\n" + body, domain)
    errors += hard
    warnings += warn
    for name, pat in CODE_PATTERNS:
        for m in re.finditer(pat, title + "\n" + body):
            ctx = (title + "\n" + body)[max(0, m.start() - 10):m.end() + 10].replace("\n", " ")
            errors.append(f"[{name}] …{ctx}…")
    ok, issues = check_content(title + body, domain)
    errors += [f"[杂域名后缀] {i}" for i in issues]
    errors += check_negative(title + "\n" + body, domain)
    errors += check_sourcing(title + "\n" + body)
    paras = [p.strip() for p in body.split("\n") if p.strip()]
    if not paras:
        errors.append("[空正文]")
    else:
        if not re.search(r"[。！？…」”]$", paras[-1]):
            errors.append(f"[结尾突兀] …{paras[-1][-25:]}…")
        for p in paras:
            if len(p) < 10:
                warnings.append(f"[过短段落{len(p)}字] …{p[:25]}…")
    if strict:
        if prose_char_count(body) < MIN_NEWS_BODY_CHARS:
            errors.append(f"[正文不足{MIN_NEWS_BODY_CHARS}字：当前{prose_char_count(body)}字]")
        first = paras[0] if paras else ""
        lead_markers = (
            f'{company}官网启用“{domain}”',
            f'{company}官网启用"{domain}"',
            f"{company}官网启用{domain}",
        )
        if company and domain and not any(marker in first for marker in lead_markers):
            errors.append("[首段缺要素：须同时出现企业全称、官网启用和域名]")
        if len(paras) < 4:
            errors.append("[新闻结构不足：至少需要事件、企业资料、行业背景和总结四个自然段]")
        if not ("中文域名" in body or ".网址" in body):
            errors.append("[缺少.网址背景：须解释中文域名/中文官网入口的意义]")
        if not any(term in body for term in ("行业", "数字化", "线上", "互联网")):
            errors.append("[缺少行业数字化段：须把企业所在行业与.网址使用场景联系起来]")
        share = domain_topic_share(body)
        if share < DOMAIN_TOPIC_SHARE_MIN:
            warnings.append(f"[中文域名篇幅不足] 相关段落约占 {share:.0%}，提示词要求约三分之一")
        for term in META_LEAK_TERMS:
            if term in body:
                errors.append(f"[写作元话语泄漏：{term}]")
        names = common.title_name_variants(company, *name_aliases)
        if names:
            missing = common.title_missing_core(title, names)
            if missing == "缺少企业名称":
                # 极端长企业名无法在平台上限内完整放入时，LLM 会用名称前缀；
                # 这是硬上限下的可审计例外，不把一个可识别的标题误判为空企业名。
                prefix = re.split(r"[，,：:。 ]", title, 1)[0]
                if len(prefix) >= 4 and any(name.startswith(prefix) for name in names):
                    warnings.append("[企业名称过长，标题按平台上限保留名称前缀]")
                else:
                    errors.append(f"[标题缺要素：{missing}]")
            elif missing:
                errors.append(f"[标题缺要素：{missing}]")
    if re.search(r"[，、（：；]$", title.strip()):
        errors.append(f"[标题结尾突兀] …{title.strip()[-15:]}…")
    return errors, warnings


def search_company(company: str, profile: str = "") -> str:
    """联网检索企业资料：先用表格资料派生查询，再用公司名兜底。

    表格“企业信息资料”是首要依据：取其中前 40 字实质内容拼一轮
    定向查询（如产品、行业、地区词），命中该企业自身的公开页面；
    再用公司名查介绍与主营业务作补充。调用方汇总时仍以表格为准，
    冲突时丢弃检索侧的矛盾说法。"""
    out = []
    try:
        anysearch = common.resolve_anysearch()
    except FileNotFoundError as e:
        return f"(检索命令不可用: {e})"
    queries = []
    cleaned = re.sub(r"\s+", "", (profile or "").strip())[:40]
    if cleaned:
        queries.append(f"{company} {cleaned}")
    queries += [f"{company} 介绍", f"{company} 主营业务"]
    for q in queries[:3]:
        try:
            r = subprocess.run(anysearch + ["search", q, "--max_results", "4"],
                               capture_output=True, text=True, timeout=60)
            out.append(f"【查询:{q}】\n" + r.stdout)
        except Exception as e:
            out.append(f"(搜索失败:{q}: {e})")
    return "\n\n".join(out)


def _run_llm_stage(command: str, prompt: str, label: str) -> str | None:
    """运行一个 stdin→stdout 的写作阶段；失败不把半成品交给下一阶段。"""
    if not command:
        return None
    try:
        result = subprocess.run(shlex.split(command), input=prompt,
                                capture_output=True, text=True, timeout=300)
    except Exception as exc:
        print(f"   ({label}失败：{exc})")
        return None
    output = result.stdout.strip()
    if result.returncode != 0 or not output:
        detail = result.stderr.strip()[-240:]
        print(f"   ({label}无输出：{detail})")
        return None
    return output


def humanization_commands(llm_cmd: str = "", human_writing_cmd: str = "",
                          humanizer_zh_cmd: str = "") -> tuple[str, str]:
    """解析两阶段命令。专用命令优先，缺少的一段复用已有命令。"""
    fallback = llm_cmd or os.environ.get("MEDIA_LLM_CMD", "")
    writing = human_writing_cmd or os.environ.get("MEDIA_HUMAN_WRITING_CMD", "") or fallback
    humanizer = humanizer_zh_cmd or os.environ.get("MEDIA_HUMANIZER_ZH_CMD", "") or fallback
    # 只配置一个专用命令时也跑完两步，但在输出中明确它被复用，避免误以为跳过了阶段。
    writing = (writing or humanizer).strip()
    humanizer = (humanizer or writing).strip()
    return writing, humanizer


def _human_writing_prompt(company: str, domain: str, material: str, variant: int = 0) -> str:
    rerun = f"（这是第 {variant} 次重做，请换一个切入角度和写法。）\n" if variant else ""
    return (f"你现在执行 human-writing 第一阶段，只负责把材料写实，不做第二阶段清腔。\n"
            f"为“{company}”的新闻稿写正文，围绕官网启用“{domain}”这一已知事实展开。\n"
            f"{rerun}"
            f"新闻稿正文至少写到{MIN_NEWS_BODY_CHARS}字（不含标题）；不足时只能展开已有事实，不能为了凑字数补造内容。\n"
            "采用接近人民网新闻报道的事实优先、克制叙述笔法，但不要声称人民网采访或发稿。不加编号小标题。\n"
            "【篇幅配比】正文按下面三块大致均分，每块占三分之一左右：\n"
            "（一）企业情况与行业位置：先站在企业所在行业的高度写这类企业的线上信息、采购或服务场景，"
            "再带出该企业少量公开事实。企业自身介绍只能占很小一部分，不要用大段文字罗列工商登记项目。\n"
            "（二）中文域名与 `.网址` 后缀本身：这个后缀由中文词语构成，怎么写怎么读、在浏览器地址栏如何输入、"
            "与线下名称写法如何对应，属于识别和访问层面的常识。\n"
            "（三）该企业启用中文域名的价值与意义：结合企业名称、品牌写法和对外传播物料，说明入口识别、"
            "访问路径和资料统一带来的实际作用。\n"
            "开头第一段先交代企业官网启用该中文域名的事实，并解释它对官网识别和访问入口的直接意义，"
            "首段必须连续出现“企业全称官网启用‘域名’”这一事实短语；末段回到已确认事实和待核验边界，简短收束。"
            "没有采访或现场材料时，不写成现场采访口吻，不虚构引语。只谈识别、访问和资料统一，"
            "不承诺流量、排名、销量或法律结果。\n"
            "【来源限制】企业事实的来源统一写成“公开工商登记信息”或“企业公开资料”，不得出现任何网站、"
            "平台、数据库或招聘网站的名称（例如企查查、爱企查、天眼查、启信宝、水滴信用、顺企网、黄页、"
            "BOSS直聘、乡振网、职友集、猎聘、智联招聘、各类百科等）；全文不得出现任何联系方式，"
            "包括电话、座机、手机号、邮箱、微信、QQ，也不要写法定代表人或联系人姓名。\n"
            "【材料优先级】表格提供的企业信息资料为准，联网检索材料只作补充；两者冲突时以表格资料为准，"
            "检索材料里读不到的内容宁可不写。不能编造数字、客户、现场、体验、引语、未来计划或第一人称亲历。"
            "材料不足就缩短，不用重复解释凑字数。每段增加新事实、新动作、新区别或新后果，主语和动作尽早出现，"
            "白话打底，句长有变化。只输出可以直接发布的正文，不要写提纲、标题、来源列表、核验步骤或写作过程，"
            "不要出现“本文、本稿、稿件、检索结果、信息边界、提示词”等自我说明。\n"
            "硬限制：正文只能出现 .网址 后缀，不得出现 .com/.cn 等其他后缀，不得出现“英文域名”或“国际域名”。"
            "禁止翻案腔、三项同构排比、破折号、提示性冒号、汇报黑话和宏大升华。直接输出段落间空行分隔的纯正文。\n"
            f"可参考的 `.网址` 背景资料（用于第二、三块，不得当成企业事实）：\n{url_material_context()}\n"
            f"可用的材料（表格资料优先）：\n{material[:6000]}")


def _humanizer_prompt(company: str, domain: str, draft: str,
                      mode: str = "新闻稿", allow_comparison: bool = False) -> str:
    comparison = ("观点文中原有的比较白名单句（含传统英文后缀的对比）必须原样保留，不能新增或扩大比较。"
                  if allow_comparison else
                  "不得引入 .com/.cn 等其他域名后缀或“英文域名/国际域名”字样。")
    identity = ("这是独立观点文，不能把当前企业或具体域名带入正文，除非它已经出现在初稿中；保留原稿的"
                "中文.网址核心结论。"
                if mode == "观点文" else
                f"必须原样保留的企业与域名：{company} / {domain}")
    preserve = ("保留初稿中的事实、数字和核心结论；不要新增企业名称、具体域名、案例、体验、引语或承诺。"
                if mode == "观点文" else
                "保留所有事实、数字、单位、公司名称、来源归属、用户指定的核心结论和域名原样；不新增任何企业资料、案例、体验、引语、承诺或第一人称经历。")
    length_rule = (f"新闻稿正文修订后仍不得少于{MIN_NEWS_BODY_CHARS}字（不含标题）；只能删掉重复和套话，"
                   "不能把已有事实压缩到门槛以下，也不能用新事实凑长度。保留首段的企业全称、‘官网启用’和域名，"
                   "并保留三块篇幅配比：企业情况与行业位置、中文域名与 `.网址` 后缀本身、该企业启用中文域名的"
                   "价值与意义，每块各占三分之一左右；不要把任何一块压扁，也不要把三段合成一段，"
                   "不要改成广告文或观点文。\n"
                   if mode == "新闻稿" else "")
    source_rule = ("企业事实的来源只能写成“公开工商登记信息”或“企业公开资料”：删掉正文里出现的任何网站、"
                   "平台、数据库或招聘网站名称，也删掉任何联系方式（电话、座机、手机号、邮箱、微信、QQ）"
                   "以及法定代表人或联系人姓名。\n"
                   if mode == "新闻稿" else "")
    return (f"你现在执行 humanizer-zh 第二阶段，只编辑下面已经写好的{mode}初稿。\n"
            f"{preserve}"
            f"{length_rule}"
            f"{source_rule}"
            "删除夸大意义、宣传式形容词、模糊归因、"
            "AI 黑话、名词化、固定连接词、假金句、万能结尾和协作话术，打破同长句与三连排比。把翻案腔"
            f"改成正面陈述，删掉破折号和提示性冒号，保留克制的新闻编辑口吻。{comparison}删掉“本文、本稿、稿件、"
            "检索结果、信息边界、上述资料、上述材料、本次报道、写入本文”等写作过程说明；指代前文时直接用具体名词"
            "（如“企业公开资料”“这些信息”），只输出修订后的"
            "纯正文，不要标题、评分、解释或 Markdown。\n"
            f"{identity}\n"
            f"--- 初稿开始 ---\n{draft}\n--- 初稿结束 ---")


def check_humanized(text: str) -> tuple[bool, str]:
    """运行人化硬规则门禁，返回 (是否通过, 检查器输出)。

    检查器只负责硬错，风格提醒仍需人工判断；输出用于让 humanizer-zh
    带着具体硬错重跑一轮，而不是直接丢掉整行。"""
    if not HUMANIZED_CHECK.exists():
        print("   (找不到 check_humanized.py，跳过人化门禁)")
        return True, ""
    try:
        result = subprocess.run(
            [sys.executable, str(HUMANIZED_CHECK), "-"],
            input=text, capture_output=True, text=True, timeout=30, cwd=str(BASE),
        )
    except Exception as exc:
        print(f"   (人化门禁执行失败：{exc})")
        return False, ""
    output = result.stdout.strip()
    if output:
        print("   " + output.replace("\n", "\n   "))
    if result.returncode != 0:
        print("   ✗ 人化门禁未通过")
        return False, output
    return True, output


def repolish_humanized(company: str, domain: str, body: str, humanizer_cmd: str,
                       detail: str, mode: str = "新闻稿",
                       allow_comparison: bool = False) -> str | None:
    """人化门禁未过时，把硬错清单交回 humanizer-zh 重跑一轮。"""
    if not humanizer_cmd or not detail:
        return None
    hard = [line for line in detail.splitlines() if line.startswith("- ")]
    note = ""
    if hard:
        note = ("\n\n上一稿未通过硬规则检查，请逐条修正后重新输出纯正文：\n"
                + "\n".join(hard[:20])[:1200])
    return _run_llm_stage(
        humanizer_cmd,
        _humanizer_prompt(company, domain, body, mode, allow_comparison) + note,
        "humanizer-zh（按门禁清单重跑）",
    )


def gen_news_llm(company: str, domain: str, search_note: str,
                 writing_cmd: str, humanizer_cmd: str,
                 variant: int = 0, profile: str = "", attempts: int = 2) -> str | None:
    """用两阶段外部 LLM 生成新闻稿正文：先 human-writing，再 humanizer-zh。

    正文完全由 LLM 写，代码里没有任何预设段落；表格提供的企业信息资料优先于联网检索。
    不合格（字数、红线、来源与联系方式、首段缺少启用事实）就带着具体原因重跑，
    仍不合格就跳过该行，不回退到模板。标题由 gen_title_llm 单独拟。"""
    if not writing_cmd and not humanizer_cmd:
        return None
    material = f"已知启用事实：{company}官网启用“{domain}”。\n"
    if profile.strip():
        material += f"表格提供的企业信息资料（以此为准，优先于联网检索）：\n{profile[:4000]}\n"
    material += f"联网检索材料（只作补充，与表格资料冲突时以表格为准）：\n{search_note[:5000]}"
    lead_markers = (f"{company}官网启用“{domain}”", f"{company}官网启用\"{domain}\"",
                    f"{company}官网启用{domain}")
    note = ""
    for attempt in range(1, attempts + 1):
        draft = _run_llm_stage(
            writing_cmd,
            _human_writing_prompt(company, domain, material, variant) + note,
            "human-writing",
        )
        body = None
        if draft:
            body = _run_llm_stage(
                humanizer_cmd,
                _humanizer_prompt(company, domain, draft, "新闻稿") + note,
                "humanizer-zh",
            )
        reasons = []
        if body:
            short = prose_char_count(body) < MIN_NEWS_BODY_CHARS
            ok, issues = check_content(company + domain + body, domain)
            sourcing = check_sourcing(body)
            first_para = body.split("\n\n")[0] if body else ""
            lead_ok = any(marker in first_para for marker in lead_markers)
            if not short and ok and lead_ok and not sourcing:
                return body
            if short:
                reasons.append(
                    f"正文只有{prose_char_count(body)}字，不足{MIN_NEWS_BODY_CHARS}字"
                    "（只能展开已有事实，不得编造）")
            if not ok:
                reasons.append(f"出现不合规写法：{issues[:2]}（只能使用 .网址 后缀）")
            if sourcing:
                reasons.append("；".join(sourcing[:2]))
            if not lead_ok:
                reasons.append(f"首段没有出现“{company}官网启用‘{domain}’”这一事实短语")
        else:
            reasons.append("上一轮没有产出可用正文")
        print(f"   (两阶段稿件不合格：{reasons[0]})")
        if attempt < attempts:
            note = ("\n\n上一稿未通过检查，请本次修正后重新输出纯正文："
                    + "；".join(reasons) + "。")
            print(f"   (带修正说明重跑两阶段 {attempt}/{attempts - 1})")
    print("   (两阶段稿件仍未合格，跳过该行)")
    return None


# ── 第二篇观点文：“企业注册.网址有没有用” ──
# 只给提示词：必须覆盖的结论写成写作要求，句子全部由 LLM 自己写，正文里没有任何预设文本。
OPINION_POINTS = (
    "中文域名的注册与保护已经是企业绕不开的一项工作",
    "主流浏览器都已经支持中文域名",
    "中文域名的使用场景与传统英文后缀相比没有实际差别",
    "中文域名可以直接作为企业官网的主域名使用",
    "在 AI 问答场景里，中文域名更容易让平台把品牌词和官网对应起来，因此更建议使用",
)


def _opinion_writing_prompt(background: str, variant: int = 0) -> str:
    points = "\n".join(f"{i}. {point}" for i, point in enumerate(OPINION_POINTS, 1))
    rerun = f"（这是第 {variant} 次重做，请换一个切入角度和写法。）\n" if variant else ""
    return (f"你现在执行 human-writing 第一阶段，写一篇独立的中文“企业注册.网址有没有用”科普文章，"
            f"不与任何具体企业、品牌或客户案例挂钩。\n"
            f"{rerun}"
            f"正文不少于{MIN_OPINION_BODY_CHARS}字。文章必须覆盖下面这些结论，但句子、结构、例子和用词"
            f"全部由你自己写，不要照搬任何现成表述：\n{points}\n"
            "每段推进一件新事实或新区别，主语和动作尽早出现，白话打底，句长有变化；"
            "没有来源就不写数字、引语、客户案例、现场细节和个人经历，也不要承诺流量、排名、销量或法律结果。\n"
            "正文只能出现 .网址 后缀，不得出现 .com/.cn 等其他后缀，也不得出现“英文域名”“国际域名”字样，"
            "与传统后缀做对比时用“传统英文后缀”这类说法。\n"
            "禁止翻案腔、三项同构排比、破折号、提示性冒号、汇报黑话和宏大升华；"
            "不要出现“本文、本稿、稿件、检索结果、提示词”等自我说明。\n"
            "只输出段落间空行分隔的纯正文，不要标题、提纲或来源列表。\n"
            f"可参考的 .网址 背景资料（只是背景，不是要照抄的文字）：\n{background[:5000]}")


def gen_opinion_llm(company: str, domain: str, writing_cmd: str, humanizer_cmd: str,
                    background: str, variant: int = 0, attempts: int = 2) -> str | None:
    """观点文正文也全部由 LLM 写：human-writing 出稿 → humanizer-zh 清腔，没有预设文案。"""
    if not writing_cmd and not humanizer_cmd:
        return None
    note = ""
    for attempt in range(1, attempts + 1):
        draft = _run_llm_stage(writing_cmd, _opinion_writing_prompt(background, variant) + note,
                               "human-writing（观点文）")
        body = None
        if draft:
            body = _run_llm_stage(humanizer_cmd,
                                  _humanizer_prompt(company, domain, draft, "观点文",
                                                    allow_comparison=True) + note,
                                  "humanizer-zh（观点文）")
        reasons = []
        if body:
            short = prose_char_count(body) < MIN_OPINION_BODY_CHARS
            ok, issues = check_content(mask_comparatives(body), "")
            decoupled = company not in body and domain not in body
            if not short and ok and decoupled:
                return body
            if short:
                reasons.append(f"正文只有{prose_char_count(body)}字，不足{MIN_OPINION_BODY_CHARS}字")
            if not ok:
                reasons.append(f"出现不合规写法：{issues[:2]}（只能使用 .网址 后缀）")
            if not decoupled:
                reasons.append("正文里出现了具体企业名或具体域名，观点文要与具体企业脱钩")
        else:
            reasons.append("上一轮没有产出可用正文")
        print(f"   (观点文不合格：{reasons[0]})")
        if attempt < attempts:
            note = ("\n\n上一稿未通过检查，请本次修正后重新输出纯正文："
                    + "；".join(reasons) + "。")
            print(f"   (带修正说明重跑观点文 {attempt}/{attempts - 1})")
    print("   (观点文多次未过，跳过本篇)")
    return None


def _title_from_body(body: str, limit: int) -> str:
    """LLM 拟题连续失败时，从正文首句里取一个可用的观点文标题（不是预设文案）。"""
    first = re.split(r"(?<=[。！？?!])", (body or "").strip(), maxsplit=1)[0].strip()
    title = _clean_title(first, limit)
    return title if title and len(title) <= limit else ""


# LLM 常见的非标题包装：Markdown 标题、序号，以及“标题：/这是为您拟的标题：”这类引导语
TITLE_LEAD_PAT = re.compile(r"^\s*(?:#{1,6}\s*|\d{1,2}\s*[.、)）:：]|[-*•]\s+)")
TITLE_META_PAT = re.compile(
    r"^\s*(?:[^：:\n]{0,20}(?:标题内容|标题|题目|以下|这是|为您|如下|拟|结果|回答|答案)"
    r"[^：:\n]{0,12}\s*[:：]\s*)")


def _clean_title(raw: str, limit: int = 0) -> str:
    """从 LLM 回复里取出标题：只取第一行内容，剥掉引导语、序号、Markdown 和引号。

    模型经常回“标题：xxx”“## xxx”“1. xxx”“这是为您拟的标题：xxx”，或者先写一句
    解释再给标题；不处理就会把包装词一起当成标题发布。超长且带冒号时取冒号后的部分。"""
    line = next((l.strip() for l in (raw or "").splitlines() if l.strip()), "")
    if limit and len(line) > limit and "：" in line:
        tail = line.rsplit("：", 1)[-1].strip()
        if tail:
            line = tail
    for _ in range(2):
        line = TITLE_LEAD_PAT.sub("", line)
        line = TITLE_META_PAT.sub("", line)
    line = line.strip().strip("“”\"'《》【】[]（） ")
    return line.rstrip("。． ,，;；:：、").strip()


def gen_title_llm(company: str, short_name: str, domain: str, body: str, limit: int, cmd: str,
                  used_titles=(), no_company: bool = False, attempts: int = 3) -> str | None:
    """标题由 LLM 读正文拟写，是标题的唯一正常来源。

    A 文只硬要求一件事：标题里出现企业名称或简称（全称/简称/缩写、企业名里的自然简称皆可），
    且不能只是名称本身；完整域名、“.网址”和“启用”都不要求进标题。
    B 文（no_company）依据正文自由发挥、不得含具体企业名。只卡字数上限与重名规避；
    连续不合格就返回 None，由 main 报警并跳过该行（不套用模板标题）。"""
    if not cmd:
        return None
    used_set = {t for t in used_titles if t}
    used = [t for t in list(used_titles)[:8] if t]
    avoid = f"\n\n以下标题已被使用，请避开重名：{'、'.join(used)}" if used else ""

    if no_company:
        prompt = (f"为下面这篇科普文章拟一个标题，读正文后自己命题；"
                  f"标题围绕中文网址/中文域名主题，不要出现具体企业、品牌或产品名称；"
                  f"要像正经媒体发的标题，主谓清楚、信息具体，不要堆名词、不要罗列标签；"
                  f"严格不超过{limit}个字；只输出标题本身这一行，不要引号、序号、前缀或任何说明。"
                  f"\n\n{body[:4000]}{avoid}")
        note = ""
        for attempt in range(1, attempts + 1):
            out = _run_llm_stage(cmd, prompt + note, "标题拟写")
            title = _clean_title(out, limit) if out else ""
            if title and len(title) <= limit and company not in title and title not in used_set:
                return title
            print(f"   (LLM标题第{attempt}次不合格：{title[:40] if title else '(空)'}，不得含企业名且≤{limit}字)")
            note = (f"\n\n注意：上一条回复不合格。只输出一行标题，不得出现企业名称“{company}”，"
                    f"总长不超过{limit}个字。")
        return None

    names = common.title_name_variants(company, short_name)
    name_note = "、".join(names) if names else company
    prompt = (f"为下面这篇企业新闻拟一个标题。\n"
              f"1. 读正文，抓住这篇文章真正在讲的那件事，用一句话把它写成标题；\n"
              f"2. 要像正经媒体发的新闻标题：主谓清楚、信息具体、读起来是一句通顺的话；"
              f"不要堆名词、不要用逗号罗列标签、不要写成口号或问句堆砌；\n"
              f"3. 硬要求只有一条：标题里出现企业名称或它的简称，可用：{name_note}。"
              f"完整域名、“.网址”和“启用”都不要求写进标题；\n"
              f"4. 句式、用词、切入角度完全自由，不要套固定格式，同一批标题之间不要雷同；\n"
              f"5. 总长严格不超过{limit}个字。只输出标题这一行，不要引号、序号、前缀或任何说明。"
              f"\n\n{body[:4000]}{avoid}")
    note = ""
    for attempt in range(1, attempts + 1):
        out = _run_llm_stage(cmd, prompt + note, "标题拟写")
        title = _clean_title(out, limit) if out else ""
        missing = common.title_missing_core(title, names, limit)
        if not missing and title not in used_set:
            return title
        if not missing:
            missing = "与已发布标题重复"
        print(f"   (LLM标题第{attempt}次不合格：{title[:40] if title else '(空)'}，{missing})")
        note = (f"\n\n注意：上一条回复不合格（{missing}）。只输出一行标题，必须出现企业名称或简称"
                f"（{name_note}），不能只写企业名称，总长不超过{limit}个字；"
                f"句式角度自由拟写即可。")
    return None






# ── GEO 问答页：同样由 LLM 写，只给格式和边界要求 ──

def _qa_prompt(company: str, domain: str, body: str) -> str:
    return (f"为“{company}”的官网中文域名“{domain}”写 {QA_PAGES} 组网页 FAQ 问答，每组 3 个问答。\n"
            "每组围绕中文域名本身：它是什么、怎么访问、和传统英文后缀在输入与识别上的差别、"
            "适合用在哪些对外场景、企业为什么要注册、访问异常怎么办。三组之间不要重复同一个角度。\n"
            "答案写成两三句可核查的白话，只讲识别、输入和使用层面的价值；不得承诺流量、排名、销量、"
            "防伪效果或法律结果，不得使用“数字身份证”“零风险”这类宣传说法。\n"
            "只能出现 .网址 后缀，不得出现 .com/.cn 等其他后缀，不得出现“英文域名”“国际域名”字样。\n"
            "严格按下面格式输出，不要任何解释或多余文字：\n"
            "【页1】\n问：……\n答：……\n问：……\n答：……\n问：……\n答：……\n"
            "【页2】\n（同样三组问答）\n【页3】\n（同样三组问答）\n\n"
            f"已生成的新闻稿（仅供了解企业情况，不要照抄）：\n{body[:2000]}")


def _parse_qa_pages(raw: str) -> list[list[dict]]:
    """解析 LLM 的 FAQ 输出；格式不对就返回已解析到的页，由调用方决定重试。"""
    pages = []
    for chunk in re.split(r"【\s*页\s*\d+\s*】", raw or "")[1:]:
        pairs = re.findall(r"问\s*[:：]\s*(.+?)\s*\n\s*答\s*[:：]\s*(.+?)(?=\n\s*问\s*[:：]|\Z)",
                           chunk, re.S)
        qas = [{"q": q.strip(), "a": re.sub(r"\s+", "", a)} for q, a in pairs if q.strip() and a.strip()]
        if qas:
            pages.append(qas)
    return pages


def gen_qa_llm(company: str, domain: str, body: str, cmd: str,
               pages: int = QA_PAGES, attempts: int = 2) -> list[list[dict]] | None:
    """用 LLM 生成 GEO FAQ 问答；拿不到可用结构就放弃 QA 页，不回退到预设文案。"""
    if not cmd:
        return None
    note = ""
    for attempt in range(1, attempts + 1):
        out = _run_llm_stage(cmd, _qa_prompt(company, domain, body) + note, "QA 问答")
        parsed = [page for page in _parse_qa_pages(out or "") if len(page) >= 2]
        if len(parsed) >= pages:
            if all(check_content("".join(x["q"] + x["a"] for x in page), domain)[0]
                   and not check_sourcing("".join(x["q"] + x["a"] for x in page))
                   for page in parsed[:pages]):
                return parsed[:pages]
            print("   (QA 问答红线未过，重写一次)")
        else:
            print(f"   (QA 问答格式不合格：解析到 {len(parsed)} 页，需要 {pages} 页)")
        if attempt < attempts:
            note = (f"\n\n注意：上一条回复不合格。只输出【页1】到【页{QA_PAGES}】共 {QA_PAGES} 段，"
                    "每段 3 组“问：/答：”，不要解释或多余文字。")
    return None


HTML_TMPL = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__TITLE__</title>
<meta name="description" content="__DESC__">
<meta name="keywords" content="__KEYWORDS__">
<script type="application/ld+json">
__JSONLD__
</script>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: "Songti SC", "SimSun", serif; color: #222; background: #fff; line-height: 1.9; }
header { background: #c00; color: #fff; padding: 14px 20px; }
header .site { font-size: 22px; font-weight: bold; letter-spacing: 2px; }
header .sub { font-size: 12px; opacity: .85; margin-top: 2px; }
main { max-width: 760px; margin: 0 auto; padding: 30px 20px 60px; }
h1 { font-size: 26px; line-height: 1.5; color: #111; margin-bottom: 14px; }
.meta { color: #888; font-size: 13px; border-bottom: 1px solid #eee; padding-bottom: 12px; margin-bottom: 22px; }
.meta .src { color: #c00; font-weight: bold; }
article p { text-indent: 2em; margin-bottom: 18px; font-size: 16.5px; }
h2 { font-size: 20px; color: #c00; margin: 26px 0 14px; padding-left: 10px; border-left: 4px solid #c00; }
.qa { background: #fafafa; border: 1px solid #eee; border-radius: 6px; padding: 18px 20px; margin-bottom: 16px; }
.qa h2 { margin-top: 0; }
.qa .a { font-size: 16px; }
footer { text-align: center; color: #999; font-size: 12px; padding: 20px; border-top: 1px solid #eee; }
.crumbs { font-size: 13px; color: #888; margin-bottom: 18px; }
.crumbs a { color: #c00; text-decoration: none; }
</style>
</head>
<body>
<header><div class="site">企业品牌资讯</div><div class="sub">品牌·域名·数字化观察</div></header>
<main>
<div class="crumbs"><a href="./__INDEX__">首页</a> &gt; 企业动态 &gt; 正文</div>
__BODY__
</main>
<footer>本页由 __DOMAIN__ 官方发布 · 内容仅供参考</footer>
</body>
</html>"""


def render(html_title: str, desc: str, keywords: str, jsonld: str, body: str, domain: str) -> str:
    return (HTML_TMPL
            .replace("__TITLE__", html_title)
            .replace("__DESC__", desc)
            .replace("__KEYWORDS__", keywords)
            .replace("__JSONLD__", jsonld)
            .replace("__BODY__", body)
            .replace("__DOMAIN__", domain))


def news_html(domain: str, title: str, paras: list[str], company: str) -> str:
    jsonld = json.dumps({
        "@context": "https://schema.org", "@type": "NewsArticle",
        "headline": title, "inLanguage": "zh-CN",
        "author": {"@type": "Organization", "name": company},
        "publisher": {"@type": "Organization", "name": company},
        "datePublished": datetime.now().strftime("%Y-%m-%d"),
        "about": {"@type": "Thing", "name": domain},
    }, ensure_ascii=False, indent=2)
    body = (f"<h1>{title}</h1>\n"
            f"<div class='meta'><span class='src'>来源：{company}</span> · {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>\n"
            f"<article>" + "".join(f"<p>{p}</p>" for p in paras) + "</article>")
    return render(title, paras[0][:150], f"{domain},{company},.网址", jsonld, body, domain)


def qa_html(domain: str, company: str, qas: list[dict], n: int) -> tuple[str, str]:
    """返回 (文件名, html)。QA 文件名只标识企业官网，不套用新闻标题模板。"""
    base = f"{company}官网"
    fname = f"{base}-QA{n}"
    jsonld = json.dumps({
        "@context": "https://schema.org", "@type": "FAQPage",
        "mainEntity": [{"@type": "Question", "name": qa["q"],
                        "acceptedAnswer": {"@type": "Answer", "text": qa["a"]}} for qa in qas],
    }, ensure_ascii=False, indent=2)
    body = (f"<h1>{base} 常见问题（{n}）</h1>\n"
            f"<div class='meta'><span class='src'>来源：{company}</span> · {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>")
    for i, qa in enumerate(qas, 1):
        body += (f"\n<section class='qa'><h2>{i}. {qa['q']}</h2>"
                 f"<div class='a'><p style='text-indent:0'>{qa['a']}</p></div></section>")
    return f"{fname}.html", render(f"{base} 常见问题（{n}）", qas[0]["a"][:150],
                                   f"{domain},{company},.网址,常见问题", jsonld, body, domain)


def publish_one(media: str, code: str, title: str, body_file: Path, dryrun: bool) -> dict:
    cmd = [sys.executable, str(MEDIA_PUBLISH[media]), "--title", title,
           "--body-file", str(body_file), "--code", code]
    if not dryrun:
        cmd.append("--go")
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=900, cwd=str(BASE))
    ok = r.returncode == 0 and ("发布提交成功" in r.stdout or "发布提交：" in r.stdout
                               or "DRYRUN" in r.stdout)
    tail = (r.stdout[-500:] + r.stderr[-200:]).replace("\n", " | ")
    return {"ok": ok, "output": tail}


def update_xlink(xlsx_path: Path, row_num: int, col_name: str, value: str) -> None:
    """按列名（媒体编码 a-1/b-1/c-1/c-2）写入对应列；列不存在时在表尾创建。"""
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.worksheets[0]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    if col_name in headers:
        col = headers.index(col_name) + 1
    else:
        col = ws.max_column + 1
        ws.cell(row=1, column=col, value=col_name)
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == row_num:
            ws.cell(row=r, column=col, value=value)
            break
    wb.save(xlsx_path)


def publish_and_record(title: str, body_file: Path, codes: list[str], xlsx: Path,
                       row_num: int, col_suffix: str, dryrun: bool) -> None:
    """逐媒体发布→回查→回写。col_suffix '' 写编码列，'-有用' 写“编码-有用”列。"""
    for code in codes:
        media = MEDIA_OF.get(code)
        col = f"{code}{col_suffix}"
        if not media:
            print(f"   未知编码 {code}，跳过")
            continue
        print(f"⑤ 发布到 {media}（{code}）…")
        res = publish_one(media, code, title, body_file, dryrun)
        print("   " + ("✓ " if res["ok"] else "✗ ") + res["output"][-180:])
        if not res["ok"]:
            update_xlink(xlsx, row_num, col, "发布失败")
            continue
        link = ""
        if not dryrun:
            subprocess.run([sys.executable, str(BASE / "scripts" / "collect_links.py"), "--code", code,
                            "--pending", str(PENDING), "--links", str(LINKS_CSV)],
                           capture_output=True, text=True, timeout=600, cwd=str(BASE))
            if LINKS_CSV.exists():
                for row in csv.DictReader(open(LINKS_CSV, encoding="utf-8-sig")):
                    # 必须同编码同标题（跨列同标题会串，必须卡 code）
                    if row["code"] == code and row["title"] == title and row["link"]:
                        link = row["link"]
                        break
        if link:
            update_xlink(xlsx, row_num, col, link)
            print(f"   链接已回写({col}): {link}")
        else:
            update_xlink(xlsx, row_num, col, "(审核中，稍后回查)")
            print("   写入占位: (审核中，稍后回查) —— 稍后重跑 collect_links.py 会覆盖")


# 表格里“企业信息资料”列的表头关键词：这张列的内容是写作的首要依据，优先于联网检索
PROFILE_HEADER_KEYS = ("企业信息", "企业资料", "公司资料", "公司介绍", "企业介绍", "简介", "经营范围")


def read_rows(xlsx_path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.worksheets[0]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    # 媒体编码列：表头是"媒体"的列（可能有多列，每列一个编码），或表头直接是编码（a-1/b-1/c-1/c-2）
    # 注意：同名“媒体”列有多列，必须按列下标读（headers.index 只返回第一列）
    media_idx = [i for i, h in enumerate(headers)
                 if h in MEDIA_OF or (h and str(h).strip() == "媒体")]
    # 可选“简称”列：标题用简称（如 中国长江三峡集团），正文仍用全称
    short_col = next((i + 1 for i, h in enumerate(headers)
                      if h and "简称" in str(h)), None)
    # 可选“企业信息资料”列：写作用这份表格资料为准，联网检索只作补充
    profile_col = next((i + 1 for i, h in enumerate(headers)
                        if h and any(k in str(h) for k in PROFILE_HEADER_KEYS)), None)
    rows = []
    for r in range(2, ws.max_row + 1):
        num = ws.cell(row=r, column=1).value
        if num is None:
            continue
        codes = []
        for i in media_idx:
            v = ws.cell(row=r, column=i + 1).value
            if v:
                # 一格可能含多个编码（顿号/逗号/分号分隔）
                for c in re.split(r"[、，,;；\s]+", str(v).strip()):
                    c = c.strip()
                    if c:
                        codes.append(c)
        codes = [c for c in codes if c in MEDIA_OF]
        codes = list(dict.fromkeys(codes))  # 去重（多列重复的 a-1 只发一次）
        short = ""
        if short_col:
            v = ws.cell(row=r, column=short_col).value
            short = str(v).strip() if v else ""
        profile = ""
        if profile_col:
            v = ws.cell(row=r, column=profile_col).value
            profile = str(v).strip() if v else ""
        rows.append({"row": r, "num": num, "domain": str(ws.cell(row=r, column=2).value).strip(),
                     "company": str(ws.cell(row=r, column=3).value).strip(),
                     "short": short, "profile": profile, "codes": codes})
    return rows


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--row", type=int, help="只处理指定编号")
    ap.add_argument("--dryrun", action="store_true")
    ap.add_argument("--no-publish", action="store_true")
    ap.add_argument("--llm-cmd", default="", help="新闻稿 LLM 命令（默认 $MEDIA_LLM_CMD，为空用内置模板）")
    ap.add_argument("--variant", type=int, default=0, help="变体种子：换一套标题/角度/引用/事实句，用于重做不同内容")
    ap.add_argument("--fresh", action="store_true", help="清空输出目录后重做（删掉该域名旧文件）")
    ap.add_argument("--human-writing-cmd", default="",
                    help="human-writing 第一阶段命令（默认 $MEDIA_HUMAN_WRITING_CMD）")
    ap.add_argument("--humanizer-zh-cmd", default="",
                    help="humanizer-zh 第二阶段命令（默认 $MEDIA_HUMANIZER_ZH_CMD）")
    a = ap.parse_args()

    writing_cmd, humanizer_cmd = humanization_commands(
        a.llm_cmd, a.human_writing_cmd, a.humanizer_zh_cmd,
    )
    humanization_enabled = bool(writing_cmd or humanizer_cmd)
    if not humanization_enabled:
        print("✗ 报警：未配置 LLM 写作命令。稿件全部由 LLM 写，没有兜底稿可写，本次不生成、不发布。")
        print("  请设置 MEDIA_HUMAN_WRITING_CMD + MEDIA_HUMANIZER_ZH_CMD（或 MEDIA_LLM_CMD），")
        print("  或用 --llm-cmd / --human-writing-cmd / --humanizer-zh-cmd 传入命令。")
        return 1
    print("已启用两阶段稿件人化：human-writing → humanizer-zh")

    xlsx = Path(a.xlsx)
    global OUTPUT_ROOT
    OUTPUT_ROOT = xlsx.parent  # 生成内容放表格所在文件夹
    rows = read_rows(xlsx)
    if a.row:
        rows = [r for r in rows if int(r["num"]) == a.row]
    print(f"待处理 {len(rows)} 行")
    ok_rows = 0

    for i, r in enumerate(rows):
        domain, company = r["domain"], r["company"]
        out_dir = OUTPUT_ROOT / domain  # 文件夹名=域名原样
        if a.fresh and out_dir.exists():
            import shutil as _sh
            _sh.rmtree(out_dir)
        print(f"\n═══ [{r['num']}] {domain} / {company} → {'+'.join(r['codes'])} ═══")

        # ① 检索
        print("① anysearch 检索企业资料（表格资料派生查询 + 公司名补充）…")
        note = clean_text(search_company(company, r.get("profile", "")))  # 洗搜索结果的 HTML/实体

        # ② 新闻稿（正文与标题都由 LLM 写，代码里没有预设文案）
        print("② 生成新闻稿…")
        body = gen_news_llm(company, domain, note, writing_cmd, humanizer_cmd, a.variant,
                            r.get("profile", ""))
        if not body:
            print("   ✗ 报警：LLM 未产出合格正文，跳过该行（不生成任何文件）")
            continue
        print("   (human-writing → humanizer-zh)")
        body = clean_text(body, domain)
        # 标题：LLM 读正文拟写（唯一正常来源，硬要求只有含企业名/简称）；无 LLM 或连续不合格才用兜底
        limits = [common.TITLE_LIMIT.get(MEDIA_OF[c], common.DEFAULT_TITLE_LIMIT)
                  for c in r["codes"] if c in MEDIA_OF]
        limit = min(limits) if limits else common.DEFAULT_TITLE_LIMIT
        # 已发布标题库：重做时自动错开，避免新旧重名串链接
        used_titles = set()
        if LINKS_CSV.exists():
            for _row in csv.DictReader(open(LINKS_CSV, encoding="utf-8-sig")):
                if _row["link"]:
                    used_titles.add(_row["title"])
        if PENDING.exists():
            for _x in json.loads(PENDING.read_text(encoding="utf-8")):
                used_titles.add(_x["title"])
        title = gen_title_llm(company, r.get("short", ""), domain, body, limit,
                              writing_cmd or humanizer_cmd, used_titles)
        if not title:
            print("   ✗ 报警：LLM 没能给出合格标题，跳过该行（不套用模板标题）")
            continue
        ok, issues = check_content(title, domain)
        if not ok:
            print(f"   ✗ 标题红线未过，跳过: {issues[:2]}")
            continue
        print(f"   标题({len(title)}字≤{limit}): {title}")
        hok, hdetail = check_humanized(title + "\n\n" + body)
        if not hok:
            print("   ⚠ 人化门禁未过，带硬错清单重跑一次 humanizer-zh")
            fixed = repolish_humanized(company, domain, body, humanizer_cmd, hdetail)
            if fixed:
                fixed = clean_text(fixed, domain)
                hok, _ = check_humanized(title + "\n\n" + fixed)
                if hok and domain in fixed and prose_char_count(fixed) >= MIN_NEWS_BODY_CHARS:
                    body = fixed
                else:
                    hok = False
        if not hok:
            print("   ✗ 标题或正文未通过人化门禁，跳过")
            continue
        ok, issues = check_content(title + body, domain)
        if not ok:
            print("   ⚠ 红线检查未过，自动修正：", issues[:3])
            for pat in [r"[\w-]+\.com\b", r"[\w-]+\.cn\b", r"[\w-]+\.net\b", r"[\w-]+\.org\b"]:
                body = re.sub(pat, domain, body)
            ok, issues = check_content(title + body, domain)
            if not ok:
                print("   ✗ 仍不合规，跳过", issues[:3])
                continue
        # ③ 稿件审核：代码字符/烂尾/杂后缀/负面→硬错跳过；通顺可疑→告警
        print("③ 稿件审核…")
        errors, warnings = review_article(
            title, body, domain, company=company,
            name_aliases=(r.get("short", ""),),
        )
        if errors:
            print(f"   ✗ 审核未过，跳过: {errors[:4]}")
            continue
        for w in warnings[:5]:
            print(f"   ⚠ 审核告警{w}")
        print("   审核通过 ✓")
        out_dir.mkdir(parents=True, exist_ok=True)
        ok_rows += 1
        md_file = out_dir / f"{common.safe_filename(title)}.md"
        md_file.write_text(f"{title}\n\n{body}", encoding="utf-8")
        print(f"   {md_file.name}")

        # ③ GEO 网页
        print("④ 生成人民网风格 GEO 网页…")
        (out_dir / f"{common.safe_filename(title)}.html").write_text(
            news_html(domain, title, body.split("\n\n"), company), encoding="utf-8")
        qa_pages = gen_qa_llm(company, domain, body, writing_cmd or humanizer_cmd)
        if qa_pages:
            for n, qas in enumerate(qa_pages, 1):
                fname, html = qa_html(domain, company, qas, n)
                (out_dir / fname).write_text(html, encoding="utf-8")
            print(f"   1 新闻页 + {len(qa_pages)} QA 页 → {out_dir}")
        else:
            print("   ⚠ QA 页未生成（LLM 未返回可用问答），仅保留新闻页")

        # 第二篇：观点文（同样只给提示词，正文由 LLM 写；链接记入“编码-有用”列）
        print("⑥ 生成观点文（企业注册.网址有没有用）…")
        b_ok = True
        btitle = ""
        bbody = gen_opinion_llm(company, domain, writing_cmd, humanizer_cmd,
                                url_material_context(), a.variant)
        if not bbody:
            print("   ✗ 观点文未生成（LLM 不可用或稿件不合格），跳过本篇")
            b_ok = False
        else:
            bbody = clean_text(bbody, domain)
            btitle = gen_title_llm(company, "", domain, bbody, limit,
                                   writing_cmd or humanizer_cmd, used_titles,
                                   no_company=True) or _title_from_body(bbody, limit)
            if not btitle:
                print("   ✗ 观点文标题未生成，跳过本篇")
                b_ok = False
        if b_ok:
            hok, hdetail = check_humanized(btitle + "\n\n" + bbody)
            if not hok:
                print("   ⚠ 观点文人化门禁未过，带硬错清单重跑一次 humanizer-zh")
                fixed = repolish_humanized(company, domain, bbody, humanizer_cmd, hdetail,
                                           mode="观点文", allow_comparison=True)
                if fixed:
                    fixed = clean_text(fixed, domain)
                    hok, _ = check_humanized(btitle + "\n\n" + fixed)
                    if hok and ".网址" in fixed:
                        bbody = fixed
                    else:
                        hok = False
            if not hok:
                print("   ✗ 观点文未通过人化门禁，跳过本篇")
                b_ok = False
        if b_ok:
            bmasked = mask_comparatives(btitle + bbody)  # 比较原话白名单，其余按红线
            ok, issues = check_content(bmasked, domain)
            if not ok:
                print("   ⚠ 观点文红线未过，跳过本篇：", issues[:3])
                b_ok = False
        if b_ok:
            print("⑥ 观点文审核…")
            errors, warnings = review_article(btitle, mask_comparatives(bbody),
                                              domain, strict=False)
            if errors:
                print(f"   ✗ 观点文审核未过，跳过本篇: {errors[:4]}")
                b_ok = False
            for w in warnings[:5]:
                print(f"   ⚠ 审核告警{w}")
        if b_ok:
            print(f"   观点文标题({len(btitle)}字≤{limit}): {btitle}")
            (out_dir / f"{common.safe_filename(btitle)}.md").write_text(f"{btitle}\n\n{bbody}", encoding="utf-8")
            (out_dir / f"{common.safe_filename(btitle)}.html").write_text(
                news_html(domain, btitle, bbody.split("\n\n"), company), encoding="utf-8")
            print(f"   观点文 + GEO新闻页 → {out_dir}")

        if a.no_publish:
            print("⑤ 跳过发布（--no-publish）")
            continue

        # ⑤ 逐媒体发布（A 文记编码列，B 文记“编码-有用”列）
        body_file = out_dir / "_publish_body.txt"
        body_file.write_text(body, encoding="utf-8")
        publish_and_record(title, body_file, r["codes"], xlsx, r["num"], "", a.dryrun)
        if b_ok:
            bbody_file = out_dir / "_publish_body_有用.txt"
            bbody_file.write_text(bbody, encoding="utf-8")
            publish_and_record(btitle, bbody_file, r["codes"], xlsx, r["num"], "-有用", a.dryrun)

    if not ok_rows:
        print("\n✗ 报警：本次没产出任何稿件（LLM 不可用或全部不合格），未写入、未发布。")
        return 1
    print(f"\n全部完成，共生成 {ok_rows} 行。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
