#!/usr/bin/env python3
"""检查 human-writing -> humanizer-zh 后的中文稿硬规则。

这是 media-publish-skill 的轻量门禁，不自动改文。风格统计只作提醒，硬错误返回 1。
规则来源：KKKKhazix/human-writing 1.1.0 与 op7418/Humanizer-zh。
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path


HARD_TERMS = (
    "说白了", "说穿了", "先说结论", "赋能", "抓手", "商业闭环", "价值闭环",
    "能力沉淀", "拉通", "底层逻辑", "顶层设计", "认知跃迁", "价值释放",
    "能力建设", "降本增效", "内容矩阵", "全链路", "组合拳", "打开想象空间",
    "结构性机会", "关键命题", "深层逻辑", "技术底座", "公共底座", "技术主权",
    "单点风险", "主脊柱", "材料锚点", "认知增量", "迭代闭环",
)
ROAD_SIGNS = (
    "更微妙的是", "还有一层", "只说对了一半", "值得注意的是",
    "需要指出的是", "从某种意义上说", "希望这对您有帮助",
)
META_LEAK_TERMS = (
    "本文", "本稿", "稿件", "提示词", "提示内容", "检索结果", "检索页面", "公开检索",
    "信息边界", "写入本文", "本次报道", "上述资料", "上述材料", "作为企业情况的来源",
)
PIVOT_PATTERNS = (
    re.compile(r"(?:并)?不是[^。！？\n]{0,90}而是"),
    re.compile(r"不仅(?:仅仅)?是[^。！？\n]{0,90}而是"),
    re.compile(r"并非[^。！？\n]{0,90}而是"),
    re.compile(r"不在于[^。！？\n]{0,90}而在于"),
    re.compile(r"与其说[^。！？\n]{0,90}(?:不如|毋宁|倒不如)"),
    re.compile(r"表面(?:上)?[^。！？\n]{0,90}(?:其实|实际|实则)"),
    re.compile(r"看似[^。！？\n]{0,90}(?:其实|实际|实则)"),
    re.compile(r"回头(?:看|一看)?才(?:发现|明白|知道)"),
    re.compile(r"答案(?:是否定的|恰恰相反)|恰恰相反"),
    re.compile(r"[^，。！？\n]{1,12}不重要，(?:重要|要紧)的是"),
)
NOMINALIZATION_PATTERNS = (
    re.compile(r"进行(?:了|一次|一场|着)?[^。，！？\n]{0,12}(?:调整|优化|升级|分析|讨论|沟通|梳理|复盘|迭代|探索|尝试|思考|规划|布局)"),
    re.compile(r"实现了?[^。，！？\n]{0,16}的?[^。，！？\n]{0,8}(?:提升|增长|突破|转变|跃升|落地)"),
    re.compile(r"完成了?对[^。，！？\n]{0,18}的"),
    re.compile(r"起到了?[^。，！？\n]{0,14}的?作用"),
)
CONJUNCTIONS = ("因为", "所以", "但是", "然而", "同时", "此外", "而且", "并且", "因此", "不仅")


def read_text(value: str) -> str:
    if value == "-":
        return sys.stdin.read()
    return Path(value).read_text(encoding="utf-8")


def mask_non_prose(text: str) -> str:
    """屏蔽代码、链接和 HTML，避免机器字段污染风格门禁。"""
    patterns = (
        re.compile(r"```.*?```", re.DOTALL),
        re.compile(r"`[^`\n]*`"),
        re.compile(r"\]\([^\n)]*\)"),
        re.compile(r"https?://[^\s)>]+"),
        re.compile(r"<[^>\n]+>"),
    )
    for pattern in patterns:
        text = pattern.sub(lambda m: re.sub(r"[^\n]", " ", m.group()), text)
    return text


def matches(text: str, terms: tuple[str, ...]) -> list[tuple[int, str]]:
    result = []
    for term in terms:
        result.extend((m.start(), term) for m in re.finditer(re.escape(term), text))
    return sorted(result)


def main() -> int:
    parser = argparse.ArgumentParser(description="检查 human-writing 与 humanizer-zh 硬规则")
    parser.add_argument("path", help="Markdown/文本文件；使用 - 从标准输入读取")
    args = parser.parse_args()
    try:
        raw = read_text(args.path)
    except (OSError, UnicodeError) as exc:
        print(f"无法读取稿件：{exc}", file=sys.stderr)
        return 2
    text = mask_non_prose(raw)
    if not re.search(r"[\u4e00-\u9fff]", text):
        print("没有检测到中文正文。", file=sys.stderr)
        return 2

    failures: list[str] = []
    warnings: list[str] = []
    line = lambda pos: raw.count("\n", 0, pos) + 1

    for symbol, label in (("—", "破折号"), ("–", "连接号式破折号")):
        for m in re.finditer(re.escape(symbol), text):
            failures.append(f"{label}：第 {line(m.start())} 行")
    for symbol, label in (("：", "中文冒号"), (":", "英文冒号")):
        for m in re.finditer(re.escape(symbol), text):
            tail = text[m.end():m.end() + 2].lstrip()
            if tail[:1] not in ("「", "『", "“", "‘", '"'):
                failures.append(f"提示性{label}：第 {line(m.start())} 行")
    for pos, term in matches(text, HARD_TERMS):
        failures.append(f"硬禁词“{term}”：第 {line(pos)} 行")
    for pos, term in matches(text, META_LEAK_TERMS):
        failures.append(f"写作过程话语“{term}”：第 {line(pos)} 行")
    for pattern in PIVOT_PATTERNS:
        for m in pattern.finditer(text):
            excerpt = re.sub(r"\s+", " ", m.group())[:44]
            failures.append(f"翻案腔：第 {line(m.start())} 行“{excerpt}”")
    for pos, term in matches(text, ROAD_SIGNS):
        failures.append(f"模型路标“{term}”：第 {line(pos)} 行")

    nominal = []
    for pattern in NOMINALIZATION_PATTERNS:
        nominal.extend(pattern.finditer(text))
    if nominal:
        warnings.append(f"名词化句式 {len(nominal)} 处，改成直接动词并确认有实际结果")
    conjunction_count = sum(text.count(term) for term in CONJUNCTIONS)
    han_count = len(re.findall(r"[\u4e00-\u9fff]", text))
    if han_count >= 600 and conjunction_count * 1000 / han_count > 7:
        warnings.append(f"连词密度偏高：每千字约 {conjunction_count * 1000 // han_count} 个")
    if re.search(r"(?:此外|同时|因此|综上|总的来说).{0,12}(?:未来|新篇章|新征程|未来可期)", text):
        warnings.append("结尾可能是通用积极收束，回到当前事实或已知动作")

    print(f"人化门禁：硬错误 {len(failures)}，风格提醒 {len(warnings)}")
    for item in failures:
        print(f"- {item}")
    for item in warnings:
        print(f"⚠ {item}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
