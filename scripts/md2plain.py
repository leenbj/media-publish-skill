#!/usr/bin/env python3
"""把文章 md 转为头条编辑器纯文本段落文件（每段一行）。"""
import re
from pathlib import Path

src = Path("/Users/ethan/Desktop/企业注册中文域名.网址有用吗？.md").read_text(encoding="utf-8")

paras = []
for raw in src.split("\n"):
    line = raw.strip()
    if not line:
        continue
    # 标题行
    if line.startswith("# "):
        title = line[2:].strip()
        continue
    line = re.sub(r"^#{1,6}\s*", "", line)      # 小节标题去 #
    line = re.sub(r"^- ", "", line)              # 列表项去 -
    line = line.replace("**", "")                # 去粗体标记
    line = re.sub(r"^[一二三四五六七八九十]+、", "", line) if line and re.match(r"^[一二三四五六七八九十]+、", line) and "##" in raw else line
    paras.append(line)

out = Path("/tmp/article-plain.txt")
out.write_text("\n".join(paras), encoding="utf-8")
print("标题:", title, f"({len(title)}字)")
print("段落数:", len(paras), "总字数:", sum(len(p) for p in paras))
for i, p in enumerate(paras):
    print(f"[{i}] {p[:40]}{'…' if len(p)>40 else ''}")
