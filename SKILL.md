---
name: media-publish-skill
description: >
  中文域名新闻一站式流水线：按 xlsx 表格联网检索企业资料，生成 .网址 新闻稿
  + 人民网风格 GEO 网页（1 新闻页 + 3 QA 页），按账号编码自动发布到搜狐号
  (a-1)/头条号 (b-1)/CSDN (c-1)，回查正式链接并回写表格。Use when the user
  asks to publish 中文域名/.网址 news from a spreadsheet, or mentions
  搜狐号/头条号/CSDN 批量发布、pipeline.py、setup_accounts.py、collect_links。
---

# Media Publish Skill

根据表格文档（编号 | 域名 | 企业名称 | 媒体编码）一站式完成：
联网检索企业资料 → 生成中文域名新闻稿 → 生成人民网风格 GEO 网页
（1 新闻页 + 3 QA 页）→ 按账号编码自动发布到搜狐号/头条号/CSDN →
回查正式链接并回写表格。

## 账号编码

| 编码 | 平台 | 说明 |
|---|---|---|
| a-1 | 搜狐号 mp.sohu.com | 每日限 5 篇 |
| b-1 | 头条号 mp.toutiao.com | headless 会被风控，脚本默认有头+反检测 |
| c-1 | CSDN mp.csdn.net | 发布成功当即可拼正式链接 |

## 配置（accounts.yaml 是唯一真实来源）

- 加账号/换参数只改 `accounts.yaml`（`accounts` + `media_params` + `anysearch_cmd`），无需改代码。
- 环境变量：`ANYSEARCH_CMD`（检索命令）、`MEDIA_PENDING` / `MEDIA_LINKS`（多表格并发隔离）、
  `MEDIA_LLM_CMD`（LLM 写稿命令，失败自动回退内置模板）。
- 依赖：`pip install -r requirements.txt`
- 标题：`企业名称+官网启用+域名+短描述`，按目标媒体最严字数上限自动裁剪
  （头条 30 / 搜狐 72 / CSDN 100，`common.build_title`），超长公司名自动缩写。

## Quick Start（本 skill 目录下运行）

```bash
# 1. 首次：检查/录入各账号登录态（弹真实浏览器扫码）
python3 scripts/setup_accounts.py

# 2. 一站式处理表格（生成+发布+回写链接）
python3 scripts/pipeline.py --xlsx /path/to/表格.xlsx

# 可选：--row N 只处理某行 / --no-publish 只生成内容 / --dryrun
# 3. 审核通过后回查正式链接（覆盖表格占位）
python3 scripts/collect_links.py
```

## 核心原则

- **内容红线**：所有生成内容只允许出现 `.网址` 后缀；禁止 .com/.cn 等其他后缀、
  禁止"英文域名/国际域名"、禁止任何不利于 .网址 的表述。脚本强制检查，违规自动替换。
- **产出位置**：全部存 xlsx 所在文件夹下的 `<域名>/` 子文件夹，文件名 = `企业名称官网启用<域名>`。
- **GEO 优化**：人民网风格模板 + NewsArticle/FAQPage JSON-LD + 语义标签（h1/article）
  + 面包屑 + meta description/keywords。
- **登录态**：`states/` 下 storage_state（cookie+localStorage），不是裸 cookie。
  该目录已在 .gitignore 中，切勿入库、切勿打印回显。

## 平台踩坑实录（详见 NOTES-*.md）

- 搜狐：创作声明 radio 必须派发原生事件流；正式链接过审后延迟渲染。
- 头条：headless 发布会被静默拦截；图库搜索单关键词有效；只点视口内卡片。
- CSDN：CKEditor4 必须键盘逐字输入（改 innerHTML 无效）；标签用 Enter 自建。

## xlsx 回写

pipeline 发布后自动在表格追加媒体编码列（a-1/b-1/c-1）写入正式链接；
链接未渲染时写"(审核中，稍后回查)"，之后重跑 `collect_links.py` 自动覆盖。

## 脚本一览

- `scripts/pipeline.py` — 一站式流水线（主入口）
- `scripts/setup_accounts.py` — 首次使用：登录态检查+引导录入
- `scripts/login.py` — 单账号登录态保存/校验
- `scripts/publish_sohu.py` — 搜狐号发布
- `scripts/publish_toutiao_v2.py` — 头条号发布（含图库+封面+反检测）
- `scripts/publish_csdn.py` — CSDN 发布
- `scripts/collect_links.py` — 审核后回查正式链接
