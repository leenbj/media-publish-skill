---
name: media-publish-skill
description: >
  中文域名新闻一站式流水线：按 xlsx 表格联网检索企业资料，生成 .网址 新闻稿
  + 人民网风格 GEO 网页（1 新闻页 + 3 QA 页），按账号编码自动发布到搜狐号
  (a-1)/头条号 (b-1)/CSDN (c-1/c-2)，回查正式链接并回写表格。生成稿件时强制
  先按 human-writing 写实，再按 humanizer-zh 清理残留 AI 腔。Use when the user
  asks to publish 中文域名/.网址 news from a spreadsheet, submit local articles
  to 头条号/CSDN, or mentions 生成新闻稿、写新闻、中文域名新闻、发布文章、
  搜狐号/头条号/CSDN 批量发布、pipeline.py、setup_accounts.py、collect_links、
  human-writing、Humanizer-zh。Replaces the former chinese-domain-news and fabu skills.
---

# Media Publish Skill

根据表格文档（编号 | 域名 | 企业名称 | 媒体编码）一站式完成：
联网检索企业资料 → 生成中文域名新闻稿 → 生成人民网风格 GEO 网页
（1 新闻页 + 3 QA 页）→ 按账号编码自动发布到搜狐号/头条号/CSDN →
回查正式链接并回写表格。

## 稿件人化流水线（强制顺序）

所有要发布的新闻稿和观点文都按下面顺序处理，不能把两个阶段调换，也不能只做第二阶段：

1. **先写实：human-writing。** 检索后先整理材料来源、时间、数字、业务和已知限制。只用可核验材料，
   每段推进一件新事实或新动作；材料不足就缩短，不编现场、客户、数据、引用或个人经历。白话打底，
   主语和动作尽早出现，避免翻案腔、三连排比、破折号、提示性冒号和汇报黑话。详见
   [references/human-writing.md](references/human-writing.md)。
2. **再清腔：humanizer-zh。** 只编辑上一步的初稿，保留全称、`.网址`、数字、单位、来源归属和用户指定
   的核心结论，不添加新事实。删除夸大意义、宣传词、模糊归因、名词化、模板化连接词、假金句和万能结尾，
   调整长短句节奏；企业新闻保持克制的编辑口吻，不伪造“我亲眼见过”。详见
   [references/humanizer-zh.md](references/humanizer-zh.md)。
3. **门禁。** 对经过外部两阶段处理的稿件，先运行 `scripts/check_humanized.py`，再运行本 skill 的 `.网址`
   红线和稿件审核。任一硬错误未清零都不发布；风格告警要人工回看，不能靠同义词替换掩盖。

当通过 `MEDIA_LLM_CMD` 或两个阶段命令调用外部模型时，`scripts/pipeline.py` 会强制执行这两步。只有
`MEDIA_LLM_CMD` 时会复用同一命令跑两次；单阶段输出失败或门禁不通过时跳过该行，不把未清腔稿件送上媒体。
未配置任何 LLM 命令时只能使用内置的确定性兜底稿，不能宣称它经过了外部两阶段人化；需要正式发布时应配置
阶段命令并重新运行。

## 企业新闻写作结构（现行规则）

企业新闻统一采用接近人民网新闻报道的事实优先、克制叙述笔法，但不得声称人民网采访、发稿或背书。
放弃旧的固定标题和旧的宣传式写法。每篇新闻按以下四个自然部分组织，不加编号小标题：

1. **首段，启用事实和入口意义。** 直接写“企业全称官网启用‘域名’”，说明启用 `.网址` 对官网识别、
   用户访问和品牌名称对应关系的直接意义。
2. **第二段，企业资料。** 以联网检索到的公开页面、正式文件或企业自述为来源，介绍成立时间、所在地、
   业务、产品、设施等已核实情况。来源归属要清楚，不能把搜索摘要之外的经营数据、客户评价或现场细节补进来。
3. **第三部分，行业数字化和 `.网址`。** 先联系企业所在行业的线上信息、服务或供应链场景，再用
   [references/url-material.md](references/url-material.md) 中用户提供的 `.网址` 背景资料解释中文入口的识别、访问和资料统一价值，
   最后回到本企业注册使用该域名的意义。只写可观察的使用价值，不承诺流量、排名、销量、平台效果或法律结果。
4. **末段，总结和边界。** 回到本次启用这一事实，使用适度的描述性语言收束，并明确产品、资质、服务和经营数据
   仍以官网或企业正式发布为准。

文章正文不少于 1000 字。首段的企业全称、`官网启用` 和完整域名是审核硬条件；正文还必须出现中文域名背景和
行业数字化联系。成稿不得出现“本文、本稿、稿件、检索结果、提示词、审核通过、信息边界”等工作过程话语。
没有采访或现场材料时，不能写成现场采访口吻或虚构引语。

## 账号编码

| 编码 | 平台 | 说明 |
|---|---|---|
| a-1 | 搜狐号 mp.sohu.com | 每日限 5 篇 |
| b-1 | 头条号 mp.toutiao.com | headless 会被风控，脚本默认有头+反检测 |
| c-1 | CSDN mp.csdn.net | 发布成功当即可拼正式链接 |
| c-2 | CSDN mp.csdn.net | 第二账号，登录态独立 |

## 配置（accounts.yaml 是唯一真实来源）

- 加账号/换参数只改 `accounts.yaml`（`accounts` + `media_params` + `anysearch_cmd`），无需改代码。
- 环境变量：`ANYSEARCH_CMD`（检索命令）、`MEDIA_PENDING` / `MEDIA_LINKS`（多表格并发隔离）、
  `MEDIA_LLM_CMD`（兼容的双阶段 LLM 命令）、`MEDIA_HUMAN_WRITING_CMD`（第一阶段命令）、
  `MEDIA_HUMANIZER_ZH_CMD`（第二阶段命令）。命令均从 stdin 读取提示词、向 stdout 输出纯正文；
  失败或门禁未过时不发布该行。两个专用命令未同时设置时，缺少的一段复用已有命令。
- 依赖：`pip install -r requirements.txt`
- 标题：放弃旧的固定模板。企业新闻标题只需出现企业全称或简称，不强制出现“官网”或完整域名；标题还必须有自然的新闻动作
  或事实角度，例如“企业简称官网入口更新”“企业名称主营产品调整”。完整域名和“官网启用”事实放在正文首段说明。
  角度由正文已出现的事实动态规划，未被正文支持的角度不得硬塞。超出平台字数上限时才使用 xlsx 简称或官方缩写；仍放不下时
  保留可识别的名称前缀并告警（`common.build_title`）。
  独立观点文标题同样按文章内容（所选观点角度）命题，且必须含企业全称或简称，格式“企业名：观点角度”；字数放不下时
  依次换下一个角度/更短名称，全部超限时才截断并告警（`pipeline.gen_useful_article`）。
- 新闻稿正文至少 1000 字，按“启用事实与入口意义 → 企业公开资料 → 行业数字化与 `.网址` 背景 → 总结边界”组织，
  第一段必须出现全称、`官网启用` 和域名三要素，例如 `公司全称官网启用“域名”。`。标题只要求企业全称或简称加新闻动作，
  没有采访或现场材料时，
  不写成现场采访口吻，不虚构引语。

## Quick Start（本 skill 目录下运行）

```bash
# 1. 首次：检查/录入各账号登录态（弹真实浏览器扫码）
python3 scripts/setup_accounts.py

# 2. 一站式处理表格（生成+发布+回写链接）
python3 scripts/pipeline.py --xlsx /path/to/表格.xlsx

# 可选：--row N 只处理某行 / --no-publish 只生成内容 / --dryrun
# 可选：--llm-cmd / --human-writing-cmd / --humanizer-zh-cmd 覆盖对应环境变量
# 3. 审核通过后回查正式链接（覆盖表格占位）
python3 scripts/collect_links.py

# 单独检查一篇已经过两阶段处理的稿件（硬错返回非 0）
python3 scripts/check_humanized.py /path/to/稿件.md
```

## 核心原则

- **内容红线**：所有生成内容只允许出现 `.网址` 后缀；禁止 .com/.cn 等其他后缀、
  禁止"英文域名/国际域名"、禁止任何不利于 .网址 的表述。脚本强制检查，违规自动替换。
- **人化顺序**：先 human-writing 写实，后 humanizer-zh 清腔；两阶段都不得补造事实。
- **新闻稿规格**：正文不少于 1000 字，首段交代官网启用事实，全文按“发生了什么、公开资料确认了什么、
  用户如何访问、哪些信息尚待核验”的新闻结构推进；正文不得泄漏写作提示和审核过程。
- **产出位置**：全部存 xlsx 所在文件夹下的 `<域名>/` 子文件夹，新闻稿文件名与动态标题一致；QA 文件使用
  企业名称官网-QA 编号，文件名不是标题规则的额外约束。
- **GEO 优化**：人民网风格模板 + NewsArticle/FAQPage JSON-LD + 语义标签（h1/article）
  + 面包屑 + meta description/keywords。
- **登录态**：`states/` 下 storage_state（cookie+localStorage），不是裸 cookie。
  该目录已在 .gitignore 中，切勿入库、切勿打印回显。

## 平台踩坑实录（详见 NOTES-*.md）

- 搜狐：创作声明 radio 必须派发原生事件流；正式链接过审后延迟渲染。
- 头条：headless 发布会被静默拦截；图库搜索单关键词有效；只点视口内卡片。
- CSDN：CKEditor4 必须键盘逐字输入（改 innerHTML 无效）；标签用 Enter 自建。

## xlsx 回写

pipeline 发布后自动在表格追加媒体编码列（a-1/b-1/c-1/c-2）写入正式链接；
链接未渲染时写"(审核中，稍后回查)"，之后重跑 `collect_links.py` 自动覆盖。

## 脚本一览

- `scripts/pipeline.py` — 一站式流水线（主入口）
- `scripts/setup_accounts.py` — 首次使用：登录态检查+引导录入
- `scripts/login.py` — 单账号登录态保存/校验
- `scripts/publish_sohu.py` — 搜狐号发布
- `scripts/publish_toutiao_v2.py` — 头条号发布（含图库+封面+反检测）
- `scripts/publish_csdn.py` — CSDN 发布
- `scripts/collect_links.py` — 审核后回查正式链接
- `scripts/check_humanized.py` — 两阶段人化后的硬规则门禁
