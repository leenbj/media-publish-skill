# 媒体多账号发布 & .网址 GEO 内容流水线 Skill

根据表格文档（编号 | 域名 | 企业名称 | 媒体编码）一站式完成：
联网检索企业资料 → 生成中文域名新闻稿 → 生成人民网风格 GEO 网页（1 新闻页 + 3 QA 页）→
按账号编码自动发布到搜狐号/头条号/CSDN → 回查正式链接并回写表格。

## 账号编码

| 编码 | 平台 | 说明 |
|---|---|---|
| a-1 | 搜狐号 mp.sohu.com | 每日限 5 篇 |
| b-1 | 头条号 mp.toutiao.com | headless 会被风控，脚本默认有头+反检测 |
| c-1 | CSDN mp.csdn.net | 发布成功当即可拼正式链接 |

## 目录结构

```
├── scripts/
│   ├── pipeline.py          # 一站式流水线（主入口）
│   ├── setup_accounts.py    # 首次使用：登录态检查+引导录入
│   ├── login.py             # 单账号登录态保存/校验
│   ├── publish_sohu.py      # 搜狐号发布
│   ├── publish_toutiao_v2.py# 头条号发布（含图库+封面+反检测）
│   ├── publish_csdn.py      # CSDN 发布
│   ├── collect_links.py     # 审核后回查正式链接
│   └── show_proxy_config.py # 代理配置现状查看
├── states/                  # 各账号 storage_state（不入库，.gitignore）
├── accounts.yaml            # 编码→媒体+账号+登录态映射（可扩展 proxy:）
├── NOTES-*.md               # 三平台 DOM 实测笔记（重要踩坑记录）
└── pending-links.json / links.csv  # 回查状态
```

## 快速开始

```bash
pip install -r requirements.txt

# 1. 首次：检查/录入各账号登录态（弹真实浏览器扫码）
python3 scripts/setup_accounts.py

# 2. 一站式处理表格（生成+发布+回写链接）
python3 scripts/pipeline.py --xlsx /path/to/表格.xlsx

# 可选：--row N 只处理某行 / --no-publish 只生成内容 / --dryrun
# 3. 审核通过后回查正式链接（覆盖表格占位）
python3 scripts/collect_links.py
```

## 配置（accounts.yaml 是唯一真实来源）

- 加账号/换账号：只改 `accounts` 列表（`code/media/account/state/note`），各脚本自动生效。
- 平台参数：`media_params`（搜狐 account_id、头条 user_id、CSDN 用户名），回查链接用。
- 检索命令：`anysearch_cmd`，为空则按 `$ANYSEARCH_CMD` → PATH 中的 `anysearch_cli.py` → 本机历史路径自动探测。
- 待回查/结果路径：`$MEDIA_PENDING` / `$MEDIA_LINKS`（默认 skill 根目录），多表格并发时分别指定即可隔离。
- LLM 写稿：`export MEDIA_LLM_CMD="..."`（stdin 吃 prompt、stdout 吐正文）或 `--llm-cmd`；失败或红线未过自动回退内置模板。

## 核心原则

- **标题规则**：`名称 + 官网启用 + 域名 + 短描述`，按本行目标媒体的最严上限裁剪
  （头条 30 / 搜狐 72 / CSDN 100 字）。名称优先级：xlsx“简称”列 > 全称 > 去尾缀
  （…有限公司→…集团）> 地域品牌截断；描述位按正文主题选描述性语言
  （品牌保护/数字化升级/互联网转型/中文直达/安全升级/服务升级），多命中轮换；
  不用广告语，装不下直接省略。
- **正文首段**：须含 `公司全称 + 启用 + 域名` 三要素，位置不限（缺失才补）。
- **稿件审核**（③）：代码字符（HTML/Markdown/链接/JSON/占位符/私用区残留）、
  烂尾断句、.网址外域名后缀（含 com.cn/info/biz/cc/tv/io/ai 等）、.网址定向负面
  （保护语境如防钓鱼不误伤）→ 硬错整行跳过；通顺可疑只告警。
- **第二篇观点文**：每行额外生成一篇独立科普文，与企业案例脱钩，只谈.网址本身；
  6 角度 × 2 变体，标题必含中文.网址，12 连标题正文不重样；正文约千字，
  核心原文必达（必选项、浏览器全支持、与英文域名.com/.cn无区别、主域名、
  AI远超更建议用）；比较原话走白名单放行，其余仍按红线；
  引用 `assets/销售彩页参考.txt`（吴东平、甘绍宁、82%/43%等）；
  新增必含点（2 种表述轮换）：用户习惯普及、百度权重无差异更吸点击、
  Google 词汇权重优先、国际化不弱于.cn、媒体露出利AI抓取；
  同流程发布，链接记入“编码-有用”列。
- **内容红线**：所有生成内容只允许出现 `.网址` 后缀；禁止 .com/.cn 等其他后缀、
  禁止"英文域名/国际域名"、禁止任何不利于 .网址 的表述。脚本强制检查，违规自动替换。
- **产出位置**：全部存 xlsx 所在文件夹下的 `<域名>/` 子文件夹，文件名 = `企业名称官网启用<域名>`。
- **GEO 优化**：人民网风格模板 + NewsArticle/FAQPage JSON-LD + 语义标签（h1/article）
  + 面包屑 + meta description/keywords。
- **登录态**：storage_state（cookie+localStorage），不是裸 cookie。

## 平台踩坑实录（详见 NOTES-*.md）

- 搜狐：创作声明 radio 必须派发原生事件流；正式链接过审后延迟渲染。
- 头条：headless 发布会被静默拦截；图库搜索单关键词有效；只点视口内卡片。
- CSDN：CKEditor4 必须键盘逐字输入（改 innerHTML 无效）；标签用 Enter 自建。

## 更新 xlsx 回写

pipeline 发布后自动在表格追加媒体编码列（a-1/b-1/c-1）写入正式链接；
链接未渲染时写"(审核中，稍后回查)"，之后重跑 `collect_links.py` 自动覆盖。
