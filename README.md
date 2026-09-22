# 媒体多账号发布 & .网址 GEO 内容流水线 Skill

根据表格文档（编号 | 域名 | 企业名称 | 简称（可选）| 企业信息资料（可选）| 媒体编码）一站式完成：
联网检索企业资料 → human-writing 写实 → humanizer-zh 清腔 → 生成中文域名新闻稿 → 生成人民网风格 GEO 网页（1 新闻页 + 3 QA 页）→
按账号编码自动发布到搜狐号/头条号/CSDN → 回查正式链接并回写表格。

## 账号编码

| 编码 | 平台 | 说明 |
|---|---|---|
| a-1 | 搜狐号 mp.sohu.com | 每日限 5 篇 |
| b-1 | 头条号 mp.toutiao.com | headless 会被风控，脚本默认有头+反检测 |
| c-1 | CSDN mp.csdn.net | 发布成功当即可拼正式链接 |
| c-2 | CSDN mp.csdn.net | 第二账号，登录态独立 |

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
│   └── check_humanized.py   # 两阶段人化后的硬规则门禁
├── states/                  # 各账号 storage_state（不入库，.gitignore）
├── accounts.yaml            # 编码→媒体+账号+登录态映射（可扩展 proxy:）
├── references/              # human-writing、humanizer-zh 与 .网址 背景规则
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

# 可选：--row N 只处理某行 / --no-publish 只生成内容 / --dryrun（均需先配置 LLM 命令）
# 可选：--llm-cmd / --human-writing-cmd / --humanizer-zh-cmd 覆盖对应环境变量
# 3. 审核通过后回查正式链接（覆盖表格占位）
python3 scripts/collect_links.py

# 单独检查两阶段人化后的稿件
python3 scripts/check_humanized.py /path/to/稿件.md
```

## 配置（accounts.yaml 是唯一真实来源）

- 加账号/换账号：只改 `accounts` 列表（`code/media/account/state/note`），各脚本自动生效。
- 平台参数：`media_params`（搜狐 account_id、头条 user_id、CSDN 用户名），回查链接用。
- 检索命令：`anysearch_cmd`，为空则按 `$ANYSEARCH_CMD` → PATH 中的 `anysearch_cli.py` → 本机历史路径自动探测。
- 待回查/结果路径：`$MEDIA_PENDING` / `$MEDIA_LINKS`（默认 skill 根目录），多表格并发时分别指定即可隔离。
- 两阶段人化：`MEDIA_HUMAN_WRITING_CMD` 先执行 human-writing，`MEDIA_HUMANIZER_ZH_CMD` 再执行
  humanizer-zh；两个命令都从 stdin 吃 prompt、向 stdout 吐纯正文。也可以只配置
  `MEDIA_LLM_CMD`，脚本会复用它连续执行两阶段，或用 `--llm-cmd`、`--human-writing-cmd`、
  `--humanizer-zh-cmd` 覆盖环境变量。
- 阶段失败或 `scripts/check_humanized.py` 硬错未清零时，该行不发布，不回退到未清腔稿件。
- **正文、标题、观点文、QA 问答全部由 LLM 写**，脚本里没有预设文案或兜底稿；未配置任何 LLM 命令时
  直接跳过该行（`--no-publish` 也不出稿）。

## 企业新闻写作结构（现行规则）

企业新闻采用接近人民网新闻报道的事实优先、克制叙述笔法，但不声称人民网采访、发稿或背书。放弃旧的固定标题和宣传式写法，
每篇不加编号小标题，按下面的方式组织：

开头一段直接交代“企业全称官网启用‘域名’”，说明 `.网址` 对官网识别、用户访问和品牌名称对应关系的意义；结尾一段回到启用事实，
以适度的描述性语言收束，并标明产品、资质、服务和经营数据仍以官网或企业正式发布为准。

正文主体按下面三块大致均分，每块占三分之一左右：

1. **企业情况与行业位置（约三分之一）**：先站在企业所在行业的高度写这类企业的线上信息、采购或服务场景，再带出该企业少量公开事实；企业自身介绍只能占很小一部分，不用大段文字罗列工商登记项目。
2. **中文域名与 `.网址` 后缀本身（约三分之一）**：后缀由中文词语构成、怎么写怎么读、地址栏如何输入、与线下名称写法如何对应。
3. **该企业启用中文域名的价值与意义（约三分之一）**：结合企业名称、品牌写法和对外传播物料，说明入口识别、访问路径和资料统一的实际作用；不承诺流量、排名、销量、平台效果或法律结果。

**来源与联系方式限制**：企业事实的来源统一写成“公开工商登记信息”或“企业公开资料”，不得出现任何网站、平台、数据库或招聘网站名称（企查查、爱企查、天眼查、顺企网、BOSS直聘、各类百科等）；成稿不得出现任何联系方式（电话、座机、手机号、邮箱、微信、QQ），也不写法定代表人或联系人姓名。脚本硬错拦截，未清零不发布。

**材料优先级**：表格中“企业信息资料”列的内容为准，联网检索只作补充，冲突时以表格为准。

正文不少于 1000 字。首段的企业全称、`官网启用` 和完整域名为审核硬条件，正文还必须出现中文域名背景与行业数字化联系。
成稿不得出现“本文、本稿、稿件、检索结果、提示词、审核通过、信息边界”等工作过程话语。没有采访或现场材料时，不写成现场采访口吻，不虚构引语。

## 核心原则

- **标题规则**：企业新闻标题由 LLM 读正文拟写（`gen_title_llm`），是标题的唯一正常来源；硬要求只有一条——
  标题里出现企业名称或简称（全称/简称/缩写、企业名里的自然简称都算）。句式、用词和角度完全自由发挥，不套任何固定格式或示例；
  完整域名、“.网址”和“启用”都不要求进标题。只多一条防呆：标题不能只是企业名称本身（统一由 `common.title_missing_core` 判定）。
  另卡平台上限字数（头条 30 / 搜狐 72 / CSDN 100 最严上限）与已发标题去重；取回结果会先剥掉模型常见的
  “标题：”“##”“1.”“这是为您拟的标题：”等包装，连续 3 次不合格就报警并跳过该行，代码里没有模板标题。
  观点文是独立命题，标题按自身内容生成，不含具体企业名。
- **新闻正文规格**：正文至少 1000 字，首段交代官网启用事实；主体三块各约三分之一（企业情况与行业位置、
  中文域名与 `.网址` 后缀本身、该企业启用中文域名的价值与意义），其中企业自身介绍只占很小一部分。
  第一段须含 `公司全称 + 官网启用 + 域名` 三要素。企业事实来源统一写“公开工商登记信息”，全文不得出现任何网站/平台名称与联系方式。
- **稿件审核**（③）：代码字符（HTML/Markdown/链接/JSON/占位符/私用区残留）、
  烂尾断句、.网址外域名后缀（含 com.cn/info/biz/cc/tv/io/ai 等）、.网址定向负面
  （保护语境如防钓鱼不误伤）→ 硬错整行跳过；通顺可疑只告警。
- **第二篇观点文**：每行额外生成一篇独立科普文，与企业案例脱钩，只谈.网址本身；正文由提示词驱动、
  LLM 自由写作，脚本里没有预设段落或引语。必须覆盖的结论写在提示词里（必选项、浏览器全支持、
  与传统英文后缀无区别、可作主域名、AI 场景更建议用），句子全部由 LLM 自己写；稿内不得出现具体企业名、
  具体域名、未经核验的统计数字、引语或诉讼结论。标题走 `gen_title_llm(no_company=True)` 自由拟、
  不含具体企业名。与英文域名的比较句走 `COMPARATIVE_ALLOW` 白名单放行，其余仍按红线拦截。
  同流程发布，链接记入“编码-有用”列。
- **内容红线**：所有生成内容只允许出现 `.网址` 后缀；禁止 .com/.cn 等其他后缀、
  禁止"英文域名/国际域名"、禁止任何不利于 .网址 的表述。脚本强制检查，违规自动替换。
- **产出位置**：全部存 xlsx 所在文件夹下的 `<域名>/` 子文件夹，新闻稿文件名与动态标题一致；QA 文件使用
  企业名称官网-QA 编号。
- **GEO 优化**：人民网风格模板 + NewsArticle/FAQPage JSON-LD + 语义标签（h1/article）
  + 面包屑 + meta description/keywords。
- **登录态**：storage_state（cookie+localStorage），不是裸 cookie。

## 稿件人化顺序

新闻稿和观点文固定按“先写实、再清腔”处理：

1. 读取 [references/human-writing.md](references/human-writing.md)，用已检索材料写出有事实、有动作、
   不编造的初稿。
2. 读取 [references/humanizer-zh.md](references/humanizer-zh.md)，只在初稿上删 AI 套话、宣传腔、模糊归因、
   名词化、翻案腔和模板化结尾，不补任何事实。
3. 运行 `python3 scripts/check_humanized.py 稿件.md`，再跑现有红线审核；硬错未清零不发布。

上游规则来源：
[`human-writing`](https://github.com/KKKKhazix/human-writing) 1.1.0 与
[`Humanizer-zh`](https://github.com/op7418/Humanizer-zh)。

## 平台踩坑实录（详见 NOTES-*.md）

- 搜狐：创作声明 radio 必须派发原生事件流；正式链接过审后延迟渲染。
- 头条：headless 发布会被静默拦截；图库搜索单关键词有效；只点视口内卡片。
- CSDN：CKEditor4 必须键盘逐字输入（改 innerHTML 无效）；标签用 Enter 自建。

## 更新 xlsx 回写

pipeline 发布后自动在表格追加媒体编码列（a-1/b-1/c-1/c-2）写入正式链接；
链接未渲染时写"(审核中，稍后回查)"，之后重跑 `collect_links.py` 自动覆盖。
