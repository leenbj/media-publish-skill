# 搜狐号（a-1）真发流程实测记录

2026-09-08 用桌面可见浏览器（ego lite）人工式跑通全流程，文章《中文域名.网址对企业有用吗？品牌保护与知识产权不可或缺的一环》已成功发布（toast"审核中"，列表状态"审核中"）。

## 关键 DOM 事实

### 1. 入口
- 后台首页：`https://mp.sohu.com/mpfe/v4/contentManagement/first/page`（不是 /index）
- 点顶部黄色按钮 `.mt-button.publish-btn`（文本"发布内容"）→ 跳编辑页
- 编辑页：`https://mp.sohu.com/mpfe/v4/contentManagement/news/addarticle?contentStatus=1`
- ⚠️ 首页有腾讯验证码 iframe（turing.captcha.gtimg.com）但通常自动过，不影响登录态访问

### 2. 编辑页结构
- 首次进编辑页有"文章排版新升级"提示条（"我知道了"按钮）→ 关掉
- 标题：`input[placeholder="请输入标题（5-72字）"]`
- 正文：`.ql-editor[contenteditable=true]`（Quill 编辑器），focus 后键盘输入，Enter 分段
- 底部固定栏（LI 元素，不是 button！）：
  - 发布：`li.publish-report-btn.active.positive-button`
  - 定时发布：`li.publish-report-btn.normal.timeout-pub`
  - 存草稿：`li.publish-report-btn.normal.negative-button`
  - 预览：`li.normal.negative-button`
  - 配额提示："您今天还能发 5 篇文章"（每日限 5 篇！）

### 3. 发布必选项：创作声明
- 不选声明点发布会 toast"请添加必选声明"，发布失败
- 声明区：`.statement-panel` 内 `label.el-radio.info-radio-label`（ElementUI radio）
- 选项：无需声明 / 含有虚构演绎内容 / 含有AI生成内容 / 含有营销信息 / 内容为转载 / 内容为个人观点
- ⚠️ 选中方式：必须对 `input.el-radio__original` 派发原生 mousedown/mouseup/click/change 事件序列
  （直接 input.click() 或 label.click() 无效！Vue 绑定需要原生事件流）
- 选中态：label 加 `is-checked` 类，input.checked=true
- 文章发布默认选"无需声明"（原创文章）；AI 生成内容必须选"含有AI生成内容"

### 4. 发布流程
1. 填标题、正文
2. 选创作声明（必选）
3. 点 `li.publish-report-btn.active`
4. 成功 → 自动跳回 `first/page`，toast"审核中"
5. 列表里文章状态"审核中"，点击标题 → `articlepreview?id=xxx`（后台预览链接，非正式链接）

### 5. 正式链接规则
- 审核中：只有 articlepreview 后台预览链接
- 已发布：列表里同一行出现 `https://www.sohu.com/a/<id>_<accountId>` 正式链接（参照已发布历史文章）
- accountId=121426417（本账号固定）

### 6. 封面
- 封面区提示"封面图片尺寸应大于450*300"，不传图也能发布（实测跳过封面直接发成功）
- 无免费图库（与头条不同），如需封面只能本地上传

## 与头条号的差异要点
| 项 | 头条号 | 搜狐号 |
|---|---|---|
| 发布按钮 | button"预览并发布" | li.publish-report-btn |
| 二次确认 | 有"确认发布"弹窗 | 无 |
| 免费图库 | 有（关键词搜索） | 无（本地上传） |
| 封面 | 必选（单图/三图/无封面） | 可跳过 |
| 创作声明 | 无 | 必选（el-radio，原生事件流） |
| 发布后跳转 | 文章列表页 | 内容管理首页 |
| 每日限额 | 无明显限制 | 5 篇/天 |

## 待办
- [ ] 写 publish_sohu.py（Playwright 版）+ dry-run 验证
- [ ] CSDN（c-1）同流程探路
- [ ] 全部完成后写 collect_links.py 回查正式链接
