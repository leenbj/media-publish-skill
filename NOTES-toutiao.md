# 头条号（b-1）真发流程实测记录

2026-09-08 用桌面可见浏览器（ego lite）人工式跑通全流程，文章《中文域名.网址对企业有用吗？品牌保护与知识产权不可或缺的一环》已成功发布并过审（展现3/阅读0）。

## 关键 DOM 事实（写成脚本要用的）

### 1. 关 AI 助手抽屉（挡点击）
- 关闭按钮：`.ai-assistant-drawer svg.close-btn`，原生 click 事件
- 残留遮罩 `.byte-drawer-mask` 不挡按钮（elementFromPoint 验证）

### 2. 填稿
- 标题：`textarea[placeholder*="文章标题"]`（限30字）
- 正文：`.ProseMirror[contenteditable=true]`，`focus()` 后键盘输入，段落间单次 Enter
- 自动保存为草稿（"草稿已保存"）

### 3. 插入免费图库图片（右侧抽屉，非 modal！）
- 触发：`.upload-image-panel`（隐藏容器），工具栏按钮无标题无 icon 名，选择器不可靠
  - 已验证方案：视觉坐标点工具栏图片按钮；或 JS 点击 `.ProseMirror` 后用 `document.elementFromPoint` 探测
- 抽屉结构：`.byte-drawer-wrapper`（z=1002）内 `.byte-tabs`：
  - tab 切换：`.byte-tabs-header-title`（文本="免费正版图片"），active 类挂同一元素
  - 三个 pane：`.byte-tabs-content-item`（active 类挂同一元素），免费图库 pane x≈285 w=960
- 图库搜索框：pane 内 `input[placeholder*="关键词"]`
  - ⚠️ 组合词"域名 互联网"返回空；单关键词"互联网"有结果（30张）
- 结果卡片：`li.item`（图片非 img 标签，是 svg/canvas 渲染）
- 选中：真实鼠标点击卡片中心 → 卡片内出现 `.pop-layer .pop-number`（数字=选中序号）
- 底部确认：按钮文本"取消"/"确定"（确定在 x≈1176, y≈668），"最多可选择 20 张，已选择 N 张图片"
- 点"确定"后图片插入正文，抽屉自动关闭（`.upload-image-panel` 消失）

### 4. 封面
- 位置：`.form-container` 内 `.article-cover`，radio 组（byte-radio）
- 选项：单图 / 三图 / 无封面；选中态：`input.checked` + `.byte-radio-inner.checked`
- 选"单图"用正文刚插入的图做封面；本流程人工验证通过

### 5. 发布
- 点击 `button:has-text("预览并发布")`（playwright 用 get_by_role；ego 用 `xpath=//button[contains(.,"预览并发布")]`）
- 弹确认框（按钮："返回编辑" / "确认发布"）
- 点"确认发布" → 跳转 `/profile_v4/graphic/articles`，toast"审核中"

### 6. 回查链接（collect_links 用）
- 列表页：`https://mp.toutiao.com/profile_v4/graphic/articles`
- 筛选tab文本：全部/已发布/审核中/审核未通过/仅我可见
- 文章行含：标题 + 时间 + 状态（已发布/审核中）+ 数据
- 正式链接：文章行内 `<a>` 的 href 指向标题 → `https://www.toutiao.com/item/<id>/`
- ⚠️ 审核期间状态="审核中"，链接可能未生效；要等状态变"已发布"再抓链接

## 待办
- [ ] 用以上事实重写 publish_toutiao.py（Playwright 版）并 dry-run + 真发各验证一次
- [ ] collect_links.py（发布列表抓状态+链接，审核中的跳过下轮再查）
- [ ] 搜狐号（a-1）、CSDN（c-1）同流程探路
