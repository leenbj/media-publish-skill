# CSDN（c-1）真发流程实测记录

2026-09-08 ego 可见浏览器 + Playwright 双重验证。文章《中文域名.网址对企业有用吗？品牌保护与知识产权不可或缺的一环》人工式发布成功（success 页 id=164629383），脚本版发布成功（id=164629565）。

## 关键 DOM 事实

### 1. 编辑器
- 编辑页：`https://mp.csdn.net/mp_blog/creation/editor`（需登录态，失效会弹 passport iframe）
- 标题：`textarea[placeholder*="文章标题"]`（5~100字）
- 正文：**CKEditor 4**（iframe.cke_wysiwyg_frame）
  - ⚠️ 直接改 contentDocument.body.innerHTML 不触发编辑器同步（页面字数永远"共 0 字"）
  - ✅ 唯一可靠方式：点击 iframe body 区域聚焦后 `keyboard.type` 逐字真实键入
  - 校验方式：页面计数器"共 N 字"，N>0 才算填入成功
- 页面可能恢复旧草稿（localStorage），发布前要核对标题内容是否本篇

### 2. 文章标签（必填）
- 入口：`button.tag__btn-tag`（文本"添加文章标签"）
- 面板：左侧分类 tabs（el_mcm-tabs__item），右侧该分类推荐标签（span.el_mcm-tag__content）
- 搜索：`input[placeholder*="自定义标签"]`
  - ⚠️ 接口 search-recommend-tag 由 Vue watch 真实输入触发；JS setValue 无效
  - ⚠️ 实测搜"域名"返回推荐列表无结果（接口按内容标签库返回）
- ✅ 最稳做法：搜索框真实键入标签词 + **Enter 创建自定义标签**（placeholder 明示支持）
- 选中校验：`[class*=tag-box]` 下出现该标签文本节点

### 3. 发布
- 底部橙色按钮：`button`（文本"发布博客"）
- 成功：跳转 `https://mp.csdn.net/mp_blog/creation/success/<articleId>`，toast"发布成功！正在审核中"
- toast"请选择创作活动/创作话题"仅提示，不阻塞发布（实测无活动时直接成功）
- ⚠️ **每日发文额度有限**（实测提示"今日发文额度还有 1 篇"），大量发布需间隔/提升额度

### 4. 正式链接
- 发布成功页 URL 含 articleId：`creation/success/164629383`
- 正式文章链接规则：`https://blog.csdn.net/<用户名>/article/details/<articleId>`
- 本账号用户名：`yumingpingce`（ego success 页"我的主页"链接确认）
- 即：发布成功当刻即可拼出正式链接，无需等审核！
  - 但审核未通过时该链接 404，回查时仍需验证 HTTP 200

### 5. 登录态导出（ego → Playwright）
- CSDN 登录态过期后用户在 ego 里重新登录
- 用 ego CDP `Network.getCookies` 取 csdn.net 全域 cookie，转 Playwright storage_state 格式写入 states/csdn-默认账号.json（42 条 cookie）
- origins: [] 即可（cookie 已够，无需 localStorage）

## 三平台差异速查
| 项 | 头条 b-1 | 搜狐 a-1 | CSDN c-1 |
|---|---|---|---|
| 编辑器 | ProseMirror | Quill | CKEditor4 iframe |
| 正文填入 | 键盘输入 | 键盘输入 | **必须键盘输入** |
| 图片 | 免费图库 | 本地上传（可跳过） | 可跳过 |
| 必选项 | 封面三选一 | 创作声明 | 文章标签 |
| 发布按钮 | 预览并发布→确认 | li.publish-report-btn | 发布博客 |
| 成功标志 | 列表页+审核中 | 管理页+审核中 | success/<id> 页 |
| 正式链接 | 列表页标题链接（审核后） | sohu.com/a/<id>_<acc>（审核后） | blog.csdn.net/<user>/article/details/<id>（**当即可拼**） |
| 限额 | 无感 | 5篇/天 | 有额度（当日1篇时实测） |
