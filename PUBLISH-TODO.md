# Office 插件上架商城 TODO 清单

> 插件名称: Office Paste Width  
> 当前版本: v26 (完美+完美)  
> 创建日期: 2026-01-14

---

## 一、Microsoft 合作伙伴中心账号准备

### 中国开发者发布全球市场说明
> ✅ **中国账号可以注册并发布到全球市场**  
> ✅ **中国开发者可以通过 PayPal 或银行电汇收款**  
> ⚠️ 中国大陆用户使用世纪互联版 Office 365，如需覆盖需额外提交

- [ ] **1.1 准备 Microsoft 账号**
  - 推荐使用 Microsoft 工作/学校账号
  - 或使用个人 Microsoft 账号（Outlook/Hotmail）

- [ ] **1.2 注册 Partner Center 账号**
  - 访问: https://partner.microsoft.com/dashboard
  - 注册 **Microsoft 365 and Copilot 程序**

- [ ] **1.3 选择账号类型**
  | 类型 | 费用 | 适用场景 |
  |------|------|----------|
  | 个人账号 | **免费** | 独立开发者 |
  | 公司账号 | **600 CNY** | 企业/团队，更专业 |

- [ ] **1.4 完成身份验证**
  - 邮箱验证
  - 身份验证（个人：身份证；公司：营业执照）
  - 同意 Microsoft Publisher Agreement

- [ ] **1.5 设置收款账户**
  | 收款方式 | 到账时间 | 说明 |
  |----------|----------|------|
  | PayPal | 1 个工作日 | 推荐，需要 PayPal 账户 |
  | 银行电汇 | 7-10 个工作日 | 需要提供银行账户信息 |
  
  - 最低提现门槛: **$50 USD**
  - [ ] 绑定 PayPal 账户（推荐）
  - [ ] 或填写银行账户信息（SWIFT 代码等）

- [ ] **1.6 选择目标市场**
  - [ ] 全球市场（推荐）
  - [ ] 或指定特定国家/地区
  - ⚠️ 如需覆盖中国大陆用户（世纪互联版），需单独提交

---

## 二、插件清单文件 (manifest.xml) 完善

- [ ] **2.1 更新基本信息**
  - [ ] `DisplayName`: 确认插件显示名称
  - [ ] `Description`: 完善插件描述（英文，150字以内）
  - [ ] `Version`: 设置正式版本号（如 1.0.0）
  - [ ] `ProviderName`: 开发者/公司名称
  - [ ] `SupportUrl`: 技术支持页面 URL
  - [ ] `AppDomains`: 添加所有使用的域名

- [ ] **2.2 更新 ID**
  - [ ] 生成新的 `Id` GUID（正式发布用）
  - [ ] 确保 ID 唯一且不会与测试版冲突

- [ ] **2.3 配置托管 URL**
  - [ ] 将 `localhost:3002` 替换为正式托管地址
  - [ ] 所有资源 URL 必须使用 HTTPS

- [ ] **2.4 设置权限和要求**
  - [ ] `Requirements`: 确认最低 Office 版本要求
  - [ ] `Permissions`: 设置最小必要权限

---

## 三、静态资源托管

- [ ] **3.1 选择托管方案**
  - 选项 A: Azure Static Web Apps（推荐）
  - 选项 B: GitHub Pages
  - 选项 C: 自有服务器
  - 选项 D: CDN 服务（如 Cloudflare Pages）

- [ ] **3.2 部署静态文件**
  - [ ] 运行 `npm run build` 生成生产版本
  - [ ] 上传 dist 目录到托管服务
  - [ ] 配置 HTTPS 证书
  - [ ] 测试所有资源可访问

- [ ] **3.3 配置域名（可选）**
  - [ ] 注册自定义域名
  - [ ] 配置 DNS 解析
  - [ ] 配置 SSL 证书

---

## 四、图标和视觉资源

- [ ] **4.1 准备插件图标**
  - [ ] 16x16 px（任务栏小图标）
  - [ ] 32x32 px（任务栏图标）
  - [ ] 80x80 px（插入对话框）
  - [ ] 128x128 px（商城展示）
  - 格式: PNG，透明背景

- [ ] **4.2 准备商城展示图片**
  - [ ] 主图: 1366x768 px（必需）
  - [ ] 截图: 至少 1 张，最多 5 张
  - [ ] 展示插件实际使用效果
  - 格式: PNG 或 JPEG

- [ ] **4.3 准备宣传视频（可选）**
  - [ ] 30-90 秒演示视频
  - [ ] 上传到 YouTube 或其他平台

---

## 五、文档和支持页面

- [ ] **5.1 创建隐私政策页面**
  - [ ] 说明收集哪些数据
  - [ ] 说明数据如何使用
  - [ ] 托管到可访问的 URL
  - 当前文件: `add-in/privacy.html`

- [ ] **5.2 创建支持页面**
  - [ ] 常见问题 FAQ
  - [ ] 联系方式
  - [ ] 问题反馈渠道
  - 当前文件: `add-in/support.html`

- [ ] **5.3 完善 README 文档**
  - [ ] 功能介绍
  - [ ] 使用说明
  - [ ] 系统要求

- [ ] **5.4 准备商城描述文案**
  - [ ] 简短描述（100字以内）
  - [ ] 详细描述（功能特点）
  - [ ] 关键词/标签
  - [ ] 支持的语言列表

---

## 六、测试和验证

- [ ] **6.1 功能测试**
  - [ ] Word 桌面版测试
  - [ ] Word 网页版测试（如支持）
  - [ ] Excel 测试（如支持）
  - [ ] PowerPoint 测试（如支持）

- [ ] **6.2 兼容性测试**
  - [ ] Windows 10/11
  - [ ] macOS（如支持）
  - [ ] 不同 Office 版本（365, 2021, 2019）

- [ ] **6.3 使用 Office Add-in Validator 验证**
  ```bash
  npx office-addin-manifest validate add-in/manifest.xml
  ```

- [ ] **6.4 安全性检查**
  - [ ] 无敏感信息硬编码
  - [ ] 所有外部请求使用 HTTPS
  - [ ] 无恶意代码或追踪器

---

## 七、提交审核

- [ ] **7.1 登录合作伙伴中心**
  - 访问: https://partner.microsoft.com/dashboard

- [ ] **7.2 创建新的 Office 插件提交**
  - [ ] 选择插件类型: Office Add-in
  - [ ] 上传 manifest.xml
  - [ ] 填写商城信息

- [ ] **7.3 填写提交表单**
  - [ ] 插件名称和描述
  - [ ] 类别选择
  - [ ] 定价（免费/付费）
  - [ ] 目标市场/地区
  - [ ] 年龄分级

- [ ] **7.4 上传资源**
  - [ ] 图标
  - [ ] 截图
  - [ ] 隐私政策 URL
  - [ ] 支持 URL

- [ ] **7.5 提交审核**
  - 审核周期: 通常 3-7 个工作日
  - 可能需要根据反馈修改后重新提交

---

## 八、发布后维护

- [ ] **8.1 监控用户反馈**
  - [ ] 商城评论
  - [ ] 支持邮件

- [ ] **8.2 版本更新流程**
  - [ ] 更新代码和 manifest 版本号
  - [ ] 重新部署静态资源
  - [ ] 在合作伙伴中心提交更新

- [ ] **8.3 数据分析（可选）**
  - [ ] 安装量统计
  - [ ] 使用情况分析

---

## 快速参考链接

| 资源 | 链接 |
|------|------|
| Partner Center 登录 | https://partner.microsoft.com/dashboard |
| Partner Center 注册 | https://partner.microsoft.com/dashboard/account/v3/enrollment/introduction/partnership |
| 开发者账号指南 | https://learn.microsoft.com/partner-center/marketplace/open-a-developer-account |
| 提交指南 | https://learn.microsoft.com/partner-center/marketplace/add-in-submission-guide |
| Manifest 验证 | https://learn.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest |
| 认证政策 | https://learn.microsoft.com/partner-center/marketplace/certification-policies |
| 设计指南 | https://learn.microsoft.com/office/dev/add-ins/design/add-in-design |
| 收款详情 | https://learn.microsoft.com/azure/marketplace/partner-center-portal/payment-thresholds-methods-timeframes |

---

## 详细操作指南

### 📝 Step 1: 注册 Partner Center 账号

1. **访问注册页面**
   - URL: https://partner.microsoft.com/dashboard/account/v3/enrollment/introduction/partnership
   - 使用 Microsoft 账号登录（Outlook/Hotmail 或工作账号）

2. **选择程序类型**
   - 在 Programs 页面找到 **"Microsoft 365 and Copilot"**
   - 点击 **"Get Started"**

3. **填写发布者信息**
   - Publisher display name（发布者显示名称）
   - Contact email（联系邮箱）
   - 同意 Microsoft Publisher Agreement

4. **完成验证**
   - 邮箱验证
   - 等待账号审批（通常 1-3 个工作日）

### 📝 Step 2: 提交插件

1. **登录 Partner Center**
   - URL: https://partner.microsoft.com/dashboard

2. **创建新 Offer**
   - 点击 **"Marketplace offers"**
   - 选择 **"Microsoft 365 and Copilot"** 标签
   - 点击 **"+ New offer"** → **"Office Add-in"**

3. **填写产品信息**
   - Offer alias（内部名称）
   - Publisher（选择发布者账号）

4. **上传 Manifest**
   - 在 Packages 页面上传 `manifest.xml`
   - 系统会自动验证

5. **填写商城信息**
   - 类别选择（最多 3 个）
   - 隐私政策 URL
   - 支持文档 URL
   - 多语言描述
   - 图标和截图

6. **设置可用性**
   - 选择发布日期
   - 选择目标市场

7. **提交审核**
   - 填写测试说明（Notes for certification）
   - 点击 **"Review and publish"**

### 📝 Step 3: 审核流程

- **审核周期**: 3-5 个工作日
- **审核内容**:
  - 自动化检查（Manifest 格式、安全性）
  - 人工审核（功能、用户体验）
- **审核结果**:
  - ✅ 通过 → 自动发布到商城
  - ❌ 需修改 → 查看报告，修改后重新提交

---

## 当前进度

| 阶段 | 状态 | 完成日期 |
|------|------|----------|
| 功能开发 | ✅ 完成 | 2026-01-14 |
| 账号准备 | ⬜ 待开始 | - |
| Manifest 完善 | ⬜ 待开始 | - |
| 资源托管 | ⬜ 待开始 | - |
| 图标资源 | ⬜ 待开始 | - |
| 文档准备 | ⬜ 待开始 | - |
| 测试验证 | ⬜ 待开始 | - |
| 提交审核 | ⬜ 待开始 | - |
| 正式发布 | ⬜ 待开始 | - |

---

*最后更新: 2026-01-14*
