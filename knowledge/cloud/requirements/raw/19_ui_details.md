# UI 精确细节（2026-05-18 实测）

> 本文档记录通过 Playwright 自动化脚本（headless 模式）对 renjie.jiao 账号的全模块探索结果。
> 探索时间：2026-05-18，补充更新：2026-05-19（框架混用、删除弹窗、Facility 表单实测）。
> 所有按钮名称、表头、过滤器均为截图/DOM 实测值。

---

## ⚠️ Installation 模块 UI 框架混用（2026-05-19 实测）

| Tab | 框架 | Playwright 选择器特征 |
|-----|------|--------------------|
| Facilities | Element Plus | `.el-button`、`.el-dialog`、`.el-table__row` |
| **Devices** | **Ant Design** | `.ant-table-row`、`.ant-table-filter-trigger`、`.ant-btn` |
| Meter Points | **Ant Design**（顶部过滤 + 表格） | 过滤栏用 `.ant-select-multiple`（⚠️ 不是 `el-select`）；表格行用 `.ant-table-row` |
| Alerts | Element Plus | `.el-button`、`.el-dialog` |

---

## Org View / Facility View 切换机制（2026-05-19 实测）

顶部导航右上角有一个视图切换按钮，两种状态：

| 状态 | 显示文字 | 含义 |
|------|---------|------|
| Org View（组织视图） | 灰色文字 "Organization View" | 当前在组织视图，可见所有设施数据 |
| Facility View（设施视图） | 绿色标签 "Facility View" | 已切换到某设施视角，部分页面缺少跨设施过滤器 |

**影响自动化的关键差异**：
- **Meter Points 过滤栏的 "Facilities" 下拉仅在 Org View 下存在**；Facility View 下该过滤器消失
- **Playwright 检测方法**：
  ```python
  # 在 Org View：
  page.locator("text=Organization View").count() > 0   # True
  # 在 Facility View：
  page.locator("text=Facility View").count() > 0        # True
  ```
- **切换回 Org View**：点击绿色 "Facility View" 文字即可切换回组织视图

---

## 登录账号

| 字段 | 值 |
|------|-----|
| 邮箱 | renjie.jiao@accuenergy.com |
| 当前组织 | AG PROYECTOS Y SERVICIOS, S.A.（orgId: 431） |
| 当前设施 | receiver_20250807（facilityId: 5341） |
| 角色 | SUPER |

---

## 顶部导航栏（全局）

| 元素 | 说明 |
|------|------|
| 🌐 English ▾ | 语言切换（English/Español/Français 等） |
| 🏠 Home | 快速跳转首页 |
| ❓ Help | 帮助文档入口 |
| ✉️ | 消息通知图标 |
| 👤 ▾ | 用户菜单（个人信息、修改密码、退出） |
| Organization View（右上开关） | 切换组织视图/设施视图 |

---

## Home（首页）

**URL**：`/#/home`

| 按钮 | 功能 |
|------|------|
| `Configure Offline Alerts` | 进入告警配置（首页离线设备告警横幅内） |
| `Keyboard shortcuts` | 显示快捷键帮助 |

---

## Installation / Facilities

**URL**：`/#/installation/facilities`（或 `/#/installation/devices` → Facilities tab）
**UI 框架**：Element Plus

| 按钮 | 功能 |
|------|------|
| `+Add Facility` | 新增设施（右上角，绿色） |
| 行内 👁 | 查看设施详情 |
| 行内 ✏️ | 编辑设施 |
| 行内 🗑️ | 删除设施 |

**表头**：Facility · Type · Devices · offline Devices · Action

### Facility 列列头搜索（2026-05-19 实测）

⚠️ Facilities 表格使用 **Element Plus**，Facility 列有独立的列头搜索功能（不是 Ant Design）。

| 元素 | 选择器 | 说明 |
|------|--------|------|
| 列头搜索触发图标 | `thead th:has-text("Facility")` 内的 `.el-table__column-filter-trigger` 或 `[class*='search']` | 点击弹出搜索面板 |
| 搜索输入框 | `.el-table-filter:visible input` / `.el-popper:visible input` | 普通文本 input |
| 确认按钮 | `button:has-text('Search'):visible` | 点击后列表过滤为匹配行 |
| 重置按钮 | `button:has-text('Reset'):visible` | 清除过滤 |

**Playwright 自动化步骤**：
```python
# 1. 找 Facility 列头
for th in page.locator("thead th").all():
    if "Facility" in th.text_content():
        th.locator(".el-table__column-filter-trigger").last.click()
        break
# 2. 填入搜索值
page.locator("input:visible").first.fill(facility_name)
# 3. 点击 Search 确认
page.locator("button:has-text('Search'):visible").first.click()
page.wait_for_timeout(800)
```

**删除弹窗（2026-05-19 实测）**：
- 有关联设备时：弹出提示，按钮为 **"Got It"**（不是 Yes/Confirm）
- 无关联设备时：正常确认弹窗
- **必须先删设备再删 Facility**，否则被限制

**Facility 表单自动化（2026-05-19 实测）**：
- City 字段：用 `get_by_role("textbox")` 定位，避免与 Electricity 等字段冲突
- Country → Province：Country 选完后等 **≥500ms** 再操作 Province（级联加载）

---

## Installation / Devices

**URL**：`/#/installation/devices`
**UI 框架**：**Ant Design**（⚠️ 与其他 Tab 的 Element Plus 不同）

| 按钮 | 功能 |
|------|------|
| `Expand` | 展开/折叠层级树结构 |
| `+Add Device` | 新增物理设备 |
| `+Add Calculated Meter` | 新增计算电表 |
| `+Add Single Parameter Device` | 新增单参数设备 |
| `+Add Manual Metering Device` | 新增手动抄表设备 |
| `+Export CSV` | 导出设备列表 CSV |

**表头**：Device · Facility · Type · Utility Type · Model · Serial Number · Alert Count · Last Updated · Status · Wiring Issue · Time Zone

⚠️ **Delete 按钮仅在设备详情页，列表页无 Delete**。删除流程：列表点名称 → 详情页左下角 Delete。

### Facility 列列头过滤器（2026-05-19 实测）

Devices 表格（Ant Design）的 **Facility 列**有 `.ant-table-filter-trigger` 列头过滤器，点击后弹出包含 `ant-select-multiple` 的过滤面板。

| 元素 | 选择器 | 说明 |
|------|--------|------|
| 过滤触发图标 | `thead th:has-text("Facility")` 内的 `.ant-table-filter-trigger` | 点击弹出过滤面板 |
| 过滤面板 | `.ant-table-filter-dropdown:visible` | 含 ant-select-multiple + Search/Reset |
| 选择组件 | 面板内 `.ant-select-multiple` | 与 Meter Points 顶栏同款组件 |
| 确认按钮 | `button:has-text('Search')` 或 `.ant-btn-primary` | 点击后表格按选中值过滤 |
| 重置按钮 | `button:has-text('Reset')` | 清除过滤 |

**Playwright 自动化步骤**：
```python
# 1. 找 Facility 列过滤触发器
for th in page.locator(".ant-table-thead th").all():
    if "Facility" in th.text_content():
        th.locator(".ant-table-filter-trigger").click()
        break
# 2. 在面板内的 ant-select-multiple 中选值（同 Meter Points 过滤栏）
dropdown = page.locator(".ant-table-filter-dropdown:visible")
wrapper = dropdown.first.locator(".ant-select-multiple")
wrapper.locator(".ant-select-selector").click()
wrapper.locator(".ant-select-selection-search-input").fill(facility_name)
page.wait_for_timeout(600)
# 选项
page.locator(".ant-select-dropdown:visible .ant-select-item-option").filter(has_text=facility_name).first.click()
page.keyboard.press("Escape")  # 关闭子下拉
# 3. 点击 Search 确认
dropdown.first.locator("button:has-text('Search')").click()
page.wait_for_timeout(800)
```

---

## Installation / Device Detail

通过点击设备名称链接进入。

**详情页 Tab**：OverView · Latest Readings · Wiring · Documents · Alert · Photos · Replacement Records

**OverView 字段**：Device Model · Is Gateway · Facility · Utility Type · Device Source · Data Interval · Location · Serial Number · Installation Time · First Reading Time · Last Updated · Token · Time Zone · Charts · Gateway · Interface · Protocol

**操作按钮**：
- `Delete`（左下，红色）
- `Edit`（右下，绿色）
- `< Back to Devices List`（顶部链接）

---

## Installation / Meter Points

**URL**：`/#/installation/devices` → Meter Points tab

**顶部过滤器**（3 个 **Ant Design** `ant-select-multiple`，⚠️ 不是 `el-select`）：

| 位置索引 | 字段 | 说明 |
|----------|------|------|
| 0 | Facilities | 按设施过滤（Org View 下才有） |
| 1 | Devices | 按设备过滤 |
| 2 | Subscription Tier | 按订阅层级过滤 |

**DOM 结构（2026-05-19 实测）**：
```
div.search-col
  div.search-label        → 文字标签 "Facilities" / "Devices" / "Subscription Tier"
  div.c_common_table_header_select
    div.ant-select.ant-select-multiple.ant-select-show-search   ← wrapper
      div.ant-select-selector                                    ← 点击此处打开下拉
        div.ant-select-selection-overflow
          input.ant-select-selection-search-input               ← 搜索输入框（始终在 DOM 中）
```

**表头**：Meter Point · Facility · Device · Device Model · Utility Type · Number of Channels · Input Channels · Subscription Tier

**表格列索引（Ant Design `.ant-table-row td`，从 0 起）**：

| 列索引 | 字段 |
|--------|------|
| 0 | Meter Point（可点击） |
| 1 | Facility |
| 2 | Device |
| 3 | Device Model |
| 4 | Utility Type |

**按钮**：`+Add Meter Point`

### ant-select-multiple 过滤栏自动化交互（2026-05-19 实测）

**正确交互方式**：
1. 等待 `.ant-select-multiple` 可见（`wait_for(state="visible", timeout=8000)`）
2. 用**位置索引**定位 wrapper：`page.locator(".ant-select-multiple").nth(0)`（Facilities=0）
3. 点击 `.ant-select-selector`（wrapper 内部）打开下拉
4. 等待 `.ant-select-dropdown:visible` 出现
5. 在 `wrapper.locator(".ant-select-selection-search-input")` 输入搜索文字
6. 点击 `.ant-select-dropdown:visible .ant-select-item.ant-select-item-option` 中匹配项
7. **按 `Escape` 关闭下拉**（ant-select-multiple 选中后下拉不会自动关，必须手动关才能触发表格过滤）
8. 等待 800ms，表格刷新

**关键点**：
- ⚠️ **选中后必须按 Escape**，否则下拉遮挡表格，且过滤不一定生效
- ⚠️ 不要用 XPath 查找 wrapper（class 名含动态 hash，XPath 可能找到嵌套子元素而非外层 wrapper）
- ⚠️ Facilities 过滤器**只在 Org View 下存在**；切换到 Facility View 后此过滤器消失

**失败处理**：过滤器失败时直接抛出 `RuntimeError`，不使用列值兜底（兜底方案已删除）。

---

## Installation / Alerts

**URL**：`/#/installation/devices` → Alerts tab

**子 Tab**：Alert Event（默认）· Alert Rules

**Alert Event 表头**：Alarm ID · Facilities · Device Name · Device Model · Alarm Type · Start Time · Last Sent · Frequency · Status · Action

**Alarm Type 枚举**：Parametric · Offline · Flatline

**Status 枚举**：Open（绿色）· Close（灰色）

**行内 Action**：👁 查看 · ✏️ 编辑 · 🔔 手动触发通知

---

## Installation / Portfolio

**URL**：`/#/installation/devices` → Portfolio tab

**按钮**：`Download`（导出 Portfolio 数据）

**表头**：Facility Number · Device · Meter Point · Number/ID · Rate Configured · Rate Structure · Status · Tenants · Full Tenant's Billing Address · Meter Address : Country · Meter Location : Province · Meter Location: City

---

## Billing / Accounts

**URL**：`/#/billing/accounts`

**Tab 导航**：Accounts（默认）· Auto Billing · Manual Billing · Previous Bills · Billing Analysis

**右上角按钮**：`Rates`（进入费率配置）

**过滤器**：Facilities（"Please Select"）

**表头**：Facility · Meter Point · Electricity Rate Structure · Water Rate Structure · Gas Rate Structure · Action

---

## Billing / Auto Billing

**按钮**：`+Add Auto Billing` · `- Electricity`（折叠） · `- Water`（折叠）

**表头**：Facility · Meter Point · Recipients · Billing Period · Start from · Send Raw Data · Action

---

## Billing / Manual Billing

**按钮**：`Cancel` · `Create`

**子 Tab**：Electricity · Water · Gas

**表头**：Meter Point · Meter Point Type · Unit · Rate Structure · Meter Number · Site Code + Candidate · Tenant · Action

---

## Billing / Previous Bills

**按钮**：`Download all selected bills in Excel format` · `- Electricity` · `- Water`

**表头**：Facility · Meter Points · Utility Type · Recipients · Billing Date · Reading Start Date · Reading End Date · Status · Bill Type · Send Raw Data · Action

---

## Billing / Billing Analysis

**按钮**：`update chart`

**子 Tab**：Trend Analysis Tab · TOU Allocation Analysis · Submeter Rate Comparison

---

## Utility Bills / Account

**URL**：`/#/utilityBills/utilityBillsAccount`

**按钮**：`+Add Account`

**表头**：Facility · Account · Meter Number · Utility Type · Action

---

## Utility Bills / Analysis

**URL**：`/#/utilityBills/utilityBillsAnalysis`

**按钮**：`Update Chart`

**子 Tab**：Energy Graph · Costs Graphs · Energy Trend (EUI) · Proportional split

---

## Carbon Model

**URL**：`/#/carbonModel/carbonAnalysis`

**按钮**：`Update Chart`

**子 Tab**：Intensity & Emission · Facility Emissions

---

## Report / Configs

**URL**：`/#/report/configs/0`

**按钮**：`+Add Report Config`

**子 Tab**：Energy Report · Billing Report · Portfolio Report · BETA  Power Quality Report

**表头（Energy Report）**：Facility · Energy Type · Enable · Reports · Recipients · Action

---

## Report / Generate

**按钮**：`Generate Report`

---

## Data / Data Exports

**URL**：`/#/data/dataExports`

**按钮**：`+Add Data Export` · `-`（折叠）

**表头**：Devices · Requested Time · From · To · Status · Time Interval · File Format · Parameters · Action

---

## Data / Data Forwards

**URL**：`/#/data/dataForwards`

**按钮**：`+Add Data Fowrards`（⚠️ 系统 UI 原文，含拼写错误：Fowrards 而非 Forwards）

**表头**：Server Address · Devices · Meter Point · Interval · Last Forward · Last Success · Next Forward · Status · Action

---

## Data / Data Import

**按钮**：`+Add Data Import`

**表头**：Devices · Status · Requested Time · Vee Enable · Action

---

## Data / Energy Exports

**按钮**：`+Add Energy Export` · `-`

**表头**：Meter Point · Requested Time · From · To · Status · Action

---

## Data / Max Min Export

**按钮**：`+Add Max/Min Export` · `-`

**表头**：Meter Point · Created At · From · To · Status · Action

---

## Data / Data Edit

**按钮**：`+Edit All` · `+Verify All`

**表头**：Timestamp · Parameter · Value(kWh) · Version · Updated · Status · Comment · Action

---

## Data / Manual VEE

**按钮**：`+Add Manual Vee` · `-`

**表头**：Facility · Meter Point · Parameter · Time Range · Created At · TimeZone · Executed By · Progress · Status

---

## Admin

**URL**：`/#/admin/users`

**Tab 导航**：Users · Organization · Tenants · Subscription · Customization

### Admin / Users

**按钮**：`+Add User`

**表头**：userName · Email · Login Count · Last Login Time · Last Time Password Changed · Two Factor Authentication · Action

**行内 Action**：⏰ 历史记录 · 👁 查看权限 · ✏️ 编辑 · 🗑️ 删除

### Admin / Tenants

**按钮**：`+Add Tenant`

**表头**：userName · Email · Login Count · Last Login Time · Last Time Password Changed · Two Factor Authentication · Action

### Admin / Subscription

**表头**：Subscription · User Organization · Status · Amount($) · Meter Point - Datalogger · Meter Point - Billing · Meter Point - Power Quality · Action

### Admin / Customization

**按钮**：`Select Image`（上传 Logo） · `Submit`（提交配置）

---

## Super Admin

**URL**：`/#/superAdmin/users`

**Tab 导航**：Users · Parameter · Device Model · Organization · Subscription · Plan · Release Center · Global Emission Factors · Global Device Search · Wiring Check

### Super Admin / Users

**按钮**：`+Add User`

**表头**：User Name · Email · Organization · Login Count · Last Login Time · Last Time Password Changed · Two Factor Authentication · Action

**行内 Action**（比 Admin/Users 多一个锁定按钮）：⏰ · 👁 · ✏️ · 🗑️ · 🔒

### Super Admin / Parameter

**按钮**：`+Add Parameter`

**表头**：Name · Description · Unit · Min Range · Max Range · Parameter Type · DI or AI · Energy · Include in total · Action

### Super Admin / Device Model

**表头**：Name · isPqEventSupported · isPqReportSupported · phases · Utility Type · Action

### Super Admin / Organization

**按钮**：`+Add Organization`

**表头**：Name · Type · Street address · Zip code · City · Country · Maximum allowed alerts · Action

### Super Admin / Subscription

**按钮**：`+Add Subscription`

**表头**：Subscription · Buyer Organization · User Organization · Plan Name · Type · Plan Code · Amount($) · Status · Zoho Status · Creation time · Created by · Action

### Super Admin / Plan

**按钮**：`Plan Details` · `+Add Plan`

**表头**：Plan Code · Type · Plan Name · Bill Every · Price · Description · Action

### Super Admin / Release Center

**按钮**：`+Add Release`

**表头**：Version · Release Time · Created By · Created Date · Notification Start · Status · Action

### Super Admin / Global Emission Factors

**按钮**：`Upload`（批量上传） · `+Add`（单条新增）

**表头**：Year · Scope Category · Energy Type · EF (kgCO₂e/unit) · Action

### Super Admin / Global Device Search

**表头**：Device · Serial Number · Organization · Facility · Type · Utility Type · Status

### Super Admin / Wiring Check

**按钮**：`Export`（导出报告） · `Wiring Check`（执行接线检测）

**表头**：Device Name · Status · Organization · Facility · Meter Point · Input Channel · Phase · Voltage L-N · Current · Diagnosis · Comments · Last Wiring Check Time

---

## 列表过滤器自动化模式汇总（2026-05-19 实测）

AcuCloud 各列表页的过滤机制不统一，共有 **3 种模式**，需按页面分别处理：

---

### 模式一：Ant Design 顶部多选过滤栏（Meter Points Tab）

**适用页面**：Installation → Meter Points

**组件**：`ant-select-multiple`（Ant Design 多选下拉）

| 步骤 | 操作 | 选择器 |
|------|------|--------|
| 导航 | 直接 URL 导航（避免 click_tab 时序问题） | `page.goto("/#/installation/meterPoints")` |
| 定位 wrapper | 按位置索引（Facilities=0, Devices=1, Tier=2） | `page.locator(".ant-select-multiple").nth(0)` |
| 打开下拉 | 点击 selector | `wrapper.locator(".ant-select-selector").click()` |
| 等待下拉 | 检查下拉可见 | `.ant-select-dropdown:visible` |
| 输入搜索 | wrapper 内搜索框 | `wrapper.locator(".ant-select-selection-search-input").fill(value)` |
| 选中选项 | 匹配文本后点击 | `.ant-select-dropdown:visible .ant-select-item.ant-select-item-option` |
| **关闭下拉** | **⚠️ 必须按 Escape** | `page.keyboard.press("Escape")` |
| 等待刷新 | 等表格重新渲染 | `page.wait_for_timeout(800)` |

**关键注意**：
- `ant-select-multiple` 选中选项后**下拉不会自动关闭**，必须手动 Escape 才能触发表格过滤
- Facilities 过滤器仅在 **Org View** 下可见；切换到 Facility View 后消失
- 用位置索引定位，不要用 XPath（class 名不稳定）

---

### 模式二：Ant Design 列头过滤器（Devices Tab）

**适用页面**：Installation → Devices（Facility 列）

**组件**：`.ant-table-filter-trigger` → 弹出 `ant-table-filter-dropdown` → 内含 `ant-select-multiple`

| 步骤 | 操作 | 选择器 |
|------|------|--------|
| 定位列头 | 遍历 `.ant-table-thead th`，找含 "Facility" 文字的 th | `page.locator(".ant-table-thead th").all()` |
| 触发过滤 | 点击列头内过滤图标 | `th.locator(".ant-table-filter-trigger").click()` |
| 等待面板 | 过滤面板弹出 | `.ant-table-filter-dropdown:visible` |
| 定位 wrapper | 面板内的 ant-select-multiple | `dropdown.locator(".ant-select-multiple")` |
| 打开子下拉 | 点击 selector | `wrapper.locator(".ant-select-selector").click()` |
| 输入搜索 | 搜索框 | `wrapper.locator(".ant-select-selection-search-input").fill(value)` |
| 选中选项 | 匹配后点击 | `.ant-select-dropdown:visible .ant-select-item-option` |
| 关闭子下拉 | Escape | `page.keyboard.press("Escape")` |
| **确认过滤** | **⚠️ 必须点 Search 按钮** | `dropdown.locator("button:has-text('Search')").click()` |
| 等待刷新 | | `page.wait_for_timeout(800)` |

**关键注意**：
- 模式二比模式一多一步：选完值后必须点 **Search** 按钮才能生效（模式一选完 Escape 即生效）
- Search/Reset 按钮在 `ant-table-filter-dropdown` 面板底部

---

### 模式三：Element Plus 列头搜索（Facilities Tab）

**适用页面**：Installation → Facilities（Facility 列）

**组件**：`.el-table__column-filter-trigger` → 弹出含普通 `<input>` 的搜索面板

| 步骤 | 操作 | 选择器 |
|------|------|--------|
| 定位列头 | 遍历 `thead th`，找含 "Facility" 文字的 th | `page.locator("thead th").all()` |
| 触发搜索 | 点击列头内搜索图标（最后一个匹配） | `th.locator(".el-table__column-filter-trigger").last.click()` |
| 等待面板 | 搜索面板弹出 | `.el-table-filter:visible` 或 `.el-popper:visible` |
| 输入搜索 | 普通文本框 | `.el-table-filter:visible input` 或 `input:visible` |
| **确认过滤** | **⚠️ 必须点 Search 按钮** | `button:has-text('Search'):visible` |
| 等待刷新 | | `page.wait_for_timeout(800)` |

**关键注意**：
- Element Plus 的搜索输入框是普通 `<input>`，不是 Select 组件，无需 Escape
- 使用 `thead th`（框架无关），不要用 `.ant-table-thead th`（Ant Design 专用）

---

### 三种模式对比表

| 模式 | 页面 | UI 框架 | 触发方式 | 选择组件 | 确认方式 | 需要 Escape |
|------|------|---------|---------|---------|---------|------------|
| 顶部多选栏 | Meter Points | Ant Design | 直接点击 `.ant-select-selector` | `ant-select-multiple` | 选中后 Escape 即生效 | ✅ 必须 |
| 列头过滤器 | Devices | Ant Design | 点击 `.ant-table-filter-trigger` | `ant-select-multiple`（面板内） | 选中 + Escape + 点 Search | ✅ 必须 |
| 列头搜索 | Facilities | Element Plus | 点击 `.el-table__column-filter-trigger` | 普通 `<input>` | 直接点 Search | ❌ 不需要 |

### 失败处理原则

所有三种模式的过滤器操作失败时，**直接抛出 `RuntimeError`**，不使用列值翻页兜底扫描：

```python
if not _apply_facility_filter(page, facility_name):
    raise RuntimeError(f"过滤器选择 '{facility_name}' 失败，请检查 DOM 结构")
```

---

## 删除操作弹窗汇总（2026-05-19 实测）

| 场景 | 弹窗标题 | 按钮 | Playwright 操作 |
|------|---------|------|----------------|
| 删除 Meter Point | **Warning** | No / **Yes** | 点击 `Yes` |
| 删除 Device（从详情页） | **Attention** | Cancel / **Delete** | 点击 `.el-button--danger`（红色 Delete） |
| 删除 Facility（有关联设备时） | 无标题限制弹窗 | **Got It** | 点击 `Got It`，然后先删设备 |
| 子 MP 被 TOTAL MP 绑定时点删除 | **Attention** | **Got It** | 点击 `Got It`，前往 TOTAL MP 解绑 |

⚠️ Device 删除确认弹窗的按钮是 **Delete**（不是 Yes），选择器 `.el-button--danger`。

---

## 操作成功通知（Toast/Notification）

AcuCloud 使用 Element Plus 的 `el-notification` 组件显示操作结果：

| 操作 | 通知文字 | 选择器 |
|------|---------|--------|
| 删除成功 | "Notice: Deleted successfully." 或类似 | `.el-notification--success` |
| 保存成功（如 TOTAL MP 解绑） | "Updated successfully." | `.el-notification--success` |
| 操作失败 | 具体错误信息 | `.el-notification--error` |

**Playwright 等待通知**：
```python
page.wait_for_selector(".el-notification--success", timeout=10000)
```

---

## 已确认访问限制的路由

以下页面在此账号下显示 404 或重定向（"Home / Back" 按钮），说明该路由需要特定订阅或权限：

| 路由 | 说明 |
|------|------|
| `/#/dashboard/list` | Dashboards — 需要订阅 |
| `/#/utilityBills/utilityBills` | Utility Bills 账单列表 — 受限 |
| `/#/log` | Logs — 受限 |
| `/#/admin/organization` | Admin/Organization — 受限 |
