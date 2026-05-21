# Installation 安装管理

**路由**：`#/installation/devices`

## 概述

Installation 模块管理物理设施、测量设备和计量点，是 AcuCloud 的数据采集基础。

## 子页面（Tab）

| Tab | 说明 |
|-----|------|
| **Facilities** | 设施管理（建筑/园区级别） |
| **Devices** | 设备管理（电表/网关） |
| **Meter Points** | 计量点管理（虚拟数据流） |
| **Alerts** | 告警规则配置 |
| **Portfolio** | 项目集管理 |

---

## Facilities（设施）页面

### 创建设施（Add Facility）

**必填字段：**

| 字段 | 说明 | 约束 |
|------|------|------|
| Facility Name | 设施名称 | 最长 200 字符，不可重复 |
| Time Zone | 时区 | 从列表选择，例如 America/Toronto |
| Type | 能源类型 | Electric / Water / Gas 等 |
| Facility Serial Number | 设施序列号 | 最长 10 字符 |

**可选字段：**

| 字段 | 说明 | 约束 |
|------|------|------|
| Address | 地址 | 最长 200 字符 |
| Email | 联系邮件 | 格式校验 |
| Latitude / Longitude | 地理坐标 | 数值格式 |
| Notes | 备注 | 最长 255 字符 |

**创建后自动生成：** TOTAL 类型的 Device 和 Meter Point（汇总设备和计量点）。

**Facility ID 唯一性：** 同名 Facility 提示 `id is duplicate`。

### 设施列表字段

| 字段 | 说明 |
|------|------|
| Facility | 设施名称（可点击查看详情） |
| Type | 设施类型 |
| Devices | 设备数量 |
| Offline Devices | 离线设备数 |
| Action | 编辑 / 删除 |

列表支持**按字母/数字排序**，排序后翻页保持排序状态，支持每页条数配置。

### 多层级 Facility（Hierarchical Facility）

- Facility 支持多层级结构（父 Facility / 子 Facility）
- 父 Facility 的 TOTAL 汇总所有子 Facility 的数据
- 支持母表绑定 Facility 的功能

### Facility UI 优化（ACU-787）

- 地图视图：Facility 支持在地图上显示地理位置（需配置经纬度）
- 批量操作增强
- 坐标字段格式校验

### Facility 表单自动化注意事项（2026-05-19 实测）

| 字段 | 自动化要点 |
|------|----------|
| City | 使用 `get_by_role("textbox")` 定位，避免与 Electricity 等同名字段冲突 |
| Country | 选完后必须等待 **≥500ms** 再操作 Province（Province 选项为级联加载） |
| Province | 依赖 Country 选择结果，需等 Country 联动完成 |

### Facility 删除限制与弹窗（2026-05-19 实测）

- **删除顺序**：必须先删除 Facility 下的所有 Device，再删 Facility；否则被系统限制（⚠️ 不是 Meter Points 优先，而是 Device 优先于 Facility）
- **有关联设备时的提示弹窗**：弹出提示按钮为 **"Got It"**，不是通用的 Yes / Confirm / Delete
- **无关联设备时**：正常删除确认弹窗

---

## Devices（设备）页面

**路由**：`#/installation/devices`

### UI 框架说明（2026-05-19 实测）

⚠️ **Installation 模块混用两套 UI 框架**：

| Tab | UI 框架 | 影响 |
|-----|---------|------|
| **Devices** | **Ant Design** | 表格用 `.ant-table-row`，列头过滤为 Ant Design 下拉（`.ant-table-filter-trigger`） |
| **Facilities** | **Element Plus** | 表格、弹窗、按钮均为 Element Plus 组件 |
| **Alerts** | **Element Plus** | 同上 |
| **Meter Points** | **Ant Design**（顶部过滤栏）+ **Ant Design**（表格） | 顶部过滤栏为 `ant-select-multiple`（⚠️ 不是 `el-select`），表格为 Ant Design |

Playwright 脚本中选择器需按 Tab 区分框架：
- Devices 列表过滤：`.ant-table-filter-trigger`（Ant Design）
- Facilities / Alerts 列表：`.el-button`、`.el-dialog` 等（Element Plus）

### 操作按钮

| 按钮 | 功能 |
|------|------|
| Expand | 展开树状结构 |
| + Add Device | 新增物理设备（Physical / Gateway） |
| + Add Calculated Meter | 添加计算表（虚拟） |
| + Add Single Parameter Device | 添加单参数设备 |
| + Add Manual Metering Device | 添加手动计量设备 |
| + Export CSV | 导出设备列表 CSV |

### 设备列表字段

| 字段 | 说明 |
|------|------|
| Device | 设备名称（可点击查看详情） |
| Facility | 所属设施名称（可点击） |
| Type | 设备类型：Physical / Gateway / Calculated 等 |
| Utility Type | 能源类型：Electricity / Gas / Water 等 |
| Model | 硬件型号 |
| Serial Number | 设备序列号 |
| Alert Count | 当前未处理告警数（0 绿色，>0 红色） |
| Last Updated | 最近数据更新时间 |
| Status | 状态：Active / Dormant / Offline 等 |

### 设备类型说明

| 类型 | 含义 |
|------|------|
| Physical | 物理硬件设备（电表等） |
| Gateway | 数据网关（AcuLink810系列） |
| Calculated | 计算表（基于公式从其他计量点计算） |
| Manual Metering | 手动录入计量设备（手动抄表） |
| Single Parameter | 单参数设备 |

### 设备详情（Device Detail）

**截图**：`screenshots/explore_renjie/49_device_detail.png`

⚠️ **Delete 按钮仅在设备详情页存在，列表页没有 Delete 按钮。**
自动化删除设备的流程：列表页点击设备名称 → 进入详情页 → 点击左下角红色 Delete 按钮。

通过点击设备名称链接进入。面包屑路径：Installation / Devices / **Device Detail**

#### 详情页 Tab（2026-05-18 实测确认）

| Tab | 说明 |
|-----|------|
| **OverView** | 基本信息（默认进入） |
| **Latest Readings** | 最新读数数据 |
| **Wiring** | 接线图/接线状态 |
| **Documents** | 设备文档（说明书等，支持 Doc/图片/PDF）又称 Manufacture 文档 |
| **Alert** | 该设备的告警记录 |
| **Photos** | 设备现场照片 |
| **Replacement Records** | 设备更换记录 |

#### OverView 字段

| 字段 | 说明 |
|------|------|
| Device Model | 设备型号 |
| Is Gateway | 是否为网关（Yes/No） |
| Facility | 所属设施（链接） |
| Utility Type | Electricity / Water / Gas / None |
| Device Source | Others 等 |
| Data Interval | 数据采集间隔（如 5 min） |
| Location | 位置描述 |
| Serial Number | 序列号 |
| Installation Time | 安装时间 |
| First Reading Time | 首次读数时间 |
| Last Updated | 最后更新时间 |
| Token | 设备 Token（UUID） |
| Time Zone | 时区 |
| Charts | 图表链接：Profile Analysis; Trend Analysis |
| Gateway | 关联网关 |
| Interface | 接口类型 |
| Protocol | 通信协议 |

#### 详情页操作按钮

| 按钮 | 位置 | 功能 |
|------|------|------|
| `Delete` | 左下，红色 | 删除设备（删除前须先删 Meter Points） |
| `Edit` | 右下，绿色 | 编辑设备信息 |
| `< Back to Devices List` | 顶部链接 | 返回设备列表 |

#### Documents（原 Manufacture）选项卡

- 显示 Document Name、Date Uploaded、File Type
- 支持按 Document Name、Date Uploaded、File Type 排序
- 文档内容与绑定的 Model 信息匹配
- 权限：所有有 Device 权限的用户可查看下载；SuperAdmin/Admin 可操作；Tenant 用户无 Installation 权限

---

## 物理设备创建（Add Device）

### 必填字段

| 字段 | 说明 |
|------|------|
| Device Name | 设备名称 |
| Facility | 所属设施 |
| Model | 硬件型号（如 Acuvim II、Acuvim III、AcuLink810） |
| Serial Number | 设备序列号（全局唯一） |
| Subscription Tier | 订阅级别 |

### SN 重复检测（Global SN Search / ACU-786）

- 序列号在全平台唯一
- 重复 SN 自动打标签提示
- 网关设备和非网关设备均检测
- **Global SN Search**（Super Admin）：跨所有组织搜索设备序列号

---

## 计算表（Calculated Meter）

### 创建流程

1. 点击 **+ Add Calculated Meter**
2. 填写设备名称、设施、数据间隔（5min / 15min）
3. 选择 Utility Type（Electricity / Water / Gas / None）
4. 配置公式行（Custom 或 None 模式）

### 公式配置

| 配置项 | 说明 |
|--------|------|
| Custom 开关 | 默认关闭（新建），历史存量默认开启且不可更改 |
| Data Interval | 5min 或 15min |
| Utility Type | Electricity / Water / Gas / None |

**公式格式：** `$deviceId_serialNumber.paramCode_unit`

例如：`$2516418_TEST2026051412PHY02.VLNa_V`

- 必须引用真实设备参数（纯常数公式如 `"0"` 会被 API 拒绝，返回 `code:501, msg:"Calculated Meter Formula is invalid."`）
- 通过 **Add Formula → Add parameter** 嵌套对话框选择设备参数来构建公式

### 多公式支持（Multi-formula，第四阶段）

- 同一计算表可配置多条公式行
- 每行独立命名，支持 5min / 15min 间隔
- None 计算方式：跳过空数据（不参与计算）
- 历史存量的 Calculated Meter 默认开启 Custom 且不可更改

---

## 手动计量设备（Manual Metering Device / 手动抄表）

### 创建流程

1. 点击 **+ Add Manual Metering Device**
2. 填写设备名称、设施
3. 配置 **Reading Cycle（读数周期）**：
   - Period：Every Month / Every Week / Every Day
   - Day of Month（月模式）：1–31
   - Time：时间点（hh:mm）
4. 配置邮件提醒接收人（Email Recipient）
5. 选择 Utility Type

### 手动抄表（ACU-832）

- 支持通过界面手动输入计量数据
- 手动录入的数据参与账单计算
- 支持批量录入和历史数据修正

---

## Meter Points（计量点）

计量点是 AcuCloud 中的数据流单元：
- 物理测量通道（Physical）
- TOTAL 汇总点（自动创建）
- 计算点（Calculated）

### 单位自动同步（MeterPoint Unit Auto-Sync）

- 计量点单位与设备参数单位自动同步
- 历史数据兼容：存量 Meter Point 在更新模型时单位自动更新

### TOTAL MP 与子 MP 绑定关系（2026-05-19 实测）

**Facility 创建时自动生成 TOTAL MP**（如 `2026051412-Electricity-Total`），所有子 MP 自动绑定到它。

**子 MP 绑定后无法直接删除**——进入子 MP 详情页会弹出 Attention 对话框：
> "This Meter Point is involved in the calculation of 2026051412-Electricity-Total.  
> If you need to delete this Meter Point, please go to the Facility Total and unbind it from the corresponding Facility Total."  
> 按钮：**Got It**

**解绑流程**：
1. 进入 TOTAL MP 详情页（名称含 `-Total`）
2. 找到子 MP checkbox 列表，取消勾选所有已选中项
3. 点击 **Save** → 等待 "Updated successfully." 通知

**TOTAL MP 详情页特征**：
- 底部有 **Save** 按钮（无 Delete 按钮）
- 有 "Meter Point:" 区块列出子 MP 的 checkbox 列表

### 删除 Meter Point

强关联下不可删除，需先解绑相关依赖项。

**MP 删除确认弹窗**（2026-05-19 实测）：
- 标题：**Warning**
- 内容："All data associated with this device will be deleted, and the operation cannot be undone. Are you sure you want to proceed?"
- 按钮：**No** / **Yes** → 点 Yes

---

## Facility + 所有数据的完整清理顺序（2026-05-19 实测）

删除一个 Facility 及其全部数据，**必须严格按以下顺序**，否则被系统限制：

| 步骤 | 操作 | 说明 |
|------|------|------|
| Step 0 | 解绑 TOTAL MP 子 MP | TOTAL MP 详情 → 取消所有 checkbox → Save |
| Step 1 | 删除所有 Meter Points | 非 TOTAL 先删，TOTAL 最后删 |
| Step 2 | 删除所有 Devices | 从设备详情页 Delete（列表页无此按钮） |
| Step 3 | 删除 Facility | 此时无关联设备，可正常删除 |

跳过任何步骤的后果：
- 跳过 Step 0：删子 MP 时弹 Attention 阻止，卡死
- 跳过 Step 1/2 直接删 Facility：弹出 "Got It" 限制弹窗，无法删除

---

## 设备删除优化（Device Deletion Optimization）

### 强关联（不可删除）

当设备满足以下条件时，**禁止删除**，弹出强关联提示：

- 设备是 **810 网关且绑定了其他设备**（有下挂设备）
- 提示语：`Device Deletion Failed for the following reason(s): ...`

### 弱关联（警告后可继续）

设备存在以下关联时，**弹出警告但允许继续**：

- 设备被账单绑定
- 设备被报告引用
- 设备在数据导出中使用
- 等其他软关联

---

## 告警配置（Alerts）

支持配置：

| 告警类型 | 说明 |
|---------|------|
| Offline | 设备离线告警（超时无数据） |
| Parametric | 数值超阈值告警（超上限/低于下限） |
| Flatline | 数据固定不变告警 |
| Rate Alarm | 费率异常告警（见 Billing 模块） |
| Meter Resealing | 计量表密封告警（水表专用） |

通知方式：邮件（Email） / SMS（短信）

**Alert Count 字段：**
- Device 列表显示 Alarm Count 列
- 0：绿色显示
- >0：红色显示，支持按数量升序/降序排序

---

## Portfolio（项目集）

Portfolio 功能用于将多个设施/计量点归组，便于跨设施的统一分析和报告生成。
