# Billing 计费管理

**路由**：`#/billing/accounts`

## 概述

Billing 模块实现基于能耗数据的自动/手动计费功能，支持电力、水务、天然气三种能源类型的费率管理，并提供账单分析和报告功能。

## 子页面（Tab）

| Tab | 路由 | 说明 |
|-----|------|------|
| **Accounts** | `/billing/accounts` | 账户配置（设施-计量点-费率绑定） |
| **Auto Billing** | `/billing/autoBilling` | 自动计费任务管理 |
| **Manual Billing** | `/billing/manualBilling` | 手动计费 |
| **Previous Bills** | `/billing/previousBills` | 历史账单查询 |
| **Billing Analysis** | `/billing/billingAnalysis` | 账单费用分析 |
| Rates（按钮） | 通过右上角按钮 | 费率结构配置 |

---

## Rates（费率配置）

### LDC（Local Distribution Company）管理

**LDC** 是费率组织单位，分为两种类型：

| 类型 | 说明 | 命名规则 |
|------|------|---------|
| **Official** | 官方 LDC（系统级，跨组织共享） | 名称全局唯一，不可重复 |
| **Unofficial** | 非官方 LDC（组织私有） | 名称在组织内唯一 |

**规则：**
- Official LDC 只能由 SuperAdmin 创建
- 在 A 组织下名称为 S 的 Unofficial LDC，B 组织创建同名时只能为 Unofficial；A 组织下创建 Official 同名 LDC 则提示重复
- 创建重复 LDC 时提示错误

### 费率类型（Utility Type）

| 类型 | 说明 |
|------|------|
| **Electricity** | 电力费率 |
| **Water** | 水费率 |
| **Gas** | 天然气费率 |
| **Sewage Water** | 污水处理费率（新增，第二阶段） |

### 费率结构类型

| 结构 | 说明 |
|------|------|
| **Fixed** | 固定费用（基本费） |
| **TOU（Time of Use）** | 分时段电价（峰/平/谷时段） |
| **Tiered** | 阶梯电价（按用量区间计费） |
| **Demand** | 需量费用（基于最大需量 kW/kVA） |
| **kVA Demand** | kVA 需量费用 |

**Billing Unit（电力）：** kWh（需手动选择）  
**Rate History（历史费率兼容）：** 编辑 currency 不影响历史关联费率

### TOU 费率时段配置（ACU-844 Schedule 支持 TOU Rate）

- Schedule 分析中支持关联 TOU Rate
- Rate 输入方式：
  - **Manual Input**：手动录入时段和费率参数
  - **Select from Rate List**：从已有费率列表选择
- 两种方式可互相切换，切换不崩溃
- **Water/Gas 类型**不支持 Select from Rate List（与生产环境一致）
- Rate 时间范围有效性校验（不允许时段重叠）

---

## Accounts（账户配置）

### 账户配置逻辑

每个 **Meter Point** 可绑定：
- Electricity Rate Structure（电力费率）
- Water Rate Structure（水费率）
- Gas Rate Structure（天然气费率）

**默认费率：** 默认费率名称为 "Default electricity Rate Structure"，未配置时列表 Rate Structure 列显示 "--"。

**Account 编辑：**
- 可更改绑定的费率
- 历史已绑定的 Meter Point 会继承原有默认费率

---

## Auto Billing（自动账单）

自动账单基于配置的计量周期和费率定时生成账单。

### 创建自动账单

| 字段 | 说明 |
|------|------|
| Facility | 目标设施 |
| Meter Point | 关联计量点 |
| Rate Structure | 费率结构 |
| Billing Period | 计费周期 |

### 电力账单精度

- 数据精度按订阅配置
- 支持 5min / 15min / 30min 数据间隔
- 历史数据自动兼容

### 水账单（Billing_Water）

- 水费率支持所有 Tier 类型
- 包含 Other Fees、Tax and Rebate
- 支持 Sewage Water 费率类型

### 天然气账单（Gas Billing）

- 基于 Gas Model 数据（TOTAL_VOLUME_m3）
- 单位：m³（立方米）

---

## 账单精度与数据校验

| 检验方法 | 说明 |
|---------|------|
| InfluxDB 原始数据对比 | API 接口数据与 InfluxDB 查询结果对比 |
| 账单金额计算逻辑 | 根据费率公式计算验证 |
| TOU 时段数据 | 峰平谷时段电量分别校验 |

---

## 历史账单（Previous Bills）

- 查询已生成的历史账单记录
- 按设施 / 时间筛选
- **批量 Excel 导出**：支持汇总 Report 和子工作簿
- **Excel Summary 格式**：含各费率结构分解

### Billing TOU 汇总 Report（新增）

- 生成包含 TOU 分时段汇总的 Excel 报告
- Peak / Off-Peak / Shoulder 各时段电量和费用

---

## Billing Analysis（账单分析）

对账单数据进行趋势分析和费用构成分解。

---

## Manual Billing（手动账单）

允许运营人员手动触发特定时段的计费任务，需设施已配置费率。

---

## Submeter Bills（子表账单，ACU-790）

- 支持 Plus Plan 和 AcuEMS Plan
- 子设施/子计量点单独生成账单
- 子账单汇总到父设施账单

---

## Utility Bills（公共事业账单）

**路由：** `#/billing/utilityBills`（或 Utility Bills 模块入口）

### 功能概述

Utility Bills 是上传和管理外部公共事业账单（水、电、气供应商账单）的模块，与自动计费账单区分。

### 子功能

| 功能 | 说明 |
|------|------|
| **Entry** | 手动输入账单数据 |
| **OCR Upload** | PDF 账单 OCR 识别上传（ACU-833） |
| **Template** | 账单模板配置 |
| **Alerts** | 账单异常告警（ACU-813/818） |

### OCR Upload（PDF 账单识别，ACU-833）

**流程：**
1. 点击 **New OCR Upload**，弹出 PDF 上传窗口
2. 上传 PDF 账单文件
3. 系统 OCR 识别账单内容，进入审核队列
4. 人工检查识别结果：
   - **Skip Review & Submit**：跳过审核直接提交
   - **Skip this file**：跳过当前文件
   - 字段逐项校验（手动更正识别错误）
5. 提交后显示二次确认弹窗
6. 成功进入 OCR 任务列表

**任务列表筛选：**
- Facility Name 关键字模糊搜索
- PDF 文件名关键字搜索
- Created By 关键字搜索

### Utility Bill Entry Update（ACU-819）

- 账单录入界面优化
- 支持更多字段类型

### Utility Bill Alerts（ACU-813/818）

- 账单金额超阈值告警
- 账单数据缺失告警

### Utility Bills Total（ACU-831）

- 跨设施账单汇总统计

### Utility Template Custom（第三阶段）

- 自定义账单模板字段
- Plus 和 AcuEMS Plan 支持

---

## Budget Tracking（预算追踪，ACU-829）

- 设置能耗/费用预算目标
- 实时对比实际值 vs 预算值
- 超预算告警提示
- 支持月度/年度预算周期

---

## Billing Zoho 集成（ACU-845）

- 账单数据同步至 Zoho CRM/Books
- 支持 Plus Plan 及以上
- SuperAdmin 需在订阅中配置 Zoho 集成信息
- 验证：Zoho 接收到账单邮件，Power Quality Addons 正常计费

---

## 订阅要求（Plan Requirements）

| 功能 | 最低 Plan |
|------|-----------|
| 基础 Billing | Lite |
| Water/Gas Billing | AcuBilling |
| Utility Bills Analysis | Plus / AcuEMS |
| Submeter Bills | Plus / AcuEMS |
| OCR Upload | Plus / AcuEMS |
| Budget Tracking | Plus / AcuEMS |
| Billing Zoho 集成 | Plus / AcuEMS |
| Utility Bills Report | Plus / AcuEMS |

---

## 数据校验方法

| 步骤 | 工具/方法 |
|------|-----------|
| InfluxDB 原始数据 | SELECT EP_IMP_kWh, "-_diff" FROM measurement |
| 账单 API 接口数据 | /api/v1/billing/* 接口 |
| 账单文件下载 | Excel/CSV 导出内容逐行校验 |
