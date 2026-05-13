# Report 报告管理

**路由**：`#/report/configs/0`

## 概述

Report 模块提供自动化报告的配置、生成和历史查询功能，支持多种报告类型和定时推送。

## 子页面（Tab）

| Tab | 说明 |
|-----|------|
| **Configs** | 报告配置管理 |
| **Generate** | 手动触发报告生成 |
| **Previous Reports** | 历史报告查询与下载 |

---

## Report Configs（报告配置）

### 报告类型 Tab

| 类型 | 说明 |
|------|------|
| **Energy Report** | 能耗报告 |
| **Billing Report** | 计费报告 |
| **Portfolio Report** | 项目集汇总报告 |
| **Power Quality Report** | 电能质量报告（BETA） |
| **Utility Bills Report** | 公共事业账单报告（Plus/AcuEMS） |

### Energy Report 配置

#### 必填字段

| 字段 | 说明 | 备注 |
|------|------|------|
| Facility | 目标设施 | 唯一性校验（同一 Facility 不可重复配置） |
| Energy Type | Electricity / Water / Gas | |
| Meter Point | 关联计量点 | 默认选择 physical device 类型 |
| Report Template | 报告模板 | |
| Recipient | 邮件接收人 | 支持多个 |

**Facility 唯一性校验：** 若同一 Facility 已有报告配置，再次添加时提示 `"Report facility is duplicated!"`

**Meter Point 默认选择：** 创建报告时，Meter Point 默认选择 physical device；其他类型 Meter Point 可手动选择。

### 报告模板列表

| 模板名 | 说明 |
|--------|------|
| Facility Energy Consumption Trend Analysis | 设施能耗趋势分析 |
| Meter Point Energy Consumption Trend Analysis | 计量点能耗趋势 |
| Meter Point Operations Analyser | 计量点运行分析 |
| Meter Point Active Power & Weekdays Analysis | 计量点有功功率与工作日分析 |
| Top Consuming Meter Points | 高能耗计量点排行 |
| Deployment | 部署/安装报告 |

---

## Utility Bills Report（第三阶段，ACU-UBRP）

### 订阅要求

| Plan | 支持情况 |
|------|---------|
| Plus | ✅ |
| AcuEMS | ✅ |
| AcuPQ | ❌ |
| Lite | ❌ |
| AcuBilling | ❌ |

### 功能说明

- 生成基于 Utility Bills 数据的报告
- 包含账单总额、各费率项分解
- 支持跨设施的账单对比

---

## Billing TOU 汇总 Report（第四阶段）

- 生成 TOU（分时段）电量和费用的汇总报告
- Excel 格式，包含 Peak / Off-Peak / Shoulder 各时段数据
- 支持多设施汇总

---

## Water Report（水报告）

- 第三阶段新增水专项报告
- 报告中包含水量消耗趋势、用水设施对比

---

## 报告推送机制

- 配置接收人邮箱后，按设定频率（日/周/月）自动生成
- 生成报告格式：PDF / Excel
- 支持多个接收人
- 历史报告可在 **Previous Reports** 页面查询下载

---

## 测试要点

| 测试项 | 验证点 |
|--------|--------|
| 能耗报告配置新增 | 配置保存成功，列表显示正确 |
| Facility 唯一性 | 重复 Facility 提示错误 |
| Meter Point 默认选择 | Physical Device 自动选中 |
| Utility Bills Report Plan | Plus/AcuEMS 可用，其他 Plan 不可用 |
| 报告推送 | 邮件按时发送，内容正确 |
| 历史报告下载 | 文件下载成功，内容准确 |
