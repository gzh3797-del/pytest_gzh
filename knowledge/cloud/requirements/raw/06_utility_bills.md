# Utility Bills 公用事业账单

**路由**：`#/utilityBills/utilityBillsAccount`

## 概述

Utility Bills 模块用于录入和管理外部公用事业（电网公司/水务/燃气）出具的实际账单，并与 AcuCloud 计量数据进行对比分析。

> 详细内容已并入 [05_billing.md](05_billing.md) 的 Utility Bills 章节。

## 子页面（Tab）

| Tab | 说明 |
|-----|------|
| **Account** | 账户管理（绑定公用事业账户） |
| **Utility Bills** | 账单录入与查看 |
| **Utility Bills Analysis** | 账单与计量数据对比分析 |
| **OCR Upload** | PDF账单OCR识别上传（第四阶段，ACU-833） |
| **Templates** | 账单模板自定义（第三阶段） |
| **Alerts** | 账单超额告警配置（ACU-813/818） |

---

## Account（账户）配置

| 字段 | 说明 |
|------|------|
| Facility | 关联设施 |
| Account | 公用事业账户号 |
| Meter Number | 电表/水表/气表编号 |
| Utility Type | 能源类型（电/水/气） |

---

## 主要功能场景

1. **账单验证**：实际账单与 AcuCloud 计量数据对比，发现计量误差
2. **多账期对比**：历史账单趋势分析
3. **预算管理**：基于历史账单的能源费用预算（Budget Tracking, ACU-829）
4. **OCR 智能录入**：PDF账单自动识别，减少手工录入工作量

## 订阅要求

| 功能 | 最低 Plan |
|------|-----------|
| 基础 Utility Bills | Lite |
| Utility Bills Analysis | Plus / AcuEMS |
| Utility Bills Report | Plus / AcuEMS |
| OCR Upload | Plus / AcuEMS |
| Submeter Bills | Plus / AcuEMS |
| Budget Tracking | Plus / AcuEMS |

---

> 更多详细内容请参考：
> - [05_billing.md](05_billing.md) — Utility Bills 完整功能（OCR/Template/Alerts/Budget）
> - [10_report.md](10_report.md) — Utility Bills Report
