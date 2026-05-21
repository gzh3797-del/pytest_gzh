# Carbon Model 碳排放模型

**路由**：`#/carbonModel/carbonAnalysis`

## 概述

Carbon Model 模块基于能耗数据和排放因子，计算并分析设施的碳排放量，支持企业碳中和目标管理。

## 订阅要求

| Plan 类型 | 支持情况 |
|-----------|---------|
| **AcuEMS** | ✅ 支持 |
| **Plus** | ✅ 支持 |
| AcuPQ | ❌ 不支持 |
| Lite | ❌ 不支持 |
| AcuBilling | ❌ 不支持 |
| Free | ❌ 不支持 |

> Carbon Model 仅在 AcuEMS 和 Plus Plan 中可用，Plan Detail 主菜单栏中显示 Carbon Model 入口。

## 子页面（Tab）

| Tab | 说明 |
|-----|------|
| **Analysis** | 碳排放分析与可视化 |
| **Emission Factor** | 排放因子配置 |

---

## Analysis（碳排放分析）

### 筛选条件

| 字段 | 选项 |
|------|------|
| Facility | 选择目标设施（必填） |
| Time Frame | Year / Quarter / Month / Custom |
| Data Source | 数据来源选择 |
| Energy Type | Electricity（默认）/ Water / Gas |
| Scope | 排放范围：Scope 1 / Scope 2 |
| Unit | 排放量单位 |

操作按钮：**Update Chart**（更新图表）

### 分析视图

| 视图 | 说明 |
|------|------|
| **Intensity & Emission** | 能源强度 vs 碳排放量双轴图 |
| **Facility Emissions** | 按设施分解的碳排放对比 |

> 初始状态：No data to display（需选择设施后查看）

---

## Emission Factor（排放因子）

### 排放因子管理

排放因子定义不同能源类型对应的碳排放系数（kg CO₂/kWh 或 kg CO₂/m³）。

| 级别 | 配置方 | 说明 |
|------|--------|------|
| **Global Emission Factors** | Super Admin | 全球各地区标准排放因子（由平台统一维护） |
| **Facility 排放因子** | Admin / 组织管理员 | 按设施自定义本地化排放因子 |

**Facility 碳排放因子管理：**
- 可为每个 Facility 单独配置排放因子
- 优先使用 Facility 级因子，未配置时使用 Global 因子

---

## 碳排放核算说明

| 排放范围 | 说明 |
|---------|------|
| **Scope 1** | 直接排放（现场燃烧天然气等直接能源使用） |
| **Scope 2** | 间接排放（购买的电力带来的上游排放） |
| Scope 3 | 价值链排放（暂不支持） |

**计算公式：**

```
碳排放量（kg CO₂）= 能耗量（kWh 或 m³）× 排放因子（kg CO₂/kWh）
```

---

## 测试要点

| 测试项 | 关键验证点 |
|--------|-----------|
| Plan 权限 | AcuPQ/Lite/AcuBilling/Free Plan 用户无 Carbon Model 入口 |
| Facility 排放因子创建 | 新建因子保存成功，列表正确展示 |
| Global Emission Factors | Super Admin 可查看/编辑全球排放因子 |
| Intensity & Emission 图 | 图表渲染正确，双轴数值准确 |
| Facility Emissions 图 | 多设施分解展示正确 |
| Scope 切换 | Scope 1 / Scope 2 切换后数据正确 |
