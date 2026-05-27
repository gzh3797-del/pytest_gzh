# Dashboard 仪表盘

**路由**：`#/dashboard`

## 概述

Dashboard 是 AcuCloud 的主页可视化面板，支持用户自定义小部件（Widget）布局，实时展示能耗、告警、分析等关键数据。

---

## Dashboard 类型

| 类型 | 说明 |
|------|------|
| **Organization View Dashboard** | 组织级仪表盘（跨所有设施） |
| **Facility View Dashboard** | 设施级仪表盘（单个设施） |

---

## Widget（小部件）类型

### Energy Analysis Widget

**功能：** 展示指定计量点的能耗趋势图

**支持参数：**
- kWh（有功电量）
- kVA（视在功率）— 第二阶段新增

**kVA 配置（第二阶段，Dashboard-Energy Analysis Widget）：**
| 配置项 | 说明 |
|--------|------|
| Show kVA 开关 | 默认勾选（打开） |
| 图表类型 | kVA 和 kWh 统一使用折线图 |
| 配置记忆 | 取消/重新开启 Show kVA 状态被持久化记忆 |

**物理设备/多通道验证：**
- 不同 Time Frame（Today/Yesterday/Last 7 Days 等）下 kVA 数据比对正确
- kVA 数据与 kWh 数据同步展示

### Data Widget

- 展示设备实时数据表格

### Alert Widget

- 展示当前活跃告警列表

### Billing Statistics Widget（Support bill statistics, ACU-789）

- 展示当期账单统计
- 费用汇总和同期对比

### Canvas Feature（ACU-796）

- 自定义画布（Canvas）功能，支持更丰富的可视化配置
- 用户可在画布上拖拽、调整布局

### Widget Update（ACU-808）

- 批量更新 Widget 配置
- 新版本 Widget 功能增强

### Sankey Diagram（ACU-799）

- 桑基图（能流图）展示能量流向
- 可视化各设施/设备间的能量分配关系
- 支持按时间段查看能流分布

---

## Dashboard 过滤

| 过滤维度 | 说明 |
|---------|------|
| 能源类型 | 可过滤展示 Electricity / Water / Gas 数据 |
| 设施 | 可按 Facility 过滤 |

---

## 自动刷新（Dashboard Auto Refresh，第四阶段）

| 配置项 | 说明 |
|--------|------|
| 自动刷新开关 | 开启后 Dashboard 数据定时自动更新 |
| 刷新间隔 | 可配置（如每 5 分钟刷新一次） |

---

## Alert 在 Dashboard

- Dashboard 中的 Alert Widget 显示当前组织/设施的活跃告警
- Facility View 下 Alert Events 列表可查看设施级告警

---

## Dashboard 支持电能质量（Dashboard_PowerQuality）

- Dashboard 可添加 Power Quality 相关 Widget
- 展示 THD、PF 等电能质量指标趋势

---

## Dashboard 支持 Gas 分析（Dashboard 支持 Gas 分析）

- 第三阶段新增 Gas 类型支持
- Dashboard Gas Widget 展示天然气消耗趋势

---

## Energy Intensity Widget（Dashboard_Energy intensity）

- 展示单位面积能耗强度（kWh/m²）
- 可配置面积参数

---

## Energy Generation Widget（Energy Generation widget）

- 展示能源发电量数据（光伏/储能等）
- 第三阶段新增

---

## Visualization（仪表盘可视化，第三阶段）

- 仪表盘支持多种可视化图表类型
- 支持自定义颜色和主题
- 支持数据导出

---

## 权限控制

| 角色 | 访问权限 |
|------|---------|
| Normal 用户 | 可访问 Dashboard（有权限设施范围内） |
| Org Admin | 可配置 Dashboard 布局和 Widget |
| Tenant 用户 | 仅 Facility View Dashboard |

> Normal 用户访问 Dashboard 和 Data Export 权限（第三阶段专项测试）

---

## 测试要点

| 测试项 | 验证点 |
|--------|--------|
| kVA Widget 添加 | Today/15min、1h 下 kVA 数据展示正确 |
| kVA 状态记忆 | 取消后重开配置保持 |
| Gas 过滤 | 切换 Gas/Water 过滤器后 Widget 数据正确 |
| 自动刷新 | 数据在配置间隔后自动更新 |
| Sankey 图 | 能流图渲染正确，各节点流量准确 |
| Alert Widget | 告警数量与实际一致 |
