# Analysis 能耗分析

**路由**：`#/analysis/energy`

## 概述

Analysis 模块是 AcuCloud 的核心数据分析功能，提供多维度的能耗数据可视化与分析工具。支持电力、水务、天然气三种能源类型。

## 子页面（Tab）

| Tab | 说明 |
|-----|------|
| **Energy** | 能耗趋势分析（柱图/折线图） |
| **Accumulative** | 累计能耗分析 |
| **Realtime** | 实时能耗及预测基线 |
| **Heatmap** | 热力图（时段×日期的能耗密度） |
| **Schedule** | 排程分析（工作/非工作时段对比） |
| **Annotations** | 标注（在时间轴添加事件注释） |
| **M&V** | 测量与验证（Measurement & Verification） |

---

## 通用控件

### 左侧计量点树

- 按设施/计量点层级展示
- 支持关键字搜索
- 可按能源类型过滤（Electricity / Water / Gas）
- **多选限制：最多同时选 5 个 Meter Point**（第二阶段从 10 个改为 5 个）
- 选择 TOTAL 类型计量点时，右侧渲染对应分析菜单

### Time Frame（时间范围）

| 快捷选项 | 说明 |
|---------|------|
| Today | 当天数据 |
| Yesterday | 昨天数据 |
| Last 7 Days | 近 7 天 |
| Last 30 Days | 近 30 天 |
| Custom | 自定义起止时间 |

切换 Time Frame 后：
- 右侧时间范围更新
- Time Interval 自动适配
- 自定义模式下时间范围可调整

### Time Interval（时间间隔聚合粒度）

| 间隔 | 适用场景 |
|------|---------|
| 5min | 今天 / 昨天 |
| 15min | 今天 / 昨天 / 近 7 天 |
| 30min | 今天 / 昨天 / 近 7 天（第二阶段新增） |
| 1hour | 近 7 天 / 近 30 天 |
| 1day | 近 30 天 / 自定义月维度 |
| 1week | 自定义周维度 |
| 1month | 年视图 |

月 Time Interval 新增 15min / 30min 选项（第二阶段优化）。

---

## Energy（能耗趋势）

- X轴：时间（按 Time Interval 聚合）
- Y轴：能耗值（kWh / kW / kVA 等）
- 支持多计量点叠加对比
- 图表类型：柱图（Bar）/ 折线图（Line）

**数据统计（含最大值/最小值/平均值/总计）：**
- Facility TOTAL 今天每 5min 聚合的输入有功能耗
- Facility TOTAL 昨天每 15min/30min 聚合
- Facility TOTAL 近 7 天每 1h 聚合
- Facility TOTAL 近 30 天每 1day 聚合

### kVA 支持（第二阶段）

- Energy Analysis 支持同时显示 kWh 和 kVA
- "Show kVA" 开关默认勾选
- 取消勾选后，状态记忆（卡片配置记忆功能）
- 图表统一使用折线图展示 kVA 数据

---

## Accumulative（累计能耗）

- 显示累积能耗曲线
- 可与其他 Time Frame 对比（Compare to 功能）

### Compare to（对比功能，ACU-837）

- 选择对比时间段（上月/上年/自定义）
- 右侧图表显示当前期 vs 对比期的叠加曲线
- 适用于 Energy、Accumulative、Realtime 视图

---

## Realtime（实时分析与预测基线，ACU-826）

### 功能概述

Realtime 页签提供实时能耗数据展示及 **Real-time Baseline（实时基线预测）** 功能。

### 实时基线预测

**触发条件：**
- 选中 Physical 类型的 Meter Point 且 Active 状态
- Time Interval ≤ 1小时（1h / 30min / 15min / 5min）时触发预测
- Time Interval > 1小时时不执行预测

**预测模型：** OLS（Ordinary Least Squares，普通最小二乘法）+ β 系数

**图表显示：**
- 实际功率曲线（Actual Active Power）
- 预测曲线（Predicted Active Power）
- 切换 Time Interval > 1h 后预测曲线实时消失

**自动优化（Th/Tc Auto Optimize）：**
- 冷热阈值（Th/Tc）支持自动优化
- 模型训练结果实时更新

### Data View（数据查看）

- 表格展示实时数据
- 支持下载（Download）
- 图表截图保存（Save Image）

### Disable Note / Hide Note

- 标注显示控制

### Restore

- 恢复到默认视图状态

### 公式刷新

- Facility 下有公式（Calculated Meter）时，MP 数据可参与预测
- 新建 MP / 历史 MP 均可预测

### 数据来源与依赖

- 单个 MP 无 Phase 数据时，system 级数据展示
- 多个 MP 选中时，支持 phase 和 system 级别数据分别展示

---

## Heatmap（热力图）

- X轴：时间（小时/天）
- Y轴：日期
- 颜色深浅表示能耗高低
- 用于识别时段型能耗规律

---

## Schedule（排程分析）

- 对比工作时间 vs 非工作时间的能耗分布
- 识别非工作时段的异常能耗（After Hours Alarm）
- 支持自定义工作时间段

### Schedule 支持 TOU Rate（ACU-844）

- Schedule 分析可关联 TOU 分时段费率
- 直接在 Schedule 视图中显示峰/平/谷时段的费用估算
- Rate 配置方式：Manual Input 或 Select from Rate List

---

## Annotations（标注）

- 在能耗曲线上的特定时间点添加事件注释
- 用于记录设备检修、生产变化等影响能耗的事件
- 支持新增、编辑、删除标注

---

## M&V（测量与验证）

M&V（Measurement & Verification）用于能效项目的效益验证。

### 访问入口

- Analysis 左侧树：选择 Total 类型计量点后，右侧渲染 M&V 菜单
- 多选时可包含 Total，切换不同 Total 时 M&V 数据对应更新

### M&V 项目流程

| 步骤 | 名称 | 说明 |
|------|------|------|
| 1 | **Project Definition** | 定义项目基本信息（名称、目标节能量） |
| 2 | **ECM（Energy Conservation Measure）** | 能效措施定义（改造项目） |
| 3 | **Baseline & Reporting Period** | 基准期和报告期的数据范围选择 |
| 5 | **Non-Routine Adjustments** | 非例行调整（如生产量变化修正） |
| 6 | **Model** | 统计模型配置（回归模型参数） |

### 页面状态

| 状态 | 说明 |
|------|------|
| 未选 Meter Point | 提示 "Please select a Meter Point to proceed" |
| 选中 Total | 渲染 M&V 菜单，可进入 M&V 详情 |
| 多选含 Total | 渲染 M&V 菜单，M&V 数据根据当前 Total 刷新 |
| 多选未含 Total | 无 M&V 菜单入口 |

---

## Live Consumption（实时用电，ACU-838）

- 实时展示当前用电量（kW/kWh）
- 区别于 Realtime 的历史趋势，Live Consumption 专注于当前时刻
- 支持 Facility 下多 MP 聚合

---

## Analysis 权限

| 权限类型 | 说明 |
|---------|------|
| 分析权限 | 有 Installation 权限的用户可访问 |
| 多选能源类型 | 可在 Electricity / Water / Gas 间切换 |
| Dashboard 过滤 | Dashboard 中可按水/气类型过滤 |

---

## Analysis 数据校验

| 数据类型 | 校验方法 |
|---------|---------|
| 能耗数据 | InfluxDB 查询 vs 图表接口数据对比 |
| M&V 节能量 | 基准期能耗 - 报告期能耗 = 节能量 |
| 数据恢复 | 删除后恢复数据完整性 |
