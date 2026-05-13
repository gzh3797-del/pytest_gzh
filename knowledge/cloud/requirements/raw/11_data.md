# Data 数据管理

**路由**：`#/data/dataExports`

## 概述

Data 模块提供原始能耗数据的导入、导出、转发和编辑功能，是 AcuCloud 数据治理的核心模块。

## 子页面（Tab）

| Tab | 路由 | 说明 |
|-----|------|------|
| **Data Exports** | `/data/dataExports` | 按需导出历史原始数据（xlsx/zip） |
| **Data Forwards** | `/data/dataForwards` | 将数据实时转发到第三方系统 |
| **Data Import** | `/data/dataImport` | 从外部文件导入历史数据 |
| **Energy Exports** | `/data/energyExports` | 导出聚合能耗数据 |
| **Max Min Export** | `/data/maxMinExport` | 导出最大/最小值统计 |
| **Data Edit** | `/data/dataEdit` | 手动编辑/修正数据值 |
| **Manual VEE** | `/data/manualVee` | 手动数据验证、估算、编辑（VEE） |

---

## Data Exports（数据导出）

### 创建导出任务

| 字段 | 说明 |
|------|------|
| Devices | 选择设备（必填，不选时提示"请选择 Devices" ） |
| Parameters | 选择导出参数（必填，不选时提示"请选择 Parameters"） |
| Time Range | 起止时间 |
| Time Interval | 数据间隔：5m / 15m / 30m / 1h / 1d / 1w 等 |
| File Format | xlsx / zip |

**重要约束：**
- 导出数据仅保留 **30天**
- 文件中数据时区与设备时区对应
- `time interval` 选择 **week** 时，数据未对齐（已知问题）

### 导出任务状态

| 状态 | 说明 |
|------|------|
| PROCESSING | 正在生成 |
| SUCCESS | 完成，可下载 |
| FAILED:no data | 指定时间段内无数据 |

### 导出参数说明（部分）

| 参数代码 | 含义 |
|---------|------|
| I1 / I2 / I3 | A/B/C 相电流 |
| DI14 / DI12 | 数字输入 14 / 12 |
| AI | 模拟输入 |
| DMD_Pb | 需量（有功） |
| kW | 有功功率 |
| EP_IMP_kWh | 输入有功电量 |

---

## Max Min Export（最大最小值导出）

- 导出指定时间段内的最大值和最小值统计
- 第四阶段（MaxMinExport）新增或优化
- 适用于需量计费和峰值分析

---

## Data Edit（数据编辑，第四阶段）

**路由：** `/data/dataEdit`

### 功能说明

- 手动修正错误的历史数据点
- 支持单点修改和批量修改
- 修改后数据重新参与 VEE 处理流程

### 操作流程

1. 选择设备和参数
2. 选择时间范围查询数据
3. 点击具体数据点进行编辑
4. 保存修改（生成修改记录）

### Data Edit 测试要点

| 测试项 | 验证点 |
|--------|--------|
| 数据查询 | 按设备/参数/时间正确查询 |
| 单点编辑 | 修改值保存后 InfluxDB 数据同步 |
| 修改记录 | 修改历史可查询 |
| 权限 | 非 Admin 用户无编辑权限 |

---

## Data Forwards（数据转发）

将 AcuCloud 采集的实时数据推送到第三方平台。

### 支持协议

- MQTT
- HTTP
- FTP
- 第三方 API（ACU-797 新增 TimeInterval 支持）

### ACU-797 TimeInterval 支持

- 数据转发支持自定义 TimeInterval（第四阶段）
- API 侧同步更新（API_797）

---

## Data Import（数据导入）

从 CSV/Excel 文件批量导入历史计量数据，用于数据补录。

### 导入格式

- 第三方 API 数据
- 历史 CSV 上传

---

## Manual VEE（手动 VEE）

VEE = Validation, Estimation, Editing（验证、估算、编辑）

| 操作 | 说明 |
|------|------|
| **Validation** | 检查数据合理性（范围检验、趋势异常检测） |
| **Estimation** | 对缺失数据进行插值/估算（基于前后数据） |
| **Editing** | 直接修改异常数据值 |

### VEE 算法类型（第二阶段）

- 电力 VEE 算法
- 水 VEE 算法（第二阶段）
- 气体 VEE 算法（第三阶段）

### Calculated VEE（Calculated_vee，第四阶段）

- Calculated Meter 数据参与 VEE 流程
- 公式计算结果的数据质量验证

### 30min Receiver VEE

- 30 分钟间隔的 Receiver 数据 VEE 处理（第四阶段）

---

## 权限控制

| 角色 | 权限 |
|------|------|
| Normal 用户 | 可查看 Data Export，不可导入/编辑 |
| Org Admin | 全部操作权限 |
| Tenant 用户 | 无 Data 模块访问权限 |

> **Normal 用户访问权限（第三阶段专项测试）：** Normal 用户可访问 Dashboard 和 Data Export，但无法执行数据编辑操作。
