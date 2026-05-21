# Gas Model 天然气模型

## 概述

Gas Model 是 AcuCloud 天然气计量的核心数据模型，负责天然气数据的采集、单位转换、存储和分析。

---

## 数据参数

### 核心测量字段

| 参数 | 单位 | 说明 |
|------|------|------|
| **TOTAL_VOLUME_m3** | m³ | 总体积（公制，立方米） |
| **TOTAL_VOLUME_cuft** | cuft | 总体积（英制，立方英尺） |
| **GAS_IMP_m3** | m³ | 进口气量（公制） |

### 单位转换规则

- 数据以 **m³** 为标准单位存储于 InfluxDB
- 若设备上报 `TOTAL_VOLUME_cuft`（立方英尺）：系统自动将 cuft 转换为 m³（1 cuft ≈ 0.0283168 m³）
- 转换后的 m³ 值存入 `GAS_IMP_m3`
- Meter Point 的单位最终为 **m³**

---

## 设备类型

### 非网关（Non-Gateway）Gas 设备

- 数据上报格式：**CSV**
- 上报字段：`TOTAL_VOLUME_m3`
- Meter Point 单位：**m³**
- 单次上报大小限制：**≤ 3MB**（超过 3M 上报失败）

**InfluxDB 校验：**
```
TOTAL_VOLUME_m3 和 GAS_IMP_m3 原始数据一致
diff 和 acc 计算正确
```

### 网关（Gateway）Gas 设备

- 数据上报格式：**JSON**
- 上报字段：`TOTAL_VOLUME_cuft`
- 系统自动转换为 m³ 存储
- Meter Point 单位：**m³**

**InfluxDB 校验：**
```
TOTAL_VOLUME_cuft 的值先转换为 m³（×0.0283168）
再与 GAS_IMP_m3 数据对比
diff 和 acc 计算正确
```

---

## 数据采集间隔

| 间隔 | 支持 |
|------|------|
| 5min | ✅ |
| 15min | ✅ |

---

## Analysis 支持

Gas Model 数据在 Analysis 模块中可用于：

| 分析类型 | 支持 |
|---------|------|
| Energy（能耗趋势） | ✅ |
| Accumulative（累计） | ✅ |
| Realtime（实时） | ✅ |
| Heatmap（热力图） | ✅ |
| Schedule（排程） | ✅ |
| Annotation（标注） | ✅ |

左侧树通过 Gas 类型过滤器选择天然气计量点。

---

## Dashboard 支持

- Dashboard 组件支持 Gas Analysis 图表
- Gas 类型 Meter Point 可在 Dashboard 中展示趋势

---

## 气表（Water + Gas）合并账单

- 第四阶段支持水电气合并账单（水电气合并账单功能）
- 同一账单中包含多种 Utility Type 的费用合并展示

---

## Gas-807 优化

第四阶段 Gas-807 测试项优化：
- 数据格式校验增强
- 超大文件处理（>3M 边界测试）
- Receiver 支持 .tar.gz 压缩包（第二阶段新增）

---

## 测试要点

| 测试项 | 验证点 |
|--------|--------|
| 非网关 CSV 上报 | TOTAL_VOLUME_m3 数据正确存入 InfluxDB |
| 非网关超限上报 | >3M 文件上报失败，返回错误 |
| 网关 JSON 上报 | cuft→m³ 转换精确 |
| Meter Point 单位 | 创建后单位显示为 m³ |
| diff/acc 计算 | 差值和累计值计算正确 |
| Gas Analysis | Energy/Heatmap/Schedule 图表数据准确 |
| Gas Billing | 基于 m³ 数据正确生成账单 |

---

## 与 Water Model 的区别

| 维度 | Gas | Water |
|------|-----|-------|
| 主单位 | m³ | m³ / Liter |
| 上报格式 | CSV（非网关）/ JSON（网关） | CSV / JSON |
| Billing 类型 | Gas Rate | Water Rate |
| Sewage Water | N/A | 支持污水处理费率 |
| VEE | 气体 VEE 算法 | 水 VEE 算法 |
| Analysis 格式转换 | cuft→m³ | 无额外转换 |
