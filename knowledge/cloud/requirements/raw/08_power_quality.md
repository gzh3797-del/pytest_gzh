# Power Quality 电能质量

**路由**：`#/powerQuality`  
**状态**：BETA 功能

## 概述

Power Quality 模块用于分析电网电能质量指标，帮助用户识别谐波、电压波动、功率因数等问题。依赖 Acuvim II / Acuvim III 等支持电能质量监测的电表。

## 页面布局

- **左侧面板**：设施/设备树状选择器（支持关键字搜索）
- **右侧主体**：选择设备后显示电能质量分析图表

> 初始状态提示：**Please select a Device to proceed**

---

## 电能质量指标

| 指标 | 英文 | 说明 |
|------|------|------|
| 谐波总失真率 | THD | Total Harmonic Distortion |
| 电压不平衡 | Voltage Unbalance | 三相电压偏差 |
| 功率因数 | Power Factor | PF |
| 电压突变/闪变 | Sag/Swell/Flicker | 短时电压偏差 |
| 频率偏差 | Frequency | 频率稳定性 |

---

## Power Quality Meter Point Tier（订阅管理）

### 背景

第四阶段新增专项订阅 Tier，用于针对电能质量数据单独计费。

### 订阅配置（Admin / SuperAdmin）

| 角色 | 操作 |
|------|------|
| **SuperAdmin** | 在 Subscription 中新建 Power Quality Tier，配置 Number 和 Rate |
| **Admin** | 为组织分配 Power Quality Tier 额度 |

**字段：**
- **Power Quality Meter Point Tier**：新增的订阅 Tier 类型
- **Number of Subscription**：订阅数量（历史默认 0）
- **Rate**：费率（历史默认 0）

**历史订阅兼容：** 历史已有订阅的 Power Quality Tier 字段默认展示 0，保持向后兼容。

### Zoho 验证

新建 Power Quality Tier 后：
1. 提交成功
2. Zoho 订阅邮件正常接收
3. 账单中 Power Quality Addons 正常计费

---

## Power Quality 类型 Meter Point 支持 Billing

- PQ 类型的 Meter Point 可参与 Billing 计费
- 与普通电力 Meter Point 计费逻辑一致
- Account 配置中可绑定 PQ Meter Point

---

## 历史设备兼容（存量 Acuvim II/III）

- 已有的 Acuvim II/III 物理设备的 Meter Point 自动兼容 Power Quality 功能
- 不需要重新创建设备
- 存量 Meter Point 在订阅激活后即可使用 PQ 功能

---

## 使用场景

1. **设备故障诊断**：谐波异常可能导致设备损坏
2. **电力合规**：确认供电质量符合 IEC/IEEE 标准
3. **功率因数优化**：降低无功功率，减少罚款
4. **敏感设备保护**：识别影响精密设备的电能质量问题

---

## 订阅要求

| 功能 | 最低 Plan |
|------|-----------|
| Power Quality 基础功能 | AcuPQ Plan |
| Power Quality Billing | 需配置 Power Quality Meter Point Tier |
