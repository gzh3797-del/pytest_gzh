# Alerts 告警管理

## 概述

AcuCloud 的告警体系覆盖设备状态、能耗异常、账单超额等多个维度，通过 Email 和 SMS 双渠道通知用户。

---

## 告警类型

### 1. Offline Alert（设备离线告警）

- **触发条件：** 设备超过配置时间无数据上报
- **告警级别：** L1（高优先级）
- **通知方式：** Email / SMS

### 2. Parametric Alert（参数阈值告警）

- **触发条件：** 指定参数值超出上限或低于下限
- **配置内容：** 参数名、阈值（上限/下限）、持续时间
- **支持设备：** Physical / Gateway / Calculated

### 3. Flatline Alert（数据固定告警）

- **触发条件：** 指定参数值在一段时间内保持不变（通常意味着设备异常）
- **典型场景：** CT 接线故障导致功率长期为 0

### 4. Rate Alarm（费率告警，ACU-793）

- **触发条件：** 实时能耗费率超过设定阈值
- **配置内容：** 费率阈值（元/kWh 或类似）、时段
- **用途：** 监控高峰时段用电费率，触发需量管理

### 5. Meter Resealing Alert（水表密封告警）

- **触发条件：** 水表密封状态异常（防止计量作弊）
- **专用于：** Water 类型设备

### 6. After Hours Alarm（非工作时段告警）

- **触发条件：** 在非工作时段检测到超阈值的能耗
- **配置依赖：** 需先配置 Schedule（工作时间段）

### 7. Utility Bill Alerts（账单告警，ACU-813/818）

- **触发条件：** 账单金额超过预算/阈值
- **通知内容：** 账单超额通知邮件/短信

---

## 告警通知渠道

### Email 通知

- 配置邮件接收人列表
- 支持多个收件人
- 告警邮件内容包含：设备名、告警类型、触发时间、触发值

### SMS 短信通知（ACU-781）

| 功能 | 说明 |
|------|------|
| 短信接收配置 | 在用户设置中配置手机号 |
| 短信模板 | 系统预设短信内容格式 |
| 已启用状态 | `smsAlertEnable: true` |
| 测试验证 | 短信发送成功率、内容准确性 |
| Plan 要求 | Plus / AcuEMS 支持 SMS |

---

## Alert 在 Device 列表（Alert Status）

### Alarm Count 字段

- Device 列表新增 **Alarm Count** 列
- 显示规则：
  - `0`：绿色
  - `>0`：红色
- 排序：支持按 Alarm Count 升序/降序排列

### 设备详情 Alert 选项卡

- 设备详情页新增 **Alert** 选项卡
- 显示该设备的历史告警列表
- 无数据时页面正确展示（不报错）

### 水类型 Alert 管理

- Water 类型设备支持告警配置
- Meter Resealing Alert 专属水表

---

## Logs（告警日志）

**路由：** `#/logs/alertLogs`

### Alert Logs 页面

| 字段 | 说明 |
|------|------|
| Alarm Type | Offline / Parametric / Flatline / Rate |
| Device | 触发告警的设备 |
| Trigger Time | 触发时间 |
| Duration | 持续时长 |
| Status | Active / Resolved |
| Action | 确认 / 查看详情 |

**过滤器：**
- 按告警类型过滤
- 按设施过滤
- 按时间范围过滤

**排序和翻页：**
- 支持按触发时间排序
- 支持翻页，正确保持排序

### Alert Logs 更新（Alert Logs Page Update，第三阶段）

- **Alert Events 列表：** 查看所有告警事件的历史记录
- **排序搜索：** 支持按多字段排序和关键字搜索
- **翻页：** 支持分页浏览大量历史告警
- **Facility View 下：** Alert Events 列表只显示当前设施的告警

### 一条告警中包含多个告警

- 告警规则可配置多个触发条件
- 一条 Offline Alert 可包含多个离线设备
- 一条 Parametric Alert 可包含多个超阈值参数

---

## 告警规则配置（Installation → Alerts Tab）

### 创建告警规则

| 字段 | 说明 |
|------|------|
| Alert Type | 选择告警类型 |
| Device | 关联设备 |
| Parameter | 目标参数（Parametric 类型） |
| Threshold | 阈值（上限/下限） |
| Duration | 持续时长（触发前需满足条件的时间） |
| Notification | Email / SMS 接收人 |

### 编辑和删除告警规则

- 支持编辑现有规则
- 支持删除规则（告警规则删除后，历史日志保留）

---

## 测试要点

| 测试项 | 验证点 |
|--------|--------|
| Offline Alert 创建 | 规则创建后设备离线触发告警 |
| 告警日志准确性 | 日志内容与触发条件一致 |
| 数据过滤 | 按类型/设施/时间过滤后结果正确 |
| Rate Alarm 触发 | 费率超阈值后告警正常生成 |
| SMS 通知 | 短信成功发送，内容正确 |
| Email 通知 | 邮件成功发送，多收件人均收到 |
| Alarm Count 显示 | Device 列表 Alarm Count 颜色和数量正确 |
| Alert Logs 翻页 | 翻页后告警记录连续无遗漏 |
