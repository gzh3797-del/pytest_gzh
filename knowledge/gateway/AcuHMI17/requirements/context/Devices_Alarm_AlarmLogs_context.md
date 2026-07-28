# Devices / Alarm / Alarm Logs — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/alarm/alarmLogs` |
| 路由名 | `alarmLogs` |
| 面包屑 | Devices / Alarm / Alarm Logs |
| 上下文 | Devices 侧 |

## 2. 页面用途

查询全部历史报警记录（含已确认/未确认），支持过滤、导出、清空。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| 二级 tab | menuitem | `getByRole('menuitem',{name:'Alarm Logs'})` | 切换（Unacknowledged 带数量角标） |
| Start/End Date | date combobox | `getByRole('combobox',{name:'Start Date'})` / `{name:'End Date'}` | Interval 过滤 |
| Serial Number | combobox | `getByRole('combobox',{name:'Serial Number'})` | 按设备过滤 |
| Monitor ID | textbox | `getByRole('textbox',{name:'Monitor ID'})` | 按监控 ID 过滤 |
| Search | button | `getByRole('button',{name:'Search'})` | 查询 |
| Reset | button | `getByRole('button',{name:'Reset'})` | 重置 |
| 报警记录表 | table | 见下列 | 分页 |
| Clear Logs | button | `getByRole('button',{name:'Clear Logs'})` | 清空日志（危险，二次确认） |

**表列**：Timestamp / Monitor ID / Device Name / Serial Number / Monitor Label / Parameter / Status / Min / Max / Value / Trigger DO Device / DO / Trigger RO Device / RO / Reason / Ack Status

- Reason：UNDERFLOW / OVERFLOW。Ack Status：Acknowledged / Unacknowledge。Min/Max/Value 为阈值与触发值。

## 4. 自动化测试要点

- 过滤（日期 + 序列号 + Monitor ID）+ Search/Reset；分页。
- 断言阈值(Min/Max)与触发值(Value)、Reason 与 Ack Status。
- **Clear Logs 破坏性**，仅验证二次确认。
- 与 activeAlarm 联动（确认后 Ack Status 变化）。

## 5. 机器可解析摘要

```json
{
  "route": "/alarm/alarmLogs",
  "name": "alarmLogs",
  "title": "Alarm Logs",
  "context_side": "devices",
  "filters": ["Start Date","End Date","Serial Number","Monitor ID"],
  "table_columns": ["Timestamp","Monitor ID","Device Name","Serial Number","Monitor Label","Parameter","Status","Min","Max","Value","Trigger DO Device","DO","Trigger RO Device","RO","Reason","Ack Status"],
  "buttons": ["Search","Reset","Clear Logs(destructive)"],
  "paginated": true
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-17 联机实测）

> 对应测试目录：`projects/RPP/tests/Alarm/`（真机 AcuHMI-1-7，192.168.3.71）。

- **Ack Status 列随 System Settings → Alarm Notification 的 Alarm Acknowledgement Enable 开关显隐**：Enable 时显示该列，Disable 时隐藏；值为 Acknowledged / Unacknowledge。
- Status 列为图标，无文本（不能按文本断言）。
- 告警日志跨轮询周期累积，断言"本轮新告警/新确认"必须按 Ack Status 状态计数，不能只看某 Monitor Label 是否存在（历史记录会一直留存）。
- 全局（Unacknowledged Alarms / Ack All Alarms）确认操作后，设备详情页 Alarm Logs 的 Ack Status 会同步变为 Acknowledged。

### 过滤区（Interval/Serial Number/Monitor ID）实测情报（2026-07-17）

> 来源：`projects/RPP/tests/Alarm/helpers_alarm.py` 的 `set_interval_filter` / `select_serial_filter` / `fill_monitor_id_filter` 等联机实测。

- **Interval 是 `el-date-editor--datetimerange`**（双输入框 `input.el-range-input`，placeholder `Start Date`/`End Date`）。
  - ⚠️ 高危坑：**直接向输入框 `fill` 文本不会同步组件 v-model**——输入框显示值看起来正常，但面板一关闭（Escape / 点击 Search）即回滚为空，Search 会按空条件执行且不报错，产生"检索通过"的假象。
  - 正确做法：点输入框打开面板 → 在左月历点选起始日期 cell（`.el-picker-panel__content.is-left td.available`，当天 cell 带 `.today`；右月历 `.is-right` 为下月）→ 再点第二个日期 cell → 点 footer 最后一个按钮完成确认（`.el-picker-panel__footer button`，footer 按钮为 Clear/OK；OK 仅在面板内点选日期后才 enabled，纯文本输入不会激活）。
  - 确认后输入框值格式如 `2026-07-17 12:00 AM`（12 小时制）。
  - Escape 会把未确认的输入整体回滚清空。
  - 面板展开时会遮挡 Search 按钮（拦截元素是面板内 placeholder=`Start Time` 的输入框），需先完成/取消面板选择再点 Search。
- Serial Number 过滤是 `el-select`（占位 `--Select Serial Number--`，页面唯一可见 select，可用 `.el-select:visible` 定位），选项为各下挂设备 SN。
- Monitor ID 过滤是文本框 `[placeholder='Enter Monitor ID']`。
- Reset 后：双日期输入框、Monitor ID 均清空，Serial Number 下拉恢复占位文本。
- Clear Logs 有二次确认弹窗，确认后列表立即清空（表行 `tr.el-table__row` 变为 0 条）；破坏性操作，自动化用例应放在批次最后执行。
- 表格行 Timestamp 为 24 小时制 `YYYY-MM-DD HH:MM:SS`（与 Interval 输入框的 12 小时制不同）；检索类断言建议遍历当前页各行对应列，逐行核对是否全部满足过滤条件（日志跨轮询周期累积，非仅新增记录）。
