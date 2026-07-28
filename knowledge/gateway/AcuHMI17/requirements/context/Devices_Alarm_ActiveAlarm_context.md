# Devices / Alarm / Unacknowledged Alarms (Active Alarm) — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/alarm/activeAlarm` |
| 路由名 | `activeAlarm` |
| 面包屑 | Devices / Alarm / Unacknowledged Alarms |
| 上下文 | Devices 侧 |

## 2. 页面用途

查看并确认(Acknowledge)当前未确认的报警。Alarm 二级 tab：Unacknowledged Alarms(带数量角标) / Alarm Logs。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| 二级 tab | menuitem | `getByRole('menuitem',{name:/Unacknowledged Alarms/})` / `{name:'Alarm Logs'}` | 切换；未确认数量角标（示例 3） |
| Start/End Date | date combobox | `getByRole('combobox',{name:'Start Date'})` / `{name:'End Date'}` | Interval 过滤 |
| Serial Number | combobox | `getByRole('combobox',{name:'Serial Number'})` | 按设备序列号过滤 |
| Search | button | `getByRole('button',{name:'Search'})` | 查询 |
| Reset | button | `getByRole('button',{name:'Reset'})` | 重置 |
| Ack All Alarms | button | `getByRole('button',{name:'Ack All Alarms'})` | 一键确认全部 |
| 报警表 | table | 列: Timestamp/Device Name/Serial Number/Monitor Label/Parameter/Status/Reason/Action | — |
| 行 Acknowledge | button | `getByRole('row',{name:/.../}).getByRole('button',{name:'Acknowledge'})` | 确认单条 |

## 4. 表列语义

Timestamp / Device Name / Serial Number / Monitor Label / Parameter(如 "Phase A Line-to-Neutral Voltage") / Status(图标) / Reason(UNDERFLOW/OVERFLOW…) / Action(Acknowledge)

## 5. 自动化测试要点

- 过滤（日期+序列号）+ Search/Reset；断言结果。
- 单条 Acknowledge → 该行移除、角标数量减少；Ack All Alarms 清空未确认。
- 确认后记录进入 Alarm Logs（联动）。

## 6. 机器可解析摘要

```json
{
  "route": "/alarm/activeAlarm",
  "name": "activeAlarm",
  "title": "Unacknowledged Alarms",
  "context_side": "devices",
  "filters": ["Start Date","End Date","Serial Number"],
  "table_columns": ["Timestamp","Device Name","Serial Number","Monitor Label","Parameter","Status","Reason","Action"],
  "reason_enum": ["UNDERFLOW","OVERFLOW"],
  "buttons": ["Search","Reset","Ack All Alarms","Acknowledge(per row)"],
  "sub_tabs": ["Unacknowledged Alarms","Alarm Logs"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-17 联机实测）

> 对应测试目录：`projects/RPP/tests/Alarm/`（真机 AcuHMI-1-7，192.168.3.71）。

- 二级 tab 为 `el-menu--horizontal c_top_navbar` 下的 menuitem：'Unacknowledged Alarms'（带未确认数量角标，`inner_text` 形如 `"Unacknowledged Alarms\n1"`，数字用正则 `(\d+)\s*$` 提取）和 'Alarm Logs'；左导航 "Alarm" 项本身**无**数量角标（角标只出现在二级 tab 上）。
- **Alarm Acknowledgement Enable=Disable 时 'Unacknowledged Alarms' tab 整体隐藏**（页面只剩 Alarm Logs），等价于"无确认入口"——自动化涉及 Ack 相关断言前须先确认该开关状态。
- 表实测确认为 8 列：Timestamp / Device Name / Serial Number / Monitor Label / Parameter / Status / Reason / Action；行内 Acknowledge 按钮定位 `row.get_by_role("button", name="Acknowledge")`；页面另有全局 "Ack All Alarms" 按钮。
- 未确认总数可从 `.el-pagination__total` 读取（无需仅靠角标数字）。
- **读表头必须限定 `.el-table__header-wrapper th`**——页面常驻一个隐藏的日期面板（`el-picker-panel`），其星期表头会污染 `page.locator("th")` 的全局匹配；同理，确认弹窗的按钮匹配必须过滤 `is_visible()`（隐藏日期面板里也有同名 OK 按钮）。
