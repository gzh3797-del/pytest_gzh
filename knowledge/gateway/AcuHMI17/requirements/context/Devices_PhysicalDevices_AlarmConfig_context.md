# Devices / Physical Devices / 设备详情 / Alarm Config — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/physicalDevices/deviceDetails/<id>/2:1?deviceModel=...`（列表页/子菜单参数随所选设备变化） |
| 面包屑 | Devices / Physical Devices / 设备详情 / Alarm / Alarm Config |
| 上下文 | Devices 侧，动态详情页 |

## 2. 页面用途

针对某台已接入设备，配置该设备的越限告警规则（Monitor：参数 + Min/Max 阈值），并可联动触发 DO/RO 输出。

## 3. 入口路径

Physical Devices 列表 → 点设备行首格进入详情 → 左侧菜单 'Alarm'（`el-sub-menu`，需先点击展开）→ 二级项 'Alarm Config' / 'Alarm Logs'。

## 4. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| 左导航 Alarm | el-sub-menu | 需先点击展开才能看到 'Alarm Config'/'Alarm Logs' | 折叠子菜单 |
| Add Alarm | button | `page.get_by_role("button", name="Add Alarm")` | 跳转到内联表单页（非弹窗），URL 追加 `?type=add` |
| 规则表 | table | 见下「表列」 | — |
| 行内编辑 | button | `button.el-button--primary`（无文本） | 跳转同表单页 `?type=edit` |
| 行内删除 | button | `button.el-button--danger`（无文本） | 二次确认弹窗后删除 |

### 表列

Monitor ID | Label | Parameter | Min | Max | Value | Status | Trigger DO Device | DO | Trigger RO Device | RO | Action

- **Status 列是图标，无文本**：告警 ON = `el-icon warning`，正常 = `el-icon success`；设备未被轮询到时 Value/Status 显示 `-`。判定方式：取 status 单元格 `innerHTML`，含 `"warning"` 即判定为触发态。
- Value 为该 Monitor 当前实测值；Min/Max 为配置阈值。

## 5. Add / Edit 表单字段（内联页，非弹窗）

| 字段 | 类型 | 定位/占位 | 说明 |
|------|------|-----------|------|
| Device Status Alarm | switch | — | 是否作为设备状态类告警 |
| Label | textbox | placeholder=`Enter Label` | ≤40 字符 |
| Parameter | el-select | 占位 `---Enter Parameter---` | 待监控参数 |
| Min Value | textbox | placeholder=`Enter Min Value` | 范围 ±2147483647 |
| Max Value | textbox | placeholder=`Enter Max Value` | 范围 ±2147483647 |
| Trigger DO Device | el-select | — | 联动 DO 所属设备 |
| DO | el-select | — | 联动 DO 通道 |
| Trigger RO Device | el-select | — | 联动 RO 所属设备 |
| RO | el-select | — | 联动 RO 通道 |
| Cancel / Save | button | — | Save 提交，Cancel 放弃返回列表 |

## 6. 自动化测试要点（2026-07-17 联机实测，来源：`projects/RPP/tests/Alarm/`，真机 AcuHMI-1-7 192.168.3.71）

- **Add Alarm 是内联表单页而非弹窗**：断言时应等待 URL 变化（`?type=add`），不能按弹窗 dialog 定位。
- **触发告警的稳定做法**：对当前轮询中 Status 已知的设备，建一条必越限规则（例如 System Frequency：Min=70 / Max=90 → 实测值落在区间外触发 UNDERFLOW/OVERFLOW）；网关轮询周期约 60s，改配置后需等待至少一个轮询周期才能看到 Status 图标翻转。
- 规则**先删后建**可保证 OFF→ON 翻转、稳定产生"新告警"事件（避免因规则已存在导致状态不变、断言误判为无新告警）。
- Physical Devices 列表页的 "Active Alarms" 列计数会随本页规则的触发/清除同步变化，可作为跨页联动断言点。
- 行内编辑/删除按钮无文本，需用 CSS class（`el-button--primary` / `el-button--danger`）区分；删除有二次确认弹窗。

## 7. 机器可解析摘要

```json
{
  "route_pattern": "/physicalDevices/deviceDetails/<id>/2:1",
  "title": "Alarm Config",
  "context_side": "devices",
  "entry": "Physical Devices 列表 → 设备详情 → 左侧 el-sub-menu 'Alarm' 展开 → 'Alarm Config'",
  "table_columns": ["Monitor ID","Label","Parameter","Min","Max","Value","Status","Trigger DO Device","DO","Trigger RO Device","RO","Action"],
  "status_icon_map": {"ON": "el-icon warning", "normal": "el-icon success", "not_polled": "-"},
  "add_form_url_param": "?type=add",
  "edit_form_url_param": "?type=edit",
  "form_fields": ["Device Status Alarm","Label","Parameter","Min Value","Max Value","Trigger DO Device","DO","Trigger RO Device","RO"],
  "row_buttons": {"edit": "button.el-button--primary(no text)", "delete": "button.el-button--danger(no text, confirm dialog)"},
  "poll_interval_seconds": 60
}
```
