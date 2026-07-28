# Devices / Physical Devices (List) — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/physicalDevices` |
| 路由名 | `physicalDevices` |
| 面包屑 | Devices / Physical Devices |
| 上下文 | Devices 侧 |
| 关联路由 | addDevice / searchDevice / deviceDetails/:deviceKey/:currentKey |

> ⚠️ 直接 goto `/#/physicalDevices` 可能被重定向到 dashboard；正确入口是 Dashboard 左导航点击 "Physical Devices"。

## 2. 页面用途

管理已接入的物理设备：列表、新增、轮询设置、下载列表、进入设备详情。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Add Device | button | `getByRole('button',{name:'Add Device'})` | 进入 addDevice |
| Port(轮询) | combobox | 顶部 combobox | 选择 COM 口（COM 1…） |
| Poll interval | textbox | `getByRole('textbox',{name:'---Enter poll interval---'})` | 该口轮询间隔（seconds） |
| Save(轮询) | button | 顶部 `getByRole('button',{name:'Save'})` | 保存该口轮询设置 |
| Download List | button | `getByRole('button',{name:'Download List'})` | 导出设备列表 |
| 设备表 | table | 见下列 | 列头可排序 |
| 行 Action | button | 行末图标 | 进入 deviceDetails/编辑 |
| 分页 | button | `getByRole('button',{name:'Go to next page'})` | 单页 disabled |

**表列**：Device Name / Interface / Protocol / Model / Serial Number / Poll Interval / Status(ON/OFF) / Active Alarms / Action

## 4. 自动化测试要点

- Add Device 入口；行 Action 进入详情。
- 轮询设置（选口+间隔+Save）；Download List 下载。
- Status(ON/OFF)、Active Alarms 数值断言；排序。

## 5. 机器可解析摘要

```json
{
  "route": "/physicalDevices",
  "name": "physicalDevices",
  "title": "Physical Devices",
  "context_side": "devices",
  "table_columns": ["Device Name","Interface","Protocol","Model","Serial Number","Poll Interval","Status","Active Alarms","Action"],
  "buttons": ["Add Device","Save(poll)","Download List"],
  "related_routes": ["/physicalDevices/addDevice","/physicalDevices/searchDevice","/physicalDevices/deviceDetails/:deviceKey/:currentKey"],
  "entry_note": "navigate via left-nav click, not direct goto"
}
```
