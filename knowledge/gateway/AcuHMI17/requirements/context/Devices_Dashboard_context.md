# Devices / Dashboard — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/dashboard` |
| 路由名 | `dashboard` |
| 面包屑 | Devices / Dashboard |
| 上下文 | Devices 侧（header 点 "Devices" 切换） |

## 2. 页面用途

设备侧首页概览：离线设备、报警设备汇总。Devices 侧左导航：Dashboard / Physical Devices / Virtual Devices / Web Devices / Alarm / Data Log。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| 左导航项 | listitem | 文本 "Physical Devices" 等 | 切换设备侧模块 |
| Offline Devices 表 | table | 列: Device Name/Interface/Protocol/Serial Number | 只读；分页 |
| Alarms 表 | table | 列: Device Name/Alarms/Interface/Protocol/Serial Number | 只读；分页 |
| 列头排序 | columnheader | `getByRole('columnheader',{name:'Device Name'})` | 排序 |
| 分页 | button | `getByRole('button',{name:'Go to next page'})` | 单页时 disabled |

- 页脚 "Up since <时间>" 显示运行起始时间。

## 4. 自动化测试要点

- 只读概览页；断言离线设备/报警表数据与排序、分页。
- 顶部 header 有 Logout / About / Devices / AcuHMI-1-7(切设置侧) 入口。

## 5. 机器可解析摘要

```json
{
  "route": "/dashboard",
  "name": "dashboard",
  "title": "Dashboard",
  "context_side": "devices",
  "tables": {
    "Offline Devices": ["Device Name","Interface","Protocol","Serial Number"],
    "Alarms": ["Device Name","Alarms","Interface","Protocol","Serial Number"]
  },
  "left_nav": ["Dashboard","Physical Devices","Virtual Devices","Web Devices","Alarm","Data Log"]
}
```
