# Devices / Physical Devices / 设备详情 / Trend Log — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/physicalDevices/deviceDetails/<id>/3:2?deviceModel=...`（同级 Data Log 为 `/3:1`） |
| 面包屑 | Devices / Physical Devices / 设备详情 / Logs / Trend Log |
| 上下文 | Devices 侧，动态详情页 |

## 2. 页面用途

RPP 需求点（`Function_RPP_025`）要求的设备趋势日志功能，含 Realtime Log / Energy Log / Management 三个子项。

## 3. 入口路径

Physical Devices 列表 → 点设备行首格进入详情 → 左侧菜单 'Logs'（`el-sub-menu`，需先点击展开）→ 'Trend Log'（`el-sub-menu`，含子项 Realtime Log / Energy Log / Management）。

## 4. 交互元素清单

| 元素 | 类型 | 说明 |
|------|------|------|
| 左导航 Trend Log | el-sub-menu | 展开后含 Realtime Log / Energy Log / Management 三个子项 |
| Realtime Log | menuitem | 见「5. 实测情报」——当前点击无跳转 |
| Energy Log | menuitem | 同上 |
| Management | menuitem | 同上 |

## 5. 自动化测试要点（2026-07-17 联机实测，来源：`projects/RPP/tests/TrendLog/`，真机 AcuHMI-1-7 192.168.3.71）

- ⚠️ **当前固件未实装 Trend Log 功能**：在两台不同设备（AcuvimIIW、AcuRev-4110-mV）上实测结果一致：
  - 点击 Realtime Log / Energy Log / Management 三个子项均**不发生路由跳转**，页面停留在 Metering（默认页）。
  - 直接改 URL hash 跳转到 `/3:2` 可以进入该路由，但**页面渲染为空白**——无任何控件、图表或报错提示。
- Trend Log 为 RPP 需求文档中的功能点（`Function_RPP_025`），当前对应自动化用例为占位状态：`projects/RPP/tests/TrendLog/` 下 12 条用例（`test_Function_RPP_025_001_case1` ~ `case12`）全部标记 `skip`。
- **待办**：RPP 真机就绪后需现场重新探查该页面实际交互结构（子项路由、控件、数据展示方式），补齐本文档「交互元素清单」与选择器细节后再将占位用例改为实际断言。

## 6. 机器可解析摘要

```json
{
  "route_pattern": "/physicalDevices/deviceDetails/<id>/3:2",
  "title": "Trend Log",
  "context_side": "devices",
  "entry": "Physical Devices 列表 → 设备详情 → 左侧 el-sub-menu 'Logs' 展开 → el-sub-menu 'Trend Log' → Realtime Log / Energy Log / Management",
  "sub_items": ["Realtime Log", "Energy Log", "Management"],
  "implementation_status": "未实装（2026-07-17 实测，AcuvimIIW 与 AcuRev-4110-mV 一致）",
  "observed_behavior": {
    "click_sub_item": "无路由跳转，停留 Metering 默认页",
    "direct_url_nav": "可进入 /3:2 路由，但页面空白（无控件/图表/报错）"
  },
  "requirement_ref": "Function_RPP_025",
  "automation_ref": "projects/RPP/tests/TrendLog/ (12 条 test_Function_RPP_025_001_case1~case12，全部 skip 占位)",
  "todo": "RPP 真机就绪后现场探查，补实际选择器与断言"
}
```
