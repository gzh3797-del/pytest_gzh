# Devices / Physical Devices / 设备详情 / Data Log — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/physicalDevices/deviceDetails/<id>/3:1?deviceModel=...`（同级 Trend Log 为 `/3:2`） |
| 面包屑 | Devices / Physical Devices / 设备详情 / Logs / Data Log |
| 上下文 | Devices 侧，动态详情页 |

## 2. 页面用途

针对某台已接入设备，按数据源+时间区间+采样间隔导出历史 Data Log（CSV，gzip 压缩），或清空该设备的日志数据。

## 3. 入口路径

Physical Devices 列表 → 点设备行首格进入详情 → 左侧菜单 'Logs'（`el-sub-menu`，需先点击展开）→ menuitem 'Data Log'（同级 'Trend Log'）。

## 4. 交互元素清单

| 元素 | 类型 | 定位策略 | 说明 |
|------|------|---------|------|
| 左导航 Logs | el-sub-menu | 需先点击展开才能看到 'Data Log'/'Trend Log' | 折叠子菜单 |
| 数据源 | el-select | — | 当前实测仅 'Current' 一个选项 |
| 间隔 | el-select | — | 选项：1 minute / 5 minutes / 10 minutes / 15 minutes / 30 minutes / 1 hour / 6 hours / 12 hours / 1 day / 7 days / 1 month |
| 日期区间 | el-date-editor--daterange | `input.el-range-input`（两个，起止各一个） | 默认区间为「昨天~今天」 |
| Download | button | — | 触发导出下载 |
| Clear Logs | button | — | 破坏性操作，有二次确认弹窗 |

页面无数据预览表，仅提供筛选条件 + 下载/清空两个动作。

## 5. 自动化测试要点（2026-07-17 联机实测，来源：`projects/RPP/tests/Datalog/`，真机 AcuHMI-1-7 192.168.3.71）

- **日期面板防护式禁选**：日期选择面板只允许点选**有 datalog 数据的日期**，无数据日（含所有未来日期）对应 `td` 均带 `disabled`；只有 `td.available` 的日期可点。即产品用禁选替代"选中无数据区间"的容错分支，自动化断言应基于此（不要假设可选任意日期后再期待报错）。
- **下载接口**：`/api/log/dataLog/export`，返回体为 gzip 压缩的 CSV；文件名形如 `AHI260110002_MAA551000027_dataLog_2026-07-16_2026-07-17.csv.gz`（`<网关SN>_<设备SN>_dataLog_<起始日期>_<结束日期>.csv.gz`）。
- 解压后 CSV 首列为 `TimeTag`，值为 ISO 格式时间戳（如 `2026-07-16T15:29:00+0800`）；文件带 BOM，需用 `utf-8-sig` 读取。
- **设计约束**（2026-07-17 需求方澄清，非缺陷）：Log Interval 必须小于设备现有数据的时间跨度；所选间隔档位大于该跨度时，点 Download 后 `/api/log/dataLog/export` 返回 **HTTP 500**，前端**无任何提示，静默失败**（无 toast、无报错弹窗）——此为设计行为，任何档位（不限于 7 days/1 month）跨度不足均会触发；前端无提示属体验可改进点，可提需求但非缺陷。自动化对策：先用 `1 minute` 档下载探明当前数据跨度（`projects/RPP/tests/Datalog/helpers_datalog.py` 的 `current_data_span_seconds`），只测跨度足够的档位，跨度不足的档位 skip；Clear Logs 执行后数据归零，之后数小时内大间隔档位 skip 属预期。对应自动化 `projects/RPP/tests/Datalog/`（跨度不足档位动态 skip）。
- Clear Logs 为破坏性操作，自动化默认不实际执行，仅验证二次确认弹窗存在。

## 6. 机器可解析摘要

```json
{
  "route_pattern": "/physicalDevices/deviceDetails/<id>/3:1",
  "title": "Data Log",
  "context_side": "devices",
  "entry": "Physical Devices 列表 → 设备详情 → 左侧 el-sub-menu 'Logs' 展开 → 'Data Log'（同级 'Trend Log' 为 /3:2）",
  "controls": ["数据源 el-select (仅 Current)", "间隔 el-select", "日期区间 el-date-editor--daterange", "Download button", "Clear Logs button (destructive, confirm dialog)"],
  "interval_options": ["1 minute", "5 minutes", "10 minutes", "15 minutes", "30 minutes", "1 hour", "6 hours", "12 hours", "1 day", "7 days", "1 month"],
  "date_picker_constraint": "仅 td.available（有数据日）可选，其余含未来日期均 disabled",
  "export_api": "/api/log/dataLog/export",
  "export_format": "gzip CSV",
  "export_filename_pattern": "<gatewaySN>_<deviceSN>_dataLog_<startDate>_<endDate>.csv.gz",
  "csv_first_column": "TimeTag (ISO timestamp, e.g. 2026-07-16T15:29:00+0800)",
  "csv_encoding": "utf-8-sig (has BOM)",
  "design_constraint": {
    "condition": "所选 Log Interval 档位 >= 设备现有数据时间跨度（任意档位，不限 7 days/1 month）",
    "symptom": "导出接口 HTTP 500，前端无提示静默失败（设计行为，非缺陷；前端无提示属体验可改进点）",
    "date_clarified": "2026-07-17",
    "automation_ref": "projects/RPP/tests/Datalog/（用 helpers_datalog.current_data_span_seconds 探明跨度，跨度不足档位动态 skip）"
  }
}
```
