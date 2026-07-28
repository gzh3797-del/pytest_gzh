# Maintenance / Event Log — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/maintenance/eventLog` |
| 路由名 | `eventLog` |
| 面包屑 | AcuHMI-1-7 / Maintenance / Event Log |
| 顶级模块 | Maintenance |

## 2. 页面用途

查询/导出/清空系统事件日志（登录、配置更改等）。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Start Date | date combobox | `getByRole('combobox',{name:'Start Date'})` | Interval 起始 |
| End Date | date combobox | `getByRole('combobox',{name:'End Date'})` | Interval 结束 |
| Level | combobox | `getByRole('combobox',{name:'Level'})` | 日志级别过滤（--Select Level--；如 Info/Warning/Error） |
| Search | button | `getByRole('button',{name:'Search'})` | 按条件查询 |
| Reset | button | `getByRole('button',{name:'Reset'})` | 重置过滤 |
| 日志表 | table | 列: Timestamp/Level/source/Message | 只读 |
| 每页条数 | combobox | 分页区 combobox | 10/page |
| 分页 | button/listitem | `getByRole('button',{name:'Go to next page'})` | 多页（示例 248 页） |
| Export Logs | button | `getByRole('button',{name:'Export Logs'})` | 导出日志 |
| Clear Logs | button | `getByRole('button',{name:'Clear Logs'})` | 清空日志（危险，通常二次确认） |

## 4. 表列语义

Timestamp / Level(Info/…) / source(如 WebServer) / Message(如 "Login: user admin, From IP ..."、"Updated Modbus Parameters Mapping")

## 5. 自动化测试要点

- 日期区间 + Level 过滤 → Search，断言结果集；Reset 清空条件。
- 分页/每页条数验证。
- Export Logs 下载；**Clear Logs 破坏性**，仅验证二次确认。
- 可用作端到端断言源（其它页保存操作会在此产生 "Updated ..." 事件）。

## 6. 机器可解析摘要

```json
{
  "route": "/maintenance/eventLog",
  "name": "eventLog",
  "title": "Event Log",
  "module": "Maintenance",
  "filters": ["Start Date","End Date","Level"],
  "table_columns": ["Timestamp","Level","source","Message"],
  "buttons": ["Search","Reset","Export Logs","Clear Logs(destructive)"],
  "paginated": true
}
```
