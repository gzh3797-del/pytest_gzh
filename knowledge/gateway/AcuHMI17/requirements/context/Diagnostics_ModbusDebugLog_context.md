# Diagnostics / Modbus Debug Log — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/diagnostics/modbusDebugLog` |
| 路由名 | `modbusDebugLog` |
| 面包屑 | AcuHMI-1-7 / Diagnostics / Modbus Debug Log |
| 顶级模块 | Diagnostics |

## 2. 页面用途

抓取/查询网关与下游设备之间的 Modbus 报文（TCP/RTU 请求与响应），用于调试。

## 3. 交互元素清单 / 表单字段

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Modbus Debug Trace | radiogroup | `getByRole('radiogroup',{name:'Modbus Debug Trace'})` | Enable/Disable（默认 Enable），控制是否抓包 |
| Start/End Date | date combobox | `getByRole('combobox',{name:'Start Date'})` / `{name:'End Date'}` | Interval 过滤 |
| Type | combobox | `getByRole('combobox',{name:'Type'})` | 报文类型（TCP_REQ/TCP_RSP/RTU_REQ/RTU_RSP 等） |
| Slave ID | textbox | `getByRole('textbox',{name:'Slave ID'})` | 按从站过滤 |
| Function Code | textbox | `getByRole('textbox',{name:'Function Code'})` | 按功能码过滤 |
| Search | button | `getByRole('button',{name:'Search'})` | 查询 |
| Reset | button | `getByRole('button',{name:'Reset'})` | 重置 |
| 报文表 | table | 列: Timestamp/Src/Dest/Type/SlaveID/Function Code/Data | 只读 |
| 分页 | button/listitem | `getByRole('button',{name:'Go to next page'})` | 多页（示例 6554 页） |
| Clear Debug Logs | button | `getByRole('button',{name:'Clear Debug Logs'})` | 清空（危险，通常二次确认） |
| Export Debug Logs | button | `getByRole('button',{name:'Export Debug Logs'})` | 导出 |

## 4. 表列语义

Timestamp / Src(meter 或 AcuHMI-1-7) / Dest / Type(TCP_REQ/TCP_RSP/RTU_REQ...) / SlaveID / Function Code / Data(hex 报文)

## 5. 自动化测试要点

- Trace Enable/Disable 抓包开关；多条件过滤（Type/SlaveID/Function Code/日期）+ Search/Reset。
- 报文十六进制内容断言；分页/每页条数。
- **Clear Debug Logs 破坏性**，仅验证二次确认；Export 下载。

## 6. 机器可解析摘要

```json
{
  "route": "/diagnostics/modbusDebugLog",
  "name": "modbusDebugLog",
  "title": "Modbus Debug Log",
  "module": "Diagnostics",
  "fields": {
    "Modbus Debug Trace": {"type":"radio","default":"Enable"},
    "filters": ["Start Date","End Date","Type","Slave ID","Function Code"]
  },
  "table_columns": ["Timestamp","Src","Dest","Type","SlaveID","Function Code","Data"],
  "buttons": ["Search","Reset","Clear Debug Logs(destructive)","Export Debug Logs"],
  "paginated": true
}
```
