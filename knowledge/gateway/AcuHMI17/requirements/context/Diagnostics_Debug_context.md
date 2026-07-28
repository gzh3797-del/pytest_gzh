# Diagnostics / Debug — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/diagnostics/debug` |
| 路由名 | `debug` |
| 面包屑 | AcuHMI-1-7 / Diagnostics / Debug |
| 顶级模块 | Diagnostics（...More） |

## 2. 页面用途

调试相关：SSH 开关与端口、总调试诊断开关、下载诊断文件、重置系统日志。

## 3. 交互元素清单 / 表单字段

| 元素 | 类型 | 定位策略 1 | 默认/示例 | 校验/说明 |
|------|------|-----------|-----------|-----------|
| SSH | radiogroup | `getByRole('radiogroup',{name:'SSH'})` | On | 必选(*) On/Off |
| port | textbox | `getByRole('textbox',{name:'port'})` | 22 | 必填(*)，默认22，范围 **6000–9999** |
| Turn On/Off all debug diagnostics | radiogroup | `getByRole('radiogroup',{name:'Turn On/Off all debug diagnostics'})` | — | On/Off（当前 "Debug Diagnostic - All On"） |
| Show Detail | checkbox | `getByRole('checkbox')` 邻接 "Show Detail" | 未勾 | 勾选展开各项调试诊断明细 |
| Download Diagnostic File | button | `getByRole('button',{name:'Download Diagnostic File'})` | — | 下载诊断文件 |
| Reset System Logs | button | `getByRole('button',{name:'Reset System Logs'})` | — | 重置系统日志（危险，二次确认） |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 页面状态与分支

| 状态 | 触发 | 结果 |
|------|------|------|
| Show Detail 勾选 | 勾选 | 展开各调试诊断项明细列表 |
| Debug diagnostics On/Off | 切换 | 批量开关所有调试诊断 |

## 5. 自动化测试要点

- SSH 端口范围校验（6000–9999；注意默认展示 22）。
- Show Detail 显隐明细断言。
- Download Diagnostic File 下载；**Reset System Logs 破坏性**，仅验证二次确认。

## 6. 机器可解析摘要

```json
{
  "route": "/diagnostics/debug",
  "name": "debug",
  "title": "Debug",
  "module": "Diagnostics",
  "fields": {
    "SSH": {"type":"radio","options":["On","Off"],"default":"On"},
    "port": {"type":"text","default":22,"range":[6000,9999]},
    "Turn On/Off all debug diagnostics": {"type":"radio","options":["On","Off"]},
    "Show Detail": {"type":"checkbox"}
  },
  "buttons": ["Download Diagnostic File","Reset System Logs(destructive)","Save"]
}
```
