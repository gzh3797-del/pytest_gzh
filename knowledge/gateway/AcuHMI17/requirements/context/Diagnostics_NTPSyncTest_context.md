# Diagnostics / NTP Sync Test — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/diagnostics/ntpSyncTest` |
| 路由名 | `ntpSyncTest` |
| 面包屑 | AcuHMI-1-7 / Diagnostics / NTP Sync Test |
| 顶级模块 | Diagnostics |

## 2. 页面用途

执行 NTP 时间同步测试并展示 ntpd 运行日志。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| NTP Sync | 只读文本块 | 文本 "NTP Sync" 后 | ntpd 输出日志 |
| Refresh | button | `getByRole('button',{name:'Refresh'})` | 刷新（测试进行中为 `disabled`） |

## 4. 页面状态与分支

| 状态 | 说明 |
|------|------|
| 测试进行中 | Refresh 按钮 disabled |
| 测试完成 | Refresh 可用，日志更新 |

## 5. 自动化测试要点

- 只读日志页；进入后自动运行 NTP 测试（Refresh disabled→enabled 的状态转换可断言）。
- 依赖 Date & Time 页 NTP 服务器配置（联动）。

## 6. 机器可解析摘要

```json
{
  "route": "/diagnostics/ntpSyncTest",
  "name": "ntpSyncTest",
  "title": "NTP Sync Test",
  "module": "Diagnostics",
  "readonly_sections": ["NTP Sync (ntpd log)"],
  "buttons": ["Refresh (disabled while running)"]
}
```
