# Diagnostics / RSTP Status — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/diagnostics/rstpStatus` |
| 路由名 | `rstpStatus` |
| 面包屑 | AcuHMI-1-7 / Diagnostics / RSTP Status |
| 顶级模块 | Diagnostics |

## 2. 页面用途

只读展示 RSTP（生成树）网桥与端口状态。未启用 RSTP 时显示 "Interface br0 does not exist."。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Bridge and Port Status | 只读文本块 | 文本 "Bridge and Port Status" 后 | RSTP 状态输出 |
| Refresh | button | `getByRole('button',{name:'Refresh'})` | 刷新 |

## 4. 自动化测试要点

- 只读；RSTP 关闭态断言 "br0 does not exist"；启用 RSTP（Network 页）后应显示网桥状态（联动用例）。

## 5. 机器可解析摘要

```json
{
  "route": "/diagnostics/rstpStatus",
  "name": "rstpStatus",
  "title": "RSTP Status",
  "module": "Diagnostics",
  "readonly_sections": ["Bridge and Port Status"],
  "buttons": ["Refresh"]
}
```
