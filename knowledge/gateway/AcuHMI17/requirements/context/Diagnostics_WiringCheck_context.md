# Diagnostics / Wiring Check — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/diagnostics/wiringCheck` |
| 路由名 | `wiringCheck` |
| 面包屑 | AcuHMI-1-7 / Diagnostics / Wiring Check |
| 顶级模块 | Diagnostics（...More） |

## 2. 页面用途

对电表接线进行检查，按电压(Voltage)与计量点(Meter Point)展示各相测量值与接线状态（正常/Missing 等）。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 默认 | 说明 |
|------|------|-----------|------|------|
| Device | combobox | `getByRole('combobox')`（Device 区） | All | 选择检查设备 |
| Export | button | `getByRole('button',{name:'Export'})` | — | 导出结果 |
| Wiring Check | button | `getByRole('button',{name:'Wiring Check'})` | — | 启动检查 |
| Stop | button | `getByRole('button',{name:'Stop'})` | disabled | 检查进行中可用 |
| Show Only Issues | switch | `getByRole('switch')` 邻接 "Show Only Issues" | 关 | 仅显示异常项 |
| Voltage 表 | table | 列: Device Name/Wiring Configuration/Phase/Measurement/Wiring Status | — | 分页 |
| Meter Point 表 | table | 列: Meter Point/Device Name/Wiring Configuration/Phase/Input Channel/Measurement/Wiring Status | — | 分页 |

## 4. 表数据语义

- Wiring Status：Not Checked / Missing / (正常态)。Measurement 形如 `0.000∠0.000°`（幅值∠相角）。
- Wiring Configuration：如 "2 Element 3 Wire 1 Phase"、"3 Element 4 Wire Y"。

## 5. 页面状态与分支

| 状态 | 说明 |
|------|------|
| 未检查 | Wiring Status = Not Checked；Stop disabled |
| 检查进行中 | Stop 可用 |
| Show Only Issues 开 | 仅列出异常(如 Missing)行 |

## 6. 自动化测试要点

- Device 选择 → Wiring Check 启动 → 等待结果（异步，Stop enable/disable 状态转换）。
- Show Only Issues 过滤断言；两张结果表分页。
- Export 下载。

## 7. 机器可解析摘要

```json
{
  "route": "/diagnostics/wiringCheck",
  "name": "wiringCheck",
  "title": "Wiring Check",
  "module": "Diagnostics",
  "controls": {"Device":{"type":"select","default":"All"},"Show Only Issues":{"type":"switch"}},
  "buttons": ["Export","Wiring Check","Stop(disabled until running)"],
  "tables": {
    "Voltage": ["Device Name","Wiring Configuration","Phase","Measurement","Wiring Status"],
    "Meter Point": ["Meter Point","Device Name","Wiring Configuration","Phase","Input Channel","Measurement","Wiring Status"]
  }
}
```
