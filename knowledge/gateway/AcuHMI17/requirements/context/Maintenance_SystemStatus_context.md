# Maintenance / System Status — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/maintenance/systemStatus` |
| 路由名 | `systemStatus` |
| 面包屑 | AcuHMI-1-7 / Maintenance / System Status |
| 顶级模块 | Maintenance |

## 2. 页面用途

展示系统资源使用（CPU/RAM/Disk）并提供重启。Maintenance 二级 tab：System Status / Event Log。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 示例 | 说明 |
|------|------|-----------|------|------|
| 二级 tab | menuitem | `getByRole('menuitem',{name:'Event Log'})` | — | 切换子页 |
| CPU | progressbar | 文本 "CPU" 后 progressbar | 1.85% | 只读 |
| RAM | progressbar | 文本 "RAM" 后 progressbar | 592/982 MB | 只读 |
| Disk | progressbar | 文本 "Disk" 后 progressbar | 590/27244 MB | 只读 |
| Reboot System | button | `getByRole('button',{name:'Reboot System'})` | — | 重启（危险操作，通常二次确认） |

## 4. 自动化测试要点

- 只读资源指标呈现断言（数值/百分比格式）。
- **Reboot System 为破坏性操作**：自动化仅验证二次确认弹框，默认不实际重启。

## 5. 机器可解析摘要

```json
{
  "route": "/maintenance/systemStatus",
  "name": "systemStatus",
  "title": "System Status",
  "module": "Maintenance",
  "readonly_metrics": ["CPU(%)","RAM(MB)","Disk(MB)"],
  "buttons": ["Reboot System(destructive)"],
  "sub_tabs": ["System Status","Event Log"]
}
```
