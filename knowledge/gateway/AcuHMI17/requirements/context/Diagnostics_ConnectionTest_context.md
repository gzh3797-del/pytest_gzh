# Diagnostics / Connection Test — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/diagnostics/connectionTest` |
| 路由名 | `connectionTest` |
| 面包屑 | AcuHMI-1-7 / Diagnostics / Connection Test |
| 顶级模块 | Diagnostics |

> Diagnostics 二级 tab：Network Status / RSTP Status / Host Lookup / Connection Test / NTP Sync Test / Modbus Debug Log /（...More→）Debug / Wiring Check。

## 2. 页面用途

尝试连接指定网络节点，测试全部网络设置并给出详细结果报告。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| 二级 tab | menuitem | `getByRole('menuitem',{name:'Network Status'})` 等 | 切换诊断子页 |
| Start Test | button | `getByRole('button',{name:'Start Test'})` | 启动连接测试，运行后展示报告 |

## 4. 自动化测试要点

- 点击 Start Test → 等待并断言结果报告出现（异步，需 wait_for）。

## 5. 机器可解析摘要

```json
{
  "route": "/diagnostics/connectionTest",
  "name": "connectionTest",
  "title": "Connection Test",
  "module": "Diagnostics",
  "buttons": ["Start Test"],
  "output": "async test report",
  "sub_tabs": ["Network Status","RSTP Status","Host Lookup","Connection Test","NTP Sync Test","Modbus Debug Log","Debug","Wiring Check"]
}
```
