# Diagnostics / Host Lookup — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/diagnostics/hostLookup` |
| 路由名 | `hostLookup` |
| 面包屑 | AcuHMI-1-7 / Diagnostics / Host Lookup |
| 顶级模块 | Diagnostics |

## 2. 页面用途

对指定主机/域名执行 nslookup / ping / traceroute 网络诊断。

## 3. 交互元素清单 / 表单字段

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Name of system or domain to lookup | textbox | `getByRole('textbox',{name:'Name of system or domain to lookup'})` | 必填(*)，目标主机/域名/IP |
| nslookup | checkbox | `getByRole('checkbox',{name:'nslookup'})` | 勾选执行 nslookup |
| ping | checkbox | `getByRole('checkbox',{name:'ping'})` | 勾选执行 ping |
| traceroute | checkbox | `getByRole('checkbox',{name:'traceroute'})` | 勾选执行 traceroute |
| Lookup | button | `getByRole('button',{name:'Lookup'})` | 执行并输出结果 |

## 4. 自动化测试要点

- 必填目标校验；至少勾选一种诊断方式。
- 各 checkbox 独立勾选（原子操作）；Lookup 后异步结果断言（wait_for）。

## 5. 机器可解析摘要

```json
{
  "route": "/diagnostics/hostLookup",
  "name": "hostLookup",
  "title": "Host Lookup",
  "module": "Diagnostics",
  "fields": {
    "target": {"type":"text","required":true},
    "methods": {"type":"checkbox","options":["nslookup","ping","traceroute"]}
  },
  "buttons": ["Lookup"]
}
```
