# Diagnostics / Network Status — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/diagnostics/networkStatus` |
| 路由名 | `networkStatus` |
| 面包屑 | AcuHMI-1-7 / Diagnostics / Network Status |
| 顶级模块 | Diagnostics |

## 2. 页面用途

只读展示网络诊断信息：接口(ifconfig)、路由表、DNS、连接统计(netstat)。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Ethernet Network | 只读文本块 | 文本 "Ethernet Network" 后 | ifconfig 输出（eth0/lan1/lan2/lo） |
| Routing Table | 只读文本块 | 文本 "Routing Table" 后 | 内核路由表 |
| DNS Server | 只读文本块 | 文本 "DNS Server" 后 | nameserver 列表 |
| Network Stat | 只读文本块 | 文本 "Network Stat" 后 | netstat（监听端口 5999/80/22/443/199/16100/49000/123 等） |
| Refresh | button | `getByRole('button',{name:'Refresh'})` | 刷新数据 |

## 4. 自动化测试要点

- 只读页；断言各区块存在与关键字（如监听端口、lan1 IP）。
- Refresh 后数据更新。

## 5. 机器可解析摘要

```json
{
  "route": "/diagnostics/networkStatus",
  "name": "networkStatus",
  "title": "Network Status",
  "module": "Diagnostics",
  "readonly_sections": ["Ethernet Network","Routing Table","DNS Server","Network Stat"],
  "buttons": ["Refresh"]
}
```
