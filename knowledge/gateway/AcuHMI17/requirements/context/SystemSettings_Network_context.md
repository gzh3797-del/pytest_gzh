# System Settings / Network — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/systemSettings/network` |
| 路由名 | `network` |
| 面包屑 | AcuHMI-1-7 / System Settings / Network |
| 顶级模块 | System Settings |

## 2. 页面用途

配置网关网络：RSTP、默认出站接口、双以太网口（Ethernet 1/2）的 DHCP/静态 IP、DNS。

## 3. 交互元素清单 / 表单字段

| 字段/元素 | 类型 | 定位策略 1 | 默认/示例 | 说明 |
|-----------|------|-----------|-----------|------|
| RSTP Enable | switch | `getByRole('switch')` 邻接 "RSTP Enable" | 关 | 开关 |
| Default Interface (Outbound Traffic) | combobox | `getByRole('combobox',{name:'Default Interface (Outbound Traffic)'})` | Ethernet 1 | 选择默认出站口 |
| Ethernet 1 面板 | 可折叠 | 文本 "Ethernet 1" + 折叠图标 | 展开 | 见下 |
| Ethernet 1 · DHCP Enable | radiogroup | `getByRole('radiogroup',{name:'DHCP Enable'})`（第1个） | Auto | Auto/Manual；Manual→显示手动 IP/掩码/网关 |
| Ethernet 1 · Interface Status | textbox(disabled) | `getByRole('textbox',{name:'Interface Status'})` | Connected | 只读 |
| Ethernet 1 · IP | textbox(disabled) | `getByRole('textbox',{name:'IP'})` | 192.168.3.71 | 只读（Auto 时） |
| Ethernet 2 面板 | 可折叠 | 文本 "Ethernet 2" | 展开 | 同 Ethernet 1 结构 |
| Ethernet 2 · Interface Status | textbox(disabled) | — | Disconnected | 只读 |
| Ethernet 2 · IP | textbox(disabled) | — | 0.0.0.0 | 只读 |
| DNS 1 | textbox | `getByRole('textbox',{name:'DNS 1'})` | 8.8.8.8 | 必填(*)，须为合法 IP 或域名 |
| DNS 2 | textbox | `getByRole('textbox',{name:'DNS 2'})` | 8.8.4.4 | 可选，合法 IP 或域名 |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 页面状态与分支 ★

| 状态 | 触发 | 结果 |
|------|------|------|
| DHCP = Auto | 默认 | Interface Status/IP 只读展示，自动获取 |
| DHCP = Manual | 切换 | **显示 IP\* / Subnet Mask\* / Gateway\* 三个必填输入框**（实测 Eth1 回显 192.168.8.101 / 255.255.255.0 / 192.168.8.1）；切回 Auto 恢复只读 |
| 接口面板折叠/展开 | 点击 Ethernet 1/2 标题 | 显隐该接口配置 |
| RSTP Enable 开 | 开关 | 启用 RSTP（生成树） |

## 5. 校验规则要点

- DNS 1 必填；DNS 1/2 须为合法 IP 或域名。
- Manual 模式下 IP/掩码/网关格式校验。

## 6. 自动化测试要点

- DHCP Auto/Manual 切换字段可编辑性联动（核心分支）。
- 双接口独立配置；折叠面板展开断言。
- DNS 格式校验；⚠️ 修改 IP 可能导致断连——自动化中谨慎实际保存。

## 7. 机器可解析摘要

```json
{
  "route": "/systemSettings/network",
  "name": "network",
  "title": "Network",
  "module": "System Settings",
  "fields": {
    "RSTP Enable": {"type":"switch"},
    "Default Interface": {"type":"select","default":"Ethernet 1"},
    "Ethernet1.DHCP": {"type":"radio","options":["Auto","Manual"],"default":"Auto"},
    "Ethernet2.DHCP": {"type":"radio","options":["Auto","Manual"],"default":"Auto"},
    "DNS 1": {"type":"text","required":true,"format":"ip_or_domain"},
    "DNS 2": {"type":"text","format":"ip_or_domain"}
  },
  "readonly": ["Interface Status","IP (when Auto)"],
  "conditional": {"when":"DHCP=Manual","shows":["IP*","Subnet Mask*","Gateway*"]},
  "buttons": ["Save"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/systemsettings/`。

### 加载态 API
- `GET /api/settings/networkConfig`、`GET /api/command/getIp`、`GET /api/settings/proxyServerConfig`

### pytest 选择器与控件
- RSTP Enable：`el-switch`，默认 OFF。
- Default Interface (Outbound Traffic)：`el-select`，`page.get_by_role("combobox", name="Default Interface (Outbound Traffic)")`，默认 `Ethernet 1`。
- Ethernet 1/2 DHCP Enable：`el-radio` Auto/Manual，默认 Auto。**Manual 时替换为 IP\*/Mask\*/Gateway\* 三必填框**（Eth1 回显 192.168.8.101/255.255.255.0/192.168.8.1），切回 Auto 恢复。
- Interface Status：Eth1=`Connected`、Eth2=`Disconnected`——**同名两个，需 scope 到对应 Ethernet 区块**。
- DNS 1*/DNS 2：`page.get_by_placeholder("Enter DNS 1")` / `("Enter DNS 2")`，默认 8.8.8.8 / 8.8.4.4。

### 校验时机（实测）
- **DNS 1/2：blur 即校验**（非法值直接报 `DNS 1 must be a valid domain or a valid IP address`，无需点 Save）。
- Manual 模式 IP/Mask/Gateway：格式类，blur 即校验。

### Element-Plus 通用坑
- `el-radio` 点 label 兜底；同名 group（Ethernet 1/2 的 Interface Status/IP）必须父容器 scope。

### 高危
- ⚠️ 修改 IP / 切 Manual 保存可能导致**当前会话断连**——自动化中谨慎实际保存，须带重连逻辑。
