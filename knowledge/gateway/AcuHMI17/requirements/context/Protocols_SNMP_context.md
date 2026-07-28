# Protocols / SNMP — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/snmp` |
| 路由名 | `snmp` |
| 面包屑 | AcuHMI-1-7 / Protocols / SNMP |
| 顶级模块 | Protocols |

## 2. 页面用途

配置网关 SNMP 代理：启用、待暴露设备选择、SNMP 版本与端口、SNMPv3 认证/加密、Trap 目标与上报缓冲参数。可下载 MIB 文件。

## 3. 交互元素清单 / 表单字段

| 字段/元素 | 类型 | 定位策略 1 | 默认/示例 | 校验/说明 |
|-----------|------|-----------|-----------|-----------|
| SNMP Enable | radiogroup | `getByRole('radiogroup',{name:'SNMP Enable'})` | Enable | 必选(*) |
| Devices Selection 表 | table | group "Devices Selection" | — | checkbox/Device Name/SN/Protocol/Online |
| SNMP Version | combobox | `getByRole('combobox',{name:'SNMP Version'})` | SNMP v3 | 必选(*)，切换影响 v3 配置显隐 |
| Port | textbox | `getByRole('textbox',{name:'Port'})` | 16100 | 必填，默认161，范围 **16100–16199** |
| Username (v3) | textbox | `getByRole('textbox',{name:'Username'})` | admin | 必填(*) |
| User Password (v3) | textbox(password+eye) | `getByRole('textbox',{name:'User Password'})` | 12345678 | 必填，**≥8 字符**，含显隐图标 |
| Auth Protocol (v3) | combobox | `getByRole('combobox',{name:'Auth Protocol'})` | SNMPV3 SHA | 必选 |
| Privacy Protocol (v3) | combobox | `getByRole('combobox',{name:'Privacy Protocol'})` | SNMPv3 DES | 必选 |
| Privacy Password (v3) | textbox(password+eye) | `getByRole('textbox',{name:'Privacy Password'})` | 12345678 | 必填，≥8 字符 |
| Trap Enable | radiogroup | `getByRole('radiogroup',{name:'Trap Enable'})` | Enable | 控制 Trap Target 显隐 |
| Trap Target 1 | textbox | `getByRole('textbox',{name:'Trap Target 1'})` | 192.168.2.9 | 必填(*)，IP 格式 |
| Trap Target 2/3/4 | textbox | `getByRole('textbox',{name:'Trap Target 2'})` 等 | 空 | 可选，IP 格式 |
| Report Buffer Size | textbox | `getByRole('textbox',{name:'Report Buffer Size'})` | 30 | 必填，范围 **0–30** |
| Report Hold Time | textbox | `getByRole('textbox',{name:'Report Hold Time'})` | 300 | 必填，范围 **0–300** |
| MIB File Download | button | `getByRole('button',{name:'MIB File Download'})` | — | 下载 MIB |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 页面状态与分支 ★

| 状态 | 触发 | 结果 |
|------|------|------|
| SNMP Version = v3 | 默认 | 显示 SNMPv3 Configuration（Username/密码/Auth/Privacy） |
| SNMP Version = v1/v2c | 切换 | v3 配置隐藏（改为 Community 等，需运行时确认） |
| Trap Enable = Enable | 默认 | 显示 Trap Target 1–4 |
| Trap Enable = Disable | 切换 | 隐藏 Trap 目标 |

## 5. 校验规则要点

- Port：16100–16199。
- User/Privacy Password：≥8 字符。
- Trap Target 1 必填且各 Target 为合法 IP。
- Report Buffer Size 0–30；Report Hold Time 0–300。

## 6. 自动化测试要点

- 版本切换导致 v3 配置显隐（核心分支）。
- 密码显隐图标切换断言（明文/掩码）。
- 端口/密码长度/IP 格式/数值范围校验用例。
- MIB File Download 触发下载。

## 7. 机器可解析摘要

```json
{
  "route": "/protocols/snmp",
  "name": "snmp",
  "title": "SNMP",
  "module": "Protocols",
  "fields": {
    "SNMP Enable": {"type":"radio","default":"Enable"},
    "SNMP Version": {"type":"select","default":"SNMP v3"},
    "Port": {"type":"text","default":161,"range":[16100,16199]},
    "Username": {"type":"text","when":"v3"},
    "User Password": {"type":"password","minlen":8,"toggle":true},
    "Auth Protocol": {"type":"select","default":"SNMPV3 SHA"},
    "Privacy Protocol": {"type":"select","default":"SNMPv3 DES"},
    "Privacy Password": {"type":"password","minlen":8},
    "Trap Enable": {"type":"radio","default":"Enable","controls":["Trap Target 1-4"]},
    "Trap Target 1": {"type":"text","format":"ip","required":true},
    "Trap Target 2-4": {"type":"text","format":"ip","required":false},
    "Report Buffer Size": {"type":"text","range":[0,30]},
    "Report Hold Time": {"type":"text","range":[0,300]}
  },
  "buttons": ["MIB File Download","Save"]
}
```
