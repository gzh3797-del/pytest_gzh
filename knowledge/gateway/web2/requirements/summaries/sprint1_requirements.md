# ACM-41-WEB2 Sprint 1 需求摘要

> 来源：《AcuRev-4100-WEB2 Sprint 1 软件需求规格说明书汇总 V1.0》（2026-02-25）
> 原件 → ../raw/（含 docx 及图片）

---

## 一、产品定位与系统模式

WEB2（ACM-41-WEB2）是 AcuRev-4100 的网络扩展模块，通过高速 USB（Modbus RTU）与表本体通信。
运行模式二选一，切换时设备重启：

| 模式 | Ethernet2 角色 | 全局设备下拉 | Settings→Device→Connection |
|------|--------------|------------|---------------------------|
| WEB Module | 普通以太网口，可独立配 DHCP/IP | 隐藏 | 隐藏 |
| Gateway | DHCP Server，自动分配 IP 给下挂设备 | 显示 | 显示 |

---

## 二、功能模块清单

### 2.1 设备与数据采集

| 功能 | 要点 |
|------|------|
| 设备添加/删除 | RS485（Modbus RTU）+ Ethernet（TCP/IP）两路，最多 3 台 4100 + 8 台 AcuIOM |
| 参数配置 | Basic/Ethernet/RS485/Modbus、User&CT（铅封保护）、PQ Event & Waveform、IO（DI/DO/RO/AI/AO） |
| AO Mapping | 4100 参数映射到 AcuIOM AO，最多 32 组，默认间隔 1 秒 |
| Reading 采集周期 | Basic 200ms、PQ/Energy 3s、Demand/TOU/Max 1min |

**Metering 展示页（9 种）：** Realtime / Energy / Demand / THD / Harmonic Spectrum / Sequence & Unbalance / Max Demand / Max Min / IO

**Advanced Commission：** 只读展示下挂设备系统参数，支持 PDF 下载（System Parameter Output / Meter Wiring Config / Demand / Network IPv4 / Device Information / DO Energy Pulse Parameter / CT Model / CT Summary）

---

### 2.2 日志管理

| 日志类型 | 最大条数 | 触发/采集 | 特性 |
|---------|---------|---------|------|
| Trend Log（Realtime） | 1d/7d/30d | Freq/VLN/VLL/I/P/Q/S：3s~1day | 支持查询、下载、推送、清除 |
| Trend Log（Energy） | 30d/1y/10y | Energy：1min~1month | |
| SOE Log | 200 条（覆盖最老） | DI 状态变化 | AcuRev-4100 + AcuIOM |
| PQ Event Log | 200 条（文档有矛盾项 1000 条） | Sag/Swell/Interrupt/Current Swell | 含 Waveform ID、Par Max/Min |
| Waveform Log | 500 条（覆盖最老） | PQ 事件/手动触发 | COMTRADE 格式 |
| System Event Log | ≥ 10 万条（循环覆盖） | 系统/用户/告警/通信/安全事件 | 单条写入 ≤5ms，查询 ≤1s |

**Data Log & Post Channel：**
- Post Channel 1~3（FTP/SFTP/HTTP/HTTPS），各自独立配置服务器
- Data Logger 最多 3 个；Log File Length 与 Log Interval 均支持 1min~1month，需 Log Interval ≤ Log File Length
- 参数范围：AcuRev-4100（Basic/Demand/PQ 不含谐波/Energy/IO）+ AcuIOM（AI/DI/DI counter/RO/DO）

**PQ Event Analysis：** ITIC Curve（ITIC 标准）、SEMI F47 Curve（半导体设备耐受标准）

**Alarm：**
- 每设备最多 32 个 Monitor（4100：Basic/Demand/PQ 不含谐波/DI status；AcuIOM：AI/DI）
- DI status 告警：min/max 不可配，DI=1 时触发
- 告警通知：邮件，最多 3 个收件人，发送间隔 1~10 整数（单位未指定）

---

### 2.3 系统配置与网络管理

**Management：** 需量/能量/告警重置；设备重启/重置运行时间/恢复出厂设置（AcuIOM 仅支持告警重置+重启+恢复出厂）

**NTP 时间同步：** 最多 3 个 NTP 服务器，时区可选，Sync Device Time 按钮手动同步到所有下挂设备

**WEB2 ↔ 4100 配置同步：** USB 通信，互写 Ethernet1/2/WiFi/RSTP/版本/SN/MAC；4100 可下发重启/恢复出厂/恢复网络配置指令

**默认网络参数：**

| 接口 | 模式 | 默认 IP | 默认 Mask | 默认 GW |
|------|------|---------|---------|---------|
| Ethernet1 | DHCP 禁用 | 192.168.8.101 | 255.255.255.0 | 192.168.8.1 |
| Ethernet2（WEB Module） | DHCP 启用 | — | — | — |
| Ethernet2（Gateway，DHCP Server）| — | 192.168.200.1 | 255.255.255.0 | — |
| Ethernet2 地址池（Gateway）| — | 192.168.200.100~200 | — | — |
| WiFi（AP 模式）| — | 192.168.100.1 | — | — |
| DNS1/DNS2 | — | 8.8.8.8 | — | 8.8.4.4 |

**WiFi：** 默认 AP 模式；SSID=`AcuRev-4100-WEB2-WIFI-{SN后6位}`，Key=`Accuenergy`；WEB2-D 型号**不支持 WiFi**
**IPv6：** 默认 Disable；支持 SLAAC/静态/DHCPv6，IPv4/IPv6 双栈

**安全：**
- Certificate Management：Import/Export/Generate CSR/Generate New Self-Signed
- Session Timeout：0~60 分钟（0=永不超时）
- Password Policy：8 个配置项（字符格式/历史/最小使用天/过期/宽限期/最小长度/最大失败次数/失败窗口/锁定等待）

**Remote Access：** URL 注册/去注册，Ping Interval 默认 60s（可选 60s/600s）

---

### 2.4 协议转换

| 协议 | 关键约束 |
|------|---------|
| MQTT | SSL/TLS；QoS=1/2；JSON 格式；周期上传；最大 12 台设备；缓存 ≥10000 条；单次 ≤50KB；最大连接时延 ≤3s |
| SNMP | v1/v2c/v3；固定 OID 不可修改；Trap 最多 4 个目标；周期最小 5 秒；最大 12 台设备 |
| AcuCloud | HTTPS；选择设备上传；固定数据点；周期/启动/手动三种触发 |
| Modbus Pass Through | 默认禁用 |
| Modbus Config | Slave ID、TCP Enable/Port 配置 |

---

### 2.5 权限管理

User Configuration / Role Configuration / Password Management / Access Control

---

### 2.6 系统运营维护

| 功能 | 要点 |
|------|------|
| Config Management | 配置备份/恢复 |
| Firmware | 支持本体及下挂设备固件更新（Sprint 2 增强，禁止降级在 Sprint 2 明确） |
| Emergency Mode | — |
| 系统诊断 | — |
| About | Information / Installation Record / Inspection Record（Sprint 2 同步推送 AcuCloud） |

---

## 三、与 Sprint 2 的分工对照

| 功能 | Sprint 1 | Sprint 2 新增/变更 |
|------|---------|-------------------|
| 协议 | MQTT / SNMP / AcuCloud | **新增** BACnet/IP / Ethernet/IP / AWS IoT / Azure IoT / Device Mirror |
| 固件升级 | 基础支持 | 禁止降级、MFEA 格式、模板版本兼容 |
| Modbus 端口 | 1~65534 | **变更** 2000~5999 |
| Alarm | 基础告警配置 | — |
| AcuCloud | 基础上传 | Advanced 后门、Installation/Inspection 自动推送 |
| Virtual Device | — | **新增** 多台 4100 能量计算映射 |
| BACnet 谐波 | — | **新增** 谐波参数 |
| 接线检查 | — | **新增** 五种接线方式检查 + CSV 导出 |
| User Channel 命名 | "User Channel" | **变更** 改名为"Meter Point" |
| 默认密码 | — | **变更** Admin@AABBCC |
| About | 基础 | **拆分** Service 页面（Troubleshooting/.a2d 诊断文件） |
