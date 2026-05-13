# RPP — Remote Power Panel 测量系统

## 项目背景

RPP（Remote Power Panel）是 Accuenergy 面向数据中心场景的远程电力面板测量系统。MH（Main Host）主控单元通过专有内部总线 AccuBus 连接多个 VMM（电压测量子模块）和 CMM（电流测量子模块），支持最多 96 路电流通道（CH001~CH096）、最多 2 个 VMM（V001/V002），以及最多 96（必须基线）/ 192（期望目标）个 Meter Point（虚拟电表）。

同时具备 Gateway 功能，可通过 Ethernet/RS-485 南向采集下挂 Modbus 设备（≥32台）。

需求文档：[requirements/summaries/v1.01.md](requirements/summaries/v1.01.md)

## 当前模块

| 模块 | 章节 | 功能 |
|------|------|------|
| 通讯 | §3 | AccuBus 私有协议、Modbus TCP/RTU Client/Server（含 Pass Through）、SNMP/BACnet/MQTT/AcuCloud 等北向 |
| 虚拟设备 | §4 | 算术派生参数（+/−/×/÷），≥16台，对北向透明 |
| 测量软件 | §5 | 五层实体（VMM/Channel/Channel Instance/MP/VMM System）、IEC 61000-4-30 Class A 聚合、接线配置/检查、PQ事件与录波、Max/Min极值 |
| 本地数据记录 | §6 | 3个 Datalog 实例（1s+，可 Post 推送）、AcuCloud（断网回推）、Trend Log（5min） |
| 告警 | §7 | 200组 Alarm Monitor（AND/OR，pickup/dropout 延迟）、设备告警、Alarm Log（≥65535条）|
| 时间同步 | §8 | NTP + PTP（PTP优先），AccuBus 下行分发 ≤0.5μs，RTC Fallback |
| 访问控制 | §9 | HTTPS，最多32用户/16角色，TLS 1.2/1.3，Password Policy |
| 固件升级 | §10 | HTTPS Web（MH）+ AccuBus 下行（VMM/CMM），A/B 双分区回滚 |
| 维护与诊断 | §11 | System Status/Log、AcuCloud Log、诊断工具、维护命令 |
| 系统设置 | §12 | 时间/网络/邮件/Alarm Notification/证书/Remote Access |
| 网页 | §13 | HTTPS，仅英文，响应式（桌面≥1280px，平板≥768px） |
| Acuview2 支持 | §14 | v1：读取/接线配置（Modbus）+ 固件升级（HTTPS） |

## 运行方式

RPP 通过 HTTPS Web 界面（内嵌 MH）进行配置与运维；北向通信：
- Modbus TCP 默认端口 502
- BACnet/IP 默认端口 47808
- SNMP 默认端口 161（可改 16100~16199）
- RS-485 Modbus RTU Server（与南向 Client 共用物理口，角色由用户配置）

## 网络配置

| 接口 | 默认 |
|------|------|
| Ethernet 1 | 手动 IP 192.168.8.101 / 255.255.255.0，Gateway 192.168.8.1 |
| Ethernet 2 | DHCP |
| DNS | 8.8.8.8 / 8.8.4.4 |

RSTP 默认禁用；启用时双口作冗余环网。

## 容量边界

| 档位 | VMM | Channel | MP | Channel Instance | 双VMM镜像 |
|------|-----|---------|-----|-----------------|----------|
| 必须基线 | 1~2 | ≤96 | ≤96 | ≤96 | 不支持 |
| 期望目标 | 1~2 | ≤96 | ≤192 | ≤192 | 支持 |

## 接线方式

1P2W / 1P3W / 2E3W Network / 2E3W Delta / 3P3W / 3P4W
（3 Element 4 Wire Delta 不支持）

## 北向协议覆盖（v1）

| 协议 | 状态 |
|------|------|
| Modbus TCP/RTU Server | 必须 |
| SNMP v2c/v3 | 必须 |
| BACnet/IP | 必须 |
| MQTT（Publisher） | 必须 |
| AcuCloud | 必须 |
| Email/HTTP POST/FTP | 必须 |
| OPC UA | 期望（下一Sprint） |
| AWS IoT Core | 期望（下一Sprint） |
| Azure IoT Hub | 期望（下一Sprint） |
| IPv6 双栈 | 期望（下一Sprint） |

## 参考文档

| 文档 | 说明 |
|------|------|
| RPP PRS v1.02 | 上游产品需求，含硬件/EMC/认证/物理测量指标 |
| RPP URS（待编写） | 下游 UI 详细规格 |
| AcuHMI-1-7 v1.04（2026-05-14） | 访问控制/维护/系统设置模块的参考依据 |
