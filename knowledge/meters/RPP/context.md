# RPP — Remote Power Panel 测量系统

## 测试设计必读（固定结构速查区）

> strategy-design / testcase-design / coverage-check 的「知识库提取表」以本区块为权威来源，正文为详情补充。正文变更涉及下列字段时必须同步更新本区块。

| 字段 | 内容 |
|------|------|
| 设备清单 | 五层实体：MH ×1（主控）+ VMM ≤2（V001/V002）+ CMM（电流通道 CH001~CH096）+ Meter Point ≤96（基线）/≤192（期望）+ Channel Instance；Gateway 功能下挂 Modbus 设备 ≥32 台；虚拟设备 ≥16 台（算术派生，对北向透明） |
| 协议与端口 | 北向：Modbus TCP 502 / Modbus RTU Server（与南向 Client 共用 RS485，角色用户配置）、SNMP v2c/v3 161（可改 16100~16199）、BACnet/IP 47808、MQTT（Publisher）、AcuCloud、Email/HTTP POST/FTP；内部 AccuBus 私有总线；**OPC UA / AWS IoT / Azure IoT / IPv6 双栈 = 下一 Sprint（本轮范围外）** |
| 接线方式 | 6 种：1P2W / 1P3W / 2E3W Network / 2E3W Delta / 3P3W / 3P4W；**3 Element 4 Wire Delta 不支持** |
| 容量与边界 | 基线档：VMM 1~2 / Channel ≤96 / MP ≤96 / 无双 VMM 镜像；期望档：MP ≤192 / 支持镜像；Alarm Monitor 200 组（AND/OR，pickup/dropout 延迟）、Alarm Log ≥65535 条；Datalog 3 实例（1s+，可 Post）；Trend Log 5min；PTP 下行分发 ≤0.5μs；用户 ≤32 / 角色 ≤16 |
| 共存/联动约束 | 时间同步 PTP 优先于 NTP、RTC Fallback；RSTP 默认禁用（启用时双口冗余环网）；固件 A/B 双分区回滚（MH 走 HTTPS、VMM/CMM 走 AccuBus 下行）；IEC 61000-4-30 Class A 聚合 |
| 安全与默认状态 | HTTPS；TLS 1.2/1.3；Password Policy；网页仅英文、响应式（桌面≥1280px/平板≥768px） |
| 高频缺陷模式 | → 详见 bugs/INDEX.md（如有） |
| 测试环境要点 | **当前为纯前端联调环境** http://192.168.2.94:3030（http 非 https、带端口），无后端，设备数据类页面/接口报错属正常；顶级导航 About/Commissioning/Monitoring/Settings/Maintenance；UI 选择器沉淀见下方「UI 选择器沉淀」节 |

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

## UI 选择器沉淀

RPP 从 AcuHMI-1-7 迁移了 BacnetIP / Wiring_check / aws_iot / data_log 四个 UI 模块。RPP 前端页面（http://192.168.2.94:3030，**http 非 https，带端口 3030**）与 HMI 同源 SPA，字段名/交互基本一致，**主要差异在顶级导航分组**：HMI 单一设备菜单（"AcuHMI"/"AcuHMI-1-7"）在 RPP 被拆成多个顶级组。当前 RPP 环境为纯前端联调、无后端，登录可用，设备数据类页面/接口报错属正常。

### 通用事实

- 登录页字段与 HMI 相同：textbox `Enter User Name` / `Enter Password` / button `Sign In`；产品标题 `Remote Power Panel`。
- 登录成功判据：顶部导航出现 `Logout`（或 About/Commissioning/Monitoring/Settings 组）。**不能再用 "AcuHMI" 文本判据**。
- 顶级导航组（`.nav-item` / `.nav-item-menu` 文本）：**About / Commissioning / Monitoring / Settings / Maintenance**（+ Logout）。
- 各协议路由总表：Modbus=`/protocols/modbus/{modbusConfig|modbusDeviceList|paramsMapping|passThrough|logicalParameterMapping}`；SNMP=`/protocols/snmp`；BACnet/IP=`/protocols/bacnet`；MQTT=`/protocols/mqtt/{general|credential|ssl|testament|deviceToPublish}`；EtherNet/IP=`/protocols/etherNetIP`；AWS IoT=`/protocols/awsIot`；Azure IoT=`/protocols/azureIot`。

### 各模块选择器沉淀文档

| 模块 | 文档 | 顶级导航组 |
|------|------|-----------|
| BACnet/IP | [requirements/context/bacnetip_context.md](requirements/context/bacnetip_context.md) | Settings → Protocols |
| AWS IoT | [requirements/context/aws_iot_context.md](requirements/context/aws_iot_context.md) | Settings → Protocols（Virtual Devices 在 Monitoring） |
| Data Log | [requirements/context/data_log_context.md](requirements/context/data_log_context.md) | Monitoring |
| Wiring Check | [requirements/context/wiring_check_context.md](requirements/context/wiring_check_context.md) | Maintenance → Diagnostics |
| User and CT（接线方式/VA/User Channel） | [requirements/context/UserAndCT_context.md](requirements/context/UserAndCT_context.md) | Physical Devices → 设备 → Settings → User and CT |

写/调试上述模块 UI 用例或派 `ui-test-engineer` 现场探查前，先查对应文档，命中即复用（团队约定 #11）。

## 自动化测试指针

Alarm Config 子模块自动化用例位于 `projects/RPP/tests/Alarm/`（30 条（Alarm Config 22 + Alarm Logs 8），当前按 AcuHMI-1-7 真机 192.168.3.71 执行），页面情报沉淀在 `knowledge/gateway/AcuHMI17/requirements/context/` 的 Alarm 相关文档（`Devices_Alarm_ActiveAlarm_context.md` / `Devices_Alarm_AlarmLogs_context.md` / `Devices_PhysicalDevices_AlarmConfig_context.md` / `SystemSettings_AlarmNotification_context.md`）。

Datalog 子模块：`projects/RPP/tests/Datalog/`（5 条（跨度不足的间隔档位按设计约束动态 skip，AcuHMI-1-7 真机执行））
Trend log 子模块：`projects/RPP/tests/TrendLog/`（12 条 skip 占位，AcuHMI-1-7 固件未实装 Trend Log 页面，待 RPP 真机）
