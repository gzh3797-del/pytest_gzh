# AcuHMI-1-7 — 网关自动化测试项目

## 项目背景

AcuHMI-1-7 是工业级 HMI 网关一体机（7" 电容触屏）。南向通过 RS485/以太网采集 Accuenergy 系列电表，北向通过 8 种协议推送至上位机/SCADA/云平台。
当前测试阶段为 **Sprint 3（三期）**；v2.00 综合版需求文档已整合所有 Sprint 需求。

**需求文档：**
- 综合版（权威）：`requirements/raw/AcuHMI-1-7_软件需求规格说明书_v2.00_20260519.docx`
- Sprint 3 规格：`requirements/raw/AcuHMI-1-7 Sprint3 软件需求规格说明书_V1.03_20260423.docx`
- Sprint 3 变更：`requirements/raw/AcuHMI-1-7 Sprint3 软件需求规格变更说明书_v1.00_20260430.docx`
- 需求摘要（综合版）：`requirements/summaries/v2.00_20260519.md`

**Bug 追踪：**
- 精简索引：`bugs/INDEX.md`（A17S-68 ~ A17S-132，共 58 条缺陷）
- 原始导出：`bugs/raw/JIRA (21).csv`

---

## 支持设备（下挂电表）

| 设备 | 固件版本 | 接线检查 | BACnet Harmonic |
|------|---------|---------|----------------|
| AcuRev-4100 | v1.03 | ✅ | ✅ |
| AcuRev-2100 | v1.20 | ✅ | ✅ |
| AcuRev-1300 | v2.24 | ✅ | ✅ |
| Acuvim II（IIW/IIR） | IIV3 v6.36 / LV4 v6.36 | ✅ | ✅ |
| Acuvim 3 | v1.03p19 | ✅ | ✅ |
| H-IO（8DI+8RO） | — | — | — |

---

## 功能模块（v2.00 综合版）

### 南向采集
- Modbus RTU（RS485 COM1/COM2）+ Modbus TCP（以太网）
- 物理设备最多 32 台；Poll Interval 1~3600 秒（默认 60）
- 虚拟设备 32 个（公式四则运算）；Web 设备 100 个（第三方 URL 聚合）

### 北向协议

| 协议 | 关键点 |
|------|-------|
| Modbus TCP | Parameters Mapping (Slave 100) / Pass Through (101~247) / Device Mirror (2~99)；端口 2000~5999 |
| BACnet/IP | 端口 47808~49000；Device Instance 0~4194302；COV；Harmonic 参数；EPICS 下载；Foreign Device/BBMD |
| SNMP | v2c+v3（MD5/SHA，DES/AES）；端口 16100~16199；Trap 4目标；MIB 下载 |
| MQTT | SSL-TLS/LWT/Topic；JSON Payload；Interval 1~600s(11档) |
| AWS IoT | 断网缓存 24~72h；URL 格式限制（小写字母+数字+"-.:/")  |
| Azure IoT | Primary/Secondary Connection String；Device Twin；SSL |
| AcuCloud | HTTPS；Log Interval **固定 5 分钟**；Advanced URL 模式 |

### 数据日志
- 3 个标准 Data Logger + 1 个 Rapid Logger（最小间隔 1 秒）
- 3 条 Post Channel（FTP/SFTP/HTTP/HTTPS）；单 Logger 设备数 ≤32；单 Channel 设备数 ≤16
- 板载存储 32 GB；单 Channel 配额 4 GB

### 告警系统
- 参数阈值 + 设备离线告警；Trigger DO/RO 联动 H-IO（RO 1~8）
- 蜂鸣器：仅有未确认告警时发声，全部确认后自动停止
- 未确认告警页（Alarm Acknowledgement Enable 时展示）
- 邮件通知：最多 3 个收件人，Email Interval 1~10 分钟

### 接线检查（Wiring Check）
- 路径：Diagnosis → Wiring Check（**仅 Web，HMI 不支持**）
- 支持 8 种接线方式（各型号支持情况见 v2.00 摘要附录）
- 导出：`wiring_check_Result.csv`
- Meter Point 命名：多回路表 → "User Channel N: {Description}"；单回路表 → 设备名

### 模板管理
- 官方（只读）+ 自定义（基于 Typical Energy Meter V2）；.def 加密导入导出
- 命名规则：`TemplateName_版本号_Firmware_版本号`（大版本如 v1.03）
- 向后兼容；网关升级后按 Firmware Version 自动匹配模板

### 安全合规（出厂默认）
- SSH / Modbus 北向 / Remote Access / 所有云协议 / Data Logger：全部 **Disable**
- admin 密码：`Admin@` + SN末6位；首次登录强制修改；EULA 强制接受
- 禁止固件降级；升级失败 → Emergency Mode 回退
- IP Allow List；配置文件加密导入导出

---

## 网络配置（默认）
- Eth1：DHCP 禁用，IP 192.168.8.101/255.255.255.0，GW 192.168.8.1
- Eth2：DHCP 自动
- DNS1: 8.8.8.8；DNS2: 8.8.4.4

---

## 当前活跃 Bug（未关闭）

| ID | 模块 | 问题 | 状态 |
|----|------|------|------|
| A17S-132 | Wiring Check | IIW/Acuvim3 接线检查显示 "No Supported" | CREATED |
| A17S-129 | BACnet/IP | 4100 上传 1869 参数耗时 40 分钟 | TO BE VERIFIED |
| A17S-128 | Data Log | Trend Log Time Interval 选项与实际不符 | TO BE VERIFIED |
| A17S-126 | Wiring Check | IIW 2E3W 1Phase 接线检查结果错误 | TO BE VERIFIED |
| A17S-121 | Wiring Check | 配置 Ia/b=1A 时 Ic 缺失，接线检查显示 "-" | CREATED |
| A17S-119 | PassThrough | Acuview2 工具概率连接失败 | SELF-TESTING |
| A17S-118 | Modbus | 系统监听两个重复 502 端口 | SELF-TESTING |
| A17S-117 | Modbus | 关闭 Modbus Config 后 Device Mirror 仍能通信 | SELF-TESTING |
| A17S-116 | BACnet/IP | BACnet 参数单位显示 "Square Meters" | TO BE VERIFIED |
| A17S-114 | AcuCloud | AcuVIM3 上传有重复参数 | CREATED |
| A17S-113 | AcuCloud | AcuCloud 显示 SN 异常 | CREATED |
| A17S-131 | SNMP | MIB 参数值未做系数转换与 Realtime 不一致 | TO BE VERIFIED |
| A17S-130 | AWS IoT | AWS 推送 JSON 设备 name/model 字段为空 | TO BE VERIFIED |
