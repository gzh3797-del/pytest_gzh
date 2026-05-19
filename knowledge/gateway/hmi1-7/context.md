# hmi1-7 — AcuHMI-1-7 网关自动化测试项目

## 项目背景

AcuHMI-1-7 是西安加中准源科技有限公司的工业物联网网关产品。Sprint 3 迭代新增北向协议支持（BACnet/IP、SNMP、MQTT）、云端上传增强（AcuCloud/AWS IoT/Azure IoT）、接线检查（Wiring Check）、模板管理、H-IO 设备支持等功能。

**需求文档：**
- 主需求规格：`requirements/raw/AcuHMI-1-7 Sprint3 软件需求规格说明书_V1.03_20260423.docx`
- 变更说明书：`requirements/raw/AcuHMI-1-7 Sprint3 软件需求规格变更说明书_v1.00_20260430.docx`
- 需求摘要：`requirements/summaries/`

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
| H-IO | — | — | — |

---

## Sprint 3 主要功能模块

### 北向协议
| 协议 | 状态 | 关键点 |
|------|------|-------|
| BACnet/IP | 新增 | 端口 47808~49000；Device Instance 0~4194302；COV 支持；支持 Harmonic 参数 |
| SNMP | 增强 | SNMPv2c + SNMPv3（认证/加密）；端口 16100~16199；Trap 4个目标；MIB 下载 |
| MQTT | 新增 | General/Credential/SSL-TLS/LWT/Topic 配置；按设备参数发布 |

### 云端上传
| 服务 | 关键变更 |
|------|---------|
| AcuCloud | Advanced URL 模式；Log Interval 固定 5min |
| AWS IoT | 按设备参数配置；断网缓存 24~72h；URL 格式限制 |
| Azure IoT | 类 MQTT 设备参数配置方式；SSL 支持 |
| Remote Access | Advanced URL 编辑 |

### 接线检查（Wiring Check）
- 路径：Diagnosis → Wiring Check
- 支持接线方式：3E4WY / 1E2W / 2E2WDelta / 2E3WDelta / 3E3WDelta / 3E4WDelta / 2E3W1Phase / 2.5E4WY / 2E3WNetwork
- 导出：`wiring_check_Reault.csv`
- Meter Point 命名：多回路表用 `UserChannel N: {Description}`，单回路表用设备名

### 告警系统
- 蜂鸣器：仅在存在未确认告警时发声
- Alarm Acknowledgement Enable / Unacknowledged Alarms 页面
- Alarm Logs 新增：Ack Status / Trigger DO/RO Device / DO/RO

### 模板管理
- 命名格式：`TemplateName_版本号_Firmware_版本号`（只显大版本如 v1.03p06）
- 向后兼容；按固件版本自动匹配；DataLog/TrendLog 备份 + Checkpoint

### 其他
- **Add to Logger**：添加设备时直接分配 Data Log 1/2/3 / Rapid Logger
- **PT1/PT2**：AcuRev-4100 General 新增（PT1: 50~1000000，PT2: 50~830）
- **H-IO**：自动添加；RO 1~8 告警触发；DI/RO 状态读取
- **可编辑 User Channel**：AcuRev-4100 Description 列（≤20 ASCII 字符）
- **虚拟设备**：AcuCloud / Data Log / AWS IoT / Azure IoT 支持虚拟设备

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
| A17S-117 | Modbus | 关闭 Modbus Config 后 Device Mirror 仍能与 Modbus Poll 通信 | SELF-TESTING |
| A17S-116 | BACnet/IP | BACnet 参数单位显示 "Square Meters" | TO BE VERIFIED |
| A17S-114 | AcuCloud | AcuVIM3 上传有重复参数 | CREATED |
| A17S-113 | AcuCloud | AcuCloud 显示 SN 异常 | CREATED |
| A17S-131 | SNMP | MIB 管理端参数值未做系数转换，与 Realtime 不一致 | TO BE VERIFIED |
| A17S-130 | AWS IoT | AWS 推送 JSON 设备 name/model 字段为空 | TO BE VERIFIED |

