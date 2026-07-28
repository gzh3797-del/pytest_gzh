# AcuHMI-1-7 Sprint3 软件需求规格说明书 — 摘要

**原文件：** AcuHMI-1-7 Sprint3 软件需求规格说明书_V1.03_20260423.docx  
**文件编号：** JZZY-A4-RD-50101  
**版本：** V1.03（正式发布）  
**生效日期：** 2026-04-23  
**图片数：** 64 张（见原始 docx）

---

## 版本历史

| 版本 | 日期 | 主要变更 |
|------|------|---------|
| v1.00 | 2026-04-02 | 初稿 |
| v1.01 | 2026-04-03 | 模板管理更新、接线检查性能、BACnet/BBMD、Remote Access Advanced |
| v1.02 | 2026-04-13 | 可编辑通道名、BACnet/IP协议、Wiring Check、AWS IoT UI、蜂鸣器逻辑、Meter Point配置 |
| v1.03 | 2026-04-23 | 虚拟设备北向支持、MeterPoint命名规则、模板管理兼容版本、删除MeterPoint配置需求、明确Dashboard/Reading |

---

## 功能模块概览

### 3.1 页脚邮件更改
- `marketing@accuenergy.com` → `support@accuenergy.com`

### 3.2 Add to Logger（添加设备时分配日志记录器）
- 新增字段：No / Data Log 1 / 2 / 3 / Rapid Logger
- 最大设备数：每个 Data Logger 限 32 台；单个 Post Channel 限 16 台
- Post Channel 内存 4G；数据库存储 32GB

### 3.3 虚拟设备改进
- 已选参数在下拉列表中不重复显示
- AcuCloud、Data Log 支持虚拟设备（默认全选参数）
- AWS IoT、Azure IoT 支持虚拟设备及参数配置
- **不支持**虚拟设备的协议：MQTT、MODBUS Data Mapping、Data Mirror、PassThrough、SNMP、BACnet/IP

### 3.4 HMI操作屏虚拟设备参数过滤
- 增加输入字符过滤功能，解决 AcuRev-4100 全参数无法完整滚动查看的问题

### 3.5 迁移 PX-EMD-G 功能
- Data Log 参数分类配置：Realtime / Energy / Demand / Harmonic Spectrum / Sequence and Unbalance
- 白名单术语：Whitelist → Access Control，Whitelist Enable → IP Allow List Enable
- HMI 显示屏时间/日期配置
- 高波特率（115200 baud）兼容性修复
- 设备导入导出功能
- 设备添加后支持修改序列号和模板

### 3.6 可编辑用户/输入通道名（AcuRev-4100）
- User Channel 1~12 新增 Description 列（≤20 ASCII 字符，默认空）
- Metering 页面显示格式：`User Channel 1: {Description}`

### 3.7 接线检查（Wiring Check）& Meter Point

#### 3.7.1 Wiring Check
- 支持设备：AcuRev-4100、Acuvim II、AcuRev-2100、Acuvim 3、AcuRev-1300
- 不在 HMI 显示屏支持
- 离线设备跳过，显示 "Offline"；未检查显示 "Not Checked"
- 结果按相位红/绿标色展示
- 支持导出 CSV（`wiring_check_Reault.csv`）
- "Faulty Only" 改名为 "Show Only Issues"
- 接线检查逻辑参考：`接线检测总表_ver1.03_20260409.xlsx`
- 支持的接线方式：3E4WY、1E2W、2E2WDelta、2E3WDelta、3E3WDelta、3E4WDelta、2E3W1Phase、2.5E4WY、2E3WNetwork

**Meter Point 命名规则：**
- 多回路表（2100/4100）：`UserChannel N: {Description}`
- 单回路表（1300、IIW、Acuvim3）：直接使用设备名称

#### 3.7.2 Meter Point Configuration — **已删除**（V1.03 删除）

### 3.8 虚拟设备参数配置
- 与 Metering 页面显示一致

### 3.9 模板升级兼容
- 增加参数不删除设备配置/告警配置；修改/删除参数特殊处理

### 3.10 协议需求

#### 3.10.1 BACnet/IP
- 作为北向从站，支持 BACnet IP 协议数据获取
- 支持同时作为 BACnet 网关读取 Modbus 设备（RS485/Modbus TCP/BACnet IP）
- 配置项：Enable、Port（47808~49000）、Network Number（1~65534）、Device Object Name、Device Instance（0~4194302）、APDU Timeout（3~60s）、APDU Retries（0~10）
- 外部设备（Foreign Device）：BBMD IP、BBMD Port（47808~49000）、Time To Live（5~1440 min）
- 参数配置：Polling Enable（支持一键全使能）、COV Enable、COV Increment（≥0.000，3位小数）
- COV Batch Update 批量操作

#### 3.10.2 SNMP
- 支持 SNMPv2c 和 SNMPv3
- 端口：默认161，范围 16100~16199
- SNMPv3 配置：Username、User Password（≥8字符）、Auth Protocol（MD5/SHA）、Privacy Protocol（DES/AES/无）、Privacy Password（≥8字符）
- Trap：支持4个 Trap Target（仅 IP 地址）；Report Buffer Size（0~30）；Report Hold Time（0~300s）
- 支持下载 MIB 文件

#### 3.10.3 MQTT
- 配置页：General、User Credential、SSL/TLS、Last Will and Testament、Topic and Parameter Selection
- General：Broker Address（1~128）、Broker Port（0~65535）、Client ID（≤40）、Keep Alive（10~600s）、Timeout（3~120s）、Clean Session
- SSL/TLS：CA File、Cert File、Key File
- Last Will：Topic、QoS（0/1/2）
- Topic：Base Topic = 基础主题 + / + 设备序列号；Interval 可选 1~600s；Retained（True/False）
- 支持设备参数按类发布：Basic Parameter、Demand、Power Quality、Energy

### 3.11 已修复 Bug（来自 v1.01p14/p33/p36，共13个）
详见各 Jira 条目（A17S-80 系列）

### 3.12 数据上传

#### 3.12.1 AcuCloud
- 支持 Advanced 模式（URL 可自定义）：`?showAdvanced=true`
- Log Interval 固定为 5 分钟
- "Devices Selection To Mapping" → "Devices Selection"

#### 3.12.2 AWS IoT
- 支持按设备参数单独配置
- URL：20~128字符，仅小写字母/数字及 `-./:`
- Interval：1~600s（下拉）
- 断网缓存 24~72 小时

#### 3.12.3 Azure IoT
- 类似 MQTT 的设备参数配置方式
- Enable SSL（X509）
- Interval：1~600s

### 3.13 Remote Access
- 新增 Advanced 模式，支持自定义 URL

### 3.14 蜂鸣器逻辑重设计
- 有未确认告警日志时发声；否则不发声
- 删除物理 Alarm Reset 按钮
- 新增 "Alarm Acknowledgement Enable" 功能
- 新增 Unacknowledged Alarms 页面（字段：Timestamp、Device Name、SN、Monitor Label、Parameter、Status、Reason、Acknowledge 按钮）
- Alarm Logs 新增：Ack Status、Trigger DO Device、DO、Trigger RO Device、RO

### 3.15 系统通信异常提示语
- 统一改为英文：`Network Error: Please check your internet connection or configuration.`

### 3.16 模板管理
- 模板名格式：`TemplateName_版本号_Firmware_版本号`（如 `v1.03p06`，只显示大版本）
- 向后兼容（新模板支持老模板功能）
- 支持按固件版本自动匹配模板
- 设备迁移通过导入自动识别版本

**支持电表版本：**
| 设备 | 支持版本 |
|------|---------|
| AcuRev-4100 | ACCUENERGY v1.03 |
| AcuRev-2100 | v1.20（MC/NMC 区分 mA/mV/RCT CT 型号） |
| AcuRev-1300 | v2.24 |
| Acuvim II | IIV3 v6.36 / LV4 v6.36 |
| Acuvim 3 | v1.03p19 |

- DataLog/TrendLog 支持备份及 Checkpoint 配置（Current / 备份时间点）
- Log Type：Energy / Realtime

### 3.17 H-IO 设备支持
- 自动添加（未接入时不显示）
- 告警可触发 H-IO RO 1~8
- Alarm Log 增加 Trigger DO/RO Device 信息
- H-IO Reading 页面：RO 1~8 状态、DI 1~8 状态
- About 页面显示 H-IO 版本信息（离线时隐藏）

### 3.18 AcuRev-4100 General 新增 PT1/PT2
- PT1 范围：50~1,000,000，默认 480
- PT2 范围：50~830，默认 480
- 约束：PT1 ≥ PT2；50 ≤ Nominal Voltage ≤ PT1
- 影响范围：电压/功率/能量计算、告警、Datalog、Waveform（以一次侧为准）
- Nominal Voltage 寄存器从 16 位改为 32 位；Nominal Current 范围 1~50000A 默认 1000A
- PT1/PT2/Nominal Voltage/Nominal Current 受铅封保护

### 3.19 Dashboard & Reading
- Dashboard：最多 32 台设备 / 每页最多 96 个 Meter Point
- Reading 页面 User Channel 下拉 → Meter Point 下拉（格式：`{DeviceName}-UserChannel n: {Description}`）
