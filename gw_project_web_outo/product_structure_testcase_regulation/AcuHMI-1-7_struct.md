# AcuHMI-1-7 产品页面结构文档

**生成时间：** 2026-05-12  
**系统地址：** https://192.168.2.199  
**固件版本：** v1.03p02  
**硬件版本：** v1.05  
**序列号：** AHI260110001  
**截图总数：** 129 张  

---

## 一、页面层级结构（树状图）

```
AcuHMI-1-7
├── [顶部导航栏]
│   ├── About                          ─ 设备信息页（独立页面）
│   ├── Devices                        ─ 切换至 Devices 区域
│   └── AcuHMI-1-7 (System Settings)  ─ 切换至 System Settings 区域
│
├── [Devices 区域] ─ 左侧边栏
│   ├── Dashboard                      ─ 首页总览
│   ├── Physical Devices               ─ 物理设备管理
│   │   └── Add Device (全页表单)      ─ 添加设备
│   ├── Virtual Devices                ─ 虚拟设备管理
│   │   └── Add Virtual Device (全页表单) ─ 添加虚拟设备
│   ├── Web Devices                    ─ Web 设备管理
│   │   └── Add Device (弹窗对话框)    ─ 添加 Web 设备
│   ├── Alarm                          ─ 报警管理
│   │   ├── Unacknowledged Alarms      ─ 未确认报警列表
│   │   └── Alarm Logs                 ─ 报警历史记录
│   └── Data Log                       ─ 数据日志
│       ├── Data Loggers 1             ─ 数据记录器配置
│       ├── Data Log Management        ─ 日志管理
│       ├── Post Historical Data       ─ 历史数据上传
│       └── AcuCloud                   ─ AcuCloud 云平台配置
│
└── [System Settings 区域] ─ 左侧边栏
    ├── System Settings                ─ 系统设置（8个标签页）
    │   ├── Date & Time                ─ NTP时间同步、时区配置
    │   ├── Network                    ─ 以太网1/2、DHCP、DNS
    │   ├── Access Control             ─ 访问控制
    │   ├── Email                      ─ 邮件服务配置
    │   ├── Alarm Notification         ─ 报警通知设置
    │   ├── Certificate Management     ─ 证书管理
    │   ├── Configuration Management   ─ 配置文件备份/恢复
    │   └── Remote Access              ─ 远程访问设置
    ├── Templates                      ─ 设备模板管理
    │   ├── Template List              ─ 模板列表
    │   ├── New Typical Energy Meter Template ─ 新建标准电表模板
    │   └── Import                     ─ 导入模板
    ├── Protocols                      ─ 协议配置
    │   ├── Modbus (下拉菜单)
    │   │   ├── Modbus Config          ─ Modbus 基础配置
    │   │   ├── Device List            ─ Modbus 设备列表
    │   │   ├── Parameters Mapping     ─ 参数映射
    │   │   ├── Pass Through           ─ 透传配置
    │   │   └── Device Mirror          ─ 设备镜像
    │   ├── SNMP                       ─ SNMP 协议配置
    │   ├── BACnet/IP                  ─ BACnet/IP 协议配置
    │   ├── MQTT (下拉菜单)
    │   │   ├── General                ─ MQTT 基础配置（Broker、端口）
    │   │   ├── User Credential        ─ 用户认证
    │   │   ├── SSL/TLS                ─ 安全传输层配置
    │   │   ├── Last Will and Testament ─ 遗嘱消息配置
    │   │   └── Topic and Parameter Selection ─ Topic 和参数选择
    │   ├── AWS IoT                    ─ AWS IoT 连接配置
    │   └── Azure IoT                  ─ Azure IoT Hub 配置
    ├── Maintenance                    ─ 系统维护
    │   ├── System Status              ─ CPU/RAM/磁盘状态（含重启按钮）
    │   └── Event Log                  ─ 系统事件日志
    ├── Diagnostics                    ─ 网络诊断工具
    │   ├── Network Status             ─ 网络接口状态
    │   ├── RSTP Status                ─ RSTP 协议状态
    │   ├── Host Lookup                ─ 主机名解析
    │   ├── Connection Test            ─ 连接测试（Ping）
    │   ├── NTP Sync Test              ─ NTP 同步测试
    │   ├── Modbus Debug Log           ─ Modbus 调试日志
    │   ├── Debug                      ─ 调试工具
    │   └── Wiring Check               ─ 接线检测
    ├── User Management                ─ 用户管理
    │   ├── General                    ─ 通用设置
    │   ├── User Configuration         ─ 用户账号管理（增删改）
    │   ├── Role Configuration         ─ 角色权限配置
    │   ├── Password Policy            ─ 密码策略设置
    │   └── Password Management        ─ 密码管理
    └── Firmware Update                ─ 固件升级（手动上传，当前 v1.03p02）
```

---

## 二、各模块功能说明

### 1. Dashboard（仪表板）
- **Offline Devices** 区域：显示离线设备列表（Device Name、Interface、Protocol、Serial Number）
- **Alarms** 区域：显示当前报警（Device Name、Alarms、Interface、Protocol、Serial Number）
- 底部显示系统启动时间（Up since）

### 2. Physical Devices（物理设备）
- 设备列表表格：Device Name、Interface、Protocol、Model、Serial Number、Poll Interval、Status、Active Alarms、Action
- **+ Add Device** 按钮 → 完整添加表单（非弹窗）：
  - 设备名称、序列号、模板选择
  - 通信协议：RTU 或 TCP
  - 端口（COM 1）、Modbus ID、波特率（38400）、数据位（8）、奇偶校验、停止位、请求超时
- 轮询间隔配置（全局）+ Save 按钮
- Download List 按钮（下载设备列表）
- 行内 Action：编辑、删除（⚠️ 危险操作）

### 3. Virtual Devices（虚拟设备）
- 与 Physical Devices 类似的列表结构
- **Add Virtual Device** 按钮 → 完整添加表单

### 4. Web Devices（Web 设备）
- 设备列表：Device Name、Serial Number、Action
- **+ Add Device** 按钮 → **弹窗对话框**（与 Physical 不同）：
  - Device Name、Serial Number、Model
  - URL（支持 https:// 前缀）

### 5. Alarm（报警）
- **Unacknowledged Alarms**：未确认报警实时列表
- **Alarm Logs**：历史报警记录，支持筛选

### 6. Data Log（数据日志）
- **Data Loggers 1**：数据记录器参数配置
- **Data Log Management**：日志文件管理
- **Post Historical Data**：历史数据推送配置
- **AcuCloud**：接入 Accuenergy AcuCloud 云平台

### 7. About（关于 - 顶部导航）
- 设备信息表单（可编辑 Name/Location/Description）：
  - Model: AcuHMI-1-7
  - Serial Number、Hardware Version、Firmware Version
  - Last Updated、Ethernet1/2 MAC Address

---

### System Settings 区域

### 8. System Settings > Date & Time
- NTP Enable/Disable（Enable/Disable 按钮）
- Device Clock 显示与手动设置
- NTP Server 1/2/3 配置
- Time Zone（世界地图选择器）

### 9. System Settings > Network
- RSTP Enable 开关
- Default Interface（Outbound Traffic）
- **Ethernet 1**：DHCP（Auto/Manual）、Interface Status、IP/Mask/Gateway
- **Ethernet 2**：同上（当前 Connected，IP: 192.168.2.199）
- DNS 1/2 配置

### 10. System Settings > Access Control
- 访问控制策略配置

### 11. System Settings > Email
- 邮件服务器 SMTP 配置

### 12. System Settings > Alarm Notification
- 报警通知接收人/触发条件配置

### 13. System Settings > Certificate Management
- SSL 证书上传与管理

### 14. System Settings > Configuration Management
- 系统配置文件备份、导出、导入恢复

### 15. System Settings > Remote Access
- 远程访问（VPN/远程桌面）开关与配置

### 16. Templates（模板）
- **Template List**：已有模板列表（Official 分类）
- **New Typical Energy Meter Template**：创建新能源表模板
- **Import**：从文件导入模板（Upload 按钮）

### 17. Protocols > Modbus
- **Modbus Config**：Enable/Disable、Modbus Port（默认502）、数据类型对照表（Bit/Uint16/Int32 等）
- **Device List**：已接入 Modbus 设备
- **Parameters Mapping**：参数地址映射配置
- **Pass Through**：Modbus 透传通道配置
- **Device Mirror**：设备镜像（从站模拟）

### 18. Protocols > SNMP
- Enable/Disable、Trap Enable/Disable
- MIB File Download 按钮

### 19. Protocols > BACnet/IP
- BACnet/IP 协议配置

### 20. Protocols > MQTT
- **General**：Broker 地址、端口、Client ID、Keep Alive 等基础配置
- **User Credential**：用户名/密码认证
- **SSL/TLS**：证书路径、TLS 版本
- **Last Will and Testament**：遗嘱 Topic/消息配置
- **Topic and Parameter Selection**：发布 Topic 与参数绑定

### 21. Protocols > AWS IoT
- AWS IoT Core 连接配置（Endpoint、证书）

### 22. Protocols > Azure IoT
- Azure IoT Hub 连接字符串配置

### 23. Maintenance
- **System Status**：实时 CPU（0.49%）、RAM（381/982 MB）、Disk（97/27244 MB）+ **Reboot System** 按钮（⚠️ 危险操作）
- **Event Log**：系统操作/事件日志

### 24. Diagnostics（诊断）
- **Network Status**：以太网接口实时状态
- **RSTP Status**：RSTP 协议运行状态
- **Host Lookup**：域名解析测试
- **Connection Test**：目标地址 Ping 测试
- **NTP Sync Test**：NTP 服务器同步测试
- **Modbus Debug Log**：Modbus 通信抓包调试
- **Debug**：系统级调试工具
- **Wiring Check**：RS485 接线极性检测

### 25. User Management
- **General**：全局用户设置
- **User Configuration**：用户账号列表（Username、Role、注册日期、过期日期、最后登录时间、状态 Active/Inactive、Lock）
  - 支持 Add User / 编辑 / 删除 / Lock 操作
- **Role Configuration**：用户角色与权限分配
- **Password Policy**：密码复杂度/有效期策略
- **Password Management**：强制改密/密码重置

### 26. Firmware Update
- 显示当前固件版本：v1.03p02
- Manual Update：上传固件文件（Upload 按钮）

---

## 三、页面跳转关系

```
顶部 [Devices]     ──→  Devices 区域（Dashboard）
顶部 [AcuHMI-1-7] ──→  System Settings 区域（Date & Time）
顶部 [About]       ──→  About 独立页面

Devices 区域各页互通：
  Dashboard ←→ Physical Devices ←→ Virtual Devices ←→ Web Devices ←→ Alarm ←→ Data Log

Physical Devices → [+ Add Device]    → Add Device 全页表单（子路由）
Virtual Devices  → [+ Add Virtual Device] → Add Virtual Device 全页表单
Web Devices      → [+ Add Device]    → Add Device 弹窗对话框（同页）

Alarm → [Unacknowledged Alarms / Alarm Logs]  → 左侧子菜单切换
Data Log → [各子页] → 左侧子菜单切换

System Settings 区域各页互通（左侧边栏 + 二级 Tab/子菜单）：
  System Settings → 8 个 Tab 页横向切换
  Templates → 3 个子菜单
  Protocols → Modbus(5) + SNMP + BACnet/IP + MQTT(5) + AWS IoT + Azure IoT
  Maintenance → 2 个 Tab 页
  Diagnostics → 8 个子菜单
  User Management → 5 个 Tab 页
  Firmware Update → 单页
```

---

## 四、截图索引（按模块）

| 模块目录 | 截图数 | 主要内容 |
|---------|--------|---------|
| `00_Login` | 6 | 登录页、填写表单、登录后首页 |
| `01_Dashboard` | 2 | 仪表板总览 |
| `02_Physical_Devices` | 6 | 设备列表、Add Device 表单（含全页） |
| `03_Virtual_Devices` | 6 | 虚拟设备列表、Add Virtual Device 表单 |
| `04_Web_Devices` | 4 | Web 设备列表、Add Device 对话框 |
| `05_Alarm` | 4 | 未确认报警、报警历史 |
| `06_Data_Log` | 7 | 数据记录器、日志管理、历史上传、AcuCloud |
| `07_TopNav_About` | 1 | 设备信息页 |
| `08_TopNav_Devices` | 1 | Devices 顶导 |
| `09_TopNav_AcuHMI-1-7` | 1 | System Settings 入口 |
| `SysSettings_01_System_Settings` | 10 | 8个系统设置Tab页 |
| `SysSettings_02_Templates` | 8 | 模板列表、新建、导入 |
| `SysSettings_03_Protocols` | 37 | Modbus(5)+SNMP+BACnet+MQTT(5)+AWS+Azure |
| `SysSettings_04_Maintenance` | 4 | 系统状态、事件日志 |
| `SysSettings_05_Diagnostics` | 13 | 8个诊断工具子页 |
| `SysSettings_06_User_Management` | 11 | 5个用户管理子页 + Add User表单 |
| `SysSettings_07_Firmware_Update` | 8 | 固件升级页 |
| **合计** | **129** | |

---

## 五、关键 URL 规律

```
# Devices 区域
https://192.168.2.199/#/dashboard
https://192.168.2.199/#/physicalDevice
https://192.168.2.199/#/virtualDevice
https://192.168.2.199/#/webDevice
https://192.168.2.199/#/alarm
https://192.168.2.199/#/dataLog

# System Settings 区域
https://192.168.2.199/#/systemSettings/dateTime
https://192.168.2.199/#/systemSettings/network
https://192.168.2.199/#/systemSettings/...

# Protocols
https://192.168.2.199/#/protocols/modbus/modbusConfig
https://192.168.2.199/#/protocols/modbus/deviceList
https://192.168.2.199/#/protocols/mqtt/general
https://192.168.2.199/#/protocols/mqtt/credential
https://192.168.2.199/#/protocols/mqtt/ssl
https://192.168.2.199/#/protocols/mqtt/testament
https://192.168.2.199/#/protocols/mqtt/deviceToPublish
```

---

## 六、注意事项（⚠️ 危险操作未点击）

以下按钮/操作截图但未执行，避免破坏测试环境：

| 位置 | 操作 | 风险 |
|------|------|------|
| Maintenance > System Status | **Reboot System** | 重启设备 |
| User Configuration 行内 | **Delete** (🗑️) | 删除用户账号 |
| Physical/Virtual Devices 行内 | **Delete** | 删除设备记录 |
| Configuration Management | **Factory Reset / Restore** | 恢复出厂设置 |

---

*文档由 AcuHMI-1-7 页面截图自动化脚本生成，结合人工分析整理。*
