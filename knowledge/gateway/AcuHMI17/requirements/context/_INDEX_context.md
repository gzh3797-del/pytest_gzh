# AcuHMI-1-7 网关 — 页面上下文索引

> 产品：AcuHMI-1-7（HMI Display and Gateway，Accuenergy）
> 站点：https://192.168.3.71（自签名证书，需在浏览器"高级→继续前往"）
> 登录：admin / Admin@080066
> 用途：页面结构上下文，供 AI 将手工用例转换为自动化用例（每个可路由子页一个文件）。
> 颗粒度：每个"可独立路由/操作"的子页一个 `Prefix_SubPage_context.md`。

## 通用说明（所有页面适用）

- **技术栈**：Vue 3 SPA + Element Plus，hash 路由（`/#/...`）。
- **Element-Plus radio 坑**：`<input type=radio>` 被 `.el-radio__inner` 遮挡，点击 input 会超时；**须点 `label`**（`page.locator('label').filter({hasText:/^Enable$/})`）。
- **两个导航上下文**：header 的 `AcuHMI-1-7`（设置侧）与 `Devices`（设备侧）切换；部分设备侧路由（如 `/physicalDevices`）直接 goto 会被重定向到 dashboard，须经左导航点击进入。
- **快照即全量**：Playwright accessibility snapshot 抓整棵 DOM（不受视口限制）；懒加载内容先滚动到底再快照。
- **破坏性操作**（自动化默认不实际执行，仅验证二次确认）：Factory Reset、Reboot System、Reset System Logs、Clear Logs、Firmware Upload、Delete。
- **实测测试情报节**：SystemSettings（8）/ DataLog（7）/ Templates（4）/ VirtualDevices（1）共 20 个子页文档已合入 `## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）` 节 —— 含 API 端点、blur-vs-Save 校验时机、Element-Plus 真实选择器与坑（radio 坐标兜底 / `.el-select` aria-controls / `.c_common_table` nth / “Post Channel N”带空格 / 自定义 Yes/No 确认 / `.el-message--success` toast）、异步 reload 轮询、跨页遍历、参考 pytest 用例路径。此前的 4 个模块级聚合文档（`systemsettings/datalog/templates/virtualdevice_context.md`）已并入并删除。

## 设置侧 (AcuHMI-1-7)

### Protocols
- `Protocols_Modbus_ParamsMapping_context.md` ★（用户指定入口）
- `Protocols_Modbus_ModbusConfig_context.md`
- `Protocols_Modbus_DeviceList_context.md`
- `Protocols_Modbus_DeviceMirror_context.md`
- `Protocols_Modbus_PassThrough_context.md`
- `Protocols_MQTT_General_context.md`
- `Protocols_MQTT_UserCredential_context.md`
- `Protocols_MQTT_SSL_TLS_context.md`
- `Protocols_MQTT_LastWillTestament_context.md`
- `Protocols_MQTT_TopicParameterSelection_context.md`
- `Protocols_BACnet_IP_context.md`
- `Protocols_SNMP_context.md`
- `Protocols_AWS_IoT_context.md`
- `Protocols_Azure_IoT_context.md`

### System Settings
- `SystemSettings_DateTime_context.md`
- `SystemSettings_Network_context.md`
- `SystemSettings_AccessControl_context.md`（whitelist）
- `SystemSettings_Email_context.md`
- `SystemSettings_AlarmNotification_context.md`
- `SystemSettings_CertificateManagement_context.md`
- `SystemSettings_ConfigurationManagement_context.md`
- `SystemSettings_RemoteAccess_context.md`

### Templates
- `Templates_TemplateList_context.md`
- `Templates_Import_context.md`
- `Templates_NewTypicalTemplate_context.md`
- `Templates_TypicalTemplateConfig_Edit_context.md`（动态编辑）

### Maintenance
- `Maintenance_SystemStatus_context.md`
- `Maintenance_EventLog_context.md`

### Diagnostics
- `Diagnostics_ConnectionTest_context.md`
- `Diagnostics_NetworkStatus_context.md`
- `Diagnostics_RSTPStatus_context.md`
- `Diagnostics_HostLookup_context.md`
- `Diagnostics_NTPSyncTest_context.md`
- `Diagnostics_ModbusDebugLog_context.md`
- `Diagnostics_Debug_context.md`
- `Diagnostics_WiringCheck_context.md`

### User Management
- `UserManagement_General_context.md`
- `UserManagement_UserConfiguration_context.md`（list/add/edit）
- `UserManagement_RoleConfiguration_context.md`（list/add/edit）
- `UserManagement_PasswordPolicy_context.md`
- `UserManagement_PasswordManagement_context.md`（list/edit）

### 顶级
- `FirmwareUpdate_context.md`
- `About_context.md`

## 设备侧 (Devices)
- `Devices_Dashboard_context.md`
- `Devices_PhysicalDevices_List_context.md`
- `Devices_PhysicalDevices_AddDevice_context.md`（add/search/edit 共用表单）
- `Devices_PhysicalDevices_AlarmConfig_context.md`（设备详情页 Alarm Config 子页）
- `Devices_PhysicalDevices_DataLog_context.md`（设备详情页 Logs → Data Log 子页）
- `Devices_PhysicalDevices_TrendLog_context.md`（设备详情页 Logs → Trend Log 子页，固件未实装）
- `Devices_VirtualDevices_context.md`（list/add/details）
- `Devices_WebDevices_context.md`
- `Devices_Alarm_ActiveAlarm_context.md`
- `Devices_Alarm_AlarmLogs_context.md`
- `Devices_DataLog_AcuCloud_context.md`
- `Devices_DataLog_DataLogger_context.md`（覆盖 Data Logger 1/2/3）
- `Devices_DataLog_DataLogParameterConfig_context.md`
- `Devices_DataLog_RapidLogger_context.md`
- `Devices_DataLog_PostChannel_context.md`（覆盖 Post Channel 1/2/3）
- `Devices_DataLog_DataLogManagement_context.md`
- `Devices_DataLog_PostHistoricalData_context.md`

## 局限与未尽项

- **动态详情页**（`deviceDetails/:deviceKey/:currentKey`、`virtualMeterDetails/:key`、`typicalTemplateConfig/:type/:id`、`passwordManagement/edit`、`roleConfiguration/add|edit`、`userConfiguration/edit`）：结构随所选实体动态生成，本索引以其 list/add 页文档为准并标注；深度字段以进入具体实体后为准。
- **代表性合并**：Data Logger 1/2/3 合并为一份文档；Post Channel 1/2/3 合并为一份文档（结构一致，仅编号不同）。
- **条件分支的二级字段**：部分 Enable 后/协议切换后出现的字段（如 Network 的 DHCP=Manual、MQTT SSL 上传、Protocol=TCP）已尽量触发并记录；个别未逐一穷尽的组合在对应文档内注明"运行时确认"。
- **超大表**（如 ParamsMapping 选中设备后 1059 参数行、Templates/EventLog/ModbusDebugLog 的数百页）仅记录结构与规律，未逐行导出。
- **测试情报覆盖范围**：`## 实测测试情报` 节已覆盖 SystemSettings / DataLog / Templates / VirtualDevices（2026-07-03）与 Alarm（含设备详情 Alarm Config，2026-07-17）共 5 个模块。**其余模块（Protocols / Maintenance / Diagnostics / UserManagement / Physical & Web Devices 等）的 API 端点、blur-vs-Save 校验时机、框架真实选择器等情报待后续联机实测补齐**（当前这些文档仅含结构与通用 Element-Plus 坑）。
