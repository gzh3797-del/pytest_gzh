# ACM-41-WEB2 Sprint2 变更需求摘要（v1.00，2026-06-25）

> 来源：《AcuRev-4100-WEB2 Sprint2 软件需求规格说明书》v1.00（生效日期 2026-06-16，发布 2026-06-24）
> 原件：requirements/raw/AcuRev-4100-WEB2 Sprint2 软件需求规格变更说明书_v1.00_20260625.docx
> 目的：同步 AcuHMI-1-7 市场反馈共性问题需求，本版本共 21 个变更项（3.1~3.21），修订类型均为 M（Modified，含部分新增子功能）
> 说明：3.5（MQTT 测试工具改用 MQTTX 客户端）不涉及产品行为变更，测试范围外，未计入下方 20 项功能变更
> 测试用例摘要 → ../../testcase/ACM-41-WEB2_Sprint2变更需求_测试用例.md（119 条用例，覆盖本文 20 个变更项）

---

## 一、日志与 Checkpoint（3.1 / 3.6 / 3.11 / 3.13 / 3.15 / 3.17）

| 章节 | 功能点 | 要点 |
|------|--------|------|
| 3.1 | Checkpoint 默认值改名 | Trend Log Management、Data Log 两个页面的 Checkpoint 默认选项名由 "Current" 改为 "Default Configuration" |
| 3.6 | Datalog CSV 表头改名 | 导出 CSV 首行首列列名由 TimeTag 改为 Time |
| 3.11 | Trend Log 默认打开 | Trend Log 功能默认状态改为开启 |
| 3.13 | Trend Log 按间隔显示能量累积值 | 图表应显示每个时间间隔内的能量**增量**，而非持续累加的读数；第一个时间点不显示；遇空值时，空值之后的时间点与空值之前最近一个有效时间点相减（不是减 0） |
| 3.15 | PQ 事件日志参数重命名 | 电压/电流最大最小值列名由 Par1/Par2/Par3 改为 Phase A / Phase B / Phase C |
| 3.17 | 能量/需量/报警重置记录 Event Log | Energy reset / Demand reset / Alarm reset 均需记录电表事件日志；Message 格式：`Setting-Management: <Action>, success/failed, device=<DeviceName>, SN=<SerialNumber>`（示例：`INFO Setting-Management: Energy reset, success, device=AcuRev-4100-1, SN=MAA1234567890`） |

## 二、设备管理（3.2 / 3.4 / 3.7）

| 章节 | 功能点 | 要点 |
|------|--------|------|
| 3.2 | 设备列表显示已配置参数数量 | Data Log、MQTT、AWS IoT、Azure IoT、BACnet IP 五个设备选择列表新增"已配置参数个数"显示 |
| 3.4 | 模板/固件版本不匹配提示（原"显示 Loading 整改"） | 新增设备时后端校验设备固件版本与所选模板是否匹配；Device List 页新增 Template、Firmware 版本号两列；不匹配时设备名前显示黄色三角标，悬停提示 "The selected template is incompatible with the current firmware version. Please verify the firmware version and select a compatible template."；设备 Setting 页与 Metering 页均需提示 "The device firmware version doesn't match the template version. Data readings may be inaccurate." |
| 3.7 | Device Mirror 和 Pass Through 测试 | 新增测试项：设备开启 Device Mirror / Pass Through 后，通过该方式添加到网关中的设备需能正常访问配置（覆盖三种组网场景，详见 context.md 设备管理节） |

## 三、SNMP（3.3 / 3.8 / 3.20 / 3.21）

| 章节 | 功能点 | 要点 |
|------|--------|------|
| 3.3 | SNMP 重复数据问题解决 | 设备上下线 Trap 消息应每次只发一条（原成对出现且时间戳相同）；修复 MIB 文件在 ManageEngine MibBrowser 中出现两个相同节点名称的问题；虚拟设备离线信息不应通过 SNMP Trap 上报 |
| 3.8 | Modbus 名称拼写错误修改 | "Mddbus" 改为 "Modbus" |
| 3.20 | Alarm Trap 支持参数告警 | 现状仅支持设备在线/离线状态的 Alarm Trap 上传；新增支持设备参数告警的 Trap 上报 |
| 3.21 | SNMP 协议 Description 修改 | 参数 Description 字段由占位文案 "reading parameter" 改为参考模板 description 列的真实物理含义描述 |

## 四、BACnet/IP（3.14）

**BACnet 下 CSV 文件动态生成下载**：BACnet/IP 页面新增 "CSV File Download" 按钮，与 "EPICS File Download" 并排。

| 约束项 | 内容 |
|--------|------|
| 数据来源 | CSV 内容须基于当前 Parameter Config 配置实时动态生成（与 EPICS 一致），仅勾选参数的设备参与导出 |
| 一致性 | CSV 与 EPICS 必须基于同一份内存对象表生成：同一参数在两文件中的 Object Type / Object Instance / Object Name 完全相同；同一时刻下载的两个文件描述同一组对象 |
| 文件命名 | `ACM-41-WEB2_SN号_BACnet_PointList_YYYYMMDD_HHMMSS.csv` |
| CSV 列 | Device Name、Serial Number、Device Instance、Parameter Category（realtimeRange/demandRange/energyRange/thdRange/diStatusRange/diCounterRange/roStatusRange/doStatusRange/harmonicRange/sequenceRange 共10类）、Parameter Name（与 UI 一致）、BACnet Unit（EngineeringUnits 枚举，与 EPICS 字段一致）、Object Type（analog-input）、Object Instance、Object Name（设备SN号＋Postlabel）、Description（≤64字符）、COV Enable（Yes/No）、COV Increment（未启用则空） |

## 五、告警配置（3.9）

**增加 Alarm State 配置项**：设备的 DI Status 类参数告警配置新增 Alarm State 配置项，支持 0 / 1 可勾选；DI Status 为所勾选值时产生告警并记录 Alarm Log，触发时 Reason 记录为 "ACTIVE"。

## 六、系统与网络（3.12 / 3.18 / 3.19）

| 章节 | 功能点 | 要点 |
|------|--------|------|
| 3.12 | 更改 WiFi 默认密码 | 默认密码改为 "accuenergy"，与现有产品线（AXM-WEB2、2100-WEB2、810）保持一致 |
| 3.18 | WEB2 设备 RSTP IP 在主表展示 | 主表 Setting 页 → WEB2 页面下，WEB2 启用 RSTP 时，将 RSTP 的 IP 地址在 ETH1 处展示 |
| 3.19 | mDNS 支持（WEB2 自身可被发现） | WEB2 新增支持 mDNS，可通过 `SN.local` 主机名免 DNS 访问设备；新增 "mDNS Enable" 开关，默认开启。**注意与已有需求的区别**：本项是 WEB2 自身对外广播可被发现，不同于 sprint2_requirements.md §1.1 中"网关扫描下挂设备"的 mDNS 功能（该功能标注"暂不做"，仍不受影响） |

## 七、界面与交互（3.10 / 3.16）

| 章节 | 功能点 | 要点 |
|------|--------|------|
| 3.10 | 显示型号名称 "ACM-41-WEB2" | 网页界面名称全局由 "AcuRev-4100-WEB2" 改为 "ACM-41-WEB2"，仅 About → Information 页保留 "-D" 型号后缀；EULA 弹窗标题/正文中的 "AcuRev-4100-WEB2" 字样同步改为 "ACM-41-WEB2" |
| 3.16 | 手动波形触发按钮通知 | 选择用户通道并点击 Trigger 手动触发波形时，页面需提示是否下发成功，文案："Waveform recording triggered successfully."（原无提示，易导致用户重复点击产生多余波形） |

---

## 21 变更项与测试范围对照

| 章节 | 变更项 | 测试范围 |
|------|--------|----------|
| 3.1~3.4 | 见上表 | 覆盖 |
| 3.5 | MQTT 测试工具调整（改用 MQTTX） | **范围外**（测试工具变更，非产品行为） |
| 3.6~3.21 | 见上表 | 覆盖 |
