# hmi1-7 Bug 索引

原始 Jira 导出 Excel 存放于 raw/ 目录（权威来源）。
本文件为精简索引，每条一行，供 Claude 快速了解历史问题。

状态说明：CREATED=待处理 | IN PROGRESS=处理中 | TO BE VERIFIED=待验证 | SELF-TESTING=自测中 | CLOSED=已关闭 | REJECTED=已拒绝

| ID | 模块 | 标题摘要 | 状态 |
|----|------|---------|------|
| A17S-132 | Wiring Check | Acuvim IIW 和 Acuvim3 接线检查 Wiring Status 显示 "No Supported" | CREATED |
| A17S-131 | SNMP | SNMP MIB 管理端参数值未做系数转换，与 Realtime 页面不一致 | TO BE VERIFIED |
| A17S-130 | AWS IoT | AWS 推送 JSON 中设备 name/model 字段为空 | TO BE VERIFIED |
| A17S-129 | BACnet/IP | BACnet/IP 上传 4100 全部 1869 个参数需要 40 分钟 | TO BE VERIFIED |
| A17S-128 | Data Log | Trend Log 1小时数据后，Realtime Log Time Interval 选项与实际不符 | TO BE VERIFIED |
| A17S-127 | Wiring Check | 2100 接线方式设为 2E3W 1Phase，接线检查结果显示错误（已关闭） | CLOSED |
| A17S-126 | Wiring Check | Acuvim IIW 接线方式设为 2E3W 1Phase，接线检查结果显示错误 | TO BE VERIFIED |
| A17S-125 | Wiring Check | 1E2W 接线方式下 Meter Point 显示 B/C 相（应只显示 A 相） | CLOSED |
| A17S-124 | SNMP | SNMP 配置页提示语 "Snmp Config Invali" 单词拼写错误（应为 Invalid） | CLOSED |
| A17S-122 | SNMP | 下载的 MIB 文件中出现 acuRev4100Web2 字段，项目名不正确 | CLOSED |
| A17S-121 | Wiring Check | 配置 Ia/b 为 1A，Ic 电流缺失，接线检查结果显示 "-" | CREATED |
| A17S-120 | Wiring Check | 2100 Wiring of Three Phase User 列表中 network 接线方式显示不全 | CLOSED |
| A17S-119 | PassThrough | 开启 PassThrough，使用 Acuview2 工具概率连接失败 | SELF-TESTING |
| A17S-118 | Modbus | 开启 Modbus 功能，系统监听两个重复的 502 端口 | SELF-TESTING |
| A17S-117 | Modbus | 关闭 Modbus Config、开启 Device Mirror，与 Modbus Poll 通信成功（预期失败） | SELF-TESTING |
| A17S-116 | BACnet/IP | BACnet 上传参数单位错误，显示 "Square Meters" | TO BE VERIFIED |
| A17S-115 | AcuCloud | AcuCloud 上传 AcuvimIIW 参数错误 | CLOSED |
| A17S-114 | AcuCloud | AcuCloud 上传 AcuVIM3 有1个重复参数，Excel 出现多个空值列 | CREATED |
| A17S-113 | AcuCloud | AcuCloud 上显示 SN 异常 | CREATED |
| A17S-112 | 用户权限 | View 用户无法正常查看设备数据，页面报错 | CLOSED |
| A17S-111 | AcuCloud | AcuCloud 上传 2100 有3个重复参数，Excel 出现多个空值列 | CLOSED |
| A17S-110 | Data Log | Trend Log Management 下载 Excel 数据间隔为 5min，页面显示 1min | CLOSED |
| A17S-109 | Data Log | 2100 UNBL_I_% 及 AcuvimIIR UNBL_V_% 远端服务器参数值错误 | CLOSED |
| A17S-108 | AcuCloud | AcuCloud 上传 2100 UNBL_I_% 及 AcuvimIIR UNBL_V_% / UNBL_I_% 数据错误 | CLOSED |
| A17S-107 | AWS IoT | AWS IoT 缓存数据推送未按时间戳顺序（已拒绝） | REJECTED |
| A17S-106 | AWS IoT | AWS IoT 勾选设备但未选参数时仍推送无参数消息体 | CLOSED |
| A17S-105 | AWS IoT | AWS IoT Interval 配置与实际上报间隔不一致 | CLOSED |
| A17S-104 | AWS IoT | AWS IoT 配置错误 URL 时 Test Connection 显示连接成功（已拒绝） | REJECTED |
| A17S-103 | BACnet/IP | 4100 新增 PT 设置功能缺失（未找到该功能） | CLOSED |
| A17S-102 | MQTT | MQTT Topic and Parameter Selection 配置刷新后设备勾选丢失 | CLOSED |
| A17S-101 | MQTT / AWS | MQTT 及 AWS IoT Topic 输入框不支持符号 `_` | CLOSED |
| A17S-100 | AWS IoT | 下挂 4100 设备未向 AWS IoT 推送数据 | CLOSED |
| A17S-99 | Data Log | 新添加设备未开始采集，修改 Data Log Parameter Config 产生备份记录 | CLOSED |
| A17S-98 | Data Log | 修改 Data Log Parameter Config 操作 Checkpoint 时设备 dump | CLOSED |
| A17S-97 | BACnet/IP | COV Increment 输入非法值后关闭窗口，影响其他设备参数选择界面报错 | CLOSED |
| A17S-96 | BACnet/IP | BACnet/IP 上传 2100 UNBL_I_% 数据错误 | CLOSED |
| A17S-95 | 模板管理 | 添加设备时 Template 名称格式与需求不一致 | CLOSED |
| A17S-94 | 固件升级 | v1.03p01 升级到 v1.03p02，Acuvim3 / IIW 设备 Realtime 显示 No Data | CLOSED |
| A17S-92 | Azure IoT | Azure IoT 上传证书后 Save，未显示证书信息 | CLOSED |
| A17S-91 | Azure IoT | Azure IoT 同一时间戳推送重复数据 | CLOSED |
| A17S-90 | AWS IoT | AWS IoT Disable 后仍缓存数据 | CLOSED |
| A17S-89 | AWS IoT | AWS IoT 取消设备勾选后仍推送数据 | CLOSED |
| A17S-88 | AWS IoT | AWS IoT 删除设备缓存后仍可推送缓存数据 | CLOSED |
| A17S-87 | AWS IoT | AWS IoT 推送参数时间戳未按顺序记录 | CLOSED |
| A17S-86 | BACnet/IP | EPICS 文件设备信息未随当前配置同步更新，保留历史设备信息 | CLOSED |
| A17S-85 | AWS IoT | AWS IoT 同一时间间隔推送多次数据 | CLOSED |
| A17S-84 | BACnet/IP | 开启 BACnet/IP，上传参数值全为 0（重复，见A17S-83） | CLOSED |
| A17S-83 | BACnet/IP | 开启 BACnet/IP，上传参数值全为 0 | CLOSED |
| A17S-82 | BACnet/IP | 修改 BACnet/IP 配置每次需重启设备才能与客户端建立通信 | CLOSED |
| A17S-81 | AcuCloud | AcuCloud 生产环境创建设备后无法连接和推送数据 | CLOSED |
| A17S-80 | 虚拟设备 | 创建虚拟设备，Name 输入空格可以创建成功 | CLOSED |
| A17S-74 | BACnet/IP | 设置 Device Instance 为 4194302，提示 "check bacnet config error" | CLOSED |
| A17S-73 | BACnet/IP | YABE 工具连接设备，Device Object Name / Device Instance 显示错误 | CLOSED |
| A17S-72 | BACnet/IP | BACnet Port 和 BBMD Port 范围与需求不一致（应为 47808~49000） | CLOSED |
| A17S-71 | BACnet/IP | 首次开启 BACnet/IP，YABE 扫描不到设备 | CLOSED |
| A17S-70 | Remote Access | Remote Access URL 注销后状态显示异常（已拒绝） | REJECTED |
| A17S-69 | Remote Access | Remote Access Ping Interval 切换为 600s 后点击 Save 问题 | CLOSED |
| A17S-68 | 告警 | 先关闭再开启 Alarm Acknowledgement Enable 后已有告警行为异常 | CLOSED |
