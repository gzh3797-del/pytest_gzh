# AcuRev-4100-WEB2 Sprint 2 需求汇总

> 来源：
> - 主文档：《软件需求规格说明书 v1.01》（2026-04-23）
> - 变更：《软件需求规格变更说明书 v1.00》（2026-04-29，同步 HMI1-7 市场反馈）
>
> 本文件仅含二期新增/增强需求，一期已实现功能不在此列。

---

## 一、设备管理

### 1.1 设备自动发现（标注"暂不做"）
- Gateway 模式 + Eth2 启用时显示 Scan Devices 按钮，触发 mDNS 扫描
- 扫描结果含 Device Name（可编辑）、SN、Model、Template、IP/Port/SlaveID
- 接入成功后自动加入设备列表，状态从"可接入"变为"已接入"

### 1.2 Virtual Device（虚拟设备）
- 支持将多台 4100 电表的有功能量数据按公式计算后映射为虚拟设备
- SN 自动生成（AEVM + 5位随机数），全局唯一
- 配置参数不少于 20 个，公式支持四则运算及常数，格式：`"$device1:param1" + 0.1 * "$device2:param2"`
- 参数只能选 4100 电表，不含 AcuIOM
- **变更（v1.00）**：新增支持出线有功能量（EP_EXP）作为参数来源
- 支持接入 Datalog Post、AWS IoT、Azure IoT、AcuCloud

### 1.3 设备固件更新
- 支持对下挂 AcuRev-4100 和 AcuIOM 进行 Firmware Update（MFEA 格式文件）
- 升级过程禁止切换页面，禁止多用户同时操作同一设备
- 禁止固件降级，尝试降级时弹窗提示

### 1.4 模板版本兼容性
- 同一网关支持同型号不同固件版本的设备同时接入
- 模板命名：TemplateName_版本号_Firmware版本号
- 编辑设备时不允许切换模板；电表升级后网关不自动更新模板

### 1.5 Checkpoint 配置
- Data Log / Trend Log Management 页面新增 Checkpoint 下拉（Current 或历史备份时间点）
- 选择非 Current 时 Time Frame 受限于该备份日志的开始/结束时间，隐藏参数选择列

---

## 二、电表配置

### 2.1 PT 配置
- Settings → Device → General 新增 PT1（50~1,000,000）、PT2（50~830），默认均 480
- 校验：PT1 ≥ PT2，50 ≤ Nominal Voltage ≤ PT1
- 新增 Nominal Current 配置（1~50,000A，默认 1000A）
- PT1/PT2 受铅封保护，修改后立即生效（不回算历史数据）
- 电表固件不支持时 WEB2 界面隐藏此配置

### 2.2 User Channel 名称（Meter Point）
- 支持对 User Channel 1~12 自定义名称（ASCII，≤20字节）
- 全局显示格式统一为"User Channel N: Description"
- 全局 UI 将"User Channel"改名为"Meter Point"

---

## 三、协议扩展

### 3.1 BACnet/IP（新增）
- 默认 Disable，端口默认 47808（范围 47808~49000）
- 配置项：Network Number、Device Object Name、Device Instance、APDU Timeout/Retries
- Foreign Device：BBMD IP/Port、Time To Live
- 每个设备参数可独立配置 Polling Enable + COV Enable + COV Increment
- 支持下载 EPICS 文件
- 映射范围：4100（Basic Parameter、Demand、Power Quality、Energy、IO）；AcuIOM（AI、DI/RO/DO status、DI counter）
- **变更（v1.00）**：新增支持 Harmonic（谐波）参数映射

### 3.2 Ethernet/IP（新增）
- 默认 Disable，显式报文端口 44818（范围 44800~44899）
- 隐式报文绑定 Eth1，UDP 2222（固定不可改）
- 支持生成并下载 EDS 文件（product code 10003），参数与 SNMP 一致
- 主要兼容罗克韦尔 PLC

### 3.3 AWS IoT（新增）
- 将 4100 / 虚拟设备数据发布到 AWS IoT Core
- 配置：URL、Topic、Interval（1~600s）
- 断网缓存 24~72 小时

### 3.4 Azure IoT（新增）
- 将 Modbus、BACnet、虚拟设备数据推送到 Azure IoT Hub
- 配置：Primary/Secondary Connection String、Interval、SSL + X509 证书
- 支持 Device Twin 从 Azure 端配置 WEB2

### 3.5 Device Mirror
- 支持第三方设备通过 Slave ID 访问指定设备参数（4100 + 虚拟设备 + WEB2 设备信息）
- 参数映射预定义，用户不可配置

### 3.6 Modbus 端口范围变更
- 端口范围从原始值修改为 2000~5999（默认 502 不变）

---

## 四、系统诊断

### Wiring Check（接线检查）
- 支持 WEB Module 和 Gateway 模式，Gateway 模式可单选或全选电表
- 结果分电压侧和电流侧两类展示，问题标红，正常标绿
- 支持 Device Name 筛选、"Show Only Issues"开关、导出 CSV
- 支持 AcuRev-4100 五种接线方式（依据接线检测总表 ver1.03）

---

## 五、系统安全

| 需求 | 类型 |
|------|------|
| 新增 EULA，所有用户登录后必须同意一次 | 新增 |
| Admin 默认密码改为 Admin@AABBCC（AABBCC=SN后六位）| 变更 |
| 新增"忘记密码"：每日临时密码，仅 Admin 账户可用 | 新增 |
| 修改密码须先验证当前密码（修改他人密码除外） | 变更 |
| 首次登录 / 使用默认密码每次登录弹窗提醒修改 | 新增 |
| 禁止固件降级 | 新增 |
| 默认禁用 Modbus Pass Through 和 SSH | 变更 |
| 禁用调试端口内核输出 | 新增 |
| DoS 攻击降级模式下维持基本功能 | 新增 |

---

## 六、AcuCloud 增强

- Advanced 模式（URL 后缀 `?showAdvanced=true`）支持修改数据传输 URL、固件更新 Server URL、Remote Access URL
- Installation Record 和 Inspection Record 每次保存时自动向 AcuCloud 推送（仅 4100 设备）

---

## 七、Web UI 改版

### Metering 页面
- AcuIOM DI 界面拆为两张表：DI Status + DI Counter（含 Edit/Clear）
- AcuIOM AO 界面：Control 按钮弹窗修改 Eng. Value 并显示对应 Output Value

### Settings 页面
- AcuIOM DI 设置：删除 Pulse Count 列，新增 Copy/Apply/Reset
- AcuIOM AO 设置：简化列，删除 AO Physical Measurement Input，新增 Copy/Apply/Reset
- AcuIOM AI 设置：总表 + Details 弹窗，新增 Signal Type（4种）和折线段配置（1~3段），新增 Copy/Apply/Reset

### About → Service 页面重构
- About 仅保留 Information
- 新增 Service 页面：Installation Record、Inspection Record、Troubleshooting（下载 .a2d 加密诊断文件）
- AcuRev-4100 Troubleshooting 支持查看/修改 CT Model 和 Direction

### **变更（v1.00）**
- Timezone → Time Zone（UI 文字）
- Post Channel 测试通道显示更详细的通道信息（参考 AXM-WEB2）

---

## 八、Data Log 默认参数（变更 v1.00 新增）

| 设备类型 | 默认参数 |
|----------|----------|
| 4100 电表 | Realtime 组（电压、电流、功率、频率）+ Energy 组（进线和出线有功能量） |
| AcuIOM | 全部参数默认选中 |

---

## 变更说明书 vs 主文档 对比

| 维度 | 主文档 v1.01 | 变更 v1.00 |
|------|-------------|------------|
| 虚拟设备参数范围 | 进线有功能量 | 新增出线有功能量（EP_EXP） |
| BACnet/IP 映射 | 不含谐波 | 扩展：支持 Harmonic |
| Data Log 默认值 | 无默认值 | 4100 电表：Realtime+Energy；AcuIOM：全部 |
| Post Channel | 未提及详细度 | 参考 AXM-WEB2 显示更详细信息 |
| UI 文字 | Timezone | 改为 Time Zone |
