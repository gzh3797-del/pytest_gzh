# 网关协议 自动化测试工具集

## 支持设备列表

| `--device` 参数 | 设备型号 | BACnet | EtherNet/IP | Modbus Mapping | SNMP | MQTT | DataLog | AcuCloud | AWS IoT | Azure IoT | Device Mirror | 备注 |
|---|---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|---|
| `acurev4100` | AcuRev 4100 | ✅ | ✅ | — | — | ✅ | ✅ | ✅ | — | — | — | |
| `acurev2100` | AcuRev 2100 | ✅ | —  | — | — | ✅ | ✅ | ✅ | — | — | — | |
| `acuvimiiw`  | Acuvim IIW  | ✅ | —  | — | — | ✅ | ✅ | ✅ | — | — | — | |
| `acuvimiir`  | Acuvim IIR  | ✅ | —  | — | — | ✅ | ✅ | ✅ | — | — | — | Modbus 地址与 IIW 相同 |
| `acuvim3`    | AcuVIM3     | ✅ | —  | — | — | ✅ | ✅ | ✅ | — | — | — | |
| `pxm350`     | PXM350      | ✅ | —  | — | — | —  | —  | —  | — | — | — | |
| `acuiom01`   | AcuIOM-01   | ✅ | ✅ | — | — | —  | —  | —  | — | — | — | 8 AI 通道（原始输入 + 物理量读数各 8 个） |
| `acuiom02`   | AcuIOM-02   | ✅ | ✅ | — | — | —  | —  | —  | — | — | — | 16 AI 通道（原始输入 + 物理量读数各 16 个） |
| `acuiom03`   | AcuIOM-03   | ✅ | ✅ | — | — | —  | —  | —  | — | — | — | 14 DI 通道，BACnet BI 对象 |
| `acuiom04`   | AcuIOM-04   | ✅ | ✅ | — | — | —  | —  | —  | — | — | — | 28 DI 通道，BACnet BI 对象 |

> 运行前请确认 `config.py` 中 `MODBUS_DEVICE_MAP` 内对应设备的 IP / Unit ID 已填写正确。

### Modbus RTU 模式（设备与电脑不同网段时）

web2 作为 DHCP Server 为下挂设备分配 IP，导致下挂设备（如 AcuRev4100）与测试电脑不在同一网段，
无法通过 Modbus TCP 直连。此时在 `config.py` 中切换为 RTU 模式，通过 RS485 串口读取：

```python
# config.py
MODBUS_MODE         = "rtu"
MODBUS_RTU_PORT     = "COM3"   # 实际串口号
MODBUS_RTU_BAUDRATE = 9600     # 与设备波特率一致
MODBUS_RTU_PARITY   = "N"
MODBUS_RTU_STOPBITS = 1
MODBUS_RTU_BYTESIZE = 8
```

RTU 模式下 `MODBUS_DEVICE_MAP` 中的 IP/Port 被忽略，仅 Unit ID（slave 地址）生效。
所有比对脚本（BACnetIP / AcuCloud / MQTT / EtherNetIP）无需修改，切换配置后直接运行。

---

## BACnetIP/comparator.py — BACnet vs 实时 Modbus 比对

**比对流程（六段式）：**
1. 从 `knowledge/shared/templates/raw/` 加载设备模板，获取应发布到 BACnet 的参数范围
   - 普通设备：按 `BACnetIP` 列过滤；AcuIOM 设备：按 `range` 列过滤（`BACNET_RANGE_MARKER` 自动切换）
2. 向网关发现实际发布的 BACnet AI / BI 对象，与模板范围对比（范围检查）
   - AcuIOM-01/02：AI 对象（模拟量）；AcuIOM-03/04：BI 对象（DI 数字量，present value = 0/1）
3. 读取每个对象的 `description` / `units` 属性，与模板元数据对比（元数据检查）
4. 并发读取 BACnet Present Value 与 Modbus 寄存器值，按容差规则判断通过/失败（数值比对）
5. **Device Object 属性**：读取 Device Object 的 12 项标准必需属性（ANSI/ASHRAE 135 §12.11）：vendorName、modelName、firmwareRevision、protocolVersion/Revision、segmentationSupported 等
6. **协议合规性**：验证非法对象请求返回错误（§16），以及 AI 必需属性（statusFlags / outOfService / units）均可读
7. **连接稳定性**：对同一 AI 对象连续读取 5 次，验证成功率与一致性

报告输出到 `reports/bacnet_<设备名>_<时间戳>.html`，包含六段。

```bash
# 全量比对（默认设备）
python Protocols/BACnetIP/comparator.py

# 指定设备
python Protocols/BACnetIP/comparator.py --device acurev4100
python Protocols/BACnetIP/comparator.py --device acurev2100
python Protocols/BACnetIP/comparator.py --device acuvimiiw
python Protocols/BACnetIP/comparator.py --device acuvimiir
python Protocols/BACnetIP/comparator.py --device acuvim3
python Protocols/BACnetIP/comparator.py --device acuiom01
python Protocols/BACnetIP/comparator.py --device acuiom02
python Protocols/BACnetIP/comparator.py --device acuiom03
python Protocols/BACnetIP/comparator.py --device acuiom04

# 快速模式：只比对前 30 个参数
python Protocols/BACnetIP/comparator.py --quick
python Protocols/BACnetIP/comparator.py --device acurev2100 --quick

# 只比对指定参数
python Protocols/BACnetIP/comparator.py --keys FREQ_Hz VLN_a_V P_kW
python Protocols/BACnetIP/comparator.py --device acuvimiiw --keys FREQ_Hz I_a_A

# 关闭元数据检查（仅范围 + 数值 + 协议测试）
python Protocols/BACnetIP/comparator.py --no-meta

# 关闭协议规范测试（仅范围 + 元数据 + 数值）
python Protocols/BACnetIP/comparator.py --no-proto
```

---

## EtherNetIP/enip_comparator.py — EtherNet/IP 协议合规性九段式检查

通过 CIP 显式消息（TCP 44818）对网关 WEB2 的 EtherNet/IP 实现执行九段式合规性检查。
EDS 文件路径由 `config.ENIP_EDS_PATH` 显式指定（多表模式）；未设置时回退到 `config.EDS_DIR` 按设备名模糊匹配（单表兼容）。

**检查流程：**
1. **范围检查**：模板 SNMP 列参数 vs EDS Assembly 参数（缺失 / 多余）
2. **单位检查**：EDS 各参数 `unit` 字段 vs 模板 `unit` 列
3. **数值比对**：通过 **UCMM 显式消息**（Get_Attribute_Single，TCP 44818）逐 Assembly 实例读取原始字节，按 CIP 数据类型（float32 / LREAL / uint16 等）解析后与 Modbus 实时值逐参数比对（容差 ±1% / ±0.05）。多表模式下只读当前设备对应的 Assembly 实例（由 EDS Connection help string/SN 定位），此方式与 Studio 5000 的隐式 I/O 连接独立，互不影响
4. **Assembly 结构合规性**：EDS 声明总字节数 vs 实际读取字节数，检测越界参数；同时检查 LREAL/DINT/UDINT 类型的字节对齐（Rockwell Logix AOP 要求）
5. **CIP Identity Object**：读取 Class=0x01 Instance=0x01 全部标准属性（Vendor ID / Device Type / Product Code / Revision / Status / Serial Number / Product Name）并解码
6. **CIP 错误响应测试**：发送四条非法请求（不存在的实例、属性、类、Assembly 实例），验证设备正确返回错误而非意外成功
7. **连接稳定性**：连续 3 次读取 Assembly，全部成功才通过
8. **EDS 文件静态合规性检查（无需 PLC）**：
   - **Connection Manager**：按 ODVA 规范 15 字段位置逐一验证（含原始字段明细表），对应 EZ-EDS 全部 6 条检查项：T→O Format / Proxy Config Format / Target Config Size 类型 / Target Config Format 引用 / Connection Name / Path
   - **Assembly 数据类型对齐**：LREAL(8B)、DINT/UDINT(4B) 参数偏移的对齐验证（Rockwell Logix AOP 要求）
   - **EDS `[File] Revision` vs CIP Identity Revision**：交叉比对一致性
   - **`[Device]` 段 vs CIP Identity 交叉比对**：VendCode / ProdCode / MajRev / MinRev / ProdName 五项与 Identity Object 实时读取值比对
   - **Assembly 字节数三路一致性**：`[Assembly]` 头部声明 = `[Params]` 成员大小之和 = Connection1 T→O Size 字段，三者不一致即报告
   - **孤儿 Param 引用**：`[Assembly]` Assem10 中引用的 ParamN 若在 `[Params]` 未定义，报告孤儿列表（静默 fallback 为 float32 会导致数据误读）
9. **Forward_Open 隐式连接建立**（Large_Forward_Open 0x5B）：
   - **仅覆盖握手阶段（TCP）**，不接收后续 UDP I/O 数据流
   - 脚本模拟 Studio 5000 的行为，手动构造 Large_Forward_Open（0x5B）报文，以 Input-Only 参数（O→T null / T→O P2P **5656B** RPI=500ms，路径 `20 04 24 0C 24 0B 24 0A`）向 Assembly 10 发起连接请求（TCP 44818）
   - 成功则记录 T→O 实际包间隔（API），随即 **立即 Forward_Close** 断开，不等待也不接收任何 UDP 2222 数据包，**说明 Studio 5000 可正常建立 I/O 扫描连接**
   - 失败则记录设备返回的 CIP 错误码，**说明 Studio 5000 会在同一步骤失败**——提前暴露固件对隐式连接的支持情况，无需等真机 PLC 联调

> **与 Section 3 的关系**：Section 3（数值比对）走显式消息，Section 9 走隐式 I/O 握手。两条路互相独立——Section 3 全部通过不代表 Studio 5000 能正常接入，必须 Section 9 也通过才能确认隐式连接可用。
>
> **握手通过后仍需人工验证**：Forward_Open 成功只代表"门开了"，实际 UDP 2222 数据流（每 500ms 一包，共 5656 字节）需用 Wireshark 抓包或在 Studio 5000 中观察 Controller Tags 刷新来确认。
>
> **Section 8 不依赖设备连接**，即使设备离线也可静态分析 EDS 文件。**Section 9 需要设备在线**，是验证固件修复效果的最终动态测试。

报告输出到 `reports/enip_compare_<设备标识>_<时间戳>.html`，包含九段可折叠区块。

**前提条件：**
- 网关 WEB2（`config.ENIP_HOST`，端口 44818）已开启 EtherNet/IP 服务
- EDS 文件已放入 `EtherNetIP/eds/`，并在 `config.ENIP_EDS_PATH` 中填写路径

**EDS 文件说明：**

| 模式 | EDS 来源 | 说明 |
|---|---|---|
| 单表（`--device`）| 按设备型号命名（如 `AcuRev-4100.eds`），放入 `eds/`，`ENIP_EDS_PATH` 填路径 | 也可不设 `ENIP_EDS_PATH`，回退到 `EDS_DIR` 按设备名模糊匹配 |
| 多表（`--all`）| web2 网关导出的配置快照 EDS | 包含所有已选设备的 Connection/Assembly，`ENIP_EDS_PATH` 必填 |

```bash
# 全量检查（默认设备 AcuRev-4100，运行所有九段）
python Protocols/EtherNetIP/enip_comparator.py

# 指定设备
python Protocols/EtherNetIP/enip_comparator.py --device AcuRev4100
python Protocols/EtherNetIP/enip_comparator.py --device AcuIOM-01
python Protocols/EtherNetIP/enip_comparator.py --device AcuIOM-02
python Protocols/EtherNetIP/enip_comparator.py --device AcuIOM-03
python Protocols/EtherNetIP/enip_comparator.py --device AcuIOM-04

# 快速模式：数值比对只跑前 30 个参数（其余八段仍完整执行）
python Protocols/EtherNetIP/enip_comparator.py --quick

# 只比对指定参数（数值比对段）
python Protocols/EtherNetIP/enip_comparator.py --keys FREQ_Hz VLN_a_V P_kW

# 多表批量测试：依次对 config.ENIP_MULTI_DEVICES 中每台设备执行全套九段检查，
# 每台生成独立 HTML 报告，完成后打印多表汇总
python Protocols/EtherNetIP/enip_comparator.py --all
python Protocols/EtherNetIP/enip_comparator.py --all --quick
```

**多表配置步骤（`--all` 模式）：**

EDS 文件由网关配置决定，选了哪些设备 EDS 里就有哪些 Connection/Assembly。每次测试前：

**Step 1：** 将网关导出的 EDS 文件放入 `EtherNetIP/eds/`，在 `config.py` 中填写路径：

```python
# config.py
ENIP_EDS_PATH = "Protocols/EtherNetIP/eds/AcuRev-4100.eds"  # 按实际文件名填写
```

**Step 2：** 填写各设备的连接参数（`eds_label` = 设备 SN，web2 生成 EDS 时固定写入各 Connection 的 help string，直接填 SN 即可）：

```python
# config.py
ENIP_MULTI_DEVICES = [
    # (eds_label,   device_name,   modbus_host,      modbus_port, modbus_unit)
    ("41002242",  "AcuRev4100",  "192.168.2.242",  502,         1),
    ("4100229",   "AcuRev4100",  "192.168.2.29",   502,         202),
    # ("AcuIOM01", "AcuIOM01",   "192.168.2.xx",   502,         xxx),
]
```

- `eds_label`：设备 SN（即 EDS ConnectionN 的 help string，web2 生成 EDS 时固定写入 SN）
- `device_name`：设备型号，对应 `devices/` 模块与参数模板
- `modbus_host`：各设备自身的 Modbus TCP IP（每台不同）
- EIP 网关地址统一使用 `config.ENIP_HOST`，无需在每条记录中重复填写
- 每台设备独立生成报告，文件名格式：`enip_compare_<eds_label>_<时间戳>.html`

> **注意：** AcuRev-4100 Modbus TCP 连接数有限，若 WEB2 网关已占用全部连接槽，
> Modbus 读取将失败。此时可暂停 WEB2 轮询，或通过 RS485 读取数据。

---

## AcuCloud/cloud_comparator.py — AcuCloud 历史快照 vs 实时 Modbus 比对

从 `Acuclouddatas/` 目录读取 AcuCloud 导出的 xlsx 文件，与设备实时 Modbus 寄存器值比对。
因快照与实时读取存在时序差异，容差较大（默认 ±5% / ±1.0）。

**注意：** xlsx 文件名须与设备名精确匹配（如 `AcuvimIIW.xlsx` 对应 `--device acuvimiiw`）。

报告输出到 `reports/cloud_<设备名>_<时间戳>.html`，包含两段：范围检查 / 数值比对。

```bash
# 自动匹配文件，取最新数据行（需先用 --device 指定设备）
python Protocols/AcuCloud/cloud_comparator.py --device acurev4100
python Protocols/AcuCloud/cloud_comparator.py --device acurev2100
python Protocols/AcuCloud/cloud_comparator.py --device acuvimiiw
python Protocols/AcuCloud/cloud_comparator.py --device acuvimiir
python Protocols/AcuCloud/cloud_comparator.py --device acuvim3

# 指定 xlsx 文件路径
python Protocols/AcuCloud/cloud_comparator.py --device acurev4100 --file "Protocols/Datas/acuclouddatas/AcuRev4100.xlsx"

# 指定数据行（1 = 第一行，默认取最新行）
python Protocols/AcuCloud/cloud_comparator.py --device acurev4100 --row 1

# 只比对指定参数
python Protocols/AcuCloud/cloud_comparator.py --device acurev4100 --keys FREQ_Hz VLN_a_V P_kW

# 组合用法
python Protocols/AcuCloud/cloud_comparator.py --device acurev2100 --file <xlsx路径> --row 2
```

---

## MQTT/mqtt_comparator.py — MQTT 快照 vs 实时 Modbus 三段式比对

从 `MQTT/` 目录读取网关推送的 MQTT JSON 快照文件，与设备实时 Modbus 寄存器值三段式比对。
JSON 的 `param` 字段直接对应 `param_key`，无需列标题映射。

**比对流程：**
1. 范围检查：模板全量参数 vs JSON `modules[].reading` 实际发布参数（缺失 / 多余）
2. 单位检查：JSON `unit` 字段 vs 模板 `unit` 列（容许 `°/deg`、大小写等常见等价写法）
3. 数值比对：JSON `value` vs 实时 Modbus 读取值，容差 ±5% / ±1.0

报告输出到 `reports/mqtt_<设备名>_<时间戳>.html`，包含三段可折叠区块。

**JSON 文件命名约定：** 文件名须包含设备型号数字（如 `4100data.json` 对应 `acurev4100`），
支持模糊匹配；也可用 `--file` 直接指定路径。

```bash
# 自动匹配文件，使用第一个 online 模块
python Protocols/MQTT/mqtt_comparator.py --device acurev4100
python Protocols/MQTT/mqtt_comparator.py --device acurev2100
python Protocols/MQTT/mqtt_comparator.py --device acuvimiiw
python Protocols/MQTT/mqtt_comparator.py --device acuvimiir
python Protocols/MQTT/mqtt_comparator.py --device acuvim3

# 指定 JSON 文件路径
python Protocols/MQTT/mqtt_comparator.py --device acurev4100 --file "Protocols/MQTT/4100data.json"

# 指定模块下标（JSON modules 数组下标，默认取第一个 online 模块）
python Protocols/MQTT/mqtt_comparator.py --device acurev4100 --module 0

# 跳过单位检查（仅范围 + 数值）
python Protocols/MQTT/mqtt_comparator.py --device acurev4100 --no-meta

# 只比对指定参数
python Protocols/MQTT/mqtt_comparator.py --device acurev4100 --keys FREQ_Hz VLN_a_V P_kW

# 组合用法
python Protocols/MQTT/mqtt_comparator.py --device acurev4100 --file "Protocols/MQTT/4100data.json" --module 0 --no-meta
```

---

## Datalog/datalog_comparator.py — Datalog 快照 vs 实时 Modbus 比对

从 `Datas/DatalogDatas/` 目录读取网关导出的 Datalog 文件，与设备实时 Modbus 寄存器值比对。
自动识别文件格式，同一设备同时有 JSON 和 CSV 时优先选 JSON。

| 格式 | 比对段数 | 说明 |
|---|---|---|
| **JSON** | 三段（范围 + 单位 + 数值） | 含 `unit` 字段，可做单位检查 |
| **CSV**  | 两段（范围 + 数值）         | 无单位元数据，跳过单位检查 |

容差均为 ±5% / ±0.05，报告输出到 `reports/datalog_<设备名>_<时间戳>.html`。

**文件命名约定：** 文件名须包含设备型号关键字（如 `AcuRev4100`、`Acuvim3`、`AcuRevIIW`），
支持模糊匹配；也可用 `--file` 直接指定路径。

```bash
# 自动匹配文件（JSON 优先），取最新数据行
python Protocols/Datalog/datalog_comparator.py --device acurev4100
python Protocols/Datalog/datalog_comparator.py --device acurev2100
python Protocols/Datalog/datalog_comparator.py --device acuvimiiw
python Protocols/Datalog/datalog_comparator.py --device acuvim3

# 指定文件路径（自动判断格式）
python Protocols/Datalog/datalog_comparator.py --device acurev4100 --file "Protocols/Datas/DatalogDatas/Logger1-AHI260110088-AcuRev4100-1778653500-1min.json"
python Protocols/Datalog/datalog_comparator.py --device acurev4100 --file "Protocols/Datas/DatalogDatas/Logger1-AHI260110088-AcuRev4100-1778651820-1min.csv"

# 指定数据行（1 = 第一行，默认取最新行；JSON 多时间戳时同样生效）
python Protocols/Datalog/datalog_comparator.py --device acurev4100 --row 1

# 只比对指定参数
python Protocols/Datalog/datalog_comparator.py --device acurev4100 --keys FREQ_Hz VLN_a_V P_kW
```

---

## 配置（config.py）

| 配置项 | 说明 |
|---|---|
| `DEVICE_NAME` / `DEVICE_MODULE` | 默认测试设备（运行时可用 `--device` 覆盖） |
| `GATEWAY_IP` / `GATEWAY_PORT` | BACnet 网关地址与端口 |
| `LOCAL_IP` / `LOCAL_PORT` | 本机 BACnet 监听地址 |
| `MODBUS_MODE` | Modbus 通信方式：`"tcp"`（默认）或 `"rtu"`（串口，设备与电脑不同网段时使用） |
| `MODBUS_DEVICE_MAP` | 各设备 Modbus 连接参数；TCP 模式三项均生效，RTU 模式仅 Unit ID 生效 |
| `MODBUS_RTU_PORT` | RTU 串口号（Windows: `"COM3"`，Linux: `"/dev/ttyUSB0"`） |
| `MODBUS_RTU_BAUDRATE` / `MODBUS_RTU_PARITY` / `MODBUS_RTU_STOPBITS` / `MODBUS_RTU_BYTESIZE` | RTU 串口参数（默认 9600/N/1/8） |
| `TEMPLATE_DIR` | 设备参数模板 xlsx 目录（指向 `knowledge/shared/templates/raw/`） |
| `BACNET_RANGE_MARKER` | AcuIOM BACnet 参数范围过滤标记；`""` 用 `BACnetIP` 列，`"8"` 用 range 列（IOM-01/02），`"10"` 用 range 列（IOM-03/04）；`--device` 时自动设置，无需手动修改 |
| `READ_TIMEOUT` / `MAX_RETRIES` / `RETRY_WAIT` | 单次读取超时、重试次数、重试间隔 |
| `CONNECT_WAIT` | BAC0 启动后等待网关就绪时间（秒） |
| `ENIP_HOST` / `ENIP_SLOT` | EtherNet/IP 网关 IP 与 CIP slot（默认 `192.168.3.9` / `0`） |
| `EDS_DIR` | EDS 文件目录（`EtherNetIP/eds/`），单表模式下按设备名模糊匹配 |
| `ENIP_EDS_PATH` | EDS 文件显式路径（网关配置快照，每次更换 EDS 时更新）；非空时优先使用，覆盖 `EDS_DIR` 匹配；`--all` 多表模式必须填写 |
| `ENIP_MULTI_DEVICES` | 多表批量测试配置，格式：`(eds_label, device_name, modbus_host, modbus_port, modbus_unit)`；`eds_label` 填设备 SN（web2 生成 EDS 时写入 help string），脚本自动定位该设备的 Assembly 实例 |
| `TOLERANCE_PERCENT` / `TOLERANCE_ABSOLUTE` | BACnet / EtherNet/IP vs Modbus 数值比对容差 |
| `CLOUD_TOLERANCE_PERCENT` / `CLOUD_TOLERANCE_ABSOLUTE` | AcuCloud 快照比对容差 |
| `CLOUD_DATA_DIR` | AcuCloud xlsx 快照文件目录 |
| `MQTT_DATA_DIR` | MQTT JSON 快照文件目录（`MQTT/`） |
| `MQTT_TOLERANCE_PERCENT` / `MQTT_TOLERANCE_ABSOLUTE` | MQTT 快照比对容差 |
| `REPORT_DIR` | HTML 报告输出目录 |

---

## 目录结构

```
Protocols/
├── BACnetIP/                  # BACnet/IP 协议
│   ├── bacnet_reader.py       # BACnet 读取模块（BAC0）
│   └── comparator.py          # BACnet vs Modbus 比对主程序
├── EtherNetIP/                # EtherNet/IP 协议
│   ├── enip_reader.py         # Assembly 读取、CIP 对象查询、EDS 静态合规检查模块（pycomm3）
│   ├── enip_comparator.py     # 九段式合规性检查主程序
│   └── eds/                   # EDS 文件目录
│       └── AcuRev-4100.eds    # 网关配置快照 EDS（每次测试前放入，ENIP_EDS_PATH 指向此文件）
├── AcuCloud/                  # AcuCloud 数据
│   └── cloud_comparator.py    # AcuCloud 快照 vs Modbus 比对主程序
├── MQTT/                      # MQTT 数据
│   ├── mqtt_comparator.py     # MQTT 快照 vs Modbus 三段式比对主程序
│   └── <设备名>data.json      # 网关推送的 MQTT JSON 快照（如 4100data.json）
├── Datas/                     # 测试数据文件
│   ├── acuclouddatas/         # AcuCloud 导出 xlsx 快照
│   └── DatalogDatas/          # 网关导出的 Datalog JSON / CSV 文件
├── modbusAddress/             # Modbus 地址表工具（xlsx 已迁移至 knowledge/shared/modbus_tables/raw/）
│   └── mapping_compile.py
├── modbus_reader.py           # Modbus TCP 读取模块（共用）
├── template_reader.py         # 模板 xlsx 解析模块（共用）
├── config.py                  # 全局配置
├── devices/                   # 各设备 Modbus 地址映射（Python 模块）
│   ├── acurev4100.py
│   ├── acurev2100.py
│   ├── acuvimiiw.py
│   ├── acuvimiir.py
│   ├── acuvim3.py
│   ├── pxm350.py
│   ├── acuiom01.py
│   ├── acuiom02.py
│   ├── acuiom03.py            # 14 DI（FC02）+ 脉冲计数（FC03）+ DO/RO（FC01）
│   └── acuiom04.py            # 28 DI（FC02）+ 脉冲计数（FC03）+ DO/RO（FC01）
└── reports/                   # HTML 比对报告输出目录
```
