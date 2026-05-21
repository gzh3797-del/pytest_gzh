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

## EtherNetIP/enip_comparator.py — EtherNet/IP 协议合规性七段式检查

通过 CIP 显式消息（TCP 44818）对网关 WEB2 的 EtherNet/IP 实现执行七段式合规性检查。
EDS 文件从 `config.EDS_DIR` 自动按设备名匹配，无需手动指定路径。

**检查流程：**
1. **范围检查**：模板 SNMP 列参数 vs EDS Assembly 参数（缺失 / 多余）
2. **单位检查**：EDS 各参数 `unit` 字段 vs 模板 `unit` 列
3. **数值比对**：并发读取 EtherNet/IP Assembly 与 Modbus，按 CIP 数据类型解析后逐参数比对（容差 ±1% / ±0.05）
4. **Assembly 结构合规性**：EDS 声明总字节数 vs 实际读取字节数，检测越界参数
5. **CIP Identity Object**：读取 Class=0x01 Instance=0x01 全部标准属性（Vendor ID / Device Type / Product Code / Revision / Status / Serial Number / Product Name）并解码
6. **CIP 错误响应测试**：发送四条非法请求（不存在的实例、属性、类、Assembly 实例），验证设备正确返回错误而非意外成功
7. **连接稳定性**：连续 3 次读取 Assembly，全部成功才通过

报告输出到 `reports/enip_compare_<设备名>_<时间戳>.html`，包含七段可折叠区块。

**前提条件：**
- 网关 WEB2（`config.ENIP_HOST`，端口 44818）已开启 EtherNet/IP 服务
- 被测设备的 EDS 文件已放置于 `config.EDS_DIR`（`EtherNetIP/eds/`）

**已有 EDS 文件的设备：**

| `--device` 参数 | EDS 文件 | 备注 |
|---|---|---|
| `AcuRev4100` | `AcuRev-4100.eds` | |
| `AcuIOM-01`  | `AcuIOM-01.eds`  | 8 AI 通道 |
| `AcuIOM-02`  | `AcuIOM-02.eds`  | 16 AI 通道 |
| `AcuIOM-03`  | `AcuIOM-03.eds`  | 14 DI 通道 |
| `AcuIOM-04`  | `AcuIOM-04.eds`  | 28 DI 通道 |

```bash
# 全量检查（默认设备 AcuRev-4100，运行所有七段）
python Protocols/EtherNetIP/enip_comparator.py

# 指定设备
python Protocols/EtherNetIP/enip_comparator.py --device AcuRev4100
python Protocols/EtherNetIP/enip_comparator.py --device AcuIOM-01
python Protocols/EtherNetIP/enip_comparator.py --device AcuIOM-02
python Protocols/EtherNetIP/enip_comparator.py --device AcuIOM-03
python Protocols/EtherNetIP/enip_comparator.py --device AcuIOM-04

# 快速模式：数值比对只跑前 30 个参数（其余六段仍完整执行）
python Protocols/EtherNetIP/enip_comparator.py --quick

# 只比对指定参数（数值比对段）
python Protocols/EtherNetIP/enip_comparator.py --keys FREQ_Hz VLN_a_V P_kW
```

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
| `ENIP_HOST` / `ENIP_SLOT` | EtherNet/IP 网关 IP 与 CIP slot（默认 `192.168.2.63` / `0`） |
| `EDS_DIR` | EDS 文件目录（`EtherNetIP/eds/`），enip_comparator 自动按设备名匹配（忽略大小写及连字符） |
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
│   ├── enip_reader.py         # Assembly 读取、CIP 对象查询、合规性检查模块（pycomm3）
│   ├── enip_comparator.py     # 七段式合规性检查主程序
│   └── eds/                   # 设备 EDS 文件（按设备名自动匹配）
│       └── AcuRev-4100.eds    # （示例，需手动放入）
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
