# 网关协议 自动化测试工具集

## 支持设备列表

| `--device` 参数 | 设备型号 | BACnet 比对 | AcuCloud 比对 | EtherNet/IP 比对 | 备注 |
|---|---|:---:|:---:|:---:|---|
| `acurev4100` | AcuRev 4100 | ✅ | ✅ | ✅ | |
| `acurev2100` | AcuRev 2100 | ✅ | ✅ | — | |
| `acuvimiiw`  | Acuvim IIW  | ✅ | ✅ | — | |
| `acuvimiir`  | Acuvim IIR  | ✅ | ✅ | — | Modbus 地址与 IIW 相同 |
| `acuvim3`    | AcuVIM3     | ✅ | ✅ | — | |
| `pxm350`     | PXM350      | ✅ | —  | — | |
| `acuiom01`   | AcuIOM-01   | ✅ | —  | — | 8 AI 通道 |
| `acuiom02`   | AcuIOM-02   | ✅ | —  | — | 16 AI 通道 |
| `acuiom03`   | AcuIOM-03   | ⚠️ | —  | — | DI 型号，FC 0x02 暂不支持 |
| `acuiom04`   | AcuIOM-04   | ⚠️ | —  | — | DI 型号，FC 0x02 暂不支持 |

> 运行前请确认 `config.py` 中 `MODBUS_DEVICE_MAP` 内对应设备的 IP / Unit ID 已填写正确。

---

## BACnetIP/comparator.py — BACnet vs 实时 Modbus 比对

**比对流程：**
1. 从 `Template/` 目录加载设备模板，获取应发布到 BACnet 的参数范围
2. 向网关发现实际发布的 BACnet AI 对象，与模板范围对比（范围检查）
3. 读取每个对象的 `description` / `units` 属性，与模板元数据对比（元数据检查）
4. 并发读取 BACnet Present Value 与 Modbus 寄存器值，按容差规则判断通过/失败（数值比对）

报告输出到 `reports/compare_<设备名>_<时间戳>.html`，包含三段：范围检查 / 元数据检查 / 数值比对。

```bash
# 全量比对（默认设备，包含范围检查 + 元数据检查 + 数值比对）
python Protocols/BACnetIP/comparator.py

# 指定设备
python Protocols/BACnetIP/comparator.py --device acurev4100
python Protocols/BACnetIP/comparator.py --device acurev2100
python Protocols/BACnetIP/comparator.py --device acuvimiiw
python Protocols/BACnetIP/comparator.py --device acuvimiir
python Protocols/BACnetIP/comparator.py --device acuvim3
python Protocols/BACnetIP/comparator.py --device acuiom01
python Protocols/BACnetIP/comparator.py --device acuiom02

# 快速模式：只比对前 30 个参数（跳过元数据检查）
python Protocols/BACnetIP/comparator.py --quick
python Protocols/BACnetIP/comparator.py --device acurev2100 --quick

# 只比对指定参数
python Protocols/BACnetIP/comparator.py --keys FREQ_Hz VLN_a_V P_kW
python Protocols/BACnetIP/comparator.py --device acuvimiiw --keys FREQ_Hz I_a_A

# 关闭元数据检查（仅范围 + 数值）
python Protocols/BACnetIP/comparator.py --no-meta
```

---

## EtherNetIP/enip_comparator.py — EtherNet/IP vs 实时 Modbus 比对

通过 CIP 协议（pycomm3）读取网关 WEB2 上的 EtherNet/IP Assembly Instance 100，
与设备 Modbus TCP 实时值进行比对。

**比对流程：**
1. 解析 EDS 文件（`/awsdatas/eds/`），获取 Assembly 参数布局
2. 读取模板 SNMP 列，与 EDS Assembly 参数范围对比（范围检查）
3. 并发读取 EtherNet/IP Assembly 字节流与 Modbus 寄存器值
4. 按 CIP 数据类型（float32/float64/UINT 等，小端序）解析 Assembly，逐参数比对

报告输出到 `reports/enip_compare_<设备名>_<时间戳>.html`，包含两段：范围检查 / 数值比对。

**前提条件：**
- 网关 WEB2（192.168.2.63:44818）已开启 EtherNet/IP 服务并配置上传目标设备参数
- 被测设备（如 AcuRev-4100）的 EDS 文件已放置于 `awsdatas/eds/` 目录

```bash
# 全量比对（AcuRev-4100，默认配置）
python Protocols/EtherNetIP/enip_comparator.py

# 快速模式：只比对前 30 个参数
python Protocols/EtherNetIP/enip_comparator.py --quick

# 只比对指定参数
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

## 配置（config.py）

| 配置项 | 说明 |
|---|---|
| `DEVICE_NAME` / `DEVICE_MODULE` | 默认测试设备（运行时可用 `--device` 覆盖） |
| `GATEWAY_IP` / `GATEWAY_PORT` | BACnet 网关地址与端口 |
| `LOCAL_IP` / `LOCAL_PORT` | 本机 BACnet 监听地址 |
| `MODBUS_DEVICE_MAP` | 各设备 Modbus TCP 连接参数（IP、端口、Unit ID） |
| `TEMPLATE_DIR` | 设备参数模板 xlsx 目录（`template/`） |
| `READ_TIMEOUT` / `MAX_RETRIES` / `RETRY_WAIT` | 单次读取超时、重试次数、重试间隔 |
| `CONNECT_WAIT` | BAC0 启动后等待网关就绪时间（秒） |
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
│   ├── enip_reader.py         # EtherNet/IP Assembly 读取模块（pycomm3）
│   └── enip_comparator.py     # EtherNet/IP vs Modbus 比对主程序
├── AcuCloud/                  # AcuCloud 数据
│   └── cloud_comparator.py    # AcuCloud 快照 vs Modbus 比对主程序
├── MQTT/                      # MQTT 数据
│   ├── mqtt_comparator.py     # MQTT 快照 vs Modbus 三段式比对主程序
│   └── <设备名>data.json      # 网关推送的 MQTT JSON 快照（如 4100data.json）
├── Datas/                     # 测试数据文件
│   ├── acuclouddatas/         # AcuCloud 导出 xlsx 快照
│   └── awsdatas/              # AWS IoT / EDS 相关数据
│       └── eds/               # 设备 EDS 文件（EtherNet/IP 用）
│           └── AcuRev-4100.eds
├── modbusAddress/             # 各设备 Modbus 地址表 xlsx
│   └── mapping_compile.py
├── template/                  # 设备参数模板 xlsx（按设备型号子目录）
│   ├── AcuRev-4100&PXB/
│   ├── AcuRev-2100/
│   ├── AcuvimIIW&PXE2/
│   ├── AcuvimIIR&PXE1/
│   ├── Acuvim3/
│   ├── AcuIOM/
│   └── AcuRev-1300&PXM350/
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
│   ├── acuiom03.py            # 占位，DI型号暂不支持
│   └── acuiom04.py            # 占位，DI型号暂不支持
└── reports/                   # HTML 比对报告输出目录
```
