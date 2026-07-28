# Modbus 地址表索引

原始 Excel 文件存放于 raw/ 目录，代码运行时直接读取（权威来源）。
Claude 参考本文件了解各设备寄存器范围，无需读原始 Excel。

| 设备 | Excel 文件（raw/） | FC | 数据类型 | 特殊说明 |
|------|------------------|----|---------|---------|
| AcuRev4100 | AcuRev4100 Modbus Address Table v1.02 20260202.xlsx | FC03 | Float32 | 端口2000，非标准 |
| AcuRev2100 | AcuRev2100_ Modbus Address_v1.02_20260406.xlsx | FC03 | Float32 | |
| AcuVIM3 | Acuvim3 User Modbus Address Table v1.08_Hongjian Zhu_260203.xlsx | FC03 | Float32 | |
| AcuvimIIW | Acuvim IIW&IIR&CL&EL Modbus Address v1.27_Haibo Song_260323.xlsx | FC03 | Float32 | 同文件含 IIR / CL(PXE1) / EL(PXE2)，后两者目前不适配 |
| AcuvimIIR | 同上 | FC03 | Float32 | 与IIW完全相同，共用文件 |
| AcuRev1300 | AcuRev1310_PXM350_ Modbus Address_v1.01_Sam Xu_260305.xlsx | FC03 | Float32 | 文件名中 AcuRev1310 为笔误（实为1300）；与 PXM350 为同一设备，共用文件 |
| AcuIOM-01 | AcuIOM Modbus Address Table v1.01 20260228 的副本.xlsx | FC03 | Float32 | 01~04 共用同一文件 |
| AcuIOM-02 | 同上 | FC03 | Float32 | |
| AcuIOM-03 | 同上 | FC02 | Bit | DI型号，框架暂不支持 |
| AcuIOM-04 | 同上 | FC02 | Bit | DI型号，框架暂不支持 |
| RPP (MH主控) | RPP Modbus Address Table v1.00 20260617.xlsx | FC03 | 混合（float / uint16 / uint32 / uint64 / double） | v1.00 / 2026-06-17；地址范围 0x1000~0x9730；含 Settings（R/W）和大量只读测量 Sheet；详见下方 Sheet 清单 |
| AcuRev-100（内部代号 RACG） | RACG Modbus Address Table.xlsx | FC03（部分 10H 写） | 混合（uint8/uint16/uint32/uint64/float） | 最新变更记录于表内 2026-06-30；仅 RS-485+USB Modbus RTU，无 BACnet Sheet；地址范围 0x1000~0xF070；含 Basic Setting（R/W）、Wiring Check（R/W）、Real Time、Energy、Calibration 等 Sheet；详见下方 Sheet 清单 |

## RPP Modbus Address Table — Sheet 清单（v1.00）

> 原件：`raw/RPP Modbus Address Table v1.00 20260617.xlsx`，共 18 个 Sheet。

| Sheet 名称 | 类型 | 地址范围（十六进制） | 说明 |
|-----------|------|------------------|------|
| Version history | 元数据 | — | 版本变更记录，无寄存器 |
| Overview | 索引 | — | 各 Block 地址分布总览，无实测寄存器行 |
| Basic Settings | 可读写 | 0x1000 ~ 0x10A0 | 通讯/时间等系统设置，FC03 读、FC10 写 |
| Wring Check | 可读写 | 0x1300 ~ 0x1300+ | 接线检查触发与状态 |
| HS(Half Cycle) | 只读 | 0x2000 起 | 半周波 RMS 测量参数（电压/电流/功率等） |
| Energy（1s） | 只读 | 0x2500 起 | 1 秒聚合电能（VMM/Channel，active/reactive） |
| Demand | 只读 | 0x2700 起 | 需量（电流/有功/无功需量） |
| 10\|12 Cycle | 只读 | 0x2900 起 | 10/12 周波 RMS 测量参数 |
| 150\|180 Cycle | 只读 | 0x4000 起 | 150/180 周波 RMS 测量参数 |
| 10 Min | 只读 | 0x5710 起 | 10 分钟 RMS 测量参数（IEC 61000-4-30 Class A） |
| 2 Hr | 只读 | 0x6E20 起 | 2 小时 RMS 测量参数 |
| 10S Freq | 只读 | 0x8530 起 | 10 秒频率测量（VMM1/VMM2） |
| MaxMin | 只读 | 0x8550 起 | 极值记录（最大值/最小值 + 时间戳） |
| Max Demand | 只读 | 0x9270 起 | 最大需量记录（电流/功率 + 时间戳） |
| Waveform\|PQ Event | 可读写 | 0x9290 起 | PQ 事件配置（标称电压、电压骤降/骤升使能等）及波形相关 |
| Information | 只读 | 0x9650 起 | 设备信息（设备类型、序列号等） |
| DeviceStatus | 只读 | 0x9720 起 | 设备运行状态（运行模式） |
| AlarmStatus | 空 Sheet | — | 当前无内容 |

**可读寄存器 Sheet（16 个）：** Basic Settings、Wring Check、HS(Half Cycle)、Energy（1s）、Demand、10\|12 Cycle、150\|180 Cycle、10 Min、2 Hr、10S Freq、MaxMin、Max Demand、Waveform\|PQ Event、Information、DeviceStatus（均 FC03 读）。Version history、Overview 为元数据/索引，AlarmStatus 为空，共 3 个 Sheet 不含可读寄存器。

## RACG（AcuRev-100）Modbus Address Table — Sheet 清单

> 原件：`raw/RACG Modbus Address Table.xlsx`，共 8 个 Sheet。RACG 为 AcuRev-100（含 AcuRev-101-mA / AcuRev-101-mV 两型号）内部代号，详见 knowledge/meters/AcuRev100/context.md。

| Sheet 名称 | 类型 | 地址范围（十六进制） | 说明 |
|-----------|------|------------------|------|
| Version history | 元数据 | — | 版本变更记录，无寄存器 |
| Overview | 索引 | 0xF000 ~ 0x301D | 各 Block 地址分布总览，无实测寄存器行 |
| Information | 只读 | 0xF000 ~ 0xF070 | 固件版本、Bootloader、产品信息（03H 读） |
| Basic Setting | 可读写 | 0x1000 ~ 0x1148 | 系统设置（RS485/USB 波特率与校验、密码等）、计量设置、Operations（03H 读 / 10H 写） |
| Wiring Check | 可读写 | 0x3000 ~ 0x301D | 接线检查开关/状态/错误码 + 各相电压电流实测值（03H 读 / 10H 写） |
| Real Time | 只读 | 0x9050 起 | 实时基础参数（频率、相电压、相角等，100ms 级） |
| Energy | 只读 | 0x4A00 起 | 分相/系统有功电能（1 秒级），含 Energy Parameter ID 列 |
| Calibration | 可读写 | 0x2000 起 | 工厂校准命令/结果/校准偏移系数（03H 读 / 10H 写） |

**可读寄存器 Sheet（6 个）：** Information、Basic Setting、Wiring Check、Real Time、Energy、Calibration。Version history 为元数据、Overview 为索引总览，不含可读寄存器行。

## 更新说明
新增设备后在本表追加一行，并将原始 Excel 放入 raw/ 目录。
