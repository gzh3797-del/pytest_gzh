# web2 — 网关自动化测试项目

## 项目背景
测试 Accuenergy 网关设备将下挂电表数据通过 BACnet/IP 协议发布到楼控系统的正确性。
同时验证 AcuCloud 云平台历史数据与实时 Modbus 读取值的一致性。

## 当前模块
| 模块 | 文件 | 功能 |
|------|------|------|
| BACnet比对 | comparator.py | BACnet Present Value vs 实时 Modbus，含范围/单位/数值三段检查 |
| AcuCloud比对 | cloud_comparator.py | AcuCloud xlsx快照 vs 实时 Modbus |
| BACnet读取 | bacnet_reader.py | BAC0 封装，对象发现、批量读值、元数据读取 |
| Modbus读取 | modbus_reader.py | Modbus TCP FC03 封装，Float32 Big Endian 解析 |
| 模板解析 | template_reader.py | 解析 blockParams xlsx，提取 BACnet/Cloud 参数范围 |

## 运行方式
```bash
# BACnet 比对（默认含范围检查 + 单位检查 + 数值比对）
python comparator.py --device acurev4100
python comparator.py --device acurev4100 --quick    # 只比对前30个参数
python comparator.py --device acurev4100 --no-meta  # 跳过单位检查

# AcuCloud 比对
python cloud_comparator.py --device acurev4100
python cloud_comparator.py --device acurev4100 --row 1  # 指定数据行
```

## 网络配置（当前测试环境）
- 网关 IP：192.168.2.209，BACnet 端口：48000
- 本机 BACnet 监听：192.168.2.45:47808
- 各设备 Modbus 配置见 config.py 的 MODBUS_DEVICE_MAP

## 报告输出
HTML 报告输出至 reports/ 目录，文件名格式：
- compare_<设备名>_<时间戳>.html（BACnet比对）
- cloud_<设备名>_<时间戳>.html（AcuCloud比对）
报告含三段可折叠区块（范围检查 / 单位检查 / 数值比对），summary行显示彩色badge。

## 目录结构
```
Protocols/                   ← 代码根目录（已并入主干）
├── BACnetIP/
│   ├── bacnet_reader.py
│   └── comparator.py
├── AcuCloud/
│   └── cloud_comparator.py
├── EtherNetIP/
│   ├── enip_reader.py
│   └── enip_comparator.py
├── modbus_reader.py
├── template_reader.py
├── config.py
├── devices/          ← Modbus地址表模块（每设备一个.py）
├── template/         ← 参数模板xlsx（blockParams sheet，按设备子目录）
├── Datas/acuclouddatas/  ← AcuCloud导出xlsx快照
└── reports/          ← HTML报告输出（自动创建）
```

## 支持设备
全部10个设备已适配，详见 CLAUDE.md 设备速查表。
IOM-03/04 为已知限制（FC02暂不支持），其余均完整支持。
