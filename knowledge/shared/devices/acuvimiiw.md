# AcuvimIIW / AcuvimIIR

## 基本信息
- 类型：多功能电表（IIW = 网络版，IIR = R型）
- 通信协议：Modbus TCP，FC03，Float32，Big Endian，2寄存器/参数
- 典型端口：502
- IIW 与 IIR 的 Modbus 寄存器地址完全相同

## 项目适配情况
| 项目 | BACnet比对 | AcuCloud比对 | Modbus模块 |
|------|-----------|-------------|-----------|
| web2 | ✅ | ✅ | devices/acuvimiiw.py / acuvimiir.py |

## 重要说明
AcuvimIIR 的 build_param_map() 和 build_cloud_col_map() 直接 import 自 acuvimiiw.py，
不单独实现，两者地址表和云端列名完全一致。

## AcuCloud 文件规则
- IIW：AcuvimIIW.xlsx
- IIR：AcuvimIIR.xlsx
