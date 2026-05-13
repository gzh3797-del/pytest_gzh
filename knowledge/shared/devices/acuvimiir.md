# AcuvimIIR

## 基本信息
- 类型：多功能电表（R型）
- 通信协议：Modbus TCP，FC03，Float32，Big Endian，2寄存器/参数
- 典型端口：502

## 重要说明
AcuvimIIR 的 Modbus 寄存器地址与 AcuvimIIW 完全相同。
代码中 `build_param_map()` 和 `build_cloud_col_map()` 直接从 acuvimiiw.py import，不单独实现。

## 项目适配情况
| 项目 | BACnet比对 | AcuCloud比对 | Modbus模块 |
|------|-----------|-------------|-----------|
| web2 | ✅ | ✅ | devices/acuvimiir.py（复用 acuvimiiw.py） |

## AcuCloud 文件规则
- xlsx 文件名：AcuvimIIR.xlsx
- 存放目录：Acuclouddatas/
