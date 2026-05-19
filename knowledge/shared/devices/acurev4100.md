# AcuRev4100

## 基本信息
- 类型：多功能电表
- 通信协议：Modbus TCP，FC03，Float32，Big Endian，2寄存器/参数
- 典型端口：2000（非标准，注意与其他设备区分）
- 典型 Unit ID：102

## 项目适配情况
| 项目 | BACnet比对 | AcuCloud比对 | Modbus模块 |
|------|-----------|-------------|-----------|
| web2 | ✅ | ✅ | devices/acurev4100.py |

## AcuCloud 文件规则
- xlsx 文件名：AcuRev4100.xlsx
- 存放目录：Acuclouddatas/

## 已知问题
- 端口 2000 为非标准端口，config.py 中 MODBUS_DEVICE_MAP 须单独配置，不能默认 502
