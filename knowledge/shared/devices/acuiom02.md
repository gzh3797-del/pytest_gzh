# AcuIOM-02

## 基本信息
- 类型：IO模块，AIAO型（模拟量输入输出）
- 通信协议：Modbus TCP，FC03，Float32，Big Endian，2寄存器/参数
- AI 通道数：16
- 寄存器范围：0x3700–0x371E（AI物理量读值）
- 典型端口：502

## 项目适配情况
| 项目 | BACnet比对 | AcuCloud比对 | Modbus模块 |
|------|-----------|-------------|-----------|
| web2 | ✅ | — | devices/acuiom02.py |

## 注意
- 模板格式与标准设备不同（无 BACnetIP / AcuCloud 列），范围检查会被跳过（graceful skip）
- 不支持 AcuCloud 比对
- 与 AcuIOM-01 协议完全相同，仅通道数不同（16 vs 8）
