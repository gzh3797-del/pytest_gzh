# AcuIOM-03

## 基本信息
- 类型：IO模块，DIDO型（数字量输入输出）
- 通信协议：Modbus TCP，FC02（Discrete Inputs）
- 典型端口：502

## 项目适配情况
| 项目 | BACnet比对 | AcuCloud比对 | Modbus模块 |
|------|-----------|-------------|-----------|
| web2 | ⚠️ 暂不支持 | — | devices/acuiom03.py（空stub） |

## 暂不支持原因

DI Status 寄存器位于 **Discrete Input 区（0x 地址段，0x0000–0x001B，共 28 通道）**，Modbus 协议规定该区域只能用 FC02 读取；FC03 只能访问 Holding Register（4x 地址段），两者不可混用。BACnet 侧对应发布为 Binary Input 对象，而非 Analog Input，整个框架结构均不匹配。

- 待办：需在读取层实现 FC02（read_discrete_inputs）路径 + Binary Input 解析
- 决策溯源：→ [decisions.md — IOM-03/04 返回空 map，暂不支持](../decisions.md)
