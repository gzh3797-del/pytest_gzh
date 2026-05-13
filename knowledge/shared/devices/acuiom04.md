# AcuIOM-04

## 基本信息
- 类型：IO模块，DIDO型（数字量输入输出）
- 通信协议：Modbus TCP，FC02（Discrete Inputs）
- 典型端口：502

## 项目适配情况
| 项目 | BACnet比对 | AcuCloud比对 | Modbus模块 |
|------|-----------|-------------|-----------|
| web2 | ⚠️ 暂不支持 | — | devices/acuiom04.py（空stub） |

## 暂不支持原因

与 AcuIOM-03 相同：DI Status 位于 **Discrete Input 区（0x 地址段）**，只能用 FC02 读取；BACnet 侧发布为 Binary Input 对象，当前框架（FC03 + Analog Input）结构不匹配。

- 待办：需在读取层实现 FC02（read_discrete_inputs）路径 + Binary Input 解析
- 完整原因：→ [acuiom03.md — 暂不支持原因](./acuiom03.md)
- 决策溯源：→ [decisions.md — IOM-03/04 返回空 map，暂不支持](../decisions.md)
