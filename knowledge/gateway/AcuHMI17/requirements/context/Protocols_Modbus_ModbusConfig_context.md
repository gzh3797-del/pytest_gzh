# Protocols / Modbus / Modbus Config — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/modbus/modbusConfig` |
| 路由名 | `modbusConfig` |
| 面包屑 | AcuHMI-1-7 / Protocols / Modbus / Modbus Config |
| 顶级模块 | Protocols → Modbus |

## 2. 页面用途

配置本网关作为 Modbus Slave/服务端对外服务的基础参数：启用开关与监听端口；并提供 Modbus 数据类型与寄存器/字节序对照参考表。

## 3. 页面结构

1. 协议标签栏：Modbus▼ | SNMP | BACnet/IP | MQTT▼ | AWS IoT | Azure IoT
2. 标题 Modbus Config
3. `Modbus Enable*`：Enable/Disable 单选（默认 Enable）
4. `Modbus Port*`：文本框，默认 502，范围 2000–5999（示例值 5999）
5. `Tips` 提示
6. **数据类型参考表**（只读）：Type / Register Length / Data Length / Sequence
7. Save 按钮

## 4. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 定位策略 2 | 说明 |
|------|------|-----------|-----------|------|
| Modbus Enable | radiogroup | `getByRole('radiogroup',{name:'Modbus Enable'})` | 文本 "Modbus Enable*" 邻接 | Enable/Disable |
| Modbus Port | textbox | `getByRole('textbox',{name:'Modbus Port'})` | placeholder "Enter Modbus Port" | 必填，默认502，范围2000-5999 |
| Save | button | `getByRole('button',{name:'Save'})` | 页面底部主按钮 | 保存配置 |

## 5. 表单字段与校验规则

- `Modbus Enable*`：必选，默认 Enable。
- `Modbus Port*`：必填；默认 502；**取值范围 2000–5999**（超范围应报错）。

## 6. 数据类型参考表（只读，用于理解参数映射的地址占用）

| Type | Register Length | Data Length | Sequence |
|------|-----------------|-------------|----------|
| Bit | 1 | 16 | A |
| Uint16 | 1 | 16 | AB |
| Int32 | 2 | 32 | ABCD |
| Uint32 | 2 | 32 | ABCD |
| Long | 2 | 32 | ABCD |
| Float | 2 | 32 | ABCD |
| Double | 4 | 64 | ABCDEFGH |

## 7. 自动化测试要点

- 端口校验为核心用例：边界值 1999(下界外)/2000/5999/6000(上界外)、非数字、空值。
- Disable 后应影响 Modbus 对外服务（可与 Parameters Mapping 联动验证）。
- 数据类型表为静态只读，通常仅做存在性断言。

## 8. 机器可解析摘要

```json
{
  "route": "/protocols/modbus/modbusConfig",
  "name": "modbusConfig",
  "title": "Modbus Config",
  "module": "Protocols/Modbus",
  "fields": {
    "Modbus Enable": {"type":"radio","options":["Enable","Disable"],"default":"Enable"},
    "Modbus Port": {"type":"text","required":true,"default":502,"range":[2000,5999]}
  },
  "reference_table": "data type: Bit/Uint16/Int32/Uint32/Long/Float/Double",
  "buttons": ["Save"]
}
```
