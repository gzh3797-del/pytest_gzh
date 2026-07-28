# Protocols / Modbus / Device List — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/modbus/modbusDeviceList` |
| 路由名 | `modbusDeviceList` |
| 面包屑 | AcuHMI-1-7 / Protocols / Modbus / Device List |
| 顶级模块 | Protocols → Modbus |

## 2. 页面用途

只读汇总列表：展示当前经由 Modbus 对外暴露的所有设备条目（含网关自身 Device Mirror、各下游设备的 Parameters Mapping / Pass Through / Device Mirror 三类映射方式），及其分配的 Slave ID。用于核对 Slave ID 分配是否冲突。

## 3. 页面结构

1. 协议标签栏：Modbus▼ | SNMP | BACnet/IP | MQTT▼ | AWS IoT | Azure IoT
2. 标题 Device List
3. **设备表**（可排序，只读）：Device Name / Protocol / Model / Serial Number / Slave ID
4. 分页器（上一页 / 页码 / 下一页；本示例单页，前后翻页禁用）

## 4. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 定位策略 2 | 说明 |
|------|------|-----------|-----------|------|
| 列头排序(Device Name等) | columnheader(可点) | `getByRole('columnheader',{name:'Slave ID'})` | 表头单元格 (cursor=pointer) | 点击排序 |
| 数据行 | row | `getByRole('row',{name:/AcuRev4100_392 Parameters Mapping/})` | 表体行 | 只读 |
| 上一页 | button | `getByRole('button',{name:'Go to previous page'})` | 分页左箭头 | 单页时 disabled |
| 下一页 | button | `getByRole('button',{name:'Go to next page'})` | 分页右箭头 | 单页时 disabled |

## 5. 表列语义

| 列 | 含义 |
|----|------|
| Device Name | 设备名（可重复，同设备可有多种映射方式） |
| Protocol | 映射方式：`Device Mirror` / `Parameters Mapping` / `Pass Through` |
| Model | 设备型号 |
| Serial Number | 序列号 |
| Slave ID | 该条目对外的 Modbus 从站地址 |

示例数据（9 行）：AcuHMI-1-7(Device Mirror, ID 1)、AcuRev4100_392(Parameters Mapping, 100 / Pass Through, 101 / Device Mirror, 98)、AcuRev2100(Pass Through, 247 / Device Mirror, 99)、Acurev1234100(Device Mirror, 4)、AcuvimIIW(5)、Acurev4100242(6)。

## 6. 页面状态与分支

| 状态 | 说明 |
|------|------|
| 有数据 | 表格列出全部映射条目 |
| 单页 | 前后翻页按钮 disabled |
| 排序态 | 点击列头切换升/降序 |

## 7. 自动化测试要点

- 纯只读页，用例以"数据呈现正确性/排序/分页"为主。
- 断言 Slave ID 唯一性/映射方式枚举值。
- 数据依赖于 Parameters Mapping、Pass Through、Device Mirror 三页的配置，可做端到端联动验证。

## 8. 机器可解析摘要

```json
{
  "route": "/protocols/modbus/modbusDeviceList",
  "name": "modbusDeviceList",
  "title": "Device List",
  "module": "Protocols/Modbus",
  "read_only": true,
  "table_columns": ["Device Name","Protocol","Model","Serial Number","Slave ID"],
  "protocol_enum": ["Device Mirror","Parameters Mapping","Pass Through"],
  "features": ["column_sort","pagination"]
}
```
