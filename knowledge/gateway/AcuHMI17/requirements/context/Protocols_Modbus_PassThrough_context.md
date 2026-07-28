# Protocols / Modbus / Pass Through — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/modbus/passThrough` |
| 路由名 | `passThrough` |
| 面包屑 | AcuHMI-1-7 / Protocols / Modbus / Pass Through |
| 顶级模块 | Protocols → Modbus |

## 2. 页面用途

Modbus 透传：将上位机对某 Slave ID 的请求原样转发给对应下游设备。为每个启用透传的设备分配一个 Slave ID（范围 **101–247**）。与 Device Mirror(2–99)、Parameters Mapping(如100) 的 Slave ID 段区分，避免冲突。

## 3. 页面结构

1. 协议标签栏
2. 标题 Pass Through
3. `*Pass Through Enable`：Enable/Disable 单选（默认 Enable）
4. **透传设备表**：Enable(表头全选) / SlaveID(可编辑) / Device Name / Interface / Protocol / Model / Serial Number
5. Save 按钮

## 4. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 定位策略 2 | 说明 |
|------|------|-----------|-----------|------|
| Pass Through Enable | radiogroup | `getByRole('radiogroup',{name:'Pass Through Enable'})` | 文本邻接 | Enable/Disable |
| 全选 Enable | checkbox | 表头 columnheader "Enable" 内 checkbox | 三态 mixed | 全选/全不选 |
| 行 Enable | checkbox | 行内 checkbox | 行首列 | 勾选启用透传并启用 SlaveID 编辑 |
| 行 SlaveID | textbox | 行 "SlaveID" 单元格内 textbox | 单元格 (Range: 101-247) | 仅启用行可编辑 |
| Save | button | `getByRole('button',{name:'Save'})` | 底部主按钮 | 保存 |

## 5. 表单字段与校验规则

- `*Pass Through Enable`：必选，默认 Enable。
- 行 `SlaveID`：整数，**范围 101–247**（每行提示 "Range: 101 - 247"）；跨设备唯一。
- 未勾选 Enable 的行，SlaveID 输入框 `disabled`（值常为默认递增或 0）。示例：AcuRev2100=247(启用)、AcuRev4100_392=101(启用)，其余 102/103/104/0 未启用为 disabled。

## 6. 页面状态与分支

| 状态 | 说明 |
|------|------|
| 行未启用 | SlaveID 输入框 disabled |
| 行已启用 | SlaveID 可编辑 |
| Disable 整体 | 关闭透传功能 |

## 7. 自动化测试要点

- SlaveID 范围校验：101/247 边界、100/248 越界、非数字、与其它段(2-99镜像 / paramsMapping)冲突。
- 勾选/取消 Enable 联动 SlaveID 输入框 enabled 状态。
- 三种映射方式（Parameters Mapping / Device Mirror / Pass Through）的 Slave ID 段位互不重叠是关键端到端断言。

## 8. 机器可解析摘要

```json
{
  "route": "/protocols/modbus/passThrough",
  "name": "passThrough",
  "title": "Pass Through",
  "module": "Protocols/Modbus",
  "fields": {
    "Pass Through Enable": {"type":"radio","default":"Enable"},
    "row.Enable": {"type":"checkbox"},
    "row.SlaveID": {"type":"text","range":[101,247],"disabled_when":"row not enabled"}
  },
  "buttons": ["Save"],
  "slaveid_segments": {"device_mirror":[2,99],"params_mapping":"e.g.100","pass_through":[101,247]}
}
```
