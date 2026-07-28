# Protocols / Modbus / Device Mirror — 页面上下文

> 路由名 `logicalParameterMapping`，UI 显示为 **Device Mirror**。

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/modbus/logicalParameterMapping` |
| 路由名 | `logicalParameterMapping` |
| 面包屑 | AcuHMI-1-7 / Protocols / Modbus / Device Mirror |
| 顶级模块 | Protocols → Modbus |

## 2. 页面用途

将下游设备"镜像"为本网关下的 Modbus 从站，为每个启用镜像的设备分配一个 Slave ID（范围 2–99），使上位机可透过网关按各自 Slave ID 直接访问对应设备寄存器。网关自身固定为 Slave ID 1（不可编辑）。

## 3. 页面结构

1. 协议标签栏
2. 标题 Device Mirror
3. `*Device Mirror Enable`：Enable/Disable 单选（默认 Enable） + `Download All` 按钮
4. **设备镜像表**：Enable(表头全选) / SlaveID(可编辑输入框) / Device Name / Interface / Protocol / Model / Serial Number
5. Save 按钮

## 4. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 定位策略 2 | 说明 |
|------|------|-----------|-----------|------|
| Device Mirror Enable | radiogroup | `getByRole('radiogroup',{name:'Device Mirror Enable'})` | 文本邻接 | Enable/Disable |
| Download All | button | `getByRole('button',{name:'Download All'})` | 含 download 图标 | 导出镜像列表 |
| 全选 Enable | checkbox | 表头 columnheader "Enable" 内 checkbox | 三态 mixed | 全选/全不选 |
| 行 Enable | checkbox | 行内 checkbox | 行首列 | 勾选后启用该设备镜像并可编辑 SlaveID |
| 行 SlaveID | textbox | 行 "SlaveID" 单元格内 textbox | 单元格 (Range: 2-99) | 可编辑；范围 2–99 |
| Save | button | `getByRole('button',{name:'Save'})` | 底部主按钮 | 保存 |

## 5. 表单字段与校验规则

- `*Device Mirror Enable`：必选，默认 Enable。
- 行 `SlaveID`：整数，**范围 2–99**（每行输入框旁提示 "Range: 2 - 99"）；多设备间应唯一。
- **特例**：网关自身行 (AcuHMI-1-7, SlaveID=1) 的 Enable 复选框 `checked+disabled`、SlaveID 输入框 `disabled`（固定不可改）。
- 未勾选 Enable 的行（如 pxm350，SlaveID=0）其 SlaveID 输入框 `disabled`。

## 6. 页面状态与分支

| 状态 | 说明 |
|------|------|
| 行未启用 | SlaveID 输入框 disabled（值常为 0） |
| 行已启用 | SlaveID 可编辑 |
| 网关自身行 | Enable checked+disabled，SlaveID=1 固定 |
| Disable 整体 | 关闭镜像功能 |

## 7. 自动化测试要点

- SlaveID 范围校验：2/99 边界、1(被网关占用)、100(超界)、非数字、重复值。
- 勾选/取消行 Enable 会切换该行 SlaveID 输入框的 enabled 状态——是重要联动断言点。
- 网关自身行不可编辑，是负向用例（尝试编辑应无效）。

## 8. 机器可解析摘要

```json
{
  "route": "/protocols/modbus/logicalParameterMapping",
  "name": "logicalParameterMapping",
  "title": "Device Mirror",
  "module": "Protocols/Modbus",
  "fields": {
    "Device Mirror Enable": {"type":"radio","default":"Enable"},
    "row.Enable": {"type":"checkbox","note":"gateway self row checked+disabled"},
    "row.SlaveID": {"type":"text","range":[2,99],"disabled_when":"row not enabled or gateway self"}
  },
  "buttons": ["Download All","Save"]
}
```
