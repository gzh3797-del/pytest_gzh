# Protocols / MQTT / Topic and Parameter Selection — 页面上下文

> 路由名 `deviceToPublish`，UI 显示 **Topic and Parameter Selection**。

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/mqtt/deviceToPublish` |
| 路由名 | `deviceToPublish` |
| 面包屑 | AcuHMI-1-7 / Protocols / MQTT / Topic and Parameter Selection |
| 顶级模块 | Protocols → MQTT |

## 2. 页面用途

配置 MQTT 发布：基础 Topic、QoS、Retained、发布间隔，并为每个待发布设备选择要上报的参数（通过弹框穿梭框）。附 Payload JSON 格式示例。

## 3. 交互元素清单（主页面）

| 元素 | 类型 | 定位策略 1 | 定位策略 2 | 默认/示例 | 说明 |
|------|------|-----------|-----------|-----------|------|
| Base Topic | textbox | `getByRole('textbox',{name:'Base Topic'})` | placeholder "Enter Base Topic" | accuenergy/pytest | 必填(*)，带清除图标 |
| Qos | combobox | `getByRole('combobox',{name:'Qos'})` | 文本 "Qos 0" 容器 | Qos 0 | 必选(*)，0/1/2 |
| Retained | radiogroup | `getByRole('radiogroup',{name:'Retained'})` | 文本邻接 | No | 必选(*) Yes/No |
| Interval | combobox | `getByRole('combobox',{name:'Interval'})` | 文本 "30 seconds" | 30 seconds | 必选(*)，发布间隔 |
| Payload Format | 只读 | group "Payload Format" | — | JSON 示例 | 展示上报报文结构 |
| 设备表全选 | checkbox | 表头 columnheader 内 checkbox | 三态 mixed | — | 全选/全不选设备 |
| 行设备勾选 | checkbox | 行内 checkbox | 行首列 | — | 选择该设备参与发布 |
| Parameter Selection | button | 行 `getByRole('row',{name:/Acurev1234100/}).getByRole('button')` | 行末列图标按钮 | — | 打开参数选择弹框 |
| Save | button | `getByRole('button',{name:'Save'})` | 底部主按钮 | — | 保存 |

## 4. 设备表列

Checkbox / Device Name / Serial Number / Protocol(Modbus TCP/RTU) / Online(ON/OFF) / Parameter Selection(按钮)

## 5. 弹框：MQTT Parameter Config ★

点击行内 Parameter Selection 按钮弹出：

| 元素 | 类型 | 定位策略 | 说明 |
|------|------|----------|------|
| 标题 | heading | `getByRole('heading',{name:'MQTT Parameter Config'})` | 弹框标题 |
| Close | button | `getByRole('button',{name:'Close this dialog'})` | 右上角关闭 |
| Parameter Type | combobox | `getByRole('combobox',{name:'Parameter Type'})` | 默认 Realtime |
| Parameters 穿梭框 | transfer(dual list) | group "Parameters*" | 左"Not Selected"→右"Selected"，含 `>` `<` `All` `Clear` 按钮 |
| All / Clear | button | `getByRole('button',{name:'All'})` / `{name:'Clear'}` | 全部移入/清空 |
| Confirm | button | `getByRole('button',{name:'Confirm'})` | 确认选择 |
| Cancel | button | `getByRole('button',{name:'Cancel'})` | 取消关闭 |

- Not Selected 列表包含该设备全部可上报参数（System Frequency、各相电压/电流/功率、Input/User Channel 1–24 等，共上百项）。

## 6. 表单字段与校验规则

- Base Topic / Qos / Retained / Interval 均必填(*)。
- 至少需选择一个设备（Devices Selection To Mapping*）及其参数。

## 7. 自动化测试要点

- 弹框穿梭框：断言 All/Clear、单项 `>`/`<` 移动后 Selected/Not Selected 列表变化。
- Parameter Type 切换可能改变可选参数集合。
- 主流程：填 Topic/QoS/Retained/Interval → 勾设备 → 打开弹框选参数 → Confirm → Save。
- Element-Plus dialog；关闭用 Close/Cancel。

## 8. 机器可解析摘要

```json
{
  "route": "/protocols/mqtt/deviceToPublish",
  "name": "deviceToPublish",
  "title": "Topic and Parameter Selection",
  "module": "Protocols/MQTT",
  "fields": {
    "Base Topic": {"type":"text","required":true},
    "Qos": {"type":"select","options":[0,1,2],"default":0},
    "Retained": {"type":"radio","options":["Yes","No"],"default":"No"},
    "Interval": {"type":"select","default":"30 seconds"}
  },
  "device_table": ["checkbox","Device Name","Serial Number","Protocol","Online","Parameter Selection(button)"],
  "dialog": {"title":"MQTT Parameter Config","controls":["Parameter Type(select,default Realtime)","dual-list transfer","All","Clear","Confirm","Cancel"]},
  "buttons": ["Save"]
}
```
