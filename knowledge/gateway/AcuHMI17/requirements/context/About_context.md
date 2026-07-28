# About — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/about` |
| 路由名 | `about` |
| 面包屑 | About |
| 入口 | 顶部 header "About" 按钮 |

## 2. 页面用途

展示设备信息（型号/序列号/版本/MAC 等），并可编辑设备的名称/位置/描述。

## 3. 交互元素清单 / 表单字段（Device Information 表：Setting/Value）

| 字段 | 类型 | 定位策略 1 | 默认/示例 | 说明 |
|------|------|-----------|-----------|------|
| Name | textbox | `getByRole('textbox',{name:'---Enter Name---'})` | (示例 aaaa...) | 可编辑，最多 40 字符 |
| Location | textbox | `getByRole('textbox',{name:'---Enter Location---'})` | (示例 bbbb...) | 可编辑，最多 40 字符 |
| Description | textbox | `getByRole('textbox',{name:'---Enter Description---'})` | (示例 cccc...) | 可编辑，最多 40 字符 |
| Model | 只读 | 行 "Model" | AcuHMI-1-7 | 只读 |
| Serial Number | 只读 | 行 "Serial Number" | AHI260110002 | 只读 |
| Hardware Version | 只读 | 行 "Hardware Version" | v1.05 | 只读 |
| Firmware Version | 只读 | 行 "Firmware Version" | v1.03p05 | 只读 |
| Last Updated | 只读 | 行 "Last Updated" | 06/08/2026 | 只读 |
| Ethernet1 MAC Address | 只读 | 行 "Ethernet1 MAC Address" | 30:7a:57:01:f5:36 | 只读 |
| Ethernet2 MAC Address | 只读 | 行 "Ethernet2 MAC Address" | 30:7a:57:01:f5:37 | 只读 |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存（仅 Name/Location/Description 可改） |

## 4. 自动化测试要点

- 编辑 Name/Location/Description（≤40 字符校验）→ Save。
- 只读字段（型号/序列号/版本/MAC）的呈现断言。

## 5. 机器可解析摘要

```json
{
  "route": "/about",
  "name": "about",
  "title": "About",
  "editable": {"Name":{"maxlen":40},"Location":{"maxlen":40},"Description":{"maxlen":40}},
  "readonly": ["Model","Serial Number","Hardware Version","Firmware Version","Last Updated","Ethernet1 MAC Address","Ethernet2 MAC Address"],
  "buttons": ["Save"]
}
```
