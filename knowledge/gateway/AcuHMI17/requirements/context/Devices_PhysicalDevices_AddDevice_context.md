# Devices / Physical Devices / Add Device — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/physicalDevices/addDevice`（`addDevice`） |
| 相关路由 | `searchDevice`（自动搜索，复用同表单）、`deviceDetails/:deviceKey/:currentKey`（编辑，预填） |
| 面包屑 | Devices / Physical Devices / Add Device |

## 2. 页面用途

新增物理设备并配置其 Modbus 连接参数。**RTU/TCP 条件表单**。

## 3. 交互元素清单 / 表单字段

| 字段 | 类型 | 定位策略 1 | 默认/示例 | 校验/说明 |
|------|------|-----------|-----------|-----------|
| Device Name | textbox | `getByRole('textbox',{name:'Device Name'})` | — | 必填(*)，≤40 字符 |
| Serial Number | textbox | `getByRole('textbox',{name:'Serial Number'})` | — | 必填(*)，设备内唯一，≤40 字符 |
| Template | combobox | `getByRole('combobox',{name:'Template'})` | ---Select Template--- | 必选(*) |
| Protocol | radiogroup | `getByRole('radiogroup',{name:'Protocol'})` | RTU | RTU/TCP（控制下方字段） |
| Port | combobox | `getByRole('combobox',{name:'Port'})` | COM 1 | RTU: COM 口 |
| Modbus ID | textbox | `getByRole('textbox',{name:'Modbus ID'})` | — | 必填，COM 内唯一，**范围 1–246** |
| Baud Rate | combobox | `getByRole('combobox',{name:'Baud Rate'})` | 38400 | 必选 |
| Data Bit | combobox(disabled) | `getByRole('combobox',{name:'Data Bit'})` | 8 | 只读 |
| Parity | combobox | `getByRole('combobox',{name:'Parity'})` | None1 | 必选 |
| Stop Bit | combobox(disabled) | `getByRole('combobox',{name:'Stop Bit'})` | 1 | 只读 |
| Request Timeout | textbox | `getByRole('textbox',{name:'Request Timeout'})` | 0.5 | 必填，**范围 0.1–5** s |
| Poll Interval | textbox | `getByRole('textbox',{name:'Poll Interval'})` | 60 | 必填，**范围 1–3600** s |
| Add to Logger | combobox | `getByRole('combobox',{name:'Add to Logger'})` | ---Select Logger--- | 必选(*) |
| Save / Cancel | button | `getByRole('button',{name:'Save'})` / `{name:'Cancel'}` | — | 提交/取消 |

## 4. 页面状态与分支 ★

| 状态 | 触发 | 结果 |
|------|------|------|
| Protocol = RTU（默认） | — | 显示串口字段：Port(COM)/Baud Rate/Data Bit/Parity/Stop Bit |
| Protocol = TCP | 切换 | 改为 IP 地址 / TCP 端口等字段（运行时确认） |

## 5. 校验规则要点

- Device Name/Serial Number ≤40，Serial Number 唯一。
- Modbus ID 1–246 且 COM 内唯一。
- Request Timeout 0.1–5；Poll Interval 1–3600。

## 6. 自动化测试要点

- RTU/TCP 切换的字段显隐（核心分支）。
- 各数值范围/唯一性校验；Template 与 Add to Logger 必选。
- searchDevice 复用此表单（自动扫描发现设备）；deviceDetails 为编辑（预填）。

## 7. 机器可解析摘要

```json
{
  "route": "/physicalDevices/addDevice",
  "name": "addDevice",
  "title": "Add Device",
  "module": "Devices/Physical Devices",
  "fields": {
    "Device Name": {"type":"text","required":true,"maxlen":40},
    "Serial Number": {"type":"text","required":true,"unique":true,"maxlen":40},
    "Template": {"type":"select","required":true},
    "Protocol": {"type":"radio","options":["RTU","TCP"],"default":"RTU"},
    "Port": {"type":"select","when":"RTU"},
    "Modbus ID": {"type":"text","range":[1,246],"unique_scope":"COM"},
    "Baud Rate": {"type":"select","default":38400},
    "Data Bit": {"type":"select","default":8,"disabled":true},
    "Parity": {"type":"select","default":"None1"},
    "Stop Bit": {"type":"select","default":1,"disabled":true},
    "Request Timeout": {"type":"text","range":[0.1,5],"unit":"s"},
    "Poll Interval": {"type":"text","range":[1,3600],"unit":"s"},
    "Add to Logger": {"type":"select","required":true}
  },
  "conditional": {"when":"Protocol=TCP","shows":["IP","TCP port"],"hides":["Baud Rate","Parity","Data/Stop Bit"]},
  "buttons": ["Save","Cancel"],
  "shared_by": ["searchDevice","deviceDetails(edit)"]
}
```
