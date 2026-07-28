# Protocols / AWS IoT — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/awsIot` |
| 路由名 | `awsIot` |
| 面包屑 | AcuHMI-1-7 / Protocols / AWS IoT |
| 顶级模块 | Protocols |

## 2. 页面用途

配置网关向 AWS IoT Core 上报数据：连接参数、证书/密钥、发布 Topic 与间隔，选择上报设备与参数。**条件分支页**（默认 Disable，仅开关；Enable 后显示完整配置）。

## 3. 交互元素清单 / 表单字段（Enable 后）

| 字段/元素 | 类型 | 定位策略 1 | 默认/示例 | 说明 |
|-----------|------|-----------|-----------|------|
| AWS IoT Enable | radiogroup | `getByRole('radiogroup',{name:'AWS IoT Enable'})` | Disable | 必选(*)，控制下方显隐 |
| Client Id | textbox | `getByRole('textbox',{name:'Client Id'})` | AHI260110002 | 必填(*) |
| URL | textbox | `getByRole('textbox',{name:'URL'})` | *.iot.cn-northwest-1.amazonaws.com.cn | 必填(*)，AWS IoT endpoint |
| Topic | textbox | `getByRole('textbox',{name:'Topic'})` | test/topic/aws | 必填(*) |
| Interval | combobox | `getByRole('combobox',{name:'Interval'})` | 30 seconds | 必选(*) |
| Cert File | upload | `group('Cert File').getByRole('button',{name:'Browse'})` | client.pem | 上传设备证书 |
| Key File | upload | `group('Key File').getByRole('button',{name:'Browse'})` | key.pem | 上传私钥 |
| Test Connection | button | `getByRole('button',{name:'Test Connection'})` | — | 测试连接 AWS |
| Devices Selection 表 | table | group "Devices Selection" | — | 见下 |
| Parameter Selection | button | 行末图标按钮 | — | 打开参数选择弹框（穿梭框） |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. Devices Selection 表列

checkbox / Device Name / **Device Type**(Modbus Device / Virtual Device) / Serial Number / Protocol / Online / Parameter Selection(按钮)

> 该表同时包含物理 Modbus 设备与大量 Virtual Device（VD_*），Virtual Device 行 Protocol/Online 显示 "-"。

## 5. 页面状态与分支 ★

| 状态 | 触发 | 结果 |
|------|------|------|
| Disable（默认） | 进入页面 | 仅 AWS IoT Enable 单选 + Save |
| Enable | 选 Enable | 显示 Client Id/URL/Topic/Interval/证书/Test Connection/设备表 |

## 6. 自动化测试要点

- 条件显隐分支（核心）。
- 证书/密钥上传（.pem）及回显。
- Test Connection 结果提示；依赖真实 AWS，测试环境需 mock。
- 设备表混合物理+虚拟设备，Parameter Selection 弹框同 MQTT 穿梭框模式。

## 7. 机器可解析摘要

```json
{
  "route": "/protocols/awsIot",
  "name": "awsIot",
  "title": "AWS IoT",
  "module": "Protocols",
  "fields": {
    "AWS IoT Enable": {"type":"radio","default":"Disable"},
    "Client Id": {"type":"text","required":true},
    "URL": {"type":"text","required":true},
    "Topic": {"type":"text","required":true},
    "Interval": {"type":"select","default":"30 seconds"},
    "Cert File": {"type":"upload","ext":".pem"},
    "Key File": {"type":"upload","ext":".pem"}
  },
  "conditional": {"when":"AWS IoT Enable=Enable","shows":["Client Id","URL","Topic","Interval","Cert File","Key File","Test Connection","Devices Selection"]},
  "device_table_columns": ["checkbox","Device Name","Device Type","Serial Number","Protocol","Online","Parameter Selection"],
  "buttons": ["Test Connection","Save"]
}
```
