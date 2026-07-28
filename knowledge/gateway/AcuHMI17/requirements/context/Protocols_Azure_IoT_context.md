# Protocols / Azure IoT — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/azureIot` |
| 路由名 | `azureIot` |
| 面包屑 | AcuHMI-1-7 / Protocols / Azure IoT |
| 顶级模块 | Protocols |

## 2. 页面用途

配置网关向 Azure IoT Hub 上报数据：连接字符串、上报间隔、SSL，选择上报设备与参数。

## 3. 交互元素清单 / 表单字段

| 字段/元素 | 类型 | 定位策略 1 | 默认/示例 | 说明 |
|-----------|------|-----------|-----------|------|
| Azure IoT Enable | radiogroup | `getByRole('radiogroup',{name:'Azure IoT Enable'})` | Enable | 必选(*)，控制显隐 |
| Primary Connection String | textbox | `getByRole('textbox',{name:'Primary Connection String'})` | 空 | Azure 主连接串 |
| Secondary Connection String | textbox | `getByRole('textbox',{name:'Secondary Connection String'})` | 空 | 备用连接串 |
| Interval | combobox | `getByRole('combobox',{name:'Interval'})` | 1 seconds | 必选(*) |
| Enable SSL | radiogroup | `getByRole('radiogroup',{name:'Enable SSL'})` | Disable | 必选(*) |
| Test Connection | button | `getByRole('button',{name:'Test Connection'})` | — | 测试连接 Azure |
| Devices Selection 表 | table | group "Devices Selection" | — | checkbox/Device Name/Device Type/Serial Number/Protocol/Online/Parameter Selection |
| Parameter Selection | button | 行末图标按钮 | — | 打开参数选择弹框（穿梭框） |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. Devices Selection 表列

checkbox / Device Name / Device Type(Modbus Device / Virtual Device) / Serial Number / Protocol / Online / Parameter Selection

## 5. 页面状态与分支

| 状态 | 触发 | 结果 |
|------|------|------|
| Enable | 默认 | 显示连接串/间隔/SSL/设备表 |
| Disable | 切换 | 隐藏配置 |
| Enable SSL | 切换 | 可能出现证书上传（需运行时确认，参照 MQTT SSL 模式） |

## 6. 自动化测试要点

- 连接串必填/格式校验；Interval/SSL 选择。
- Test Connection 依赖真实 Azure，测试环境需 mock。
- 设备表混合物理+虚拟设备；Parameter Selection 弹框同 MQTT 穿梭框。

## 7. 机器可解析摘要

```json
{
  "route": "/protocols/azureIot",
  "name": "azureIot",
  "title": "Azure IoT",
  "module": "Protocols",
  "fields": {
    "Azure IoT Enable": {"type":"radio","default":"Enable"},
    "Primary Connection String": {"type":"text"},
    "Secondary Connection String": {"type":"text"},
    "Interval": {"type":"select","default":"1 seconds"},
    "Enable SSL": {"type":"radio","default":"Disable"}
  },
  "device_table_columns": ["checkbox","Device Name","Device Type","Serial Number","Protocol","Online","Parameter Selection"],
  "buttons": ["Test Connection","Save"]
}
```
