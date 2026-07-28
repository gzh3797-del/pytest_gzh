# Protocols / BACnet/IP — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/bacnet` |
| 路由名 | `bacnet` |
| 面包屑 | AcuHMI-1-7 / Protocols / BACnet/IP |
| 顶级模块 | Protocols |

## 2. 页面用途

配置网关的 BACnet/IP 服务：端口、网络号、设备对象、APDU 参数、Foreign Device/BBMD，以及将哪些下游设备参数映射为 BACnet 对象。可下载 EPICS 文件。

## 3. 交互元素清单 / 表单字段

| 字段/元素 | 类型 | 定位策略 1 | 默认/示例 | 校验/说明 |
|-----------|------|-----------|-----------|-----------|
| BACnet Enable | radiogroup | `getByRole('radiogroup',{name:'BACnet Enable'})` | Enable | 必选(*) |
| BACnet Port | textbox | `getByRole('textbox',{name:'BACnet Port'})` | 49000 | 必填，默认47808，范围 **47808–49000** |
| Network Number | textbox | `getByRole('textbox',{name:'Network Number'})` | 65534 | 必填，范围 **1–65534** |
| Device Object Name | textbox | `getByRole('textbox',{name:'Device Object Name'})` | TestGW | 必填，最多 40 字符 |
| Device Instance | textbox | `getByRole('textbox',{name:'Device Instance'})` | 26000 | 必填，**唯一**，范围 **0–4194302** |
| Advertised APDU Timeout | combobox | `getByRole('combobox',{name:'Advertised APDU Timeout'})` | 60 seconds | 必选 |
| Advertised APDU Retries | combobox | `getByRole('combobox',{name:'Advertised APDU Retries'})` | 10 | 必选 |
| Foreign Device Function | radiogroup | `getByRole('radiogroup',{name:'Foreign Device Function'})` | Enable | 控制 BBMD 组显隐 |
| BBMD IP | textbox | `getByRole('textbox',{name:'BBMD IP'})` | 192.168.2.100 | 必填，须为 IP 地址 |
| BBMD Port | textbox | `getByRole('textbox',{name:'BBMD Port'})` | 47809 | 必填，范围 47808–49000 |
| Time To Live | textbox | `getByRole('textbox',{name:'Time To Live'})` | 1440 | 必填，单位分钟 |
| 设备映射表 | table | group "Devices Selection To Mapping" | — | checkbox/Device Name/SN/Protocol/Online/Parameter Selection |
| Parameter Selection | button | 行末图标按钮 | — | 打开参数选择弹框（同 MQTT 穿梭框模式） |
| EPICS File Download | button | `getByRole('button',{name:'EPICS File Download'})` | — | 下载 EPICS 文件 |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 页面状态与分支 ★

| 状态 | 触发 | 结果 |
|------|------|------|
| Foreign Device Function = Enable | 默认 | 显示 BBMD IP / BBMD Port / Time To Live 三字段 |
| Foreign Device Function = Disable | 切换 | 隐藏 BBMD 组 |
| BACnet Disable | Enable组选Disable | 关闭 BACnet 服务 |

## 5. 校验规则要点

- BACnet Port / BBMD Port：47808–49000。
- Network Number：1–65534。
- Device Instance：0–4194302 且唯一。
- Device Object Name：≤40 字符。
- BBMD IP：合法 IP 格式。

## 6. 自动化测试要点

- 多字段范围/唯一性/格式校验（重点用例源）。
- Foreign Device Function 显隐 BBMD 组是核心分支。
- 设备映射表 + Parameter Selection 弹框（穿梭框，参见 MQTT TopicParameterSelection 文档）。
- EPICS File Download 触发文件下载。

## 7. 机器可解析摘要

```json
{
  "route": "/protocols/bacnet",
  "name": "bacnet",
  "title": "BACnet/IP",
  "module": "Protocols",
  "fields": {
    "BACnet Enable": {"type":"radio","default":"Enable"},
    "BACnet Port": {"type":"text","default":47808,"range":[47808,49000]},
    "Network Number": {"type":"text","range":[1,65534]},
    "Device Object Name": {"type":"text","maxlen":40},
    "Device Instance": {"type":"text","range":[0,4194302],"unique":true},
    "Advertised APDU Timeout": {"type":"select","default":"60 seconds"},
    "Advertised APDU Retries": {"type":"select","default":10},
    "Foreign Device Function": {"type":"radio","default":"Enable","controls":["BBMD IP","BBMD Port","Time To Live"]},
    "BBMD IP": {"type":"text","format":"ip"},
    "BBMD Port": {"type":"text","range":[47808,49000]},
    "Time To Live": {"type":"text","unit":"minutes","default":1440}
  },
  "device_mapping_table": true,
  "buttons": ["EPICS File Download","Save","Parameter Selection(per row)"]
}
```
