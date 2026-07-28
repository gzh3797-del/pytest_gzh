# Devices / Data Log / Post Historical Data — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/dataLog/postHistoricalData` |
| 路由名 | `postHistoricalData` |
| 面包屑 | Devices / Data Log / Post Historical Data |
| 上下文 | Devices 侧 |

## 2. 页面用途

将历史日志按指定通道/格式重新补发(Post)。

## 3. 交互元素清单 / 表单字段

| 字段 | 类型 | 定位策略 1 | 默认/示例 | 校验/说明 |
|------|------|-----------|-----------|-----------|
| Post Channel | combobox | `getByRole('combobox',{name:'Post Channel'})` | --Select Post Channel-- | 必选(*) |
| Device | combobox | `getByRole('combobox',{name:'Device'})` | --Select Device-- | 必选(*) |
| Timestamp Format | radiogroup | `getByRole('radiogroup',{name:'Timestamp Format'})` | Local Time String | Local Time String / UTC Seconds / ISO8601 |
| Log File Name Format | radiogroup | `getByRole('radiogroup',{name:'Log File Name Format'})` | UTC Timestamp | UTC Timestamp / Time interval Format |
| Log File Format | combobox | `getByRole('combobox',{name:'Log File Format'})` | csv | 必选(*) |
| Log File Name Prefix | textbox | `getByRole('textbox',{name:'Log File Name Prefix'})` | — | ≤20 字符 |
| Log File Length | combobox | `getByRole('combobox',{name:'Log File Length'})` | 1 minute | 必选(*)，选设备后启用 |
| Log Interval | combobox | `getByRole('combobox',{name:'Log Interval'})` | 1 minute | 必选(*)；AcuMesh 时 ≥5 分钟 |
| Post | button | `getByRole('button',{name:'Post'})` | — | 执行补发 |

## 4. 页面状态与分支

| 状态 | 说明 |
|------|------|
| 未选 Device | Log File Length / Log Interval 等 disabled |
| 选 Device 后 | 级联启用 |

## 5. 自动化测试要点

- Post Channel + Device 必选；格式单选覆盖；Prefix ≤20。
- 级联启用（选 Device 后 Length/Interval 可用）；Post 触发补发结果提示。
- 与 Post Channel 页联动（依赖已配置通道）。

## 6. 机器可解析摘要

```json
{
  "route": "/dataLog/postHistoricalData",
  "name": "postHistoricalData",
  "title": "Post Historical Data",
  "context_side": "devices",
  "fields": {
    "Post Channel": {"type":"select","required":true},
    "Device": {"type":"select","required":true},
    "Timestamp Format": {"type":"radio","options":["Local Time String","UTC Seconds","ISO8601"]},
    "Log File Name Format": {"type":"radio","options":["UTC Timestamp","Time interval Format"]},
    "Log File Format": {"type":"select","default":"csv"},
    "Log File Name Prefix": {"type":"text","maxlen":20},
    "Log File Length": {"type":"select","default":"1 minute"},
    "Log Interval": {"type":"select","default":"1 minute","note":"AcuMesh>=5min"}
  },
  "buttons": ["Post"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/datalog/`。

### 进入路径
- 顶层子项 `.el-menu-item`（`Post Historical Data`，无父展开），或直达 `#/dataLog/postHistoricalData`。

### pytest 选择器与控件
- Post Channel / Device：`el-select`（`.el-select__input`）；Timestamp Format / Log File Name Format 为 `el-radio`（点 label 兜底）。
- 级联：选 Device 后 Log File Length / Log Interval 由 disabled→enabled。

### 结果反馈
- 点 Post 后有补发结果提示 toast（`.el-message`）；依赖 Post Channel 页已配置通道。
