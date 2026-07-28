# Devices / Data Log / Rapid Logger — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/dataLog/dataLogger/rapidLogger` |
| 路由名 | `rapidLogger` |
| 面包屑 | Devices / Data Log / Data Loggers / Rapid Logger |
| 上下文 | Devices 侧 |

## 2. 页面用途

高频（快速）数据记录器配置。**条件分支页**（默认 Disable，仅开关；Enable 后显示与 Data Logger 类似的记录配置：通道/格式/间隔/设备选择）。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 默认 | 说明 |
|------|------|-----------|------|------|
| Data Logger Rapid Enable | radiogroup | `getByRole('radiogroup',{name:'Data Logger Rapid Enable'})` | Disable | 必选(*)，控制显隐 |
| Enable/Disable radio | radio(label) | `page.locator('label').filter({hasText:/^Enable$/})` | — | Element-Plus，点 label |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

> Enable 后字段结构参考 [Devices_DataLog_DataLogger_context.md]（高频记录，间隔更短）。

## 4. 自动化测试要点

- Enable/Disable 显隐分支；启用后间隔/设备配置校验。

## 5. 机器可解析摘要

```json
{
  "route": "/dataLog/dataLogger/rapidLogger",
  "name": "rapidLogger",
  "title": "Rapid Logger",
  "context_side": "devices",
  "fields": {"Data Logger Rapid Enable": {"type":"radio","default":"Disable"}},
  "conditional": {"when":"Enable","shows":"logger config similar to dataLogger"},
  "buttons": ["Save"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/datalog/`。

### 进入路径
- 父 `div.el-sub-menu__title`（`Data Loggers`）→ 子 `.el-menu-item`（`Rapid Logger`），或**直达 URL** `#/dataLog/dataLogger/rapidLogger`（避免 hover popup）。

### pytest 选择器与控件
- Enable 表单 label 为 **`Data Logger Rapid Enable`**（与 Logger 1/2/3 不同）；`.el-radio` filter Enable/Disable，点 label 兜底。
- Enable 后字段结构同 [Devices_DataLog_DataLogger_context.md]（Post Channel/格式/间隔/设备表）。

### 保存与成功判定
```python
page.get_by_role("button", name="Save").click(); page.wait_for_timeout(1500)
assert page.locator(".el-message--error").count() == 0
```
