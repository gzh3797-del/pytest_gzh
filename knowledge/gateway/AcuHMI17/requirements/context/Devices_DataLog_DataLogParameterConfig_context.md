# Devices / Data Log / Data Log Parameter Config — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/dataLog/dataLogger/dataLogParameterConfig` |
| 路由名 | `dataLogParameterConfig` |
| 面包屑 | Devices / Data Log / Data Loggers / Data Log Parameter Config |
| 上下文 | Devices 侧 |

## 2. 页面用途

为数据记录器选择每个设备要记录的参数（穿梭框）。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 默认/示例 | 说明 |
|------|------|-----------|-----------|------|
| Device | combobox | `getByRole('combobox',{name:'Device'})` | pxm350-pxm350 | 必选(*)，选择设备 |
| Parameter Type | combobox | `getByRole('combobox',{name:'Parameter Type'})` | Realtime | 必选(*)，参数类型 |
| Parameters 穿梭框 | transfer(dual list) | group "Parameters" | — | Not Selected ↔ Selected，含 `>` `<` `All` `Clear` |
| All / Clear | button | `getByRole('button',{name:'All'})` / `{name:'Clear'}` | — | 全部移入/清空 |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

- Selected 侧示例：Phase A/B/C Line Current、各类电压/功率/功率因数/需量/电能等（上百项）。

## 4. 自动化测试要点

- Device/Parameter Type 切换刷新可选参数集合。
- 穿梭框 All/Clear、单项 `>`/`<` 移动断言 Selected/Not Selected 列表变化。

## 5. 机器可解析摘要

```json
{
  "route": "/dataLog/dataLogger/dataLogParameterConfig",
  "name": "dataLogParameterConfig",
  "title": "Data Log Parameter Config",
  "context_side": "devices",
  "fields": {
    "Device": {"type":"select","required":true},
    "Parameter Type": {"type":"select","required":true,"default":"Realtime"},
    "Parameters": {"type":"dual-list transfer","controls":["All","Clear",">","<"]}
  },
  "buttons": ["Save"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/datalog/`。

### 进入路径
- 父 `div.el-sub-menu__title`（`Data Loggers`）→ 子 `.el-menu-item`（`Data Log Parameter Config`），或直达 `#/dataLog/dataLogger/dataLogParameterConfig`。

### pytest 选择器与控件
- Device / Parameter Type：`el-select`（`.el-select__input`）；下拉选项 `.el-select-dropdown__item`，`aria-controls` 可能为 None 时用全局 + 可见性过滤。
- Parameters 穿梭框：`All` / `Clear` 按 `page.get_by_role("button", name="All"/"Clear")`；单项 `>`/`<` 移动。

### 保存与成功判定
```python
page.get_by_role("button", name="Save").click(); page.wait_for_timeout(1500)
assert page.locator(".el-message--error").count() == 0
```
