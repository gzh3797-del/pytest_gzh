# Devices / Data Log / Data Log Management — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/dataLog/dataLogManagement` |
| 路由名 | `dataLogManagement` |
| 面包屑 | Devices / Data Log / Data Log Management |
| 上下文 | Devices 侧 |

## 2. 页面用途

下载或删除已记录的历史日志。分 Download Log / Delete Log 两区。

## 3. 交互元素清单

### 3.1 Download Log
| 字段 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Device | combobox | `getByRole('combobox',{name:'Device'})` | 必选(*)，选后启用后续字段 |
| Checkpoint | combobox | `getByRole('combobox',{name:'Checkpoint'})` | 必选(*)，选设备后启用 |
| Time Frame (Start/End Date) | date combobox | `getByRole('combobox',{name:'Start Date'})` / `{name:'End Date'}` | 必选(*)，选设备后启用 |
| Log Interval | combobox | `getByRole('combobox',{name:'Log Interval'})` | 必选(*)，选设备后启用 |
| Download | button | `getByRole('button',{name:'Download'})` | 初始 disabled，条件满足后可用 |

### 3.2 Delete Log
| 字段 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Device | combobox | `getByRole('combobox',{name:'Device'})`（Delete 区） | 必选(*) |
| Delete | button | `getByRole('button',{name:'Delete'})` | 初始 disabled；选设备后可用（危险，二次确认） |

## 4. 页面状态与分支 ★

| 状态 | 说明 |
|------|------|
| 未选 Device | Checkpoint/Time Frame/Log Interval/Download 均 disabled |
| 选 Device 后 | 级联启用后续字段与 Download |

## 5. 自动化测试要点

- 级联启用：Device 选中后其余字段与 Download 由 disabled→enabled（核心断言）。
- Download 触发下载；**Delete 破坏性**，验证二次确认。

## 6. 机器可解析摘要

```json
{
  "route": "/dataLog/dataLogManagement",
  "name": "dataLogManagement",
  "title": "Data Log Management",
  "context_side": "devices",
  "sections": {
    "Download Log": {"fields":["Device","Checkpoint","Time Frame(Start/End Date)","Log Interval"],"button":"Download","cascade":"enabled after Device"},
    "Delete Log": {"fields":["Device"],"button":"Delete(destructive)"}
  }
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/datalog/`。

### 进入路径
- 顶层子项 `.el-menu-item`（`Data Log Management`，无父展开），或直达 `#/dataLog/dataLogManagement`。

### pytest 选择器与控件
- Device：`el-select`（`.el-select__input`）；选中前 Checkpoint/Time Frame/Log Interval/Download 均 `disabled`，选中后级联启用（核心断言）。

### 高危
- ⚠️ **Delete Log 破坏性**：默认不实际执行，仅验证二次确认。Download 触发文件下载。
