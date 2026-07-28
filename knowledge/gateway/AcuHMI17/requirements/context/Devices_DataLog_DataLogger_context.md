# Devices / Data Log / Data Loggers (1/2/3) — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/dataLog/dataLogger/dataLogger1`（`dataLogger1`；同构 `dataLogger2`/`dataLogger3`） |
| 父路由 | `/#/dataLog/dataLogger`（`dataLogger`，Data Loggers 分组入口） |
| 面包屑 | Devices / Data Log / Data Loggers / Data Loggers 1 |
| 上下文 | Devices 侧 |

> **本文档同时覆盖 Data Logger 1/2/3**（三页结构一致，仅编号不同）。同组还有 Data Log Parameter Config (`dataLogParameterConfig`) 与 Rapid Logger (`rapidLogger`)，见各自文档。

## 2. 页面用途

配置某个数据记录器：记录哪些设备、日志文件格式/命名/时长/间隔、时间戳格式、上报通道。**条件分支页**（Disable→仅开关）。

## 3. 交互元素清单（Enable 后）

| 字段 | 类型 | 定位策略 1 | 默认/示例 | 校验/说明 |
|------|------|-----------|-----------|-----------|
| Data Logger N Enable | radiogroup | `getByRole('radiogroup',{name:'Data Logger 1 Enable'})` | Disable | 必选(*) |
| Post Channel | combobox | `getByRole('combobox',{name:'Post Channel'})` | Post Channel 1 | 关联上报通道 |
| Timestamp Format | radiogroup | `getByRole('radiogroup',{name:'Timestamp Format'})` | Local Time String | Local Time String / UTC Seconds / ISO8601 |
| Log File Name Format | radiogroup | `getByRole('radiogroup',{name:'Log File Name Format'})` | Time interval Format | UTC Timestamp / Time interval Format |
| Log File Format | combobox | `getByRole('combobox',{name:'Log File Format'})` | csv | 文件格式 |
| Log File Name Prefix | textbox | `getByRole('textbox',{name:'Log File Name Prefix'})` | meter2_Logger1 | 必填(*)，≤20 字符 |
| Log File Length | combobox | `getByRole('combobox',{name:'Log File Length'})` | 1 minute | 必选(*) |
| Log Interval | combobox | `getByRole('combobox',{name:'Log Interval'})` | 1 minute | 必选(*)；AcuMesh 设备时不得短于 5 分钟 |
| Devices Selection 表 | table | group "Devices Selection" | — | checkbox/Device Name/Device Type/Serial Number/Protocol/Online |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 页面状态与分支

| 状态 | 触发 | 结果 |
|------|------|------|
| Disable（默认） | 进入页面 | 仅 Enable 单选 + Save |
| Enable | 选 Enable | 显示上报通道/格式/命名/间隔/设备选择等 |

## 5. 自动化测试要点

- 条件显隐；Prefix ≤20 校验；Log Interval 与 AcuMesh 约束（≥5min）。
- Timestamp/File Name Format 单选覆盖；设备表勾选（含物理+虚拟设备）。
- 三个 Logger（1/2/3）行为一致，可参数化复用。

## 6. 机器可解析摘要

```json
{
  "route": "/dataLog/dataLogger/dataLogger1",
  "name": "dataLogger1",
  "title": "Data Loggers 1",
  "context_side": "devices",
  "covers": ["dataLogger1","dataLogger2","dataLogger3"],
  "fields": {
    "Data Logger Enable": {"type":"radio","default":"Disable"},
    "Post Channel": {"type":"select"},
    "Timestamp Format": {"type":"radio","options":["Local Time String","UTC Seconds","ISO8601"]},
    "Log File Name Format": {"type":"radio","options":["UTC Timestamp","Time interval Format"]},
    "Log File Format": {"type":"select","default":"csv"},
    "Log File Name Prefix": {"type":"text","required":true,"maxlen":20},
    "Log File Length": {"type":"select","default":"1 minute"},
    "Log Interval": {"type":"select","default":"1 minute","note":"AcuMesh>=5min"}
  },
  "device_table": ["checkbox","Device Name","Device Type","Serial Number","Protocol","Online"],
  "buttons": ["Save"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/datalog/`。

### 进入路径
- 顶部为横向 `el-menu`，父项需先展开：父 `div.el-sub-menu__title`（`Data Loggers`）→ 子 `.el-menu-item`（`Data Loggers 1/2/3`）。
- **直达 URL 避免 hover popup**：`/#/dataLog/dataLogger/{dataLogger1|dataLogger2|dataLogger3|rapidLogger}`。

### pytest 选择器与控件
- Enable/Disable：`el-radio`，`page.locator(".el-radio").filter(has_text="Enable"/"Disable").first`，选中态判 class 含 `is-checked`。
- **各 Logger 的 Enable 表单 label 文案不同**：

  | 子项 | Enable label |
  |---|---|
  | Data Loggers 1 | `Data Logger 1 Enable` |
  | Data Loggers 2 | `Data Logger 2 Enable` |
  | Data Loggers 3 | `Data Logger 3 Enable` |
  | Rapid Logger | `Data Logger Rapid Enable` |

- **Post Channel 下拉仅当该 Logger 处于 Enable 时出现**（Disable 时 `.el-select` count=0）。Logger Enable 后页面有 4 个 `el-select`，顺序：Post Channel / Log File Format / Log File Length / Log Interval。
  ```python
  pc_select = page.locator(".el-form-item").filter(has_text="Post Channel").first.locator(".el-select")
  pc_select.first.click()   # 展开
  ```
  - ⚠️ 选项文案为 **`Post Channel N`（带空格！）**——曾用 `Channel1`（无空格）致空转假通过。
  - 被 **Disable** 的 Post Channel：**仍出现在下拉中但带 `is-disabled`（`aria-disabled="true"`），不移除**；点击无效。断言不可选须**先 presence 再判 disabled**，防空转。

### 框架坑与兜底
- `el-menu` popper 偶发遮挡事件链致 `click()` timeout → 降级：移开鼠标（`page.mouse.move(400,300)`）后按 `bounding_box` 坐标点击。

### 保存与成功判定
```python
page.get_by_role("button", name="Save").click(); page.wait_for_timeout(1500)
assert page.locator(".el-message--error").count() == 0
```
- 无改动时保存也会弹提示。

### 参考用例
- `postchannel` 目录 `test_TestCase_AcuHMI_003_05_case01/_case09/_case15`（PC Disable → Logger 下拉不可选）。
