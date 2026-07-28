# Devices / Virtual Devices — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 列表路由 | `/#/virtualMeter`（`virtualMeter`） |
| 新增路由 | `/#/virtualMeter/addVirtualMeter`（`addVirtualMeter`） |
| 详情路由 | `/#/virtualMeter/virtualMeterDetails/:key`（`virtualMeterDetails`，动态） |
| 相关动态 | `/virtualDevices/virtualDetails/:deviceKey`、`.../virtualReadings` |
| 面包屑 | Devices / Virtual Devices |
| 上下文 | Devices 侧 |

## 2. 页面用途

创建"虚拟设备"：通过公式对物理设备参数进行计算，产出派生参数。

## 3. 列表页 (virtualMeter)

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Add Virtual Device | button | `getByRole('button',{name:'Add Virtual Device'})` | 进入新增 |
| 列表表 | table | 列: Virtual Device Name / Serial Number / Action | 分页（4 页） |
| 行 Action | button | 行末图标 | 进入详情/编辑/删除 |

## 4. 新增页 (addVirtualMeter)

| 字段 | 类型 | 定位策略 1 | 默认 | 校验 |
|------|------|-----------|------|------|
| Virtual Device Name | textbox | `getByRole('textbox',{name:'Virtual Device Name'})` | — | 必填(*)，≤40 字符 |
| Calculated Interval | textbox | `getByRole('textbox',{name:'Calculated Interval'})` | 5 | 必填(*)，seconds |
| **Parameter 区（可重复）** | group | 文本 "Parameter" | — | 至少 1 个 |
| ⤷ Parameter Name | textbox | `getByRole('textbox',{name:'Parameter Name'})` | — | 必填(*)，≤40 |
| ⤷ Post Label | textbox | `getByRole('textbox',{name:'Post Label'})` | — | 必填(*)，≤40 |
| ⤷ Calculated Meter Formula | textbox | `getByRole('textbox',{name:'Calculated Meter Formula'})` | — | 必填(*)，表达式 |
| ⤷ Select Device Parameter | button | `getByRole('button',{name:'Select Device Parameter'})` | — | 插入 `$serial:param` 引用 |
| ⤷ Unit | textbox | `getByRole('textbox',{name:'Unit'})` | — | 必填(*)，≤40 |
| ⤷ Delete | button | `getByRole('button',{name:'Delete'})` | — | 删除该参数 |
| Add Parameter | button | `getByRole('button',{name:'Add Parameter'})` | — | 追加参数块 |
| Save / Cancel | button | `getByRole('button',{name:'Save'})` / `{name:'Cancel'}` | — | 提交/取消 |

- 公式示例：`"$device1serial:param1name"+0.1*"$device2serial:param2name"+5.00`

## 5. 自动化测试要点

- 新增：设备名唯一/长度、间隔、至少一个参数（名称/标签/公式/单位必填）。
- 动态增删参数块（Add Parameter / Delete）；Select Device Parameter 弹选择器插入引用。
- 公式语法校验；详情页 (virtualMeterDetails/:key) 编辑预填。

## 6. 机器可解析摘要

```json
{
  "routes": {"list":"/virtualMeter","add":"/virtualMeter/addVirtualMeter","details":"/virtualMeter/virtualMeterDetails/:key"},
  "name": "virtualMeter",
  "title": "Virtual Devices",
  "context_side": "devices",
  "list_columns": ["Virtual Device Name","Serial Number","Action"],
  "add_fields": {
    "Virtual Device Name": {"type":"text","required":true,"maxlen":40},
    "Calculated Interval": {"type":"text","required":true,"unit":"s","default":5},
    "parameter(repeatable)": ["Parameter Name","Post Label","Calculated Meter Formula","Unit"]
  },
  "buttons": ["Add Virtual Device","Add Parameter","Select Device Parameter","Delete","Save","Cancel"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/virtualdevice/`。

### 进入路径
- **必须停在列表页本身**（`"/#/virtualMeter" in url and "addVirtualMeter" not in url`）；先确保 Devices 视图再点左侧 `.left-nav-item` `Virtual Devices`。

### pytest 选择器速查
| 操作 | 选择器 |
|---|---|
| 新增按钮 | `page.get_by_role("button", name="Add Virtual Device")` |
| 设备名 | `page.get_by_label("Virtual Device Name", exact=True)` |
| Parameter Name | `page.get_by_placeholder("---Enter Parameter Name---")` |
| Post Label | `page.get_by_placeholder("---Enter Post Label---")` |
| Formula 输入框 | `page.get_by_placeholder("---Enter Calculated Meter Formula---")` |
| **Select Device Parameter** | `page.locator("button.el-button--Copy").first`（多行参数用 `.first` 限当前行） |
| Unit | `page.get_by_placeholder("---Enter Unit---")` |
| 追加参数行 | `page.get_by_role("button", name="Add Parameter")`（**单设备最多 20 个参数**） |
| 列表行 | `page.locator("tbody").get_by_role("row").filter(has_text=name)` |
| 删除按钮 | `row.locator(".el-button").last`（Action 列最后一个，建议 `click(force=True)`） |
| 删除确认 | `page.get_by_role("button", name="Yes")`（兜底 `"Confirm"`） |

- 表单默认已带 **一行** 参数，直接 fill，无需先点 Add Parameter。

### ★ 分页（易漏坑）
- 列表默认**每页 10 行**，新建 VD 追加到**末页**，`tbody tr` 只查当前页 → **验证/删除某 VD 必须跨页查找**（否则第 11 个及以后漏判/漏删）。
- 下一页 `page.locator(".el-pagination .btn-next")`，禁用判定 `get_attribute("aria-disabled")=="true"`。

### ★ Calculated Meter Formula 选设备参数
- 点公式框右侧 `Select Device Parameter`（`button.el-button--Copy`）弹窗：`page.locator('[aria-label="Select Device Parameter"]')`。
- 弹窗内两个 EP v2 Select（`.el-select__input`）：`.first`=Device、`.last`=Parameter；两者 `aria-controls` 均 None → 选项用全局 `.el-select-dropdown__item` + 可见性过滤。
- **插入后公式格式**：`$<设备serial>:<参数名>`，例 `$pxm350:System Active Power Demand`。
- **两参数相加**：选完第 1 个 → 公式框末尾 `press("End")` + `type("+")` → 再选第 2 个。
- 设备下拉文案格式 `<设备名>-<设备名>`。

### 参考用例
- `virtualdevice/general/test_TestCase_AcuHMI_VD_003_003.py`（新增两 VD 公式相加→删甲→跨页断言→finally 清理）、`..._VD_003_001.py`（基础保存）、`..._VD_003_005.py`（20 参数上限）。
