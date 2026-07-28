# AcuHMI-1-7 · Virtual Devices 模块 UI 选择器参考

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/virtualdevice/`
> 用途：Virtual Device 类 UI 用例（新增/删除/公式选参数/列表验证）的**选择器与交互事实沉淀**，
> 避免每次现场 Playwright 探查。本文件是**纯参考文档**——用例仍各自自包含（复制所需片段），不 import。
> 数据来源：真机 192.168.3.71 实测（Element Plus v2）。

## 1. 进入 Virtual Devices

```python
# 必须停在列表页本身（不是 addVirtualMeter 子路径）
on_list = "/#/virtualMeter" in page.url and "addVirtualMeter" not in page.url
if not on_list:
    if not any(s in page.url for s in [
        "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
        "/#/webDevice", "/#/alarm", "/#/dataLog",
    ]):
        page.locator("header span").filter(has_text="Devices").first.click()
        page.wait_for_load_state("networkidle")
    page.locator(".left-nav-item").filter(has_text="Virtual Devices").click()
    page.wait_for_load_state("networkidle")
```

## 2. 新增 / 删除 / 列表选择器速查

| 操作 | 选择器/方法 |
|------|------------|
| 新增按钮 | `page.get_by_role("button", name="Add Virtual Device")` |
| 设备名输入 | `page.get_by_label("Virtual Device Name", exact=True)` |
| Parameter Name | `page.get_by_placeholder("---Enter Parameter Name---")` |
| Post Label | `page.get_by_placeholder("---Enter Post Label---")` |
| Calculated Meter Formula 输入框 | `page.get_by_placeholder("---Enter Calculated Meter Formula---")` |
| **Select Device Parameter 按钮** | `page.locator("button.el-button--Copy").first`（文案 `Select Device Parameter`，多行参数时用 `.first` 限定当前行） |
| Unit | `page.get_by_placeholder("---Enter Unit---")` |
| 保存 | `page.get_by_role("button", name="Save")` |
| 追加参数行 | `page.get_by_role("button", name="Add Parameter")`（单设备最多 20 个参数） |
| 列表行定位 | `page.locator("tbody").get_by_role("row").filter(has_text=name)` |
| Action 列删除按钮 | `row.locator(".el-button").last`（Action 列最后一个按钮=删除，建议 `click(force=True)`） |
| 删除确认 | `page.get_by_role("button", name="Yes")`（兜底 `"Confirm"`） |

- 表单默认已带 **一行** 参数，直接 fill 即可，无需先点 Add Parameter。

## 3. ★ 分页（易漏坑）

- 列表默认**每页 10 行**，新建的 VD 追加到**末页**，`tbody tr` 只查当前页。
- **验证某 VD 是否存在、删除某 VD，都必须跨页查找**，否则第 11 个及以后的设备会漏判/漏删。
- 分页控件：下一页 `page.locator(".el-pagination .btn-next")`，禁用判定
  `next_btn.get_attribute("aria-disabled") == "true"`；页码 li `page.locator(".el-pagination .el-pager li")`。

```python
def _count_vd_across_pages(page, name):
    _nav_to_virtual_devices(page)
    total = 0
    while True:
        total += page.locator("tbody").get_by_role("row").filter(has_text=name).count()
        next_btn = page.locator(".el-pagination .btn-next")
        if next_btn.count() == 0 or next_btn.get_attribute("aria-disabled") == "true":
            break
        next_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    return total
```

## 4. ★ Calculated Meter Formula 选设备参数

公式框右侧有 `Select Device Parameter` 按钮（`button.el-button--Copy`）。点击弹出选择器弹窗，
从"已添加物理设备"中选设备 + 参数，Confirm 后把引用插入公式框。

- 弹窗定位：`page.locator('[aria-label="Select Device Parameter"]')`
- 弹窗内两个 El Plus v2 Select（`.el-select__input`）：`.first`=Device，`.last`=Parameter
- 两个 Select 的 `aria-controls` 均为 None，选项须用全局 `.el-select-dropdown__item` + 可见性过滤定位

```python
dialog = page.locator('[aria-label="Select Device Parameter"]')
sel = dialog.locator(".el-select__input")
sel.first.click()                                              # Device 下拉
page.locator(".el-select-dropdown__item").filter(has_text="pxm350").first.click()
sel.last.click()                                               # Parameter 下拉
visible = [i for i in page.locator(".el-select-dropdown__item").all() if i.is_visible()]
visible[0].click()
dialog.get_by_role("button", name="Confirm").click()
```

- **插入后公式格式**：`$<设备serial>:<参数名>`，例：`$pxm350:System Active Power Demand`
- **两参数相加**：选完第 1 个 → 在公式框末尾 `press("End")` + `type("+")` → 再选第 2 个
- **设备下拉文案**格式为 `<设备名>-<设备名>`；当前 HMI 已添加：
  `pxm350`、`Acurev1234100`、`AcuRev2100`、`AcuvimIIW`、`Acurev4100242`、`Acurev41002179`、`Acucim3`
- pxm350 参数示例：`System Active Power Demand`、`System Reactive Power Demand`、
  `System Apparent Power Demand`、`Phase A/B/C Export Active Energy`、`System Export Active Energy` …

## 5. 参考用例

- `projects/AcuHMI_1_7/tests/ui/virtualdevice/general/test_TestCase_AcuHMI_VD_003_003.py`
  ——新增两个 VD（公式选真实设备参数相加）→ 删除甲 → 跨页断言甲删除/乙保留 → finally 清理甲乙。
- 同目录 `..._VD_003_001.py`（基础字段保存）、`..._VD_003_005.py`（20 参数上限）。
