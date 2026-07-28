# AcuHMI-1-7 · Data Log 模块 UI 选择器参考

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/datalog/`
> 用途：Post Channel / Data Logger 类 UI 用例的**选择器与交互事实沉淀**，避免每次现场
> Playwright 探查。本文件是**纯参考文档**——用例仍各自自包含（复制所需片段），不 import 本文件。
> 数据来源：真机 192.168.3.71 实测（Element Plus v2）。

## 1. 进入 Data Log

```python
# 若不在 dataLog 页：先确保在 Devices 视图，再点左侧 Data Log
if "/#/dataLog" not in page.url:
    if not any(s in page.url for s in [
        "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
        "/#/webDevice", "/#/alarm", "/#/dataLog",
    ]):
        page.locator("header span").filter(has_text="Devices").first.click()
        page.wait_for_load_state("networkidle")
    page.locator(".left-nav-item").filter(has_text="Data Log").click()
    page.wait_for_load_state("networkidle")
```

## 2. 顶部菜单结构（横向 el-menu，父项需先展开再点子项）

| 父 tab（`div.el-sub-menu__title`） | 子项（`.el-menu-item`） |
|---|---|
| `Post Channels` | `Post Channel 1` / `Post Channel 2` / `Post Channel 3` |
| `Data Loggers` | `Data Loggers 1` / `Data Loggers 2` / `Data Loggers 3` / `Rapid Logger` / `Data Log Parameter Config` |
| `Data Log Management` | （无子项） |
| `Post Historical Data` | （无子项） |
| `AcuCloud` | （无子项） |

```python
# 通用：展开父 tab 再点子项
tab = page.locator("div.el-sub-menu__title").filter(has_text=tab_text)
if tab.count() > 0 and tab.first.is_visible():
    tab.first.click()
    page.wait_for_timeout(400)
item = page.locator(".el-menu-item").filter(has_text=item_text)
if item.count() > 0 and item.first.is_visible():
    item.first.click()
    page.wait_for_load_state("networkidle")
```

各 Data Logger 亦可直达 URL（避免 hover popup）：
`/#/dataLog/dataLogger/{dataLogger1 | dataLogger2 | dataLogger3 | rapidLogger}`

## 3. Enable / Disable 单选

- 均为 `el-radio`，按 label 文案过滤：`page.locator(".el-radio").filter(has_text="Enable"/"Disable").first`
- 选中态判定：元素 class 含 `is-checked`
- **坑**：`el-menu` 弹出的 popper 偶发遮挡事件链，`locator.click()` 会超时。降级方案（conventions #9 允许）——移开鼠标后按坐标点击：

```python
radio = page.locator(".el-radio").filter(has_text=label).first
radio.scroll_into_view_if_needed()
try:
    radio.click(timeout=3000)
except Exception:
    page.mouse.move(400, 300)          # 移开 hover，避免 popper 遮挡
    page.wait_for_timeout(200)
    box = radio.bounding_box()
    if box is not None:
        page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
```

各 Logger 的 Enable 表单 **label 文案不同**：

| Logger 子项 | Enable 表单 label |
|---|---|
| Data Loggers 1 | `Data Logger 1 Enable` |
| Data Loggers 2 | `Data Logger 2 Enable` |
| Data Loggers 3 | `Data Logger 3 Enable` |
| Rapid Logger | `Data Logger Rapid Enable` |

## 4. 保存与成功判定

```python
page.get_by_role("button", name="Save").click()
page.wait_for_timeout(1500)
# 成功判定：无 error 提示
assert page.locator(".el-message--error").count() == 0
```

- 无改动时保存也会弹提示（不因“未改动”而不弹）。
- 保存 Post Channel 可能触发服务重启，读取前留足等待。

## 5. Post Channel 下拉（Data Logger 配置页）

- **仅当该 Logger 处于 Enable 时才出现**（Disable 时 `.el-select` count = 0）。
- 表单项 label 文案 `Post Channel`，控件为 `el-select`，placeholder `--Select Post Channel--`。
- Logger Enable 后页面有 4 个 el-select，顺序：Post Channel / Log File Format / Log File Length / Log Interval。

```python
pc_select = page.locator(".el-form-item").filter(has_text="Post Channel").first.locator(".el-select")
pc_select.first.click()          # 展开下拉
page.wait_for_timeout(400)
```

### 下拉选项与“不可选”判定（关键）

- 选项元素：`.el-select-dropdown__item`，**文案为 `Post Channel N`（带空格！）**。
  - ⚠️ 历史坑：曾用 `"Channel1"`（无空格）匹配 → 永远匹配不到 → 用例空转假通过。必须用 `Post Channel 1`。
- 某 Post Channel 被 **Disable** 时：其选项**仍出现在下拉中**，但带 `is-disabled` 类
  （`class="el-select-dropdown__item is-disabled"`，`aria-disabled="true"`）；点击无效 = **不可选**。
  **不会从列表中移除。**
- 判定“不可选” = 选项 class 含 `disabled`。写断言时应**先要求选项被找到（presence）**，再判 disabled，
  防止“下拉没展开/选项没采到”导致的空转假通过。

```python
# 采集当前展开下拉里各 Post Channel N 选项的 class
classes = {}
for item in page.locator(".el-select-dropdown__item").all():
    if item.is_visible():
        txt = item.inner_text().strip()
        if txt.startswith("Post Channel"):
            classes[txt] = item.get_attribute("class") or ""
page.keyboard.press("Escape")     # 关闭下拉

# 断言：必须找到 + 带 disabled（不可选）
cls = classes.get("Post Channel 3")
assert cls is not None, "未找到 'Post Channel 3' 选项，无法确认其不可选"
assert "disabled" in cls, f"'Post Channel 3' 应不可选，但 class='{cls}'"
```

## 6. 相关约定

- 假通过陷阱与规避（try/except 吞断言、文案不匹配空转、presence 断言防空转）见团队约定
  conventions #9（Playwright 交互）/ #10（一函数一用例）。
- 已按本文件事实修正的用例：`test_TestCase_AcuHMI_003_05_case01 / _case09 / _case15`
  （postchannel 目录，PC Disable → Logger 下拉不可选）。
