# AcuHMI-1-7 · Templates 模块 UI 选择器参考

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/templates/`
> 用途：Template 类 UI 用例（新建/派生/删除自定义模板）的**选择器与交互事实沉淀**，避免每次
> 现场 Playwright 探查。本文件是**纯参考文档**——用例仍各自自包含（复制所需片段），不 import 本文件。
> 数据来源：真机 192.168.3.71 实测（Element Plus v2）。

## 1. 进入 Templates

```python
# 若不在 templates 页：先确保在某一顶级视图，再点左侧 Templates
if "/#/templates" not in page.url:
    if not any(s in page.url for s in [
        "/#/systemSettings", "/#/templates", "/#/protocols",
        "/#/maintenance", "/#/diagnostics", "/#/userManagement", "/#/firmwareUpdate",
    ]):
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
    page.locator(".left-nav-item").filter(has_text="Templates").click()
    page.wait_for_load_state("networkidle")
```

顶部子菜单为 `.el-menu-item`，含 `New Typical Energy Meter Template`、`Template List` 等，直接 `.filter(has_text=...)` 点击。

## 2. Template List 页面结构（★ 无 tab，两张表纵向排列）

- **没有 el-tabs 组件**（`.el-tabs__item` 数量为 0）。页面有两张 `.c_common_table`：
  - `.c_common_table` `nth(0)` = **Official/Standard** 表
  - `.c_common_table` `nth(1)` = **Customized** 表（用户自定义模板）
- 段落标题 `h2/h3` 文案分别为 `Official`、`Customized`。
- **★ Customized 表在页面最下方**：页面总高约 1649px（视口 720px），Customized 首行 y≈941px，
  超出首屏。定位/断言/点击 Customized 行前必须先滚动到底：

```python
def _get_customized_table(page):
    return page.locator(".c_common_table").nth(1)

# 滚到底部确保 Customized 行进入视口（或对目标元素用 scroll_into_view_if_needed()）
page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
```

- 按名定位某行：`_get_customized_table(page).locator("tbody tr").filter(has_text=template_name)`
- **模板创建是异步的**：`processTemplate` API 完成后，列表 GET 接口可能短暂返回旧数据。
  验证新模板出现需 **轮询 `page.reload()`** 直到目标行出现（实测最多等 ~30s），不能一次读取就断言。

## 3. Customized 行 Action 列（4 个图标，顺序固定）

| 顺序 | CSS class | 颜色 | 功能 |
|------|-----------|------|------|
| 第 1 个 | `.el-button--success` | 绿 | View（查看） |
| 第 2 个 | `.el-button--primary` | 蓝 | **Create new template from this one（派生）** |
| 第 3 个 | `.el-button--warning` | 黄 | Edit（编辑） |
| 第 4 个 | `.el-button--danger`  | 红 | Delete（删除） |

```python
row = _get_customized_table(page).locator("tbody tr").filter(has_text=name)
row.first.locator(".el-button--primary").first.click()   # 派生
row.first.locator(".el-button--danger").first.click()    # 删除
```

点"派生"后跳转 URL：`/#/templates/typicalTemplateConfig/1/<uuid>`。派生表单继承源模板的 Block
数据（无需重选 Wiring/Function），Template Name 预填为空且 `is_editable()==True`，只需重填
Template Name / Version 再点 `Create Template`。

## 4. New / 派生模板表单字段

`.el-form-item` 按 `has_text` 过滤后取 `input` / `.el-select`：
Template Name、Version、Typical Model、Wiring Configuration、Function、Address Format、Start、Count。

- Block 区：填 Function/Start/Count 后点 `get_by_role("button", name="Save Block")`，
  再点 `get_by_role("button", name="Create Template")` 提交。

### Wiring Configuration 下拉（8 项，★ 真实文案，注意大小写与无复数 s）

`3 Element 4 Wire Y`、`1 Element 2 Wire`、`2 Element 3 Wire 1 Phase`、`2 Element 3 Wire Network`、
`2 Element 3 Wire Delta`、`3 Element 3 Wire Delta`、`3 Element 4 Wire Delta`、`2 1/2 Element 4 Wire Y`

> 用例/手工步骤里写的 “3 elements 4 Wire Y”（小写复数）**不存在**，正确文案是 `3 Element 4 Wire Y`。

### Function 下拉（4 项）

`READ_HOLDING_REGISTERS`、`READ_COILS`、`READ_DISCRETE_INPUTS`、`READ_INPUT_REGISTERS`

## 5. 弹框与提示

- **创建成功 toast**：`.el-message--success`，文案 `Create Success.`（含句号）。
  基础模板与派生模板创建成功均为同款 toast。断言建议 `expect(loc.first).to_be_visible()` +
  校验 `inner_text()` 含 `Create Success`，避免旧 toast 残留导致假通过。
- **删除确认弹框**：产品自定义 Yes/No（非标准 MessageBox），按钮用
  `get_by_role("button", name="Yes")` / `"No"`。删除后断言
  `_get_customized_table(page).locator("tbody tr").filter(has_text=name).count() == 0`。

## 6. 参考用例

`projects/AcuHMI_1_7/tests/ui/templates/general/test_TestCase_AcuHMI_008_01_case03_3.py`
——新建基础模板 → 派生新模板 → 列表验证 → `finally` 清理（删派生 + 删基础）的完整实现。
