# Templates / Template List — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/templates/templateList` |
| 路由名 | `templateList` |
| 面包屑 | AcuHMI-1-7 / Templates / Template List |
| 顶级模块 | Templates |

## 2. 页面用途

管理设备模板：内置官方模板（只读）与自定义模板（可操作）。Templates 二级 tab：Template List / New Typical Energy Meter Template / Import。

## 3. 页面结构

1. 二级 tab 栏（menubar）：Template List / New Typical Energy Meter Template / Import
2. **Official** 表（只读）：Template Name / Version / Firmware Version / Last Update；分页（10/page）
3. **Customized** 表：Template Name / Version / Last Update / **Action**（每行 4 个图标按钮）；分页（多页，含页码 1..11 与 "Next 5 pages"）+ 每页条数下拉

## 4. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 定位策略 2 | 说明 |
|------|------|-----------|-----------|------|
| 二级 tab | menuitem | `getByRole('menuitem',{name:'New Typical Energy Meter Template'})` | menubar 内 | 切换子页 |
| 列头排序 | columnheader(可点) | `getByRole('columnheader',{name:'Template Name'})` | 表头 (cursor=pointer) | 排序 |
| Customized 行 Action | button×4 | `getByRole('row',{name:/TestTpl_928635/}).getByRole('button')` | 行末 4 图标 | **顺序固定：View(绿 `.el-button--success`) / 派生 Create-from(蓝 `.el-button--primary`) / Edit(黄 `.el-button--warning`) / Delete(红 `.el-button--danger`)** |
| 每页条数 | combobox | `getByRole('combobox')`（分页区） | "10/page" | 改每页数量 |
| 上/下一页 | button | `getByRole('button',{name:'Go to next page'})` | 分页箭头 | 翻页 |
| 页码 | listitem | `getByRole('listitem',{name:'page 2'})` | 分页数字 | 跳页 |

## 5. 页面状态与分支

| 状态 | 说明 |
|------|------|
| Official 单页 | 前后翻页 disabled |
| Customized 多页 | 11 页，可翻页/跳页/改每页数 |
| 排序态 | 点列头切换升降序 |

## 6. 自动化测试要点

- Official 只读呈现；Customized 行操作（View/派生/Edit/Delete）为核心用例，删除为自定义 Yes/No 二次确认。
- 分页、每页条数、排序功能验证。
- 与 New Typical Template / Import 子页联动（新增后出现在 Customized 表）。

## 7. 机器可解析摘要

```json
{
  "route": "/templates/templateList",
  "name": "templateList",
  "title": "Template List",
  "module": "Templates",
  "sub_tabs": ["Template List","New Typical Energy Meter Template","Import"],
  "tables": {
    "Official": {"columns":["Template Name","Version","Firmware Version","Last Update"],"read_only":true},
    "Customized": {"columns":["Template Name","Version","Last Update","Action"],"row_actions":["View","Create-from(派生)","Edit","Delete"],"paginated":true}
  }
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/templates/`。

### 进入路径
- 先确保在某顶级视图（`page.locator("header span").filter(has_text="AcuHMI").first.click()`），再点左侧 `.left-nav-item` `Templates`；顶部子菜单 `.el-menu-item` 直接 `.filter(has_text=...)` 点击。

### 页面结构（★ 无 tab，两张表纵向排列）
- **无 `el-tabs`**（`.el-tabs__item`=0）。两张 `.c_common_table`：`nth(0)`=**Official**，`nth(1)`=**Customized**；段落标题 `Official`/`Customized`。
- **★ Customized 表在页面底部**（页高约 1649px/视口 720px，首行 y≈941px）：定位/点击前必须先滚动到底 `page.evaluate("window.scrollTo(0, document.body.scrollHeight)")`。
  ```python
  def _get_customized_table(page): return page.locator(".c_common_table").nth(1)
  ```
- 按名定位行：`_get_customized_table(page).locator("tbody tr").filter(has_text=name)`。

### Action 列（4 图标，顺序固定）
| 顺序 | class | 功能 |
|---|---|---|
| 1 | `.el-button--success`(绿) | View |
| 2 | `.el-button--primary`(蓝) | 派生（Create new template from this） |
| 3 | `.el-button--warning`(黄) | Edit |
| 4 | `.el-button--danger`(红) | Delete |

- 点“派生”跳 `/#/templates/typicalTemplateConfig/1/<uuid>`。

### 异步/结果反馈（实测）
- **模板创建异步**：`processTemplate` 完成后列表 GET 可能短暂返回旧数据 → **轮询 `page.reload()`** 直到目标行出现（最多 ~30s），不能一次读取就断言。
- 删除为**产品自定义 Yes/No 确认框**（非标准 MessageBox）：`page.get_by_role("button", name="Yes")` / `"No"`；删后断言 `filter(has_text=name).count()==0`。

### 参考用例
- `templates/general/test_TestCase_AcuHMI_008_01_case03_3.py`（新建→派生→列表验证→finally 清理）。
