# Templates / Typical Energy Meter Template (Edit/View) — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/templates/typicalTemplateConfig/:type/:id`（动态） |
| 路由名 | `typicalTemplateConfig` |
| 面包屑 | AcuHMI-1-7 / Templates / Typical Energy Meter Template |
| 顶级模块 | Templates |
| 入口 | Template List → Customized 表某行 Action 图标（编辑/查看） |

## 2. 页面用途

编辑或查看已有的自定义典型电表模板。**页面结构与 [Templates_NewTypicalTemplate_context.md] 完全一致**（Device / Block / Save / Block Table / Parameter Table 三表），区别在于：
- 通过路由参数 `:type`（模板类型）与 `:id`（模板ID）定位具体模板；
- 进入时字段/表格已用该模板的现有数据**预填**；
- 提交按钮语义为更新（而非全新创建）。

## 3. 元素与校验

参见 `Templates_NewTypicalTemplate_context.md`（同 Device/Block 字段、Block Table、Configured/Unconfigured Parameter Table 结构、同校验规则）。

## 4. 自动化测试要点

- 需先从 Template List 的 Customized 表进入（依赖存在的模板 id）；直接构造 URL 需有效 `:type/:id`。
- 断言预填数据与所选模板一致；编辑后保存并回列表验证 Last Update 变化。
- 其余同 New 模板页。

## 5. 机器可解析摘要

```json
{
  "route": "/templates/typicalTemplateConfig/:type/:id",
  "name": "typicalTemplateConfig",
  "title": "Typical Energy Meter Template",
  "module": "Templates",
  "dynamic_params": ["type","id"],
  "same_structure_as": "newtypicalTemplateConfig",
  "entry": "Template List > Customized row action",
  "mode": "edit/view (prefilled)"
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/templates/`。

### 派生（Create-from）实测事实
- 入口：Template List → Customized 行第 2 个图标（蓝 `.el-button--primary`）。跳转 URL：`/#/templates/typicalTemplateConfig/1/<uuid>`。
- **派生表单继承源模板的 Block 数据**（无需重选 Wiring/Function）；**Template Name 预填为空且 `is_editable()==True`**，只需重填 Template Name / Version 再点 `page.get_by_role("button", name="Create Template")`。
- 成功 toast 同 New：`.el-message--success` 文案 `Create Success.`。
- 字段选择器/Wiring 8 值/Function 4 值同 [Templates_NewTypicalTemplate_context.md]。

### 参考用例
- `templates/general/test_TestCase_AcuHMI_008_01_case03_3.py`（新建基础→派生→列表验证→finally 删派生+删基础）。
