# 页面探索坑清单（SPA 导航 / 框架识别 / 表单校验 / 跨页传播）

> **来源**：本清单从**测试二组** `autotest` 阶段二的实战复盘吸收（跨组沉淀），均为带真实故障场景的工程坑。探索时**必须实测采集**，结论写入对应子页 `_context.md`「实测测试情报」节（见 [OUTPUT_STANDARD.md](OUTPUT_STANDARD.md) §四）。
>
> **本 skill 用 Playwright MCP 工具（`browser_*`）探索**，下方检测片段已按 MCP 惯用法给出；`browser_evaluate` 里跑 JS，导航用 `browser_navigate`，读落点用 `browser_snapshot` 后核对 URL。
>
> **为什么这些坑必须在探索期采集**：`_context.md` 是下游 `webtestcase_manual_to_auto` 生成脚本的唯一依据。这五类坑若探索期漏采，生成的脚本会**静默跑错**（落点错、状态没重置、框架选择器错、断言对着旧数据/未触发的校验），且报错信息完全看不出根因。

---

## §1 SPA 路由守卫（`goto` 直达是否被拦截）

**症状**：Vue/React hash SPA 里，通过菜单点击能到达的页面，直接 `goto(目标URL)` 可能被路由守卫重定向回登录页或某个默认页——下游脚本若用 `goto` 直达就会落错页面。

**检测方法（子页面跟随 / 记录路由后必做）**：

```
1. 先通过菜单点击成功到达目标页 → browser_snapshot → 记录当前 URL（如 /#/userManagement/general）
2. browser_navigate 到另一个区域（菜单里的其他顶级项）
3. browser_navigate 直接跳目标 URL → browser_snapshot → 核对落点 URL
   · 落点 == 目标 URL  → 无守卫，goto 直达可用
   · 落点 != 目标 URL  → 存在守卫，直达不可用，须走"菜单点击链"
```

> **禁止**：发现守卫后用连续多次 `goto` 重试——多次 goto 会累积守卫状态，越试越锁死。正确做法是**一次 goto 检测，失败即改用菜单点击链**。

**记录到 `_context.md`**（「实测测试情报」§进入路径 / 页面标识）：
```
goto 可达性：直达可用 / 存在守卫（须菜单点击）
守卫解法：sidebar listitem("顶级菜单") → wait → menuitem("子页") 点击链
  侧边栏展开定位：getByRole('listitem').filter({hasText:'XXX'})
  子菜单定位：getByRole('menuitem',{name:'YYY'})
```

---

## §2 同路由 `goto` 空操作陷阱（列表/分页/筛选页必做）

**症状**：hash SPA 里，若当前已在某路由，再 `goto(同一路由)` 是**空操作**——组件不重新挂载，分页/筛选/排序/展开 Tab 等页面内部状态**原样保留**。一个典型故障：helper 翻到列表最后一页读总数，设计上指望"下一步 `nav_to_list()` 会重置回第一页"，但 `nav_to_list()` 内部只是 `goto(LIST_URL)`——因已在该路由，goto 什么也没做，分页停在最后一页，后续用例读到"最后一页行数"而非"全部行数"，断言失败且报错看不出根因。

**检测方法（涉及"改状态后期望回干净默认态"的页面必做）**：
```
1. browser_navigate 进入目标列表页
2. 翻到第 2 页 / 应用某个筛选 → browser_snapshot
3. 再次 browser_navigate 同一路由（模拟 nav 函数"回默认态"）→ browser_snapshot
4. 核对：分页/筛选控件是回到默认值，还是仍停在刚才改动后的状态？
```

**记录到 `_context.md`**（「实测测试情报」§页面状态与分支）：
```
同路由 goto 状态重置：
  - 回列表是否依赖"重新挂载重置状态"：是/否
  - 同路由 goto 后状态是否真被重置：是/否
  - 若否，解法：① 页面有"重置"控件 → 显式点击（上一页点到禁用/清空筛选，附定位）
             ② 无重置控件 → 二次 goto 强制重挂载（goto 别的路由 → goto 回目标路由）
```

> **禁止**：在 `_context.md` 里写"调用 nav_to_X() 就等于回到初始态"——这个假设只在"上一步离开了该路由"时成立，同路由内连续调用会静默失效。

---

## §3 框架识别必须靠 JS DOM 计数（Ant Design vs Element Plus）

**症状**：**Ant Design Select 与 Element Plus Select 在 a11y snapshot 里都呈现为 `combobox`，文字描述相同**——只看 `browser_snapshot` 无法区分框架类型。若据快照类名猜错框架，生成的选择器全错（这是最高频错误）。

**检测方法（含下拉/选择/过滤控件的页面，识别前强制先跑）**：

```
browser_evaluate --function "() => ({
  elSelect:  document.querySelectorAll('.el-select__wrapper, .el-select').length,
  antSelect: document.querySelectorAll('.ant-select').length,
  antMulti:  document.querySelectorAll('.ant-select-multiple').length
})"
```

判定规则：

| 结果 | 结论 | 真实选择器 |
|------|------|-----------|
| `elSelect > 0` | Element Plus Select | 触发 `.el-select__wrapper`；选项 `.el-select-dropdown__item`（teleport 到 body，见下） |
| `elSelect == 0` 且 `antSelect > 0` | **Ant Design Select** | 单选 `.ant-select`；多选搜索输入 `.ant-select-selection-search-input`；选项 `.ant-select-item-option-content` |
| 两者都有（混用页） | 逐控件单独确认 | 对每个控件分别套对应模板 |

> **禁止**：JS 计数显示 `elSelect == 0`，仍在 `_context.md` 写 `.el-select__wrapper`——这是本项检查要消灭的头号错误。

### 框架通用坑（据识别结果择一/两者都写）

**Element Plus：**
- `el-radio` 原生 `<input>` 被 `.el-radio__inner` 遮挡 → 点 input 超时，**须点 label 文案**；坐标点击兜底。
- `el-select` 选项在独立 popper（teleport 到 `<body>` 最外层，**不在 `.el-dialog` 内部**）；弹窗内的 select，触发在 dialog 内点、选项在 body 层找。用 `input.getAttribute('aria-controls')` 取 listId 再查该 popper 的 `.el-select-dropdown__item`；`aria-controls` 为 None 时用全局 `.el-select-dropdown__item` + 可见性过滤，**禁止全局 querySelectorAll 乱抓**。
- 同名 group/字段（证书 Issuer/Subject、双网口 Interface Status/IP）**必须父容器 scope**。
- 已选值读取：`.el-select__selected-item` 有两个节点（显示 `<span>` + 空 input-div），**须跳过 inner_text 为空的节点**，兜底读 `.el-select__placeholder`；`.first` 可能命中空节点返回空串。
- 必填 `*` 号由 CSS `::before` 生成，**不在 `innerText` 中**——`get_by_label` 参数用不含 `*` 的 `innerText` 版本，别用 a11y name（含 `*`）。
- 下拉选项文案可能带空格（`Post Channel 1`）；禁用项**仍在下拉但带 `is-disabled`**（不移除）。
- 自定义确认框可能是产品级 **Yes/No**（非标准 MessageBox）。
- `el-date-editor` 只读态：`props.readonly` 在 Vue 层拦截输入，JS 移除 DOM `readonly` 无效——只能做读取断言（`.el-date-editor input` 的 value），不能经 UI 改值。

**Ant Design：**
- Table 行选择器**必须用 `tr.ant-table-row`**，禁止通用 `tr`——`ant-table-measure-row` / `ant-table-placeholder` 会污染通用选择器，导致加载态误判行是否存在。
- Select 已选值读 `.ant-select-selection-item` 的 inner_text；多选搜索用 `.ant-select-selection-search-input`。

---

## §4 表单校验：多轮干跑 + 依赖字段多状态

**症状**：只看静态 DOM 快照会漏掉隐性提交阻断；且**依赖字段**在依赖未触发时组件库常**跳过校验**——只探"表单全空提交"会把"从未触发的校验错误"误记进 `_context.md`。

**多轮提交干跑（含提交按钮的页面必做，属阶段3 提交按钮项的加强）**：

```
第一轮：空表单直接提交 → 记录所有错误
  · 字段级：.el-form-item__error
  · 全局/区块级：.el-message--error / .tip-error-info / [class*=alert]
    （重点抓区块级错误，如 "Please specify at least one facility."——来自整块区域，查字段 label 发现不了）
第二轮：按第一轮清单填最小有效数据再提交 → 记录成功 toast 类名+文案、成功后路由跳转、是否需 cleanup
第三轮：必填字段 label.innerText 核实（见 §3 Element Plus 的 * 号坑）
```

**多入口覆盖**：同一表单页存在"新建（空表单）"和"编辑（已有值）"两种入口时，**必须对空表单入口干跑**——已有值入口会遮蔽"其他必填为空"的校验错误（典型：某编辑页未配置入口有 8 个必填字段，已配置入口发现不了）。

**依赖字段多状态（含级联/条件禁用字段的表单必做）**：
若字段 B 的可编辑性/取值/校验依赖字段 A 的状态，**必须探两种状态**，不得只探一种就假设另一种一致：
```
状态一（依赖未触发）：A 保持默认/未选 → 看 B 是否 disabled、干跑提交时 B 的必填错误是否出现
状态二（依赖已触发）：选中 A 的一个真实值 → 看 B 是否启用、是否被自动回填、回填值是否非空
```
典型：某"Default Rate Structure"字段在 Facility 未选时 disabled，此时留空提交**不报错**；只有先选中一个未配默认费率的 Facility、字段启用后留空才报 "Rate structure must be specified!"。仅探"全空提交"会把这条从未触发的错误误写进约束表和断言。

**记录到 `_context.md`**（「实测测试情报」§校验时机 / §6 校验规则要点）：
```
校验时机：逐字段标「blur 即报错」/「仅 Save 后异步报错」，错误文案抄原文
依赖字段：<B> 依赖 <A>——状态一（A未选）B: disabled/不参与校验；状态二（A已选）B: 启用/回填<值>
必填字段清单（空表单干跑实测）：字段1 / 字段2 / ...（区块级错误单独列）
```

---

## §5 跨页传播延迟 + 表格行按钮多样性

### 5.1 跨页写入传播延迟（跨页断言场景可选，命中则必测）

**症状**：A 页操作、B 页验证效果时，写入到 B 页可见**可能有传播延迟**——断言若对着 B 页还没刷新的旧数据就会误通过或误失败。凭"应该是实时的"假设是这类场景的头号风险。

**检测方法（分场景，禁止逐跳延迟求和当链路总超时）**：
```
单跳 A→B：A 拿到成功信号（toast/API 响应）后立即导航 B 检查；未出现则每 1–2s 轮询，记录实际需几次×几秒才出现
链路 A→B→C（3+页）：端到端直接轮询最终验证页 C，记录 A 操作完成到 C 出现变化的总时长（勿逐跳相加）
长任务（报表/账单/批量）：改用更长间隔（5–10s）、更长总时长（2–5min）轮询，笔记注明"长任务轮询"，勿套用近实时上限
跨上下文（不同 BASE_URL/登录态/租户）：立即停止常规探查，标注"跨上下文，需人工确认可行性"，不硬套单上下文模板
```

**记录到 `_context.md`**（「实测测试情报」§异步/结果反馈）：
```
跨页传播：场景类型（单跳/链路/长任务/跨上下文）
  操作页<页>完成信号<toast/API> → 验证页<页>验证点<定位>
  传播延迟：同步立即 / 轮询 N次×M秒后可见 / 超时未生效（异常，需人工核实）
```

### 5.2 表格行按钮多样性（含 Action 列的表格必做）

**症状**：表格各行的操作按钮**可能数量/种类不同**（如只有部分行可编辑/可删除）——硬编码"第 N 行的按钮"会脆。

**检测方法**：
```
browser_evaluate --function "() => Array.from(document.querySelectorAll('tr.ant-table-row, .el-table__row')).slice(0,10).map(r => ({
  key: r.querySelectorAll('td')[0]?.innerText?.trim().slice(0,20),
  btnCount: r.querySelectorAll('button').length,
  hasPrimary: r.querySelectorAll('.el-button--primary').length > 0,
  hasDanger:  r.querySelectorAll('.el-button--danger').length > 0
}))"
```

**记录到 `_context.md`**（「实测测试情报」§交互元素清单 备注）：
```
表格行按钮：所有行相同 → 可固定 locator
           行间不同 → 下游须动态查找（扫描前 N 行找第一个含目标按钮的行），禁止硬编码特定行 identifier
```
