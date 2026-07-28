---
name: analyze-web-page
description: Use when the user wants to exhaustively explore a site and sediment per-subpage 页面上下文 (`_context.md`) for manual-to-automated testcase conversion / selector reuse (约定 #11), or invokes /analyze-web-page <URL> <项目知识库根>. Triggers include 分析网页, 穷尽探索页面, 生成页面上下文/选择器沉淀, 为手工用例转自动化做页面探查.
---

# analyze-web-page

## 用途

使用 Playwright MCP 穷尽式探索一个站点，**为每个可路由子页产出一份「页面上下文」文档 `<Prefix_SubPage>_context.md`**，落到项目知识库的 `requirements/context/` 目录（约定 #11「选择器沉淀」），供 `webtestcase_manual_to_auto` 把手工用例转自动化、以及后续 UI 用例直接复用选择器。

> 产物契约、命名、格式全部对齐仓库现有 `_context.md` 文件（见 [OUTPUT_STANDARD.md](OUTPUT_STANDARD.md)），**不再**产出旧的整页 10 章大报告。

## 触发条件

- **调用命令**：`/analyze-web-page <URL> <项目知识库根> [登录凭据]`
  - `<项目知识库根>` = 项目一览中该项目 `context.md` 所在目录（如 `knowledge/gateway/hmi1-7`）；交付物写入其下 `requirements/context/`。
  - `[登录凭据]` 可选（账号/密码/自签证书提示），缺省时运行时询问。
  - **开工前先读** `<项目知识库根>/requirements/context/_INDEX_context.md`（若存在）：了解已沉淀的子页、命名与通用说明，本次为增量合并而非覆盖。

## 路径约定（中间产物 vs 交付物，硬约束）

| 记号 | 含义 | 归属 |
|------|------|------|
| `{output_dir}` | **临时工作目录**：阶段1–4 的所有中间快照（`phase*_*.md` / `form_*.md` / `subpage_*.md` / `{name}_screenshot.png` / `expected_ops.json`）落这里 | 工作区 `.analyze_scratch/`（不进 git），**任务结束全部清理，零残留** |
| `{ctx_dir}` | **交付目录** = `<项目知识库根>/requirements/context/` | 只留 `<Prefix_SubPage>_context.md` + `_INDEX_context.md` |
| `{skill_dir}` | 本 skill 目录 | 用于调用 `scripts/audit_snapshots.py` |

> ⚠️ 遵 CLAUDE.md「禁止操作」：**中间调试/探查产物一律不得留在 `{ctx_dir}` 或项目任何知识库目录**。阶段1–4 提到的 `{output_dir}` 一律指临时工作目录；只有阶段5 产出的 `_context.md` / `_INDEX_context.md` 写入 `{ctx_dir}`。

## 核心原则

> **交互即探索** — 每遇到一个可交互元素，先触发它看看会发生什么，再继续前进。
> **看不到不代表没有** — 页面上凡是能折叠、能滚动、能点击展开的区域，都必须主动打开查看。
> **枢纽即目录** — 遇到导航/索引型页面（交互元素中 80% 以上为 `<a>` 链接，且表单/弹框等不足 5 个），必须进入深度递归模式，逐一点击每个同域内部链接，对目标子页面执行完整探索。
> **产物即证据** — 每个操作必须产出对应的快照文件。`[x]` 不等于"我做过了"，等于"我有快照文件证明我做过了"。
> **结构之外还有情报** — 光记录"有什么元素"不够；自动化脚本真正需要的是：调用的 **API 端点**、校验**触发时机**（blur / Save）、**框架真实选择器与坑**、结果**反馈**（toast/确认框文案）、**高危操作**标注。这些必须实测采集，不得靠推测（见阶段1/阶段3 强制采集项 + OUTPUT_STANDARD `_context.md`「实测测试情报」节）。

## 穷尽性约束机制（6项强制规则）

### 规则1：原子操作规则

**禁止在一次 MCP 调用中批量操作多个交互元素。** 每个独立的交互操作必须是独立的 MCP 调用，操作后必须紧跟一次 `browser_snapshot`。

```
禁止行为：
✗ 使用 browser_run_code_unsafe 批量选择多个option
✗ 使用 browser_evaluate 循环遍历checkbox
✗ 在一次 browser_fill_form 中填写多个字段（除非每个字段都有独立快照）
✗ 使用 for/forEach 循环操作多个元素

强制行为：
✓ 每个 option 选择 → 独立的 browser_select_option → 独立的 browser_snapshot
✓ 每个 checkbox 状态切换 → 独立的 browser_click → 独立的 browser_snapshot
✓ 每个 radio 选择 → 独立的 browser_click → 独立的 browser_snapshot
```

### 规则2：产物计数契约

探索开始前，从阶段1快照统计预期操作次数，写入 `{output_dir}/expected_ops.json`（机器可读，供 `audit_snapshots.py` 直接消费）：

```json
{
  "expected_counts": {
    "total_options": 0,
    "checkboxes": 0,
    "radios": 0,
    "submit_buttons": 0,
    "reset_buttons": 0
  },
  "sub_pages": []
}
```

> `total_options` = 各 select 的 option 数之和；`checkboxes`/`radios`/`submit_buttons`/`reset_buttons` 为各自元素数；`sub_pages` 填同域内部链接列表（用于子页面快照计数）。审计与最终计数均以本文件为预期基准。

### 规则3：缺省否定清单

所有自检项**默认为 `[ ]`（未做）**。改为 `[x]` 必须同时满足：
1. 对应的 MCP 操作已执行
2. 对应的 snapshot 文件已产出
3. snapshot 文件路径写在 `[x]` 后面

**缺少任一条 → 该项不能改为 `[x]`。**

### 规则4：穷尽→计数翻译

每个模糊需求必须翻译为可计数指标：

| 模糊表述 | 翻译为 |
|----------|--------|
| "探索所有下拉框" | S个select × 各自option数 = N次独立选择 |
| "测试所有复选框" | C个checkbox × 2种状态(选中/未选) = 2C次操作 |
| "遍历所有子页面" | L个内部链接 = L个子页面快照 |
| "点击所有提交按钮" | Sb个submit按钮 = Sb次点击+Sb个快照 |

### 规则5：独立审计Agent（阶段3.5）

主Claude完成阶段3后，**必须**启动独立审计子Agent。审计Agent拥有独立上下文（adversarial stance），逐项验证产物文件数是否匹配预期计数。任一维度 FAIL → 强制返回阶段2。

### 规则6：反转质询（Reversal Test）

阶段3自检清单全部为 `[x]` 后，进入阶段3.5审计前，**必须在对话中输出反转质询块**。

反转质询要求 Claude 站到批评者立场，列出自己**可能遗漏的N个问题**，然后逐条回应。**有未回应或回应不充分的条目 → 禁止进入阶段3.5。**

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 🔍 反转质询 — 如果我是一个挑剔的审计员，我会问：
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. "select[name='s1'] 有 onchange=alert，每次切换都处理了 dialog 吗？"
   → ✅ 已处理。证据: form_select_s1_o1~o3.md（3次各触发1次 alert 均已处理）。
2. "子页面 showModal.htm 的 5 个 button 你逐个触发了，还是只记了数量？"
   → ❌ 只记了数量、未逐个触发 → 回退阶段2重新探索 showModal.htm。
3. "select[multiple] 的全选组合都测了吗？"
   → ⚠️ 仅测部分组合；非核心交互，在「局限」说明即可。
（每条回应必须附快照路径或承认遗漏，禁止全部回"已确认"）

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 质询结果: 3个问题中 1✅ / 1❌ / 1⚠️
 → ❌ 存在 → 禁止进入阶段3.5 → 回退阶段2补充该 ❌ 项
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

**反转质询规则：**

| 规则 | 说明 |
|------|------|
| 最少质询数 | 不低于 `(select数 + checkbox数 + radio数 + 子页面数) × 10%` 条，最少3条 |
| 质询范围 | 必须覆盖至少3个不同维度（下拉框/复选框/单选组/子页面/弹框/提交重置） |
| 回应标准 | 每条回应必须引用具体的快照文件路径或承认遗漏 |
| ❌ 处理 | 任一 ❌ → 回退阶段2补充 → 补充后重新质询 |
| ⚠️ 处理 | 累计 ⚠️ ≥ 2条 → 等同于1条 ❌ |
| 质询深度 | 禁止所有回应都是"已确认"——至少1条必须是真正的自我质疑（发现一个问题或不足） |

## 执行流程（6阶段穷尽式分析）

### 阶段1：初始扫描

目标：了解页面全貌，建立交互路线图，**统计预期操作次数**。

```bash
browser_navigate --url "<URL>"
browser_wait_for --time 3
browser_snapshot --filename "phase1_initial.md"
browser_take_screenshot --filename "{output_dir}/{name}_screenshot.png" --fullPage true
```

然后**必须**从 phase1_initial.md 中提取计数，写入 `{output_dir}/expected_ops.json`（schema 见规则2）：

- 统计 select 数量和每个 select 的 option 数量（option 总数 → `total_options`）
- 统计 checkbox 数量 → `checkboxes`
- 统计 radio 数量 → `radios`
- 统计 submit/reset 按钮数量 → `submit_buttons`/`reset_buttons`
- 统计同域内部链接，列入 `sub_pages`

**同时（强制）采集两项测试情报：**

1. **加载态 API 端点**：`browser_navigate` 后调用 `browser_network_requests`，记录本页加载触发的 GET/POST 接口（method + path），供 `_context.md`「实测测试情报」节的 API 小节使用。保存/提交后再采集一次，记录写接口。
2. **UI 框架识别（含下拉/选择/过滤控件时，强制先跑 JS DOM 计数）**：**禁止仅凭 a11y 快照类名猜框架**——Ant Design Select 与 Element Plus Select 在快照里都是 `combobox`、文字相同，无法区分。含下拉/选择/过滤控件的页面，识别前必须先 `browser_evaluate` 计数 `.el-select__wrapper` vs `.ant-select`（判定规则与完整框架坑见 [PITFALLS.md](PITFALLS.md) §3）。据此在 `_context.md`「实测测试情报」节给出**框架专属真实选择器**（如 `.el-radio` / `.el-select__input` / `.el-message--success` / `.el-pagination`，Ant 则 `.ant-select` / `tr.ant-table-row` 等）与对应**框架坑清单**。

> **框架坑速查（Element Plus，完整清单+Ant Design 见 [PITFALLS.md](PITFALLS.md) §3）：**
> - `el-radio` 原生 `<input>` 被 `.el-radio__inner` 遮挡 → 点击 input 超时，**须点 label 文案**；坐标点击兜底（移开鼠标后按 `bounding_box` 点）。
> - `el-select`（`.el-select__input`）：选项在独立 popper，用 `input.getAttribute('aria-controls')` 取 listId 再查该 popper 下 `.el-select-dropdown__item`；`aria-controls` 为 None 时用全局 `.el-select-dropdown__item` + 可见性过滤；**禁止全局 querySelectorAll 乱抓**。
> - 同名 group/字段（如证书 Issuer/Subject、双网口 Interface Status/IP）**必须父容器 scope**。
> - 自定义确认框可能是产品级 **Yes/No**（非标准 MessageBox）。
> - 下拉选项文案可能带空格（如 `Post Channel 1`）；被禁用项**仍在下拉但带 `is-disabled`**（不移除）。

### 阶段2：穷尽探索

目标：触发所有隐藏内容，不让任何角落漏掉

**执行顺序（严格按此进行，不可跳过）：**

```
优先级1：滚动
  → 全页滚动到底部 → 触发懒加载 → browser_snapshot
  → 所有局部滚动容器逐屏扫描 → 每屏 browser_snapshot

优先级2：展开
  → 每个折叠面板/手风琴点击展开 → browser_snapshot
  → 每个标签页点击切换 → browser_snapshot
  → 每个下拉框打开查看全部选项（仅查看，不操作选项）→ browser_snapshot

优先级3：触发
  → 每个hover菜单用browser_hover触发 → browser_snapshot
  → 每个可能打开弹框的元素点击触发 → browser_snapshot
  → 每个抽屉/侧边栏触发打开 → browser_snapshot

优先级4：遍历
  → 分页器 → 逐页翻到最后一页 → 每页 browser_snapshot
  → "加载更多" → 持续点击到无新内容 → 每次点击后 browser_snapshot

优先级5：子页面跟随（枢纽页必须执行，默认全量遍历）
  → 统计所有同域内部链接（排除 JS伪协议、外部域、mailto/tel/javascript: 等非导航链接）
  → **对每个内部链接逐一进入子页面** → 在子页面执行优先级1-4完整探索
  → 每个子页面探索完成后保存快照: {output_dir}/subpage_{序号}_{名称}.md
  → 记录子页面 URL、标题、交互元素统计（表单数、按钮数、下拉框数、弹框数）
  → **SPA 导航实测（hash/history 路由站点必做，见 [PITFALLS.md](PITFALLS.md) §1/§2）**：
      · 路由守卫：菜单点击到达后，从其他区域 browser_navigate 直跳该 URL，核对落点是否被重定向 → 记「goto 直达可用 / 须菜单点击链」
      · 同路由 goto 空操作（列表/分页/筛选页）：改状态（翻页/筛选）后再 goto 同路由，核对状态是否重置 → 记「回列表是否须显式重置/二次 goto」
  → 返回索引页（browser_navigate_back 或重新 goto 索引URL），继续下一个链接
  → **默认全量遍历，不遗漏任何子页面。** 仅当用户明确指令"抽样X%"或子页面>100且用户同意时启用抽样
  → 探索每个子页面时，必须检测并标注响应类型：
    · 正常页面（200 + DOM内容）→ 完整探索
    · 204 No Content → 标注「特殊响应页: 204」
    · 延迟加载页（>5s 才返回内容）→ 标注「特殊响应页: 延迟加载」
    · 重定向页（3xx）→ 记录最终 URL
    · 错误页（4xx/5xx）→ 标注 HTTP 状态码
    · Alert-on-load页 → 标注「含 onload alert」，注册dialog handler后继续
```

**⚠️ 原子性警告：阶段2中每个操作后都必须 browser_snapshot。禁止合并操作。**

### 阶段3：表单深度探索

目标：覆盖所有输入状态和条件分支

**原子操作要求：以下8项操作，每一项的每个子操作都是独立的 MCP 调用 + 独立的 browser_snapshot：**

```
操作列表（原子执行，禁止批量）：
  1. 下拉框 → 依次选择每个option（每次browser_select_option后browser_snapshot），记录联动变化
  2. 单选组 → 每个option单独选中一次（每次browser_click后browser_snapshot）
  3. 复选框 → 每个checkbox分别测试选中（check）和未选中（uncheck），各一次browser_snapshot
  4. 级联选择 → 逐级展开到最末级（每级browser_select_option后browser_snapshot）
  5. 搜索框 → 输入内容（browser_type后browser_snapshot），等待建议列表（browser_snapshot）
  6. 上传/拾色器/日历 → 记录触发方式（browser_snapshot）
  7. 提交按钮 → 每个submit按钮逐一实际点击，记录跳转URL或错误提示（每次browser_click后browser_snapshot）
  8. 重置按钮 → 每个reset按钮逐一实际点击，验证所有字段恢复默认值（每次browser_click后browser_snapshot）
```

**快照文件命名规则：**

```
{output_dir}/form_select_{select标识}_{option值}.md     ← 下拉框选择
{output_dir}/form_checkbox_{标识}_{checked|unchecked}.md ← 复选框状态
{output_dir}/form_radio_{标识}_{value}.md               ← 单选选择
{output_dir}/form_submit_{标识}.md                       ← 提交点击
{output_dir}/form_reset_{标识}.md                        ← 重置点击
{output_dir}/form_search_{标识}.md                       ← 搜索输入
```

### 阶段3+：交互实测情报采集（强制，与上 8 项并行）

以下 4 类情报**必须实测采集**并写入 `_context.md`「实测测试情报」节，不得靠推测：

```
9.  校验时机（逐个必填/受限字段）：填非法/超长值 → blur → 看是否报错；再点 Save → 看是否报错。
    分类记录为「blur 即报错」或「仅 Save 后异步报错」，并抄录错误文案原文。
10. 结果反馈：每个提交/危险操作触发后，记录 toast 文案原文（如 .el-message--success "Create Success."）
    与确认框类型（标准 dialog / 产品自定义 Yes/No）及按钮文案。
11. 高危操作标注：识别 reboot / factory reset / reset logs / clear logs / firmware upload / import-export
    / delete / 改 IP 等——标注为高危（执行前须确认、禁无人值守、脚本需带重启等待+重登）。
12. 异步/分页事实：列表是否异步刷新（需 reload 轮询才出现新行）、是否需跨页遍历查找/删除。
13. 依赖字段多状态（含级联/条件禁用字段的表单必做，见 [PITFALLS.md](PITFALLS.md) §4）：字段 B 依赖 A 时，
    须探两种状态——「A 未触发」B 是否 disabled/不参与校验，「A 已触发」B 是否启用/被自动回填。
    禁止只探"全空提交"就把从未触发的校验错误当必填规则写入。多入口表单须对"空表单入口"干跑。
14. 表格行按钮多样性（含 Action 列的表格，见 [PITFALLS.md](PITFALLS.md) §5.2）：browser_evaluate 扫前 10 行按钮分布，
    记「所有行相同（可固定 locator）/ 行间不同（下游须动态查找，禁止硬编码特定行 identifier）」。
15. 跨页传播延迟（跨页断言场景，见 [PITFALLS.md](PITFALLS.md) §5.1）：A 操作拿到成功信号后立即到 B 页轮询验证点，
    记实际传播延迟；区分单跳/链路(端到端测，勿逐跳求和)/长任务(更长轮询)/跨上下文(停止，标注需人工确认)。
```

> **下拉选项全量枚举（强制）**：凡可枚举的下拉/单选，其选项值必须**逐个抄录真实文案**（注意大小写/空格/无复数 s），禁止用"运行时确认""选项若干"等占位搪塞可枚举项。

### ⚠️ 阶段3穷尽自检（缺省否定 — 初始全为 `[ ]`）

执行完阶段3全部8项操作后，**必须在对话中输出以下清单**。

**重要：以下清单初始全部为 `[ ]`。只有同时满足 (1)操作已执行 (2)快照文件存在 (3)文件路径附在后面，才能改为 `[x]`。非交互页（无表单元素）全部标 `[-]`。**

```
阶段3穷尽自检 — 缺一项都不得进入阶段4:

1. 下拉框覆盖（预期 {S}个select × {O_total}个option = {N}次操作）:
   [ ] select[{name/id}]: option[{val1}], 快照: {output_dir}/form_select_{name}_{val1}.md
   [ ] select[{name/id}]: option[{val2}], 快照: {output_dir}/form_select_{name}_{val2}.md
   ...（每个option一行）

2. 单选组覆盖（预期 {R}个radio = {R}次操作）:
   [ ] radio[{name/id}]: 选中[{val1}], 快照: {output_dir}/form_radio_{name}_{val1}.md
   ...（每个radio一行）

3. 复选框覆盖（预期 {C}个checkbox × 2状态 = {2C}次操作）:
   [ ] checkbox[{name/id}]: checked, 快照: {output_dir}/form_checkbox_{id}_checked.md
   [ ] checkbox[{name/id}]: unchecked, 快照: {output_dir}/form_checkbox_{id}_unchecked.md
   ...（每个checkbox两行）

4. 级联选择覆盖（有则列出，无则标 [-]):
   [-] 页面无级联选择元素

5. 搜索框覆盖（有则列出，无则标 [-]):
   [ ] search[{name/id}]: 输入内容, 快照: {output_dir}/form_search_{name}.md

6. 上传/拾色器/日历（有则记录，无则标 [-]):
   [-] 页面无此类元素

7. 提交按钮覆盖（预期 {Sb}个submit = {Sb}次操作）:
   [ ] submit[{name/id}]: 已点击, 快照: {output_dir}/form_submit_{name}.md
   ...（每个submit一行）

8. 重置按钮覆盖（预期 {Rb}个reset = {Rb}次操作）:
   [ ] reset[{name/id}]: 已点击, 快照: {output_dir}/form_reset_{name}.md
   ...（每个reset一行）

9. 加载态/写操作 API 端点已采集（browser_network_requests）: [ ]
10. 每个必填/受限字段校验时机已实测分类（blur / Save）: [ ]
11. 提交/危险操作的结果反馈（toast 文案 + 确认框类型）已记录: [ ]
12. 高危操作已标注; 异步/分页事实已记录; 下拉选项已全量枚举（无"运行时确认"占位）: [ ]
```

> **自检欺诈检测**：如果对话中出现过 `browser_run_code_unsafe` 用于批量表单操作 → 以上所有项强制回退为 `[ ]`，必须重新逐元素独立操作。

> **非交互页（枢纽页/纯导航页）**：8项全部标 `[-]`，豁免阶段3（但子页面内的表单仍需探索）。

### 阶段3.4：反转质询（进入阶段3.5前的强制门禁）

阶段3自检清单全部为 `[x]` 或 `[-]` 后，**必须在对话中输出反转质询块**（格式参见规则6）。

质询结果处理：
- 全部 ✅ → 进入阶段3.5
- 有 ❌ 或累计 ≥2条 ⚠️ → 回退阶段2补充 → 补充后重新质询
- 质询未输出 → 禁止进入阶段3.5

### 阶段3.5：独立审计（强制门禁 — 不可跳过）

阶段3.4反转质询通过（无 ❌，⚠️ < 2）后，**必须启动独立审计Agent**。审计Agent拥有独立上下文，其提示词为 adversarial stance（默认假设主Claude有遗漏）。

审计Agent执行以下检查并输出 JSON 报告：

```json
{
  "audit_pass": false,
  "snapshot_dir": "{output_dir}/",
  "items": [
    {"dimension": "下拉框", "expected": 9, "actual": 9, "status": "PASS", "missing": []},
    {"dimension": "复选框", "expected": 8, "actual": 4, "status": "FAIL",
     "missing": ["form_checkbox_cv2_checked.md", "form_checkbox_cv2_unchecked.md"]}
  ],
  "failures": ["复选框: 预期8个快照, 实际4个"],
  "verdict": "RETURN_TO_PHASE2"
}
```

> 上方 `items` 仅示意 PASS/FAIL 两种结构；实际须逐项列出下表全部 7 个维度。

**审计维度（7项，逐项输出 PASS/FAIL）。维度1-6 的计数由 `audit_snapshots.py` 跨平台统计（勿用 `ls`/`wc`），维度7 由审计Agent人工核查对话：**

| # | 维度 | 预期快照数 | 验证方式 |
|---|------|-----------|----------|
| 1 | 下拉框覆盖 | Σ(每个select的option数) | `audit_snapshots.py` 统计 `form_select_*` |
| 2 | 复选框覆盖 | checkbox数 × 2 | `audit_snapshots.py` 统计 `form_checkbox_*` |
| 3 | 单选组覆盖 | radio数 | `audit_snapshots.py` 统计 `form_radio_*` |
| 4 | 子页面覆盖 | 内部链接数 | `audit_snapshots.py` 统计 `subpage_*` |
| 5 | 提交按钮覆盖 | submit按钮数 | `audit_snapshots.py` 统计 `form_submit_*` |
| 6 | 重置按钮覆盖 | reset按钮数 | `audit_snapshots.py` 统计 `form_reset_*` |
| 7 | 原子操作合规 | N/A | 检查主Claude对话是否使用批量操作 |

**审计结果处理：**
- `verdict = "PASS"` → 进入阶段4
- `verdict = "RETURN_TO_PHASE2"` → 列出所有 FAIL 项 → 强制返回阶段2补充 → 补充后重新阶段3 → 再次审计
- 最多重试2次；第3次仍不通过 → 在 `_INDEX_context.md`「局限与未尽项」中详记未通过项，继续流程

**审计Agent调用模板：**

```
调用 Audit Agent:
  Agent(
    subagent_type: "general-purpose",
    prompt: """
    你是独立审计员，验证 stage2-3 的执行质量。默认假设有遗漏，用证据推翻。

    输入:
    - 预期操作计数: {output_dir}/expected_ops.json
    - 快照目录: {output_dir}/
    - 阶段3自检清单（见上方对话）

    审计步骤:
    1. 运行: python {skill_dir}/scripts/audit_snapshots.py {output_dir} {output_dir}/expected_ops.json
       （脚本跨平台统计各前缀快照文件数并逐维度对比，勿用 ls/wc）
    2. 读取脚本输出与退出码（0=PASS, 1=FAIL）
    3. 核查维度7（原子操作合规）：检查主Claude对话是否出现批量操作
    4. 按上述 JSON 格式汇总审计报告，verdict 取脚本结论与维度7的合并

    注意: 你是 adversarial 审计员，不要信任主Claude的自检清单。
    以脚本的文件系统计数为准。
    """
  )
```

### 阶段4：最终确认

目标：对比阶段1，验证无遗漏

**进入条件**：阶段3.5审计 verdict = "PASS"。

```bash
browser_snapshot --filename "phase4_final.md"
```

检查项：
- 对比阶段1的初始结构，确认是否有新增内容
- **运行快照计数验证**（跨平台，勿用 `ls`/`wc`）：
  ```
  python {skill_dir}/scripts/audit_snapshots.py {output_dir} {output_dir}/expected_ops.json
  ```
  退出码 0 = 计数对齐；退出码 1 = 列出缺口 → 返回阶段2补充
- **判断页面类型**（表单页/列表页/枢纽导航页/...），如果判定为枢纽页（链接数 > 20 且表单+弹框 < 5），确认已执行优先级5（子页面跟随）
- 如有遗漏，返回阶段2补充探索

### 阶段5：沉淀页面上下文文档（`_context.md`）

目标：把阶段1–4 的探索结果，**按可路由子页拆分**，逐页产出 `<Prefix_SubPage>_context.md` 到 `{ctx_dir}`，并维护 `_INDEX_context.md`。**完整格式规范见 [OUTPUT_STANDARD.md](OUTPUT_STANDARD.md)，产出必须与仓库现有 `_context.md` 同构。**

**逐子页产出流程：**

1. **枚举可路由子页**：从阶段2优先级5 的子页面清单 + 站点导航菜单树，列出每个「可独立路由/操作」的子页。结构一致仅编号不同的子页（如 Data Logger 1/2/3、Post Channel 1/2/3）**代表性合并为一份**，正文注明覆盖范围。
2. **命名派生**：按被测页 Web 导航菜单路径 → 每段 PascalCase、去空格/特殊字符、`_` 连接 + `_context.md`（如 `Devices / Data Log / AcuCloud` → `Devices_DataLog_AcuCloud_context.md`）。详见 OUTPUT_STANDARD §一。
3. **逐页填充** [OUTPUT_STANDARD.md](OUTPUT_STANDARD.md) 的单页骨架：页面标识 / 用途 / 交互元素清单（`getByRole` 定位）/ 状态与分支 / 校验要点 / 测试要点 / 机器可解析 JSON + **「实测测试情报」节**（用阶段1 的 API·框架识别、阶段3+ 的校验时机 blur|Save、结果反馈 toast、高危操作实测数据填充；拿不到的如「参考 pytest 用例路径」标「待补」）。
4. **写入 `{ctx_dir}`**，不写任何中间快照到该目录。
5. **维护 `_INDEX_context.md`**（OUTPUT_STANDARD §五）：按导航上下文分组列全部子页文档 + 通用说明（技术栈/框架坑/破坏性操作）+ 局限；已存在则增量合并、不删既有条目。
6. **交付计数自检（清理前跑）**：`python {skill_dir}/scripts/audit_snapshots.py {output_dir} {output_dir}/expected_ops.json {ctx_dir}` —— 中间快照逐维度对齐（退出码 0），`ctx_dir` 段信息性列出已交付 `*_context.md` 数与 `_INDEX_context.md` 状态。
7. **清理 `{output_dir}` 临时工作目录**（`phase*`/`form_*`/`subpage_*`/`expected_ops.json`/`screenshot` 等），**零残留**。
8. 逐页执行 OUTPUT_STANDARD §六 质量自检；不合格 → 修正 → 复检 → 合格后交付。

**对话内交付摘要**（不写入 `{ctx_dir}`，仅回给用户）：
- 探索完成度清单（三态 `[x]`/`[-]`/`[ ]` + 快照证据，见下）+ 阶段3.5 审计结论摘要 + 本次产出/更新的 `_context.md` 列表。

**探索完成度清单（对话输出，标记约定）：**

| 标记 | 含义 |
|------|------|
| `[x]` | 已验证存在并完成探索（附中间快照路径） |
| `[-]` | 确认不存在，无需探索（N/A） |
| `[ ]` | 应存在但未探索/遗漏（须在「局限」说明） |

```
[x] 页面已滚动到底部，懒加载已触发 — {output_dir}/phase4_final.md
[x] 所有折叠/手风琴/标签页已展开切换 — form_*/tab_*.md
[x] 所有下拉框已展开并全量枚举选项 — form_select_*.md ({N}个)
[x] 表单互斥选项已遍历 — form_checkbox_*.md + form_radio_*.md
[x] 分页/加载更多已遍历到最后一页 — pagination_*.md
[x] 每个可路由子页已逐个探索并产出 _context.md — subpage_*.md ({L}个) → {ctx_dir}/*_context.md
```

## 穷尽式探索规范

完整规范请参见 [EXHAUSTIVE.md](EXHAUSTIVE.md)，包含 9 大维度（视口与滚动 / 显式交互触发 / 弹框与浮层 / 表单深度 / 动态内容 / 多页面与导航 / CSS 隐藏内容 / 完成度检查 / 局限记录）。**枢纽/导航页的每个同域可路由子页 → 一份 `_context.md`**。

## 坑清单（SPA 导航 / 框架识别 / 表单校验 / 跨页传播）

页面探索的五类实战工程坑（路由守卫、同路由 goto 空操作、Ant/El Plus JS 计数区分、依赖字段多状态、跨页传播延迟 + 表格行按钮多样性）见 [PITFALLS.md](PITFALLS.md)。每类含**症状 / 检测方法（MCP）/ 记录到 `_context.md` 的字段**。阶段1 框架识别、阶段2 子页面跟随、阶段3+ 情报采集第 13–15 项均引用本清单。

## 输出规范

单页 `_context.md` 与 `_INDEX_context.md` 的完整格式、命名、字段、`getByRole` 定位约定、质量自检清单，见 [OUTPUT_STANDARD.md](OUTPUT_STANDARD.md)。

## 产物目录约定

```
{ctx_dir} = <项目知识库根>/requirements/context/   ← 交付物（唯一进 git 的产出）
├── _INDEX_context.md                 ← 页面上下文索引（增量维护）
├── {Prefix_SubPage}_context.md       ← 每个可路由子页一份（本 skill 产出）
└── ...

{output_dir} = .analyze_scratch/（工作区临时目录，不进 git，任务末清理零残留）
├── phase1_initial.md / phase4_final.md   ← 阶段1/4 快照
├── expected_ops.json                     ← 阶段1 预期计数（审计基准）
├── form_select_*/checkbox_*/radio_*/submit_*/reset_*/search_*.md  ← 阶段3 原子操作快照
├── subpage_{序号}_{名称}.md              ← 阶段2优先级5 子页面探索快照
└── {name}_screenshot.png                 ← 页面截图（可选）
```

产出关系：
- `analyze-web-page` → `{ctx_dir}/{Prefix_SubPage}_context.md`（逐子页）+ `{ctx_dir}/_INDEX_context.md`
- 下游 `webtestcase_manual_to_auto` 读取上述 `_context.md` 把手工用例转自动化
