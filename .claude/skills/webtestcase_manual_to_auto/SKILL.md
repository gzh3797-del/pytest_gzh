---
name: webtestcase-manual-to-auto
description: Use when the user wants to convert web-UI manual testcases (Excel) into Playwright+pytest automation end-to-end — analyze case logic (record root cause, never fabricate) → clarify → batch-generate scripts → run+auto-debug → write results back → sediment new selectors into knowledge context. Triggers include 把网页手工用例转自动化, 生成 UI 自动化用例, 手工用例转 pytest, /webtestcase_manual_to_auto. NOT for meter/Acuview cases (use acuview_auto_testcase_generate).
---

# webtestcase-manual-to-auto

## 用途

把 Web UI 手工测试用例（Excel）**端到端**转成 Playwright + pytest 自动化脚本：

```
① 分析用例逻辑（懂→绿 / 缺信息→红并记根因 / 半自动→橙，严禁编造）
② 人工澄清（用户填「用户答复」列或对话回复）
③ 批量生成脚本（先出探查约束表）→ ④ 覆盖门禁自检（可追溯 + 步骤/预期零漏项）
→ ⑤ 运行 pytest → ⑥ 自动调试失败项（≤3 轮）→ ⑦ 回填调试结果列
⑧ 把调试中确认的新选择器/坑增量沉淀回知识库 context
```

本 skill 由旧 `webtestcase-analyze_debug` + `webtestcase_manual_to_auto` 两个 skill 合并而来。
缺失页面探查复用 `analyze-web-page`（webpage_analyze）skill；不重复造探查逻辑。

> 适用 **Web UI 用例**。电表 Acuview 上位机用例请用 `acuview_auto_testcase_generate`。

---

## 调用格式

```
/webtestcase_manual_to_auto                                  # 无参 → 跑 --status 列各模块状态后停
/webtestcase_manual_to_auto --module 用户管理
/webtestcase_manual_to_auto --module 用户管理 系统设置
/webtestcase_manual_to_auto --all
/webtestcase_manual_to_auto --case TestCase_AcuHMI_009_08_case01
/webtestcase_manual_to_auto --module 用户管理 --analyze-only     # 只分析染色，不生成
/webtestcase_manual_to_auto --module 用户管理 --debug-only       # 跳过生成，跑已有脚本→回填
/webtestcase_manual_to_auto --module 用户管理 --overwrite        # 覆盖已存在脚本
/webtestcase_manual_to_auto --all --file <其他用例集.xlsx> --project <项目名>
```

参数（存于 `$ARGUMENTS`）：
- `--file <路径>`：用例集 Excel。**默认** `knowledge/gateway/hmi1-7/testcase/【AcuHMI-1-7】项目软件全功能测试用例2026060701.xlsx`
- `--project <名>`：目标项目，**默认** `AcuHMI_1_7`（见「项目映射」）
- `--module <名..>`：一个或多个模块（空格分隔）
- `--all`：所有模块
- `--case <用例编号>`：仅处理单条（调试用；交互模式下遇不清楚立即停下确认）
- `--analyze-only`：只做阶段 1，输出待澄清清单后停
- `--debug-only`：跳过生成，直接跑已有脚本→自动调试→回填
- `--overwrite`：覆盖已存在测试文件（默认跳过，保护手动修改）

---

## 项目映射（通用骨架）

维护一张映射表（当前一条，新增项目在此扩展）：

| project | code_root | knowledge_root |
|---------|-----------|----------------|
| `AcuHMI_1_7` | `projects/AcuHMI_1_7` | `knowledge/gateway/hmi1-7` |

由此派生：
- `EXCEL_PATH` = `--file` 或默认用例集
- `CONTEXT_DIR` = `<knowledge_root>/requirements/context/`
- **落盘路径** = `<code_root>/tests/<自动化类型>/test_<用例编号下划线化>.py`
  - `自动化类型` 取自 Excel 同名列（如 `ui/about/general`）
  - 例：`用例编号=TestCase_AcuHMI_009_08_case01` + `自动化类型=ui/about/general`
    → `projects/AcuHMI_1_7/tests/ui/about/general/test_TestCase_AcuHMI_009_08_case01.py`
  - 幂等创建缺失的 `__init__.py`

`excel_writer.py` 与本 SKILL.md 同目录：
`.claude/skills/webtestcase_manual_to_auto/excel_writer.py`（所有命令从**仓库根**执行）。

---

## Excel 列约定

**只读现有列**：`模块 / 子模块 / 用例编号 / 用例标题 / 预置条件 / 测试步骤 / 预期结果 / 用例级别 / 自动化(是/否) / 自动化类型`。
过滤条件：仅处理 `自动化(是/否)="是"` 的行。

**新增 3 个 Claude 列**（`--init` 幂等追加，绝不动 A–P 现有列）：
- `需补充信息(claude识别回填)`：分析结论 + 字体色 `绿006400 / 红FF0000 / 橙FFA500`
- `用户答复(澄清信息)`：用户澄清输入（二次运行读取）
- `自动化脚本调试结果`：`是` / `跳过|原因` / `调试失败|原因`（**不写现有「测试是否通过」列**）

---

## 执行流程（不要跳过任何步骤）

### Step 0 — 参数解析 & 环境

解析 `$ARGUMENTS` → `EXCEL_PATH / PROJECT / TARGET_MODULES / CASE_FILTER / ANALYZE_ONLY / DEBUG_ONLY / OVERWRITE`。
由 PROJECT 查项目映射得 `CODE_ROOT / KNOWLEDGE_ROOT / CONTEXT_DIR`。
`--init` 幂等追加 Claude 列：
```bash
python .claude/skills/webtestcase_manual_to_auto/excel_writer.py --init "<EXCEL_PATH>"
```

**若无 `--module`/`--all`/`--case`**（或含 `--status`）：跑 `--status` 展示各模块绿/红/橙/未分析分布，然后**停下**等用户重新调用。
```bash
python .claude/skills/webtestcase_manual_to_auto/excel_writer.py --status "<EXCEL_PATH>"
```

**若 `DEBUG_ONLY=True`**：跳过 Step 1–3，直接进 Step 4（对已有脚本运行+调试+回填）。

### Step 1 — 读取用例

```bash
# --all：先 --list-modules 得模块名，再逐模块 --read
python .claude/skills/webtestcase_manual_to_auto/excel_writer.py --read "<EXCEL_PATH>" --module "<模块名>"
# --case：
python .claude/skills/webtestcase_manual_to_auto/excel_writer.py --read "<EXCEL_PATH>" --case "<用例编号>"
```
`auto != "是"` 的行跳过。

### Step 2 — 定位页面 context（需求 #3，约定 #11）

**先读** `CONTEXT_DIR/_INDEX_context.md`，按用例的 `模块/子模块` → 被测页面 **Web 导航菜单路径** 语义匹配对应
`<Prefix_SubPage>_context.md`（PascalCase，如 `Devices_DataLog_DataLogger_context.md`）。
**禁止**用测试目录名直接拼文件名。命中即读取复用其选择器/进入路径/实测测试情报。

**若整页 context 缺失**：调用 `analyze-web-page`（webpage_analyze）skill 现场探查并沉淀 `_context.md` + 更新 `_INDEX`，
再回到本步。（探查交给该 skill，禁止在本 skill 内自造一次性探查脚本；零残留。）

### Step 3 — 逐条分析并染色（需求 #5，严禁编造）

对每条用例结合 `title/precondition/steps/expected` + 命中的 context 判断：

| 结论 | 颜色 | 写入 `需补充信息` 列 | 后续 |
|------|------|--------------------|------|
| 完全理解 | 绿 `006400` | `已理解：<测试点一句话> \| 断言方式：<如何 expect>` | 进 Step 4 生成 |
| 信息缺失/步骤不清 | 红 `FF0000` | `根因：<为何无法自动化，具体到缺哪个元素/哪步模糊>\n疑问：1...` | 记录后**跳过生成** |
| 半自动化 | 橙 `FFA500` | `[半自动化] 需人工介入：<哪步不能自动化及原因>` | 跳过生成 |

**「完全理解」需同时满足**：涉及 UI 元素在 context 有明确定位；导航路径清晰；预期结果有可断言现象。
**触发红色**（任一）：元素无定位依据 / 依赖系统时间·外部工具·硬件 / 预期仅「应正常」无具体现象 /
Factory Reset·固件升级·出厂状态 / 前置需不可复现设备状态。

**二次运行**：`用户答复` 有内容的红/绿行，结合答复重判——懂了改绿，仍有疑问保持红并追问。
已是绿且 `用户答复` 为空的行不覆盖（幂等）。

每模块分析完立即写入（JSON 用 `--data-file` 传 UTF-8 文件，>50 条分批）：
```bash
python .claude/skills/webtestcase_manual_to_auto/excel_writer.py --write "<EXCEL_PATH>" --data-file <tmp.json>
```

> **批量 vs 单条（需求 #5 的核心约束）**：
> - 批量（`--module`/`--all`）：红色用例**只记根因、跳过、继续**其他用例；末尾统一汇总。
> - 单条交互（`--case`）：遇不清楚**立即停下**向用户确认，不臆造。
> - `--analyze-only`：到此结束，输出红色待澄清清单后停。

### Step 4 — 生成脚本

仅对**绿色**用例（`--debug-only` 时跳过本步）。

**4.0 先产出探查约束表**（[COVERAGE_GATE.md](COVERAGE_GATE.md) §一）：写代码前，把该用例要用的每个
控件/字段/断言值整理成表，每行注明证据来源（命中的 `*_context.md` 某节或本次探查）。约束表是代码的**唯一事实源**，
留白项禁止凭推断写代码。

**4.1 生成** 按 [GENERATION_RULES.md](GENERATION_RULES.md) 生成 `test_<用例编号下划线化>.py`：
- import/fixture 用 `projects.<PROJECT>.settings` + `LoginPage` + `login_page` fixture
- 头部注释块（用例编号/标题/预置/步骤/预期 + **探查注**）
- 有副作用加 `try/finally` 清理
- **测试过程配置恢复（硬性）**：凡改动了被测配置的用例，必须"先读原值 → `finally` 中无条件还原"，把配置恢复到测试前状态（[GENERATION_RULES.md](GENERATION_RULES.md) §8）——不只是删数据/登出，开关/模式/阈值/通信参数等**每一项被改动的配置**都要逐项还原
- 落盘到 `<CODE_ROOT>/tests/<自动化类型>/`，幂等建 `__init__.py`
- 文件已存在且无 `--overwrite` → 跳过
- 生成中发现某步元素 context 未描述且未探明 → 用 `# TODO:` 占位并把该用例转红记根因，**不臆造选择器**
- 代码里每个 locator/label/断言值都必须来自 4.0 约束表
每生成一个立即写盘。

### Step 4.5 — 生成后覆盖门禁（AI 自检，硬门禁）

对每个新生成的 `test_*.py` 跑 [COVERAGE_GATE.md](COVERAGE_GATE.md) 的五道门禁（**跑 pytest 前**执行，
静态检查与运行互补）：
- **Gate A 可追溯性**：代码中每个 locator/label/断言值能回溯到约束表；回溯不到 = 编造嫌疑 → 转红/补探，不放行。
- **Gate B 步骤覆盖（零漏项）**：`测试步骤` 每行都有对应操作；任一 missing → FAIL。
- **Gate C 预期覆盖（零漏项）**：`预期结果` 每行都有对应断言（反模式 `A or B or C`、只验第一行、helper 内断言不计分 → 判 partial）；任一 missing → FAIL。
- **Gate D 配置恢复（硬门禁）**：改动了被测配置的用例，必须在 `finally` 中把每个被改项还原为测试前原值（[GENERATION_RULES.md](GENERATION_RULES.md) §8）；有配置副作用却无 `finally` 恢复、或恢复只在正常路径/未覆盖全部被改项 → FAIL。纯只读用例不适用（N/A，不扣分）。
- **Gate E 步骤↔脚本语义对照（硬门禁）**：产出对照矩阵，逐条比对步骤/预期的**关键参数**（身份/角色/输入值/目标对象/断言值）与代码是否**语义一致**（不只是"有代码"）；任一行关键参数不符（如角色 view↔edit、填错谁的密码）= L2 ❌ → FAIL。矩阵是 Step 7 回显的统一产物。

判定：Gate A/D/E 失败或 B/C 有 missing → FAIL（无视加权分）；加权分 Gate B×50%+Gate C×50%，≥95 PASS / 70–94 PASS_WITH_WARNINGS（二次自评）/ <70 FAIL。
- FAIL 且属缺步骤/弱断言/缺配置恢复/语义不符 → **补齐或改对后重评**（≤3 轮，与调试循环共享轮次预算）；Gate D 缺项直接补 `finally` 恢复动作，Gate E 缺项直接把参数改对。3 轮仍无法覆盖（如预期本身模糊）→ 转红记根因，调试结果列写 `调试失败|门禁未过:缺<步骤/预期/配置恢复/语义不符>`，**绝不删步骤或凑弱断言放行**。
- 仅 verdict ∈ {PASS, PASS_WITH_WARNINGS} 的用例进入 Step 5。

### Step 5 — 运行 + 自动调试

从仓库根运行：
```bash
python -m pytest <目标目录..> -v --tb=short
```
- PASSED/XPASSED → `是`；SKIPPED → `跳过|reason`；FAILED/XFAILED → 进调试。
- 首次通过/跳过**立即回填**结果列（不等调试结束）。
- 失败项按 [DEBUG_PATTERNS.md](DEBUG_PATTERNS.md) 自动调试，每条 ≤3 轮；仍失败标 `调试失败|原因`。

回填（`--results-file` 传 UTF-8 文件，禁用 shell 内联避免 GBK 乱码）：
```bash
python .claude/skills/webtestcase_manual_to_auto/excel_writer.py --write-result "<EXCEL_PATH>" --results-file <tmp.json>
```
JSON 每条须含 `"case_id"` 与 `"value"`。

### Step 6 — 增量沉淀 context（需求 #4）

把调试中**确认为真实**的选择器 / 框架坑 / 成功判定，**增量追加**到对应 `*_context.md` 的
「实测测试情报（pytest / Element Plus，来源：<日期> 联机实测）」节；若探明了新可路由子页，更新 `_INDEX_context.md`
并在项目 `context.md` 加指针。只**追加不删**既有条目。仅主 AI 写 `knowledge/`。

### Step 7 — 汇报（需求 #6）

输出：
1. **统计**：新生成 / 覆盖门禁 PASS / 首次通过 / 调试修复 / 调试失败（含门禁未过）/ 跳过 / 已存在跳过。
2. **逐条清单**：每条列**完整 `用例编号`** + **覆盖门禁结论（B/C 分 + D 配置恢复 restored/N/A + E 语义符合/❌）** + 一句话**测试逻辑**（做什么、断言什么），覆盖绿色已生成、skip、失败（带根因）。
3. **对照矩阵回显（Gate E）**：每条用例打印「步骤/预期 ↔ 脚本对应 ↔ L1覆盖 ↔ L2语义 ↔ 成对比较值」矩阵（格式见 [COVERAGE_GATE.md](COVERAGE_GATE.md) §二 Gate E），便于测试人员逐条核对、快速定位语义偏差；失败用例展开矩阵并高亮 ❌ 行。
4. **红色待澄清清单**：`用例编号 → 根因/疑问`，提示用户在「用户答复」列填写或对话澄清后二次运行。
5. 沉淀了哪些 `*_context.md`。

---

## 注意事项

- 从**仓库根**执行所有命令；Excel 路径含中文/空格用双引号。
- **严禁编造**任何选择器/步骤/预期（需求 #5）——无依据即标红记根因；由 Step 4.5 Gate A 可追溯性硬门禁兜底。
- **覆盖门禁**（Step 4.5）：跑得过 ≠ 覆盖到位；每条步骤/预期必须逐条覆盖（零漏项），不得以"综合测试""已有类似断言"跳过。
- 只写 3 个 Claude 列，**不改**任何原始列；不改「用例编号」列（约定 #12）。
- 函数名内嵌用例编号、一函数一用例（约定 #10），编号与 Excel 逐字符一致，仅 `-`→`_`。
- 已存在脚本默认跳过（`--overwrite` 覆盖）；每文件即时写盘，中断不丢已生成。
- 调试每条 ≤3 轮；破坏性/需硬件的用例直接 skip 不硬试。
- 探查一律走 `analyze-web-page` skill；**禁止**在项目任何目录残留一次性探查脚本（CLAUDE.md 最高优先级禁止项）。
- Python 代码零告警（约定 #8）。
