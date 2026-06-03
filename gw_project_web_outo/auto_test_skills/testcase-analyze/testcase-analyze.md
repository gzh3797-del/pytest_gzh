# testcase-analyze — 手工测试用例分析 Skill

## 用途
分析指定模块的手工测试用例，判断每条用例是否可自动化实现，
将理解程度回填到 Excel「需补充信息(claude识别回填)」列（Q 列）：
- 深绿色：已理解，可直接编写自动化脚本
- 红色：有疑问，需用户澄清后才能实现
- 橙色：半自动化用例，部分步骤需人工介入

支持二次运行：用户在「用户答复」列填写澄清后重新调用，Claude 会更新颜色。

---

## 调用格式

```
/testcase-analyze
/testcase-analyze --status
/testcase-analyze --module 用户管理
/testcase-analyze --module 用户管理 系统设置 About
/testcase-analyze --all
/testcase-analyze --all --file Manual_testcase/【AcuHMI-1-7】测试用例_foAI.xlsx
```

参数说明：
- 无参数：运行 `--status` 展示各模块分析状态，由用户选择下一步操作
- `--status`：快速统计所有模块中 `auto=是` 用例的分析状态（绿/红/橙/未分析），识别剩余待处理用例
- `--module <名称> [名称...]`：分析一个或多个模块（空格分隔）
- `--all`：分析 Excel 中所有模块（二次运行时处理有用户答复的红色用例）
- `--file <路径>`：指定 Excel 文件（默认自动在 Manual_testcase/ 下查找含 foAI 的 xlsx）

> **"剩余用例"定义**：L列(`自动化`)="是" 且 Q列字体为红色的用例。
> - 有 R列答复 → 可立即做「二次分析」，运行 `--all` 或 `--module <名称>` 处理
> - 无 R列答复 → 需先在 Excel 中填写答复后再次运行

用户传入的参数存于：$ARGUMENTS

---

## 执行流程

收到调用后，按以下步骤执行，不要跳过任何步骤。

### Step 1 — 定位 Excel 文件

解析 $ARGUMENTS：
- 若包含 `--file <路径>`，使用该路径
- 否则，用 Glob 在 `Manual_testcase/` 下查找文件名含 `foAI` 的 `.xlsx` 文件，取第一个

将找到的路径记为 `EXCEL_PATH`。

### Step 2 — 确定目标模块列表

解析 $ARGUMENTS，按以下规则确定 `TARGET_MODULES`：

**情况 A — 包含 `--all`：**
执行：
```bash
python auto_test_skills/testcase-analyze/excel_writer.py --list-modules "<EXCEL_PATH>"
```
将返回 JSON 中所有 `module` 字段提取为列表，设为 `TARGET_MODULES`。

**情况 B — 包含 `--module` 且后面有值：**
将 `--module` 后面所有空格分隔的词（直到下一个 `--` 参数或末尾）收集为 `TARGET_MODULES`。

**情况 C — 包含 `--status`，或无 `--module` 且无 `--all`：**
执行：
```bash
python auto_test_skills/testcase-analyze/excel_writer.py --status "<EXCEL_PATH>"
```
将结果格式化后展示给用户，显示各模块分析状态总览：

```
模块                  auto=是  绿(已理解)  红(待二次)  红(待澄清)  橙(半自动)
──────────────────────────────────────────────────────────────────
用户管理               126       88         38          0           0
接入设备日志管理         65       15         23          0          27
系统设置                59       37         22          0           0
系统诊断                26        5         21          0           0
模板管理                21        5         16          0           0
接入设备参数设置         16       13          0          3           0
...

剩余待处理：
  · 122 条红色用例已有用户答复（R列）→ 运行 /testcase-analyze --all 即可二次分析
  ·   3 条红色用例等待用户澄清      → 请在 Excel R 列填写答复后再次运行

请选择操作：
  /testcase-analyze --all              （处理所有待二次分析）
  /testcase-analyze --module <模块名>  （处理指定模块）
```

然后停止，等待用户重新调用。

### Step 3 — 批量检查 struct.md 覆盖情况

用 Glob 列出 `product_structure_testcase_regulation/` 下所有 `*_struct.md` 文件。

**第一轮：匹配专属文档（per-module struct.md）**

逐一读取每个 `*_struct.md`，判断是否覆盖 `TARGET_MODULES` 中的某个模块
（文件名或第一行标题含模块名关键词）。

将 `TARGET_MODULES` 暂分为两组：
- `HAS_DOC_MODULES`：找到对应专属 struct.md 的模块
- `UNMATCHED_MODULES`：暂未匹配到专属 struct.md 的模块

**第二轮：检查产品级兜底文档（product-level struct.md）**

对所有在第一轮中未被匹配到任何模块的 `*_struct.md` 文件
（即文件名前缀不对应任何 TARGET_MODULE，如 `AcuHMI-1-7_struct.md`），
视为产品级兜底文档。

若存在兜底文档，读取其内容，逐一检查 `UNMATCHED_MODULES` 中每个模块名是否在文档中有对应章节：
- 有对应章节 → 移入 `HAS_DOC_MODULES`，标记 `source=fallback`
- 无对应章节 → 移入 `NO_DOC_MODULES`

若不存在兜底文档，`UNMATCHED_MODULES` 全部移入 `NO_DOC_MODULES`。

**最终两组：**
- `HAS_DOC_MODULES`：有专属 struct.md 或在兜底文档中有章节的模块
- `NO_DOC_MODULES`：两轮均未找到文档的模块

**若 `NO_DOC_MODULES` 不为空，输出以下警告（然后继续，不停止）：**

```
⚠️  以下模块缺少页面结构文档，将全部标红：
  - <模块名1>
  - <模块名2>
  ...

建议在产品级文档中补充对应模块章节，或新建专属 struct.md。
格式参考：product_structure_testcase_regulation/usermanagement_struct.md
```

读取 `product_structure_testcase_regulation/autotest_generativerule.md`（通用自动化规则）。
读取 `HAS_DOC_MODULES` 各模块的文档内容：
- `source=per_module`：读取对应专属 struct.md 全文
- `source=fallback`：从兜底文档中仅提取该模块对应的章节内容（不读整个文档）

### Step 4 — 初始化 Excel 列（幂等）

```bash
python auto_test_skills/testcase-analyze/excel_writer.py --init "<EXCEL_PATH>"
```

确认「需补充信息(claude识别回填)」（Q 列）和「用户答复(基于需补充信息,澄清信息)」（R 列）两列存在。

### Step 5 — 快速处理无文档模块（NO_DOC_MODULES）

对 `NO_DOC_MODULES` 中每个模块，依次执行：

```bash
python auto_test_skills/testcase-analyze/excel_writer.py --read "<EXCEL_PATH>" --module "<模块名>"
```

对读取到的每条用例：
- `auto` 不是 `"是"` → 跳过
- `claude_color` 是 `006400` 且 `user_reply` 为空 → 跳过
- `semi_auto = "是"` → 写入橙色：`[半自动化] 需人工介入步骤：<说明>`，颜色 `FFA500`
- 其余 → 统一写入红色：
  ```
  [⚠️ 缺少页面结构文档] 无法判断 UI 定位方式和导航路径。
  请补充 product_structure_testcase_regulation/<模块英文名>_struct.md 后重新分析。
  当前仅记录测试点：<用一句话描述该用例在测试什么>
  ```

每个模块处理完后立即写入 Excel（不等所有模块处理完再写）：
```bash
python auto_test_skills/testcase-analyze/excel_writer.py --write "<EXCEL_PATH>" --data '<json>'
```

### Step 6 — 精细分析有文档模块（HAS_DOC_MODULES）

对 `HAS_DOC_MODULES` 中每个模块，依次执行：

```bash
python auto_test_skills/testcase-analyze/excel_writer.py --read "<EXCEL_PATH>" --module "<模块名>"
```

对读取到的每条用例，按以下优先级判断：

**跳过条件（不写入，按优先级）：**
1. `auto` 字段不是 `"是"` → 跳过
2. `claude_color` 是 `006400`（深绿）且 `user_reply` 为空 → 已确认理解，不覆盖
3. `claude_color` 是 `FF0000`（红色）且 `user_reply` 为空 → 等待用户在 R 列填写澄清，不重复覆写疑问

**半自动化处理（橙色）：**
- `semi_auto = "是"` → 写入：`[半自动化] 需人工介入步骤：<说明哪些步骤无法自动化及原因>`，颜色 `FFA500`

**二次运行处理（`user_reply` 有内容，含红色或绿色用例）：**
- 结合用户答复重新判断：
  - 若现在理解 → 颜色改为深绿 `006400`，内容更新为理解摘要
  - 若仍有疑问 → 保持红色 `FF0000`，补充追问

**首次分析：**
综合 `title`、`precondition`、`steps`、`expected` 及对应 struct.md 判断：

| 理解程度 | 颜色 | 写入内容格式 |
|---------|------|------------|
| 完全理解 | 深绿 `006400` | `已理解：<测试点一句话摘要> \| 断言方式：<简述如何用 Playwright 断言>` |
| 有疑问 | 红色 `FF0000` | `疑问：\n1. <具体问题1>\n2. <具体问题2>` |

**判断「完全理解」需同时满足：**
- ✅ 所有涉及的 UI 元素在 struct.md 或 autotest_generativerule.md 中有明确定位方式
- ✅ 操作路径（导航步骤）清晰，与 struct.md 中描述的导航结构一致
- ✅ 预期结果有具体的可断言现象（弹框文本、元素状态、URL 变化等）

**触发「需澄清」的情形（任一即判为红色）：**
- UI 元素在文档中未描述，无法确定定位方式
- 步骤依赖外部环境（系统时间修改、外部工具、硬件操作）
- 预期结果模糊（仅写"应正常"、"应成功"，无具体现象描述）
- 涉及 Factory Reset / 固件升级 / 设备出厂状态
- 前置条件需要特定设备状态，当前测试环境无法复现

每个模块分析完后立即写入 Excel（若用例数超过 50 条，分批写入，每批 50 条）：
```bash
python auto_test_skills/testcase-analyze/excel_writer.py --write "<EXCEL_PATH>" --data '<json>'
```

### Step 7 — 输出合并摘要

**若存在 NO_DOC_MODULES，在摘要顶部输出文档补充指引：**

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⚠️  文档缺失模块（已全部标红）：<模块名1>、<模块名2>
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
补充步骤：
  1. 参考格式：product_structure_testcase_regulation/usermanagement_struct.md
  2. 新建文件：product_structure_testcase_regulation/<模块英文名>_struct.md
  3. 补充完成后执行：/testcase-analyze --module <模块名>
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

**各模块统计（每个模块一行）：**

```
模块                  总计   已理解  待澄清  半自动化  跳过
──────────────────────────────────────────────────────
用户管理               149     82      40       5       22
系统设置               100      0     100⚠️      0        0  ← 无文档
About                   12     10       2       0        0
──────────────────────────────────────────────────────
合计                   261     92     142       5       22
```

**待澄清用例汇总（有文档模块，需用户在「用户答复」列填写后重新调用）：**
```
[用户管理]
- <case_id>：<核心疑问一句话>

[About]
- <case_id>：<核心疑问一句话>
```

---

## 注意事项

- Excel 文件路径若包含中文或空格，用双引号包裹
- `--write` 的 `--data` 参数若 JSON 过长，先写入临时文件再传路径（由 Claude 判断）
- 分析过程中不修改用例的任何原始列（A-P），只写入「需补充信息(claude识别回填)」（Q 列）和「用户答复」（R 列）
- 每个模块写入完成后立即保存，避免中途中断导致数据丢失
- `--all` 模式下，无 struct.md 的模块会被快速批量标红，有 struct.md 的模块才做逐条精细分析
- **"剩余用例"定义**：L列(`自动化`)="是" 且 Q列字体为红色（`claude_color="FF0000"`）的用例。用 `--status` 查看各模块分布；有 R列答复的直接运行 `--all` 做二次分析，无答复的需先填写 R 列
