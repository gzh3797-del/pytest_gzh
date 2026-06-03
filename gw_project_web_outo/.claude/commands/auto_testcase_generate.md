# auto_testcase_generate — 自动化测试脚本生成 Skill

## 用途
读取 Excel 中已被 `/testcase-analyze` 标注为深绿色（已理解）的用例，
结合模块的 `*_struct.md` 页面结构文档，为每条用例生成独立的 Playwright + pytest 测试脚本；
生成后立即运行测试，对失败脚本自动诊断并修复，最终将调试结果回填至 Excel P 列。

**整体流程（无需人工干预）：**
```
生成脚本 → 运行测试 → 自动调试失败项 → 回填 P 列结果
```

**前提**：需先运行 `/testcase-analyze`，Excel 第 13 列有绿色标注后再调用本 Skill。

| Skill | 阶段 | 频率 |
|-------|------|------|
| `/testcase_auto_init` | 项目启动，生成框架骨架 | 一次性 |
| `/testcase-analyze` | 分析手工用例，标注绿/红/橙 | 反复使用 |
| `/auto_testcase_generate` | 生成脚本 → 运行 → 自动调试 → 回填 P 列 | 反复使用 |
| `/auto_testcase_generate --debug-only` | 跳过生成，直接运行已有脚本 → 自动调试 → 回填 P 列 | 重新调试时 |

---

## 调用格式

```
/auto_testcase_generate
/auto_testcase_generate --module 用户管理
/auto_testcase_generate --module 用户管理 系统设置
/auto_testcase_generate --all
/auto_testcase_generate --all --file Manual_testcase/其他.xlsx
/auto_testcase_generate --module 用户管理 --case case02_01
/auto_testcase_generate --module 用户管理 --overwrite
```

参数说明：
- 无参数：列出各模块绿色用例数，等待用户重新调用
- `--module <名称> [名称...]`：生成指定模块的测试脚本
- `--all`：生成所有模块的测试脚本
- `--file <路径>`：指定 Excel 文件（默认自动在 Manual_testcase/ 下查找含 foAI 的 xlsx）
- `--case <case_id>`：仅生成指定用例（调试单条用）
- `--overwrite`：覆盖已存在的测试文件（默认跳过）
- `--debug-only`：跳过脚本生成，直接运行已有测试脚本，收集结果并回填 Excel P 列

用户传入的参数存于：$ARGUMENTS

---

## 执行流程

收到调用后，按以下步骤执行，**不要跳过任何步骤**。

### Step 0 — 参数解析

从 $ARGUMENTS 中提取：
- `TARGET_MODULES`（来自 --module 或 --all）
- `EXCEL_PATH`（来自 --file 或自动查找）
- `CASE_FILTER`（来自 --case，可选，默认为空表示不过滤）
- `OVERWRITE`（来自 --overwrite，默认 False）
- `DEBUG_ONLY`（来自 --debug-only，默认 False）

**若 `DEBUG_ONLY=True`**：完成 Step 1、Step 2 后，直接跳至 **Step 7**，跳过 Step 3–6（不生成脚本，直接运行已有脚本）。

### Step 1 — 定位 Excel 文件

解析 $ARGUMENTS：
- 若包含 `--file <路径>`，使用该路径
- 否则，用 Glob 在 `Manual_testcase/` 下查找文件名含 `foAI` 的 `.xlsx` 文件，取第一个

将找到的路径记为 `EXCEL_PATH`。

### Step 2 — 确定目标模块列表

**情况 A — 包含 `--all`：**
```bash
python auto_test_skills/testcase-analyze/excel_writer.py --list-modules "<EXCEL_PATH>"
```
将返回 JSON 中所有 `module` 字段提取为列表，设为 `TARGET_MODULES`。

**情况 B — 包含 `--module` 且后面有值：**
将 `--module` 后面所有空格分隔的词（直到下一个 `--` 参数或末尾）收集为 `TARGET_MODULES`。

**情况 C — 无参数 / `--module` 后无值：**
执行：
```bash
python auto_test_skills/testcase-analyze/excel_writer.py --list-modules "<EXCEL_PATH>"
```
遍历返回的模块列表，对每个模块读取用例数据，统计其中 `claude_color=006400` 且 `auto=是` 的用例数量，
格式化后展示给用户：
```
请指定要生成的模块（使用 /auto_testcase_generate --module <名称> 或 --all）：

  1. 用户管理          可生成：82 条（红色：40，橙色：5，跳过：22）
  2. 系统设置          可生成：0 条（无文档，全部标红）
  3. About            可生成：10 条（红色：2）
  ...
```
然后停止，等待用户重新调用。

### Step 3 — 读取参考文档

用 Glob 列出 `product_structure_testcase_regulation/` 下所有 `*_struct.md` 文件。

**第一轮：匹配专属文档（per-module struct.md）**

逐一读取每个 `*_struct.md`，判断是否覆盖 `TARGET_MODULES` 中的某个模块
（文件名或第一行标题含模块名关键词）。

暂分为：
- `HAS_DOC_MODULES`：找到专属 struct.md 的模块
- `UNMATCHED_MODULES`：暂未匹配的模块

**第二轮：检查产品级兜底文档（product-level struct.md）**

对第一轮未被匹配到任何模块的 `*_struct.md` 文件（如 `AcuHMI-1-7_struct.md`），
视为产品级兜底文档，读取其内容，检查 `UNMATCHED_MODULES` 中每个模块名是否有对应章节：
- 有对应章节 → 移入 `HAS_DOC_MODULES`，标记 `source=fallback`
- 无对应章节 → 移入 `NO_DOC_MODULES`

**若 `NO_DOC_MODULES` 不为空，输出警告（然后继续，不停止）：**
```
⚠️  以下模块缺少页面结构文档，无法生成测试脚本（已跳过）：
  - <模块名>
  请在产品级文档中补充对应模块章节，或新建专属 struct.md 后重新调用。
```

读取以下文件（后续生成时作为参考）：
- `product_structure_testcase_regulation/autotest_generativerule.md`
- `HAS_DOC_MODULES` 各模块的文档：
  - `source=per_module`：读取专属 struct.md 全文
  - `source=fallback`：从兜底文档中仅提取该模块对应的章节内容

### Step 4 — 读取绿色用例

对每个 `HAS_DOC_MODULES` 中的模块，执行：
```bash
python auto_test_skills/testcase-analyze/excel_writer.py --read "<EXCEL_PATH>" --module "<模块名>"
```

对返回的每条用例按以下规则过滤：

| 条件 | 动作 | 计入 |
|------|------|------|
| `auto != "是"` | 跳过 | 非自动化计数 |
| `claude_color != "006400"` | 跳过 | 非绿色计数 |
| `--case` 已指定 且 `case_id` 不匹配 | 跳过 | — |
| 对应文件已存在 且 `OVERWRITE=False` | 跳过 | 已存在计数 |
| 以上均不满足 | 进入 Step 5 生成 | 生成计数 |

### Step 5 — 确定输出路径

**模块目录名**（`tests/<module_dir>/`）：
- 从对应 struct.md 文件名前缀提取英文名
  - 例：`usermanagement_struct.md` → `usermanagement`

**子模块目录名**（`tests/<module_dir>/<submodule_dir>/`）：
- 读取 struct.md 中的 Tab/子模块列表（英文名）
- 将用例的 `子模块`（中文）与 struct.md 中描述的 Tab 名进行语义匹配
- 匹配成功 → 取英文 Tab 名转小写无空格（如 `User Configuration` → `userconfiguration`）
- 匹配失败 → 用 `子模块` 原值转小写无空格；若含中文则保留原值，在 Step 7 摘要中列出警告

**文件名**：`test_<case_id>.py`
- `case_id` 来自 Excel `用例编号` 列，保留原值（已含 `TestCase_AcuHMI_xxx` 格式）

完整路径示例：
```
tests/usermanagement/userconfiguration/test_TestCase_AcuHMI_007_01_case02_01.py
```

确保 `tests/<module_dir>/<submodule_dir>/__init__.py` 存在（幂等创建，已存在则跳过）。

### Step 6 — 生成测试脚本

对每条过滤后的绿色用例，生成测试文件内容，然后写入磁盘。

#### 6.1 解析 claude_note

`claude_note` 格式：`已理解：<测试点一句话摘要> | 断言方式：<如何用 Playwright 断言>`

从中提取：
- `TEST_POINT`：测试点摘要
- `ASSERT_DESC`：断言描述（指导生成 expect/assert 代码）

#### 6.2 判断是否加 @pytest.mark.skip

满足以下任一条件时，生成 skip 版（函数体仅含 `pass`）：

- 步骤或标题含「Factory Reset」、「恢复出厂」、「固件升级」、「出厂状态」、「EULA」首次弹出
- 前置条件要求特殊硬件或无法在测试环境复现的设备状态
- `semi_auto = "是"`（半自动化用例）
- 等待时间超过 30 分钟

#### 6.3 生成文件内容

**标准版模板**（可完全自动化的用例）：

```python
import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


# 用例编号：<case_id>
# 用例标题：<title>
# 预置条件：<precondition>
# 测试步骤：
#   1. <step_line_1>
#   2. <step_line_2>
#   ...（原始步骤文本，每行一条）
# 预期结果：<expected>
def test_<case_id>(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # ── 操作步骤 ────────────────────────────────────────────
    <根据 steps 生成的 Playwright 操作代码>

    # ── 断言 ────────────────────────────────────────────────
    <根据 ASSERT_DESC 和 expected 生成的断言代码>
```

**有副作用时（创建/修改数据）加 try/finally 清理**：

```python
def test_<case_id>(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    try:
        # ── 操作步骤 ──────────────────────────────────────
        <Playwright 操作代码>

        # ── 断言 ──────────────────────────────────────────
        <断言代码>

    finally:
        # ── 清理（恢复环境）──────────────────────────────
        <删除测试数据的代码，确保下次运行环境一致>
```

**skip 版模板**：

```python
import pytest
from pages.login_page import LoginPage


# 用例编号：<case_id>
# 用例标题：<title>
# 预置条件：<precondition>
# 测试步骤：（略，见 Excel）
# 预期结果：<expected>
@pytest.mark.skip(reason="<跳过原因，如：涉及恢复出厂/固件升级，需手动执行>")
def test_<case_id>(login_page: LoginPage):
    pass
```

**步骤无法完全转换时（部分 TODO）**：

```python
    # ── 操作步骤 ──────────────────────────────────────────
    <能转换的步骤正常生成>
    # TODO: 步骤 N — <原始步骤文本>（UI 元素在 struct.md 中未描述，需手动补充定位）
```

#### 6.4 代码生成规范（来自 autotest_generativerule.md）

**定位优先级（高 → 低）：**
1. `page.get_by_role("button"/"textbox"/"menuitem", name="...")`
2. `page.get_by_label("...", exact=True)`
3. `page.locator("#id")`
4. `page.get_by_text("...", exact=True)`
5. `page.get_by_placeholder("...")`
6. 简单 CSS：`page.locator("input[name='...']")`

**禁止使用：** XPath、复杂 CSS 多级选择器、`.nth(N)` 硬编码索引

**ElementUI 特殊组件：**
```python
# Dropdown（el-select，非原生 select）
page.get_by_text("--Select Role--", exact=True).click()
page.get_by_role("option", name="admin").click()

# Checkbox（el-checkbox）
page.locator(".el-form-item").filter(has_text="Multiple Login").locator(".el-checkbox__inner").click()
```

**导航后等待：**
```python
page.wait_for_load_state("networkidle")
page.wait_for_timeout(500)   # ElementUI 渲染缓冲
```

**弹框可选处理：**
```python
try:
    page.get_by_role("button", name="Cancel").click(timeout=3000)
except Exception:
    pass
```

**跨上下文登录验证（验证密码修改后能否登录）：**
```python
ctx = login_page.page.context.browser.new_context(ignore_https_errors=True)
p = ctx.new_page()
p.goto(BASE_URL + "/#/login")
# ... 执行登录操作
ctx.close()
```

#### 6.5 写入文件

每生成一个文件立即写入磁盘，不等所有用例处理完。

### Step 7 — 确定测试目录并运行首次测试

#### 7.1 确定测试目录

对 `TARGET_MODULES` 中每个模块，推断对应 `tests/` 子目录：

- 用 Glob 列出 `product_structure_testcase_regulation/` 下所有 `*_struct.md`，取文件名前缀作为目录名
  （如 `systemsettings_struct.md` → `tests/systemsettings/`）
- 若无专属 struct.md，则在 `tests/` 下按模块名关键词查找现有目录（忽略大小写、忽略空格）
- 若仍找不到，输出警告并跳过该模块

将所有找到的目录合并为 `TEST_DIRS`。

#### 7.2 首次运行 pytest

```bash
python -m pytest <TEST_DIRS> -v --tb=short
```

捕获完整输出，按文件路径记录每条测试的状态：
- `PASSED` / `XPASSED` → 标记为已通过
- `SKIPPED` → 标记为已跳过
- `FAILED` / `XFAILED` → 加入 `FAILED_LIST`（待自动调试）

**立即将已通过和已跳过的用例写入 P 列（不等调试完成）：**

⚠️ **必须用 `--results-file` 方式，禁止用 `--results` 内联 JSON**（Windows shell 会将 UTF-8 中文用 GBK 重编码导致乱码）：

1. 用 Write 工具将结果 JSON 写入 `scripts/results_p_pass.json`（UTF-8）
   ⚠️ **JSON 每条记录必须包含 `"value"` 键**（不是 `p_value`，不是 `result`），例如：
   ```json
   [{"case_id": "TestCase_AcuHMI_007_01_case01", "value": "是"},
    {"case_id": "TestCase_AcuHMI_007_01_case02", "value": "跳过 | 涉及恢复出厂"}]
   ```
2. 再调用：
```bash
python auto_test_skills/testcase-analyze/excel_writer.py \
  --write-p "<EXCEL_PATH>" --results-file scripts/results_p_pass.json
```

输出首次运行汇总：
```
首次运行：通过 33 / 跳过 2 / 失败 12 — 开始自动调试失败项...
```

---

### Step 8 — 自动调试循环

对 `FAILED_LIST` 中每条失败测试，执行自动调试循环（最多 **3 轮**）。

每轮流程：
1. 读取失败测试的完整错误信息（来自 pytest `--tb=short` 输出）
2. 读取测试文件内容
3. 按下方**错误模式识别表**匹配错误，选择修复策略
4. 修改测试文件（Edit 工具）
5. 单独重新运行该测试：
   ```bash
   python -m pytest <test_file_path> --tb=short -q
   ```
6. 若通过 → 标记"是"，停止本条迭代
7. 若仍失败 → 轮次 +1，继续下一轮
8. 3 轮后仍失败 → 标记"调试失败"，记录最终错误原因

#### 8.1 错误模式识别表

| 错误特征 | 判断 | 修复策略 |
|---------|------|---------|
| `TimeoutError: ...locator("a").first` | VD 列表行无 `<a>` 链接 | 改为 `locator("td").first.click()`；若进入详情后找不到表单，在点击后加 `.el-tabs__item` 中 Configuration tab 的点击 |
| `StrictModeViolation: ...resolved to N elements` | 定位到多个元素 | 在定位器末尾加 `.first`；或用 `.filter(has_text="...")` 缩小范围 |
| `TimeoutError: get_by_role("button", name="Yes, continue")` | 确认按钮文字不同 | 依次尝试 `"Yes"` → `"Yes,continue"`（无空格）→ `"Confirm"` |
| `TimeoutError: get_by_role("button", name="Confirm")` 在 popconfirm 中 | popconfirm 使用 primary button | 改为 `page.locator(".el-popconfirm__action .el-button--primary").click()` |
| `TimeoutError: get_by_placeholder("X")` | placeholder 文字与实际不符 | 读取页面 HTML（`page.content()` 或截图），找到实际 placeholder 后修正 |
| `AssertionError: locator(".el-form-item__error").count() == 0` | 产品未做该输入校验 | 加 `@pytest.mark.xfail(strict=False, reason="产品未校验该字段")` |
| `AssertionError: 元素仍然存在` / 删除后仍显示 | 删除后未等待刷新 | 在删除确认后加 `page.wait_for_timeout(1000)` 再查询 |
| `AssertionError: 结果应为空，当前显示 N 行` | 产品不过滤该条件 | 加 `@pytest.mark.xfail(strict=False, reason="产品不过滤该参数边界值")` |
| `TimeoutError: ...filter(has_text="Enable")` | 功能开关名称不同 | 截图确认按钮实际文字；或改为 `locator(".el-form-item").filter(has_text="...").locator(".el-radio").filter(has_text="Enable").click()` |
| `Error: element(s) not found` 在 `expect(...).to_have_value()` | 目标字段在 tab 内未激活 | 在操作前加点击对应 tab 的步骤 |
| `TimeoutError` 出现在功能入口点击 | 导航路径不正确 | 读取 struct.md 重新确认导航步骤；补充 `wait_for_load_state("networkidle")` |
| 重复添加导致「已达上限」 | 前置数据残留 | 在 try 块前加清理步骤；或先统计已有条目数再计算要添加的数量 |
| `ImportError` / `ModuleNotFoundError` | 依赖缺失或路径错误 | 检查 import 语句，与其他同模块测试文件对齐 |

#### 8.2 无法自动修复的情形

满足以下任一条件时，**停止尝试，直接标记"调试失败"并说明原因**：

- 错误涉及 Factory Reset、固件升级、设备重启等破坏性操作
- 需要真实硬件或外部设备才能复现
- 3 轮修复后错误类型未改变（说明根本原因未被识别）
- 修复方向不明确（错误信息过于通用，需人工排查）

此类用例可在测试文件中加 `@pytest.mark.skip(reason="...")` 后标记"跳过"，或保持"调试失败"。

**标记"跳过"或"调试失败"时必须写明 reason：**
- skip：`@pytest.mark.skip(reason="<说明为何无法自动化>")`
- xfail：`@pytest.mark.xfail(strict=False, reason="<产品行为说明>")`
- reason 是 P 列原因的唯一来源，**缺少 reason 则 P 列只写状态不写原因**

#### 8.3 每轮修复后输出进度

```
[自动调试] test_TestCase_AcuHMI_005_02_case02_02 — 第 1 轮
  错误：TimeoutError on get_by_placeholder("Enter IP").nth(1)
  识别：Ethernet 2 在 Auto 模式下 IP 字段被禁用
  修复：在填写 IP 前加切换 Manual 步骤
  重跑：PASSED ✓ → 标记"是"

[自动调试] test_TestCase_AcuHMI_005_04_case06_2 — 第 1 轮
  错误：AssertionError: el-form-item__error count == 0（期望校验空密码失败）
  识别：产品未校验空密码
  修复：加 @pytest.mark.xfail(strict=False)
  重跑：XFAILED → 标记"调试失败"（产品行为与规格不符）
```

---

### Step 9 — 写入剩余 P 列结果

调试循环结束后，将调试阶段产生的结果（"是" / "调试失败" / "跳过"）写入 P 列。

**case_id 格式对齐**：Excel 用例编号统一以 `TestCase_` 开头，写入前自动补全缺少前缀的 case_id：
```
ACUREV4100WEB2_VD_002_001  →  TestCase_ACUREV4100WEB2_VD_002_001
```

**写入步骤（固定流程，不得变更）：**

1. 用 **Write 工具**将结果 JSON 写入 `scripts/results_p.json`（Write 工具直接写 UTF-8，不经过 shell）
   ⚠️ **JSON 每条记录必须包含 `"value"` 键**（不是 `p_value`，不是 `result`），例如：
   ```json
   [{"case_id": "TestCase_AcuHMI_007_01_case01", "value": "是"},
    {"case_id": "TestCase_AcuHMI_007_01_case02", "value": "调试失败 | <原因>"}]
   ```
2. 调用命令时使用 `--results-file`，**禁止使用 `--results` 内联 JSON**：

```bash
python auto_test_skills/testcase-analyze/excel_writer.py \
  --write-p "<EXCEL_PATH>" --results-file scripts/results_p.json
```

> **原因**：Windows shell（bash/cmd）在 `$()` 命令替换或参数传递时，会将 UTF-8 中文用 GBK 重编码，
> openpyxl 收到的是 GBK 字节流，写入 Excel 后产生乱码（如 `是` → `ÊÇ`）。
> `--results-file` 由 Python 直接以 UTF-8 读文件，完全绕开 shell 编码问题。

若 `not_found` 不为空，逐一列出并说明可能原因。

---

### Step 10 — 输出最终摘要

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
auto_testcase_generate 完成
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【生成阶段】
  新生成：45 个脚本
  已存在（跳过）：7 个

【运行 & 调试阶段】
模块            首次通过  调试修复  调试失败  跳过   合计
─────────────────────────────────────────────────
用户管理           30       8        4       3     45
系统设置           20       5        2       2     29
─────────────────────────────────────────────────
合计               50      13        6       5     74

【P 列写入】已写入 80 行

调试失败（需人工介入）：
  TestCase_AcuHMI_005_04_case06_2  — 产品未校验空密码（xfail）
  TestCase_AcuHMI_009_05_case01_13 — 产品不过滤超范围 slaveid（xfail）
  ...

跳过：
  TestCase_AcuHMI_005_04_case06    — Email 密码显示按钮未实现（skip）
  ...
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

---

## 注意事项

- **默认模式**（不含 `--debug-only`）：生成脚本 → 立即运行 → 自动调试 → 回填 P 列，全程无需人工干预
- **`--debug-only` 模式**：跳过生成，直接对已有脚本运行 → 自动调试 → 回填 P 列（用于重新调试）
- 已存在的测试文件默认跳过生成（保护手动修改），加 `--overwrite` 强制重新生成
- 无对应 struct.md 的模块跳过生成，不中断其他模块处理
- 每生成一个文件立即写入磁盘，中途中断已生成文件不丢失
- 自动调试每条失败用例最多尝试 **3 轮**，超出后标记"调试失败"
- 此 Skill 在 `/testcase-analyze` 标注绿色后使用；两者通过 Excel 数据松耦合，无代码依赖
- P 列写入值规范：
  - `是` — 调试通过
  - `跳过 | <原因>` — 有 `@pytest.mark.skip`，原因取 `reason=` 参数值（最长 80 字符）
  - `调试失败 | <原因>` — xfail 失败或 3 轮仍失败，原因取 `reason=` 参数值或断言错误首行
- 已通过和跳过的用例在首次运行后**立即写入** P 列，不等调试循环结束
- 原因提取需处理 reason 字符串内含单引号的情况（如 `"System shows a 'reboot' dialog"`）
