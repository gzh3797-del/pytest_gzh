# auto_testcase_generate — 自动化测试脚本生成 Skill

## 用途
读取 Excel 中已被 `/testcase-analyze` 标注为深绿色（已理解）的用例，
结合模块的 `*_struct.md` 页面结构文档，为每条用例生成独立的 Playwright + pytest 测试脚本。

**前提**：需先运行 `/testcase-analyze`，Excel 第 13 列有绿色标注后再调用本 Skill。

| Skill | 阶段 | 频率 |
|-------|------|------|
| `/testcase_auto_init` | 项目启动，生成框架骨架 | 一次性 |
| `/testcase-analyze` | 分析手工用例，标注绿/红/橙 | 反复使用 |
| `/auto_testcase_generate` | 读取绿色用例，生成测试脚本 | 反复使用 |

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

### Step 7 — 输出摘要

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
auto_testcase_generate 完成
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

模块           新生成  已存在(跳过)  非绿色  非自动化  无文档(跳过)
──────────────────────────────────────────────────────────
用户管理          45        0          40       22         0
系统设置           0        0         100        0       100
──────────────────────────────────────────────────────────
合计              45        0         140       22       100

新生成文件：
  tests/usermanagement/general/test_TestCase_AcuHMI_007_01_case01_1.py
  tests/usermanagement/userconfiguration/test_TestCase_AcuHMI_007_01_case02_01.py
  ...（每条一行）

⚠️  子模块目录名自动推断（请确认是否正确）：
  - 「密码策略」→ passwordpolicy
  - 「角色配置」→ roleconfiguration

下一步：
  pytest tests/usermanagement/ -v          # 运行该模块所有生成用例
  pytest tests/ -m smoke -v               # 运行 smoke 测试
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

---

## 注意事项

- 本 Skill 只读取 Excel（不写入任何列），生成的脚本写入 `tests/` 目录
- 已存在的测试文件默认跳过（保护手动修改），加 `--overwrite` 强制覆盖
- 无对应 struct.md 的模块跳过，不中断其他模块处理
- 每生成一个文件立即写入，中断后已生成文件不丢失
- 此 Skill 在 `/testcase-analyze` 标注绿色后使用；两者通过 Excel 数据松耦合，无代码依赖
