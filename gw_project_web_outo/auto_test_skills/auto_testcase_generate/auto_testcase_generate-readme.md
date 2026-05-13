# auto_testcase_generate Skill 使用说明

## 功能简介

`auto_testcase_generate` 是一个 Claude Code 自定义 Skill，
读取 `/testcase-analyze` 已标注为深绿色（已理解）的手工用例，
结合模块页面结构文档（struct.md），自动生成 Playwright + pytest 测试脚本。

每条用例生成一个独立的 `test_<case_id>.py` 文件，遵循项目已有的代码规范。

### 工作流中的位置

```
【分析阶段】/testcase-analyze --all
    ↓ Excel 第 21 列标注绿/红/橙
【生成阶段】/auto_testcase_generate --module 用户管理
    ↓ 读取绿色用例 → 生成 tests/<module>/<submodule>/test_<case_id>.py
【执行阶段】pytest tests/usermanagement/ -v
```

| Skill | 阶段 | 频率 |
|-------|------|------|
| `/testcase_auto_init` | 项目启动，生成框架骨架 | 一次性 |
| `/testcase-analyze` | 分析手工用例，标注绿/红/橙 | 反复使用 |
| `/auto_testcase_generate` | 读取绿色用例，生成测试脚本 | 反复使用 |

### 与 testcase-analyze 的耦合关系

**松耦合，通过 Excel 数据间接关联，无代码层面依赖。**

- `testcase-analyze` 写入 Excel 第 21 列（`claude_color=006400` 表示已理解）
- `auto_testcase_generate` 读取该列过滤绿色用例
- 两者共享 `excel_writer.py` 工具，但可以独立部署和调用

---

## 目录结构

```
auto_test_skills/auto_testcase_generate/
├── auto_testcase_generate.md          # Skill 主体内容（执行流程）
└── auto_testcase_generate-readme.md   # 本说明文档

.claude/commands/
└── auto_testcase_generate.md          # Skill 调用入口（内容与上面相同）
```

---

## 一、部署方式

### 方式 A：通过 Git（推荐）
```bash
git pull
```

确认以下文件存在：
```
.claude\commands\auto_testcase_generate.md
auto_test_skills\auto_testcase_generate\（2 个文件）
```

### 方式 B：手动复制
```
.claude\commands\auto_testcase_generate.md        # 必须
auto_test_skills\testcase-analyze\excel_writer.py  # 依赖（通常已有）
```

---

## 二、调用方式

> **重要**：必须先运行 `/testcase-analyze` 完成用例分析（Excel 第 21 列有绿色标注），
> 再调用本 Skill 生成脚本。

### 参数说明

| 参数 | 说明 | 默认值 |
|------|------|-------|
| `--module <名称> [名称...]` | 指定模块 | — |
| `--all` | 处理所有模块 | — |
| `--file <路径>` | 指定 Excel 文件 | 自动查找 Manual_testcase/ 下含 foAI 的 xlsx |
| `--case <case_id>` | 仅生成指定用例（调试用） | — |
| `--overwrite` | 覆盖已存在的测试文件 | 跳过已存在文件 |

### 典型调用场景

**场景 1：生成单个模块**
```
/auto_testcase_generate --module 用户管理
```

**场景 2：生成多个模块**
```
/auto_testcase_generate --module 用户管理 系统设置
```

**场景 3：生成全部模块**
```
/auto_testcase_generate --all
```

**场景 4：调试单条用例**
```
/auto_testcase_generate --module 用户管理 --case case02_01
```

**场景 5：重新生成（覆盖已有文件）**
```
/auto_testcase_generate --module 用户管理 --overwrite
```

**场景 6：无参数（查看各模块可生成用例数）**
```
/auto_testcase_generate
```

---

## 三、生成文件结构

每条绿色用例生成一个独立测试文件：

```
tests/
└── <module_dir>/                         # 来自 struct.md 文件名前缀
    └── <submodule_dir>/                  # 来自 struct.md Tab 英文名
        ├── __init__.py                   # 自动创建（若不存在）
        └── test_<case_id>.py             # 每条用例一个文件
```

示例：
```
tests/
└── usermanagement/
    ├── general/
    │   └── test_TestCase_AcuHMI_007_01_case01_1.py
    ├── userconfiguration/
    │   ├── test_TestCase_AcuHMI_007_01_case02_01.py
    │   └── test_TestCase_AcuHMI_007_01_case02_02.py
    └── passwordpolicy/
        └── test_TestCase_AcuHMI_007_03_case01_1.py
```

---

## 四、生成脚本格式

### 标准版（可完全自动化）

```python
import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_007_01_case02_01
# 用例标题：添加用户 - 正常流程
# 预置条件：设备可访问，admin 密码为 Admin@110001
# 测试步骤：
#   1. 进入 User Management > User Configuration
#   2. 点击 Add User 按钮
#   3. 填写 Username、Password、Repeat Password
#   4. 选择 Role
#   5. 点击 Save
# 预期结果：新用户出现在用户列表中
def test_TestCase_AcuHMI_007_01_case02_01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    try:
        # ── 操作步骤 ──────────────────────────────────────
        page.get_by_text("User Management").first.click()
        page.wait_for_load_state("networkidle")
        page.get_by_role("menuitem", name="User Configuration").click()
        page.wait_for_load_state("networkidle")

        page.get_by_role("button", name="Add User").click()
        page.get_by_label("Username", exact=True).fill("testuser01")
        page.get_by_label("Password", exact=True).fill("Test@110001")
        page.get_by_label("Repeat Password", exact=True).fill("Test@110001")
        page.get_by_text("--Select Role--", exact=True).click()
        page.get_by_role("option", name="viewer").click()
        page.get_by_role("button", name="Save").click()
        page.wait_for_load_state("networkidle")

        # ── 断言 ──────────────────────────────────────────
        expect(page.locator("tbody").get_by_role("row").filter(has_text="testuser01")).to_be_visible()

    finally:
        # ── 清理（删除测试用户）──────────────────────────
        row = page.locator("tbody").get_by_role("row").filter(has_text="testuser01")
        if row.is_visible():
            row.get_by_role("button").last.click()
            page.get_by_role("button", name="Yes, continue").click()
            page.wait_for_load_state("networkidle")
```

### Skip 版（需人工干预）

```python
import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_007_05_case01_1
# 用例标题：恢复出厂设置后验证默认密码
# 预置条件：设备可访问
# 预期结果：恢复出厂后密码恢复为 Admin@110001
@pytest.mark.skip(reason="涉及恢复出厂，需手动执行")
def test_TestCase_AcuHMI_007_05_case01_1(login_page: LoginPage):
    pass
```

### 部分 TODO 版（步骤无法完全转换）

```python
    # ── 操作步骤 ──────────────────────────────────────────
    page.get_by_role("button", name="Add User").click()
    page.get_by_label("Username", exact=True).fill("testuser")
    # TODO: 步骤 3 — 设置过期时间（Expiration Date 控件在 struct.md 中未描述，需手动补充定位）
    page.get_by_role("button", name="Save").click()
```

---

## 五、幂等性说明

| 情况 | 默认行为 | `--overwrite` 行为 |
|------|---------|------------------|
| 文件不存在 | 正常生成 | 正常生成 |
| 文件已存在 | **跳过**（保留手动修改） | 覆盖重新生成 |
| 用例非绿色（红/橙/空） | 跳过 | 跳过（颜色不符） |
| 用例 `auto != "是"` | 跳过 | 跳过 |
| 模块无 struct.md | 跳过（输出警告） | 跳过（同上） |

---

## 六、常见问题

**Q：调用后输出「可生成 0 条」？**
A：有两种可能：
- 该模块尚未运行 `/testcase-analyze`，Excel 第 21 列无绿色标注
- 所有绿色用例的文件已存在且未加 `--overwrite`

**Q：子模块目录名显示警告？**
A：当 Excel 子模块列的中文值无法与 struct.md 的 Tab 英文名匹配时，Skill 会自动推断并在摘要中列出，
请人工确认目录名是否正确。如需修改，重命名对应目录后用 `--overwrite` 重新生成即可。

**Q：生成的文件运行报错？**
A：常见原因：
1. 登录选择器与实际 UI 不符 → 检查 `pages/login_page.py`
2. 导航路径变化 → 对照 struct.md 检查 menuitem name
3. 等待时间不足 → 在相应步骤后增加 `page.wait_for_timeout(1000)`

**Q：如何只重新生成某条用例？**
A：
```
/auto_testcase_generate --module 用户管理 --case TestCase_AcuHMI_007_01_case02_01 --overwrite
```

**Q：模块无 struct.md 能生成吗？**
A：不能。Skill 需要 struct.md 提供 UI 元素定位方式和导航路径，
缺少文档时无法生成有效的 Playwright 代码。请先补充 struct.md 再调用。

---

## 七、相关文件

| 文件 | 说明 |
|------|------|
| `auto_test_skills/testcase-analyze/excel_writer.py` | 读取 Excel 用例数据（`--read --module`） |
| `product_structure_testcase_regulation/autotest_generativerule.md` | 自动化代码规范 |
| `product_structure_testcase_regulation/*_struct.md` | 各模块页面结构文档 |
| `pages/login_page.py` | 登录 Page Object（生成脚本的基础 fixture） |
| `conftest.py` | `login_page` fixture 定义 |
