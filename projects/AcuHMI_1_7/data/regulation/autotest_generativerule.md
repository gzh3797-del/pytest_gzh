# AcuHMI-1-7 User Management 自动化用例规则

> 本文档覆盖用例目录规划、代码生成规则、元素定位规范、特殊场景处理及已验证调试经验。  
> 产品结构（UI 字段、URL、权限说明等）见：[usermanagement_struct.md](usermanagement_struct.md)  
> 调试脚本存放目录：`../claude_debug/`；产品探索截图存放路径：`explore/`（与本文件同目录）

---

## 一、用例目录规划

```
tests/
└── usermanagement/
    ├── general/
    │   └── test_session_timeout.py
    ├── userconfiguration/
    │   ├── test_add_user.py
    │   ├── test_edit_user.py
    │   ├── test_delete_user.py
    │   └── test_lock_user.py
    ├── roleconfiguration/
    │   ├── test_add_role.py
    │   ├── test_edit_role.py
    │   └── test_delete_role.py
    ├── passwordpolicy/
    │   ├── test_complexity.py
    │   ├── test_password_expiry.py
    │   └── test_failed_login.py
    ├── passwordmanagement/
    │   └── test_change_password.py
    ├── passwordchange/
    │   ├── test_TestCase_ARM_XXL_002_04_case1_04.py
    │   ├── test_TestCase_ARM_XXL_002_04_case1_05.py
    │   └── test_TestCase_ARM_XXL_002_04_case1_06.py
    ├── passwordconfiguration/
    │   ├── test_TestCase_AcuHMI_007_04_case01.py
    │   └── test_TestCase_AcuHMI_007_04_case01_1.py
    ├── passwordreset/
    │   └── test_TestCase_ARM_XXL_002_04_case12_02.py
    ├── eula/
    │   └── test_TestCase_AcuHMI_007_01_case06_6.py
    └── defaultpasswordlogin/
        └── test_default_password_login.py
```

---

## 二、自动化用例生成规则

### 2.1 目录与文件命名规则

| Excel 列 | 对应自动化结构 | 示例 |
|----------|---------------|------|
| **模块** | 一级目录（文件夹名） | `usermanagement` |
| **子模块** | 二级目录（文件夹名） | `defaultpasswordlogin` |
| **用例名称** | Python 文件名，格式 `test_用例编号.py` | `test_TestCase_AcuHMI_007_01_case06_9.py` |

完整路径示例：
```
tests/usermanagement/defaultpasswordlogin/test_TestCase_AcuHMI_007_01_case06_9.py
```

### 2.2 用例文件内容模板

每条自动化用例文件须包含以下中文注释头，紧跟测试函数：

```python
# 用例编号：TestCase_AcuHMI_007_01_case06_9
# 用例标题：<从 Excel "用例标题" 列填入>
# 预置条件：<从 Excel "预置条件" 列填入>
# 测试步骤：
#   1. <步骤1>
#   2. <步骤2>
#   ...
# 预期结果：<从 Excel "预期结果" 列填入>
def test_TestCase_AcuHMI_007_01_case06_9(login_page):
    ...
```

### 2.3 完整用例文件示例

```python
import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_007_01_case06_9
# 用例标题：使用默认密码登录后弹出修改密码提示，点击取消可正常进入系统
# 预置条件：设备可访问，admin 账号密码为默认密码 Admin@110001
# 测试步骤：
#   1. 打开登录页 https://192.168.2.199/#/login
#   2. 输入用户名 admin，密码 Admin@110001，点击 Sign In
#   3. 登录成功后，弹出"修改默认密码"提示框
#   4. 点击 Cancel 按钮
# 预期结果：弹框关闭，页面正常进入 User Management 主页，AcuHMI 导航菜单可见
def test_TestCase_AcuHMI_007_01_case06_9(login_page: LoginPage):
    login_page.open()
    login_page.login()
    assert login_page.is_logged_in(), "登录失败：AcuHMI 菜单未出现"
```

### 2.4 页面元素定位规则

#### 优先级（从高到低）

| 优先级 | 定位方式 | 示例 | 说明 |
|--------|---------|------|------|
| ★★★★★ | `get_by_role` + name | `page.get_by_role("button", name="Save")` | 最稳定，语义化，推荐首选 |
| ★★★★★ | `get_by_label` | `page.get_by_label("Username")` | 表单元素首选 |
| ★★★★☆ | `id` 属性 | `page.locator("#userId")` | 唯一 id 时极稳定 |
| ★★★★☆ | `data-*` 自定义属性 | `page.locator("[data-testid='add-user-btn']")` | 前端专为测试预留时使用 |
| ★★★☆☆ | `get_by_text` | `page.get_by_text("User Configuration", exact=True)` | 文本稳定时可用 |
| ★★★☆☆ | `get_by_placeholder` | `page.get_by_placeholder("Enter User Name")` | 输入框 placeholder 稳定时可用 |
| ★★☆☆☆ | 简单 CSS（单层） | `page.locator("input[name='username']")` | 仅允许单层属性选择器 |
| ✗ 禁止 | CSS 层级复杂选择器 | `.el-icon.el-select__caret > svg` | **禁止使用**，易随版本失效 |
| ✗ 禁止 | XPath 绝对路径 | `//div[1]/ul/li[2]/span` | **禁止使用**，DOM 变动即失效 |
| ✗ 避免 | class 定位 | `page.locator(".el-button--primary")` | 避免使用，UI 改版易失效 |

#### 使用原则

1. **唯一性优先**：使用定位器前须人工确认该定位器在当前页面唯一，验证过程不体现在用例代码中。
2. **组合缩窄范围**：单一定位器不唯一时，允许通过父容器缩窄范围，但层级不超过 2 层。
   ```python
   # 允许：父容器 + role/name 组合
   page.get_by_role("dialog").get_by_role("button", name="Save")
   # 允许：表格行内定位
   page.locator("tbody").get_by_role("row").filter(has_text="admin")
   ```
3. **filter 辅助**：多个同名元素时用 `.filter(has_text=...)` 区分，不用 `.nth()` 硬编码序号。
   ```python
   # 推荐
   page.get_by_role("button").filter(has_text="Delete")
   # 避免
   page.get_by_role("button").nth(3)
   ```
4. **稳定性验证**：定位器写入用例前，在 Playwright Inspector 或 codegen 中验证其唯一性和跨刷新稳定性。

### 2.5 规则约束

1. **注释必填**：每条用例文件必须包含"用例编号、用例标题、预置条件、测试步骤、预期结果"五项中文注释。
2. **函数名唯一**：函数名 = `test_` + 用例编号，不同模块下编号不能重复。
3. **fixture 复用**：登录统一使用 `login_page` fixture，不在用例内手动创建浏览器。
4. **断言中文化**：`assert` / `expect` 的失败提示信息使用中文，便于报告阅读。
5. **目录结构**：模块、子模块目录下须有 `__init__.py`（空文件），pytest 才能正确收集。
6. **清理原则**：测试中若新增数据（用户、角色等），用例结束前须清理，保证环境幂等。

### 2.6 特殊用例处理规则

| 场景 | 处理方式 |
|------|----------|
| 涉及**恢复出厂设置**的用例 | 生成完整代码，但在函数前加 `@pytest.mark.skip(reason="涉及恢复出厂，需手动执行")`，不纳入 CI 自动运行 |
| **不纳入自动化**的用例 | 生成注释文件（仅保留注释头 + `pass`），标注 `@pytest.mark.manual` |
| **时间依赖**（等待 1 分钟）| 用 `page.wait_for_timeout(61000)` 实现 |
| **时间依赖**（等待 30 分钟以上）| 不纳入自动化 |
| **外部工具依赖**（临时密码生成）| 不纳入自动化 |
| **密码值未明确**的用例 | 密码符合策略规则即可随机生成，在自动化注释中备注生成逻辑 |

---

## 三、自动化调试经验（已验证规则）

> 本节记录调试 3 条已通过用例期间发现的坑点与修正方案，编写后续用例时直接遵循。

### 3.1 SPA 导航：禁止 `page.goto()` 直接跳转子页面

**问题**：直接调用 `page.goto("/#/userManagement/passwordManagement")` 会被 SPA 重定向到 System Settings 主页，无法到达目标页面。

**根因**：SPA 需要先在顶部 header 建立设备上下文，直接 `goto` 不会触发该流程。

**正确做法**：封装 `_nav_to_submenu(page, submenu)` 辅助函数，通过 UI 点击导航：

```python
def _nav_to_submenu(page, submenu: str):
    if "/userManagement/" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_text("User Management").first.click()
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
```

**适用子菜单名称**：`"User Configuration"`、`"Role Configuration"`、`"Password Policy"`、`"Password Management"`

---

### 3.2 Add User 表单：非标准 dialog，用 `get_by_label` 定位

**问题**：Add User 表单不是标准 ARIA `role="dialog"` 元素（是 drawer/panel），`page.get_by_role("dialog")` 会超时或定位到错误容器。

**正确做法**：直接在 page 层级用 `get_by_label` 定位表单字段，无需套 dialog 容器：

```python
page.get_by_role("button", name="Add User").click()
page.wait_for_timeout(1000)
page.get_by_label("Username", exact=True).fill(username)
page.get_by_label("Password", exact=True).fill(password)
page.get_by_label("Repeat Password", exact=True).fill(password)
```

---

### 3.3 角色下拉（Element UI el-select）：点击文本而非 combobox

**问题**：`page.get_by_role("combobox").click()` 会被 `<span>--Select Role--</span>` 拦截指针事件，导致下拉不展开。

**正确做法**：直接点击占位文本，再选择 option：

```python
page.get_by_text("--Select Role--", exact=True).click()
page.get_by_role("option", name="view").click()
```

---

### 3.4 Password Management 编辑表单：实际字段标签名

**问题**：测试用例文档中写的字段名为 "New Password" / "Confirm Password"，但页面实际标签为 **"Password"** / **"Repeat Password"**，导致 `get_by_label` 定位失败。

**正确字段标签**：

| 文档中的名称 | 页面实际标签（用于定位） |
|-------------|------------------------|
| New Password | `Password` |
| Confirm Password | `Repeat Password` |
| Current Password | **不在编辑表单中**（见 §3.9） |

**正确做法**：

```python
page.get_by_label("Password", exact=True).fill(new_pwd)
page.get_by_label("Repeat Password", exact=True).fill(new_pwd)
```

> `exact=True` 必须加，否则 `get_by_label("Password")` 会同时匹配到 "Repeat Password"。

---

### 3.5 跨用户登录验证：独立 browser context

验证改密后新密码能否登录，必须用独立 context（不能复用当前 admin session）：

```python
def _can_login(browser, username: str, password: str) -> bool:
    ctx = browser.new_context(ignore_https_errors=True)
    p = ctx.new_page()
    try:
        p.goto(BASE_URL + "/#/login")
        p.wait_for_load_state("networkidle")
        p.get_by_role("textbox", name="Enter User Name").fill(username)
        p.get_by_role("textbox", name="Enter Password").fill(password)
        p.get_by_role("button", name="Sign In").click()
        p.wait_for_load_state("networkidle")
        # 处理 EULA（新用户首次登录）
        for btn in ["Accept", "I Accept", "确认"]:
            try:
                p.get_by_role("button", name=btn).click(timeout=2000)
                p.wait_for_load_state("networkidle")
            except Exception:
                pass
        # 处理默认密码提示
        try:
            p.get_by_role("button", name="Cancel").click(timeout=2000)
        except Exception:
            pass
        return "/#/login" not in p.url
    finally:
        ctx.close()
```

通过 `page.context.browser` 获取 browser 对象：`browser = page.context.browser`

---

### 3.6 成功提示断言

密码修改成功后页面弹出 Toast，文本匹配用 `exact=False`（不区分大小写前缀）：

```python
expect(page.get_by_text("password changed", exact=False)).to_be_visible(timeout=5000)
```

---

### 3.7 openpyxl 保存含中文路径的 Excel 文件

**问题**：`wb.save(path)` 在路径含中文时触发 `PermissionError`（Windows ZipFile 编码问题）。

**正确做法**（BytesIO 绕过）：

```python
import io
buf = io.BytesIO()
wb.save(buf)
with open(path, 'wb') as f:
    f.write(buf.getvalue())
```

---

### 3.8 调试脚本管理规则

- 临时调试脚本**统一放入 `../claude_debug/`**（项目根下的 `claude_debug/`），不放项目根目录
- 脚本内引用 Excel 等文件时，**用 `Path(__file__).parent.parent / ...`** 而非硬编码相对路径，确保从任何目录运行均可定位文件
- `claude_debug/` 目录不纳入测试收集（`pytest.ini` 中 `testpaths` 未包含该目录）

---

### 3.9 非 admin 用户修改密码：Current User Password 弹框

**问题**：非 admin 用户（拥有 User=edit 权限的角色）在 Password Management 点击编辑按钮后，会弹出独立的"Current User Password"弹框拦截，无法直接点击编辑表单内的 Save 按钮。

**根因**：弹框是一个 `role="dialog"` 的 el-overlay，覆盖在编辑表单之上，拦截所有鼠标事件。弹框要求输入**当前登录用户的密码**（与被修改的目标用户无关）。

**正确做法**：点击编辑按钮后，先处理弹框，再操作表单：

```python
row.get_by_role("button").click()
page.wait_for_timeout(500)
# 填入 Current User Password 弹框
page.get_by_placeholder("Please input").fill(current_password)
page.get_by_role("button", name="Confirm").click()
page.wait_for_timeout(500)
# 弹框消失后正常操作编辑表单
page.get_by_label("Password", exact=True).fill(new_pwd)
page.get_by_label("Repeat Password", exact=True).fill(new_pwd)
page.get_by_role("button", name="Save").click()
```

**适用场景**：
- 非 admin 用户修改自己的密码
- 非 admin 用户修改其他（非 admin）用户的密码
- admin 用户**不触发**此弹框

---

### 3.10 Password Management 权限说明

| 角色 | 是否有编辑按钮 | Current User Password 弹框 |
|------|--------------|--------------------------|
| admin | 有（除 admin 自身行保留按钮） | **不弹出** |
| User=edit 自定义角色 | 有（admin 行无按钮） | **弹出** |
| view 内置角色 | **无按钮**，显示 "no operation permission" 提示 | 不适用 |

**内置角色只有两种**：`admin`（User=edit）和 `view`（User=view）。  
如需测试非 admin 的编辑操作，须在 Role Configuration 新建 User=edit 的自定义角色。

---

### 3.11 EULA 弹框按钮

首次登录触发的 EULA 弹框按钮：

| 按钮 | 行为 |
|------|------|
| `Accept` | 接受协议，允许进入系统 |
| `Close` | 不接受，关闭弹框，停留在登录页 |

- **所有新建用户首次登录**都会触发 EULA
- 点击 Close 后再次登录，EULA 仍会出现（需再次选择）
- 点击 Accept 后不再出现

```python
# 接受 EULA
page.get_by_role("button", name="Accept").click()
# 不接受 EULA（停留登录页）
page.get_by_role("button", name="Close").click()
```

---

### 3.12 Forgot Password 功能

登录页存在 "Forgot password" 文本链接/按钮，点击后弹框提示：  
**"Please contact your administrator for assistance"**

```python
# 点击 Forgot password
page.get_by_text("Forgot password", exact=False).click()
page.wait_for_timeout(1000)
# 验证弹框提示
expect(page.get_by_text("Please contact your administrator for assistance",
                          exact=False)).to_be_visible(timeout=5000)
# 关闭弹框（尝试常见按钮名）
for btn_name in ["OK", "Ok", "Close", "确认", "关闭"]:
    try:
        page.get_by_role("button", name=btn_name).click(timeout=2000)
        break
    except Exception:
        pass
```
