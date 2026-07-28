# 测试脚本生成规范（GENERATION_RULES）

本文件是 `webtestcase_manual_to_auto` 阶段 2 生成 Playwright + pytest 脚本时的硬性规范。
对齐仓库现网真实样例：`projects/AcuHMI_1_7/tests/ui/usermanagement/passwordchange/test_TestCase_ARM_XXL_002_04_case1_0[3-6].py`，
并遵守 CLAUDE.md 约定 #9（Playwright 编码约定）、#10（函数名内嵌编号、一函数一用例）、#12（编号逐字符一致）。

---

## 1. import / fixture 模板（按 PROJECT 映射）

```python
import pytest
from playwright.sync_api import Page, expect
from projects.<PROJECT>.settings import BASE_URL, DEFAULT_PASSWORD, DEFAULT_USERNAME
from projects.<PROJECT>.pages.login_page import LoginPage
```

- `<PROJECT>` 来自 `--project`（默认 `AcuHMI_1_7`）。
- 只 import 实际用到的符号（禁止未使用 import —— 约定 #8 零告警）。
- 凭据一律从 `settings` 取，**禁止在脚本硬编码密码**。

## 2. 测试函数骨架

```python
# 用例编号：<用例编号原串>（函数名/文件名因 Python 不能含 '-' 用下划线）
# 用例标题：<title>
# 预置条件：<precondition>
# 测试步骤：
#   1. <step_line_1>
#   2. <step_line_2>
# 预期结果：
#   1. <expected_line_1>
# 探查注：
#   - <本用例依赖的 context 选择器/坑，来源 *_context.md 或本次探查>
def test_<用例编号下划线化>(login_page: LoginPage) -> None:
    login_page.open()
    login_page.login()
    page = login_page.page

    try:
        # ── 操作步骤 ──────────────────────────────────────
        <Playwright 操作代码>

        # ── 断言 ──────────────────────────────────────────
        <expect 断言代码>
    finally:
        # ── 清理 & 配置恢复（无条件执行，确保可重复运行）──
        <删除/还原测试数据 + 还原本用例改动过的所有配置项（见 §8）>
```

- **函数名**：`test_` + `用例编号`，将 `-` 替换为 `_`（Python 标识符限制）。其余字符**逐字符保留**（约定 #10/#12）。
- **文件名**：`test_<用例编号下划线化>.py`。
- **改了配置必还原**：凡修改被测设备/系统配置的用例，必须"先读原值 → `finally` 中无条件还原"，把配置恢复到测试前状态（见 §8）；这是硬性要求，不是可选清理。
- 无副作用的只读用例可省略 `try/finally`。
- 需要多用户/跨上下文时用内联 `_helper` 函数（前缀 `_`），如 `_create_user` / `_login_with` / `_nav_to_submenu`，与现网样例一致。

## 3. 定位优先级（高 → 低）

1. `page.get_by_role("button"/"textbox"/"menuitem"/"radiogroup"/"combobox", name="...")`
2. `page.get_by_label("...", exact=True)`
3. `page.locator("#id")`
4. `page.get_by_text("...", exact=True)`
5. `page.get_by_placeholder("...")`
6. 简单 CSS：`page.locator("input[name='...']")`

**禁止**：XPath、多级复杂 CSS、`.nth(N)` 硬编码索引。多元素用 `.filter(has_text="...")` 或 `.first` 收敛。

## 4. Element Plus 组件（约定 #9）

优先 `locator` + `expect`（保留 auto-wait 与 actionability）。仅当 Element Plus 合成事件链拦截（如 el-radio）才降级坐标点击，且先 `scroll_into_view_if_needed()` 再取坐标。

```python
# el-select 下拉（非原生 select）
page.get_by_text("--Select Role--", exact=True).click()
page.get_by_role("option", name="admin").click()

# el-checkbox
page.locator(".el-form-item").filter(has_text="Multiple Login").locator(".el-checkbox__inner").click()

# el-radio（合成事件被拦截时才降级）
radio = page.locator(".el-form-item").filter(has_text="Enable").locator(".el-radio").filter(has_text="Enable")
radio.scroll_into_view_if_needed()
radio.click()

# JS 批量点按钮须加可见性过滤
page.evaluate("() => [...document.querySelectorAll('button')].filter(b => b.offsetParent !== null).forEach(b => ...)")
```

**禁止**在测试体内使用 `asyncio.run()`（若需异步用线程模式 `_run_coro`）。

## 5. 导航 / 等待 / 弹框

```python
page.wait_for_load_state("networkidle")
page.wait_for_timeout(500)          # Element Plus 渲染缓冲

# 可选弹框（如默认改密提示）
try:
    page.get_by_role("button", name="Cancel").click(timeout=3000)
except Exception:
    pass
```

- 导航路径以 context 文档「进入路径」为准，禁止臆造菜单名。
- 跨上下文登录验证（如改密后能否登录）：
  ```python
  ctx = page.context.browser.new_context(ignore_https_errors=True)
  p = ctx.new_page(); p.goto(BASE_URL + "/#/login")
  # ... 登录 ...
  ctx.close()
  ```

## 6. skip / xfail（保守，必带 reason）

生成 skip 版（函数体仅 `pass`）当满足任一：Factory Reset / 恢复出厂 / 固件升级 / EULA 首次弹出 / 需特殊硬件或不可复现设备状态 / 等待 > 30 分钟。

```python
@pytest.mark.skip(reason="涉及恢复出厂，需手动执行")
def test_<...>(login_page: LoginPage) -> None:
    pass
```

产品行为与规格不符（非脚本 bug）用 `@pytest.mark.xfail(strict=False, reason="...")`。
`reason` 是「自动化脚本调试结果」列原因的唯一来源，缺 reason 则只写状态不写原因。

## 7. 严禁编造（对应需求 #5）

- 若某步骤所需 UI 元素在 context 文档中**没有**、且本次未现场探明 → **不得臆造选择器**。
- 该用例应在阶段 1 标红并记录根因，不进入生成；已开始生成才发现缺元素的，用
  `# TODO: 步骤 N — <原始步骤文本>（context 未描述，需补充定位）` 占位并将该用例记为待澄清，
  绝不用猜测的选择器凑数。

## 8. 测试过程配置恢复（硬性要求）

**凡是测试过程中修改了被测设备/系统配置的用例，必须在用例结束时把配置恢复到测试前的状态。**
这是每条有配置副作用用例的必备动作，不是可选项——用例可重复运行、不污染后续用例的前置条件全靠它。

原则：
- **先读后改、改后必还**：修改任一配置项之前，先读取并保存其**原始值**（在进入 `try` 之前或步骤起始处）；
  在 `finally` 中把它还原回去。禁止用臆造的"默认值"当原值——原始值来自实际读取，还原目标值来自约束表（§4.0）或 context。
- **恢复放 `finally`，无条件执行**：无论断言是否失败、中途是否抛异常，恢复动作都必须跑到 →
  一律写在 `try/finally` 的 `finally` 块，**不能**只在正常路径末尾还原（失败即泄漏配置）。
- **覆盖所有被改项**：登出、删除新建用户/数据只是**数据清理**；配置恢复指把开关、模式、阈值、通信参数、
  权限、时间/日期等**本用例改动过的每一项设置**逐项还原为原值。数据清理与配置恢复两者都要做，不能相互替代。
- **恢复动作推荐可断言**：还原后再 `expect`/`assert` 一次确认已回到原值，避免"以为恢复了其实没生效"。
- 无配置副作用的**纯只读用例**不受此约束（可省略 `finally`）。

骨架：

```python
def test_<用例编号下划线化>(login_page: LoginPage) -> None:
    login_page.open()
    login_page.login()
    page = login_page.page

    # 1) 先读并保存将被修改配置的原始值（还原基准）
    original_value = _read_config(page, "<配置项>")

    try:
        # 2) 修改配置 + 执行测试步骤 + 断言
        _set_config(page, "<配置项>", "<测试值>")
        # ... 操作 & expect ...
    finally:
        # 3) 无条件恢复原始配置（测试过程配置恢复）
        _set_config(page, "<配置项>", original_value)
        # 可选：expect 确认已回到 original_value
```
