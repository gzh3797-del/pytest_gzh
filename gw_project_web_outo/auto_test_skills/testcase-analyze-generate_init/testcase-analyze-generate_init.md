# testcase-analyze-generate_init — 自动化测试项目初始化 Skill

## 用途
为新产品/新项目一键生成 Playwright + pytest 自动化测试框架骨架，
复用 gw_project_web_outo 已验证的目录结构、配置文件和 Page Object 模式。

**使用时机**：项目第 0 天，只需运行一次。初始化完成后与源项目完全独立。

| Skill | 使用时机 | 使用频率 |
|-------|---------|---------|
| `/testcase-analyze-generate_init` | 新产品/新项目启动时 | **一次性** |
| `/testcase-analyze` | 每次分析手工用例时 | 反复使用 |

---

## 调用格式

```
/testcase-analyze-generate_init --product AcuHMI2 --url https://192.168.2.200
/testcase-analyze-generate_init --product AcuHMI2 --url https://192.168.2.200 --password Admin@654321
/testcase-analyze-generate_init --product AcuHMI2 --url https://192.168.2.200 --dir C:\projects\acuhmi2_auto
/testcase-analyze-generate_init --product NewGW --url https://192.168.1.100 --username-field "Username" --password-field "Password" --login-button "Log In" --logged-in-indicator "Dashboard"
```

参数说明：
- `--product`：产品名（**必填**，缺失立即停止）
- `--url`：设备地址（**必填**，缺失立即停止）
- `--password`：默认测试密码（可选，默认 `Admin@110001`）
- `--dir`：目标目录（可选，默认 `<当前项目同级>/<product小写>_auto`）
- `--username-field`：登录页用户名输入框的 name 文本
- `--password-field`：登录页密码输入框的 name 文本
- `--login-button`：登录按钮的 name 文本
- `--logged-in-indicator`：登录成功后 header 中显示的文字

用户传入的参数存于：$ARGUMENTS

---

## 执行流程

收到调用后，按以下步骤执行，**不要跳过任何步骤**。

### Step 0 — 参数解析

从 $ARGUMENTS 中提取：
- `PRODUCT`（必填）
- `TARGET_URL`（必填）
- `TARGET_PASSWORD`（可选，默认 `Admin@110001`）
- `TARGET_DIR`（可选，默认为当前目录的上一级拼接 `<product小写>_auto`，例如若当前在 `C:\autotest_local\autotest\gw_project_web_outo`，则默认为 `C:\autotest_local\autotest\<product小写>_auto`）
- `USERNAME_FIELD`、`PASSWORD_FIELD`、`LOGIN_BUTTON`、`LOGGED_IN_INDICATOR`（均可选）

设置标志位：
- `HAS_LOGIN_SELECTORS = True`：四个登录参数（USERNAME_FIELD、PASSWORD_FIELD、LOGIN_BUTTON、LOGGED_IN_INDICATOR）**全部提供**
- `HAS_LOGIN_SELECTORS = False`：任意一个缺失

**缺失 `--product` 或 `--url` 时，输出以下内容后立即停止：**
```
❌ 缺少必填参数。

用法：
  /testcase-analyze-generate_init --product <产品名> --url <设备地址>

示例：
  /testcase-analyze-generate_init --product AcuHMI2 --url https://192.168.2.200
  /testcase-analyze-generate_init --product NewGW --url https://192.168.1.100 \
      --username-field "Username" --password-field "Password" \
      --login-button "Log In" --logged-in-indicator "Dashboard"
```

边界处理：
- `--product` 含空格时，目录名替换为下划线，代码注释中保留原值
- `--url` 不含协议头（`http://` 或 `https://`）时，输出警告后继续（不中断）

### Step 1 — 目标目录检查

用 Bash 检查 `TARGET_DIR` 是否已存在且非空：
```bash
ls "<TARGET_DIR>"
```

- 不存在 / 空目录 → 继续
- 非空 → 暂停，输出警告并等待用户确认：
  ```
  ⚠️  目标目录已存在且非空：<TARGET_DIR>
  继续执行将覆盖已有文件。是否继续？(yes/no)
  ```
  用户回复 `yes` 继续，否则停止。

### Step 2 — 创建目录结构

用 Bash 创建以下目录（Windows 路径，含中文/空格时加双引号）：
```bash
mkdir "<TARGET_DIR>\config"
mkdir "<TARGET_DIR>\pages"
mkdir "<TARGET_DIR>\fixtures"
mkdir "<TARGET_DIR>\utils"
mkdir "<TARGET_DIR>\test_data"
mkdir "<TARGET_DIR>\tests"
mkdir "<TARGET_DIR>\Manual_testcase"
mkdir "<TARGET_DIR>\product_structure_testcase_regulation"
mkdir "<TARGET_DIR>\reports"
mkdir "<TARGET_DIR>\screenshots"
mkdir "<TARGET_DIR>\.claude\commands"
mkdir "<TARGET_DIR>\auto_test_skills\testcase-analyze"
mkdir "<TARGET_DIR>\auto_test_skills\testcase-analyze-generate_init"
```

### Step 3 — 写入通用样板文件

以下文件内容固定不变，直接写入（不做变量替换）。

**`<TARGET_DIR>\config\__init__.py`**（空文件，写入空字符串）

**`<TARGET_DIR>\pages\__init__.py`**（空文件）

**`<TARGET_DIR>\fixtures\__init__.py`**（空文件）

**`<TARGET_DIR>\utils\__init__.py`**（空文件）

**`<TARGET_DIR>\test_data\__init__.py`**（空文件）

**`<TARGET_DIR>\tests\__init__.py`**（空文件）

**`<TARGET_DIR>\Manual_testcase\.gitkeep`**（空占位文件）

**`<TARGET_DIR>\product_structure_testcase_regulation\.gitkeep`**（空占位文件）

**`<TARGET_DIR>\conftest.py`**：
```python
import pytest
from pathlib import Path
from dotenv import load_dotenv
load_dotenv(Path(__file__).parent / ".env")
from config.settings import BROWSER, HEADLESS, SLOW_MO, BASE_URL

@pytest.fixture(scope="session")
def browser_type_launch_args():
    return {"headless": HEADLESS, "slow_mo": SLOW_MO}

@pytest.fixture(scope="session")
def browser_context_args(browser_context_args):
    return {
        **browser_context_args,
        "base_url": BASE_URL,
        "ignore_https_errors": True,
        "viewport": {"width": 1280, "height": 720},
    }

@pytest.fixture
def login_page(page):
    from pages.login_page import LoginPage
    return LoginPage(page)

@pytest.fixture(autouse=True)
def screenshot_on_failure(request, page):
    yield
    if request.node.rep_call.failed if hasattr(request.node, "rep_call") else False:
        from config.settings import SCREENSHOT_DIR
        from utils.helpers import timestamp
        name = f"FAIL_{request.node.name}_{timestamp()}"
        page.screenshot(path=str(SCREENSHOT_DIR / f"{name}.png"))

@pytest.hookimpl(tryfirst=True, hookwrapper=True)
def pytest_runtest_makereport(item, call):
    outcome = yield
    rep = outcome.get_result()
    setattr(item, f"rep_{rep.when}", rep)
```

**`<TARGET_DIR>\pages\base_page.py`**：
```python
from playwright.sync_api import Page, expect
from config.settings import TIMEOUT

class BasePage:
    def __init__(self, page: Page):
        self.page = page
        self.timeout = TIMEOUT

    def navigate(self, url: str):
        self.page.goto(url, timeout=self.timeout)

    def click(self, selector: str):
        self.page.click(selector, timeout=self.timeout)

    def fill(self, selector: str, value: str):
        self.page.fill(selector, value, timeout=self.timeout)

    def get_text(self, selector: str) -> str:
        return self.page.text_content(selector, timeout=self.timeout)

    def is_visible(self, selector: str) -> bool:
        return self.page.is_visible(selector)

    def wait_for_selector(self, selector: str):
        self.page.wait_for_selector(selector, timeout=self.timeout)

    def take_screenshot(self, name: str):
        from config.settings import SCREENSHOT_DIR
        path = SCREENSHOT_DIR / f"{name}.png"
        self.page.screenshot(path=str(path))
        return path

    def wait_for_load(self):
        self.page.wait_for_load_state("networkidle", timeout=self.timeout)
```

**`<TARGET_DIR>\utils\helpers.py`**：
```python
import time
import json
from pathlib import Path
from datetime import datetime

def timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def load_json(file_path: str | Path) -> dict:
    with open(file_path, "r", encoding="utf-8") as f:
        return json.load(f)

def save_json(data: dict, file_path: str | Path):
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def wait_seconds(seconds: float):
    time.sleep(seconds)
```

**`<TARGET_DIR>\pytest.ini`**：
```ini
[pytest]
pythonpath = .
testpaths = tests
python_files = test_*.py
python_classes = Test*
python_functions = test_*
addopts =
    --browser chromium
    --html=reports/report.html
    --self-contained-html
    -v
markers =
    smoke: smoke tests
    regression: regression tests
    login: login related tests
```

**`<TARGET_DIR>\requirements.txt`**：
```
pytest>=7.4.0
pytest-playwright>=0.4.0
playwright>=1.40.0
pytest-html>=4.0.0
pytest-xdist>=3.3.0
python-dotenv>=1.0.0
allure-pytest>=2.13.0
openpyxl>=3.1.0
```

**`<TARGET_DIR>\.gitignore`**：
```
.env
__pycache__/
*.pyc
*.pyo
.pytest_cache/
reports/
screenshots/
playwright-report/
test-results/
~$*
```

### Step 4 — 写入产品定制文件

以下文件含变量替换，将占位符替换为实际参数值后写入。

**`<TARGET_DIR>\config\settings.py`**
将 `{{URL}}` 替换为 TARGET_URL，`{{PASSWORD}}` 替换为 TARGET_PASSWORD：
```python
import os
from pathlib import Path
BASE_DIR = Path(__file__).parent.parent
BROWSER = os.getenv("BROWSER", "chromium")
HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"
SLOW_MO = int(os.getenv("SLOW_MO", "300"))
TIMEOUT = int(os.getenv("TIMEOUT", "30000"))
BASE_URL = os.getenv("BASE_URL", "{{URL}}")
SCREENSHOT_DIR = BASE_DIR / "screenshots"
REPORT_DIR = BASE_DIR / "reports"
SCREENSHOT_DIR.mkdir(exist_ok=True)
REPORT_DIR.mkdir(exist_ok=True)
DEFAULT_USERNAME = os.getenv("WEB_USERNAME", "admin")
DEFAULT_PASSWORD = os.getenv("WEB_PASSWORD", "{{PASSWORD}}")
```

**`<TARGET_DIR>\.env.example`**
将 `{{URL}}` 和 `{{PASSWORD}}` 替换为实际值：
```
# 复制此文件为 .env 并填写实际值
BASE_URL={{URL}}
WEB_USERNAME=admin
WEB_PASSWORD={{PASSWORD}}
HEADLESS=false
SLOW_MO=300
BROWSER=chromium
```

**`<TARGET_DIR>\pages\login_page.py`**

若 `HAS_LOGIN_SELECTORS = True`（四个登录参数全部提供），替换以下模板中的占位符后写入
（`{{USERNAME_FIELD}}`、`{{PASSWORD_FIELD}}`、`{{LOGIN_BUTTON}}`、`{{LOGGED_IN_INDICATOR}}` 替换为对应参数值）：
```python
from playwright.sync_api import Page
from config.settings import DEFAULT_USERNAME, BASE_URL, DEFAULT_PASSWORD
from pages.base_page import BasePage

class LoginPage(BasePage):
    def __init__(self, page: Page):
        super().__init__(page)
        self.url = BASE_URL + "/#/login"

    def open(self):
        self.navigate(self.url)
        self.wait_for_load()

    def login(self, username: str = DEFAULT_USERNAME, password: str = DEFAULT_PASSWORD):
        self.page.get_by_role("textbox", name="{{USERNAME_FIELD}}").fill(username)
        self.page.get_by_role("textbox", name="{{USERNAME_FIELD}}").press("Tab")
        self.page.get_by_role("textbox", name="{{PASSWORD_FIELD}}").fill(password)
        self.page.get_by_role("button", name="{{LOGIN_BUTTON}}").click()
        self.wait_for_load()
        try:
            self.page.get_by_role("button", name="Cancel").click(timeout=3000)
        except Exception:
            pass

    def is_logged_in(self) -> bool:
        return self.page.locator("header span").filter(has_text="{{LOGGED_IN_INDICATOR}}").is_visible()
```

若 `HAS_LOGIN_SELECTORS = False`（任意登录参数缺失），写入存根版本
（将 `{{URL}}` 替换为 TARGET_URL；记录缺失的参数名称列表用于 Step 8 摘要）：
```python
from playwright.sync_api import Page
from config.settings import DEFAULT_USERNAME, BASE_URL, DEFAULT_PASSWORD
from pages.base_page import BasePage

# TODO: 使用 playwright codegen {{URL}} 查看实际登录元素选择器
# 操作步骤：
#   1. 运行：playwright codegen {{URL}}
#   2. 在打开的浏览器中手动操作登录
#   3. 从生成的代码中复制正确的选择器
#   4. 替换下方带 TODO 注释的行

class LoginPage(BasePage):
    def __init__(self, page: Page):
        super().__init__(page)
        self.url = BASE_URL + "/#/login"  # TODO: 确认登录页 URL 路径

    def open(self):
        self.navigate(self.url)
        self.wait_for_load()

    def login(self, username: str = DEFAULT_USERNAME, password: str = DEFAULT_PASSWORD):
        # TODO: 替换为实际的用户名输入框选择器
        self.page.get_by_role("textbox", name="TODO_USERNAME_FIELD").fill(username)
        self.page.get_by_role("textbox", name="TODO_USERNAME_FIELD").press("Tab")
        # TODO: 替换为实际的密码输入框选择器
        self.page.get_by_role("textbox", name="TODO_PASSWORD_FIELD").fill(password)
        # TODO: 替换为实际的登录按钮选择器
        self.page.get_by_role("button", name="TODO_LOGIN_BUTTON").click()
        self.wait_for_load()
        try:
            self.page.get_by_role("button", name="Cancel").click(timeout=3000)
        except Exception:
            pass

    def is_logged_in(self) -> bool:
        # TODO: 替换为实际的登录成功标志选择器
        # 提示：通常是 header 中显示的用户名或产品名
        # 配置完成后删除此行并取消下一行注释：
        # return self.page.locator("header span").filter(has_text="TODO_INDICATOR").is_visible()
        return True
```

**`<TARGET_DIR>\tests\test_smoke.py`**

若 `HAS_LOGIN_SELECTORS = True`，写入正常版本：
```python
import pytest
from pages.login_page import LoginPage

@pytest.mark.smoke
def test_login_success(page):
    login_page = LoginPage(page)
    login_page.open()
    login_page.login()
    assert login_page.is_logged_in(), "登录失败：未检测到登录成功标志"

@pytest.mark.smoke
def test_login_wrong_password(page):
    login_page = LoginPage(page)
    login_page.open()
    login_page.login(password="WrongPassword123")
    assert not login_page.is_logged_in(), "预期登录失败，但检测到登录成功标志"
```

若 `HAS_LOGIN_SELECTORS = False`，写入带 skip 的版本：
```python
import pytest
from pages.login_page import LoginPage

@pytest.mark.skip("TODO: 先完成 pages/login_page.py 的选择器配置")
@pytest.mark.smoke
def test_login_success(page):
    login_page = LoginPage(page)
    login_page.open()
    login_page.login()
    assert login_page.is_logged_in(), "登录失败：未检测到登录成功标志"

@pytest.mark.skip("TODO: 先完成 pages/login_page.py 的选择器配置")
@pytest.mark.smoke
def test_login_wrong_password(page):
    login_page = LoginPage(page)
    login_page.open()
    login_page.login(password="WrongPassword123")
    assert not login_page.is_logged_in(), "预期登录失败，但检测到登录成功标志"
```

### Step 5 — 复制 testcase-analyze Skill 文件

检查当前工作目录下是否存在 `auto_test_skills/testcase-analyze/` 路径。

**若路径存在**，读取以下文件内容并写入目标项目对应位置：
- `auto_test_skills/testcase-analyze/testcase-analyze.md` → `<TARGET_DIR>/auto_test_skills/testcase-analyze/testcase-analyze.md`
- `auto_test_skills/testcase-analyze/testcase-analyze-readme.md` → `<TARGET_DIR>/auto_test_skills/testcase-analyze/testcase-analyze-readme.md`
- `auto_test_skills/testcase-analyze/excel_writer.py` → `<TARGET_DIR>/auto_test_skills/testcase-analyze/excel_writer.py`
- `.claude/commands/testcase-analyze.md` → `<TARGET_DIR>/.claude/commands/testcase-analyze.md`

同时，若 `product_structure_testcase_regulation/autotest_generativerule.md` 存在，读取并写入：
- `<TARGET_DIR>/product_structure_testcase_regulation/autotest_generativerule.md`

**若路径不存在**，输出警告并跳过（不中断流程）：
```
⚠️  未找到 testcase-analyze Skill 源文件，跳过复制步骤。
    请手动将以下文件复制到目标项目：
    - auto_test_skills/testcase-analyze/（3 个文件）
    - .claude/commands/testcase-analyze.md
```

### Step 6 — 生成 readme.md

写入 `<TARGET_DIR>/readme.md`（将 `{{PRODUCT}}`、`{{URL}}`、`{{PASSWORD}}` 替换为实际值）：

内容模板：
```
# {{PRODUCT}} 自动化测试项目

基于 Playwright + pytest 的 Web UI 自动化测试框架。

## 快速开始

### 1. 安装依赖

    pip install -r requirements.txt
    playwright install chromium

### 2. 配置环境

    copy .env.example .env

编辑 .env 文件，填写实际的设备地址和密码：

    BASE_URL={{URL}}
    WEB_PASSWORD={{PASSWORD}}

### 3. 运行 Smoke 测试

    pytest tests/test_smoke.py -v

### 4. 分析手工测试用例

1. 将手工测试用例 Excel 放入 Manual_testcase/ 目录
2. 在 Claude Code 中打开本项目目录
3. 运行：/testcase-analyze --all

## 目录结构

    .
    ├── config/          # 配置文件
    ├── pages/           # Page Object（base_page.py、login_page.py）
    ├── fixtures/        # pytest fixtures
    ├── utils/           # 工具函数
    ├── test_data/       # 测试数据
    ├── tests/           # 测试脚本
    ├── Manual_testcase/ # 手工测试用例 Excel 文件
    ├── product_structure_testcase_regulation/  # 模块页面结构文档
    ├── .claude/commands/   # Claude Code Skill 入口
    ├── auto_test_skills/   # Skill 存档文件
    ├── conftest.py      # pytest 全局配置
    ├── pytest.ini       # pytest 配置
    └── requirements.txt # Python 依赖

## 常用命令

    # 运行 smoke 测试
    pytest -m smoke -v

    # 运行所有测试并生成 HTML 报告
    pytest -v
```

### Step 7 — 写入 testcase-analyze-generate_init Skill 存档

读取 `.claude/commands/testcase-analyze-generate_init.md` 内容，写入目标项目：
- `<TARGET_DIR>/.claude/commands/testcase-analyze-generate_init.md`
- `<TARGET_DIR>/auto_test_skills/testcase-analyze-generate_init/testcase-analyze-generate_init.md`

若 `auto_test_skills/testcase-analyze-generate_init/testcase-analyze-generate_init-readme.md` 存在，也写入：
- `<TARGET_DIR>/auto_test_skills/testcase-analyze-generate_init/testcase-analyze-generate_init-readme.md`

### Step 8 — 输出初始化摘要

输出以下格式的摘要（根据 HAS_LOGIN_SELECTORS 状态调整「还需手动完成」部分）：

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
项目初始化完成：<TARGET_DIR>
产品：<PRODUCT> | 设备地址：<TARGET_URL>
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
生成文件（共 XX 个）：
  目录：config/ pages/ fixtures/ utils/ test_data/ tests/
        Manual_testcase/ product_structure_testcase_regulation/
  通用：conftest.py / pages/base_page.py / utils/helpers.py
        pytest.ini / requirements.txt / .gitignore
  定制：config/settings.py（BASE_URL=<TARGET_URL>）
        .env.example / pages/login_page.py（完整版 or 存根版）/ tests/test_smoke.py
  Skill：.claude/commands/testcase-analyze.md
         .claude/commands/testcase-analyze-generate_init.md
         auto_test_skills/testcase-analyze/（3 个文件）
         auto_test_skills/testcase-analyze-generate_init/（2 个文件）
  其他：readme.md / product_structure_testcase_regulation/autotest_generativerule.md
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
下一步
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  1. cd <TARGET_DIR>
  2. pip install -r requirements.txt
  3. playwright install chromium
  4. copy .env.example .env  （然后编辑 .env 填写实际密码）
  5. pytest tests/test_smoke.py -v
  6. 将手工用例 Excel 放入 Manual_testcase/
  7. 在 Claude Code 中打开 <TARGET_DIR>，运行：
       /testcase-analyze --all
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
还需手动完成
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  # HAS_LOGIN_SELECTORS = True 时：
  ✅ login_page.py 已生成完整选择器，运行 smoke 测试确认与实际 UI 一致

  # HAS_LOGIN_SELECTORS = False 时（当前状态）：
  ⚠️  缺失登录参数：<列出具体缺失的参数名>
  - 运行：playwright codegen <TARGET_URL>
  - 按 TODO 注释更新 pages/login_page.py 的选择器
  - 删除 tests/test_smoke.py 中的 @pytest.mark.skip
  - 重新运行：pytest tests/test_smoke.py -v

  # 通用（所有产品）：
  - 为每个测试模块补充 product_structure_testcase_regulation/<module>_struct.md
    参考格式：product_structure_testcase_regulation/autotest_generativerule.md
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

---

## 注意事项

- 路径含中文或空格时，Bash 命令中用双引号包裹路径
- 生成的项目与 gw_project_web_outo **完全独立**，无运行时依赖
- 目标目录已存在文件时，覆盖前会询问确认
- `HAS_LOGIN_SELECTORS = False` 时，smoke 测试加 `@pytest.mark.skip`，不影响框架其余部分
- `autotest_generativerule.md` 运行时从源项目读取，确保版本与源项目一致
- 此 Skill 只需在新项目启动时运行一次；后续使用 `/testcase-analyze` 分析用例
