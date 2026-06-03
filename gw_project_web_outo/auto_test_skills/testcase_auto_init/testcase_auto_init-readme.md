# testcase_auto_init Skill 使用说明

## 功能简介

`testcase_auto_init` 是一个 Claude Code 自定义 Skill，
用于为新产品/新项目**一键生成** Playwright + pytest 自动化测试框架骨架，
复用 gw_project_web_outo 已验证的目录结构、配置文件和 Page Object 模式。

**使用时机**：项目第 0 天，只需运行一次。初始化完成后新项目与源项目**完全独立**，无运行时依赖。

### 工作流中的位置

```
【新产品启动】项目第 0 天，运行一次
    ↓
/testcase_auto_init --product X --url https://...
    ↓
框架骨架就绪（目录、配置、Page Object、Skill 文件全部生成）
    ↓
【用例分析阶段】反复使用
    ↓
/testcase-analyze --all
    ↓
【脚本生成阶段】（未来 Skill）
```

| Skill | 使用时机 | 使用频率 |
|-------|---------|---------|
| `/testcase_auto_init` | 新产品/新项目启动时 | **一次性** |
| `/testcase-analyze` | 每次分析手工用例时 | 反复使用 |
| `/testcase-generate`（未来） | 每次生成自动化脚本时 | 反复使用 |

---

## 目录结构

```
auto_test_skills/testcase_auto_init/
├── testcase_auto_init.md          # Skill 主体内容（执行流程）
└── testcase_auto_init-readme.md   # 本说明文档

.claude/commands/
└── testcase_auto_init.md          # Skill 调用入口（内容与上面相同）
```

---

## 一、部署方式（分享给同事）

### 方式 A：通过 Git（推荐）

```bash
git pull
```

拉取后确认以下文件存在：
```
.claude\commands\testcase_auto_init.md
```

### 方式 B：手动复制

将以下文件复制到同事的 gw_project_web_outo 目录：
```
.claude\commands\testcase_auto_init.md
```

### 环境依赖

无额外依赖（文件生成是 Claude 原生能力，不需要额外的 Python 脚本）。

---

## 二、调用方式

> **重要**：Skill 必须用斜杠命令触发，普通提问不会启动 Skill。
> 必须在 Claude Code 中打开 gw_project_web_outo 目录后调用。

### 参数说明

| 参数 | 是否必填 | 说明 | 默认值 |
|------|---------|------|-------|
| `--product` | **必填** | 产品名称 | — |
| `--url` | **必填** | 设备/产品地址 | — |
| `--password` | 可选 | 默认测试密码 | `Admin@110001` |
| `--dir` | 可选 | 目标目录 | `<当前项目同级>/<product小写>_auto` |
| `--username-field` | 可选 | 登录页用户名输入框 name 文本 | — |
| `--password-field` | 可选 | 登录页密码输入框 name 文本 | — |
| `--login-button` | 可选 | 登录按钮 name 文本 | — |
| `--logged-in-indicator` | 可选 | 登录成功后 header 中显示的文字 | — |

### 四种典型场景

**场景 1：最简调用（存根模式）**
```
/testcase_auto_init --product AcuHMI2 --url https://192.168.2.200
```
- 生成完整框架骨架
- `login_page.py` 含 TODO 注释（选择器待填写）
- `test_smoke.py` 含 `@pytest.mark.skip`
- 需手动补全选择器后才能运行 smoke 测试

**场景 2：指定密码**
```
/testcase_auto_init --product AcuHMI2 --url https://192.168.2.200 --password Admin@654321
```

**场景 3：指定输出目录**
```
/testcase_auto_init --product AcuHMI2 --url https://192.168.2.200 --dir C:\projects\acuhmi2_auto
```

**场景 4：完整登录选择器（完整模式）**
```
/testcase_auto_init --product NewGW --url https://192.168.1.100 --username-field "Username" --password-field "Password" --login-button "Log In" --logged-in-indicator "Dashboard"
```
- `login_page.py` 直接生成可运行版本
- `test_smoke.py` 无 skip，可直接运行

---

## 三、生成文件清单

初始化完成后，目标目录结构如下：

```
<TARGET_DIR>/
├── config/
│   ├── __init__.py
│   └── settings.py             ← 产品定制（BASE_URL、DEFAULT_PASSWORD）
├── pages/
│   ├── __init__.py
│   ├── base_page.py            ← 通用 Page Object 基类
│   └── login_page.py           ← 产品定制（完整版 or 存根版）
├── fixtures/
│   └── __init__.py
├── utils/
│   ├── __init__.py
│   └── helpers.py              ← 通用工具函数
├── test_data/
│   └── __init__.py
├── tests/
│   ├── __init__.py
│   └── test_smoke.py           ← 产品定制（正常版 or skip 版）
├── Manual_testcase/
│   └── .gitkeep
├── product_structure_testcase_regulation/
│   ├── .gitkeep
│   └── autotest_generativerule.md  ← 从源项目复制
├── .claude/
│   └── commands/
│       ├── testcase-analyze.md
│       └── testcase_auto_init.md
├── auto_test_skills/
│   ├── testcase-analyze/
│   │   ├── testcase-analyze.md
│   │   ├── testcase-analyze-readme.md
│   │   └── excel_writer.py
│   └── testcase_auto_init/
│       ├── testcase_auto_init.md
│       └── testcase_auto_init-readme.md
├── conftest.py                 ← 通用 pytest 配置
├── pytest.ini                  ← 通用 pytest 配置
├── requirements.txt            ← 通用依赖
├── .gitignore                  ← 通用忽略规则
├── .env.example                ← 产品定制（URL、密码示例）
└── readme.md                   ← 产品定制（快速开始文档）
```

| 文件类型 | 说明 |
|---------|------|
| **通用文件** | 内容固定，所有产品相同（base_page.py、helpers.py、conftest.py 等） |
| **产品定制文件** | 含变量替换（settings.py、.env.example、login_page.py、test_smoke.py、readme.md） |
| **Skill 文件** | 从源项目复制，确保版本一致 |

---

## 四、登录选择器策略

### 完整模式（HAS_LOGIN_SELECTORS = True）

四个登录参数**全部提供**时，生成可直接运行的 `login_page.py`：

```python
def login(self, ...):
    self.page.get_by_role("textbox", name="Username").fill(username)
    self.page.get_by_role("textbox", name="Password").fill(password)
    self.page.get_by_role("button", name="Log In").click()

def is_logged_in(self) -> bool:
    return self.page.locator("header span").filter(has_text="Dashboard").is_visible()
```

### 存根模式（HAS_LOGIN_SELECTORS = False）

任意参数缺失时，生成含 TODO 注释的存根，smoke 测试加 skip：

```python
def login(self, ...):
    # TODO: 替换为实际的用户名输入框选择器
    self.page.get_by_role("textbox", name="TODO_USERNAME_FIELD").fill(username)
    ...

def is_logged_in(self) -> bool:
    return True  # 临时返回 True，配置完成后改为实际断言
```

**获取选择器的方法：**
```bash
playwright codegen https://192.168.2.200
```
在弹出的浏览器中手动操作登录，右侧面板会实时生成 Playwright 选择器代码，
从中复制对应的 `get_by_role(...)` 或 `locator(...)` 写法。

---

## 五、初始化后必要手动步骤

### 通用步骤（所有产品）

```bash
cd <TARGET_DIR>
pip install -r requirements.txt
playwright install chromium
copy .env.example .env   # 然后编辑 .env 填写实际密码
```

### 若使用完整模式（HAS_LOGIN_SELECTORS = True）

```bash
# 直接运行 smoke 测试验证选择器是否与实际 UI 一致
pytest tests/test_smoke.py -v
```

若测试失败，用 `playwright codegen <URL>` 查看实际选择器，修正 `login_page.py`。

### 若使用存根模式（HAS_LOGIN_SELECTORS = False）

```bash
# 1. 查看实际选择器
playwright codegen <TARGET_URL>

# 2. 按 TODO 注释编辑 login_page.py

# 3. 删除 tests/test_smoke.py 中的 @pytest.mark.skip 行

# 4. 重新验证
pytest tests/test_smoke.py -v
```

### 补充页面结构文档

为每个需要自动化的模块补充结构文档，供 `/testcase-analyze` 精细分析使用：

```
product_structure_testcase_regulation/<module>_struct.md
```

参考格式：`product_structure_testcase_regulation/autotest_generativerule.md`（已复制到新项目）

文档内容需包含：
- 导航路径（如何从首页进入该模块）
- 页面 Tab / 子页面结构
- 关键 UI 元素的 Playwright 定位方式
- 表单字段名称及说明
- 弹框/提示文案

---

## 六、与 testcase-analyze 的协作流程

```
【初始化，一次性】
/testcase_auto_init --product X --url https://...
    ↓
框架骨架就绪 + testcase-analyze Skill 已复制到新项目
    ↓
【切换到新项目目录】
cd <TARGET_DIR>
在 Claude Code 中打开新项目
    ↓
【用例分析，反复使用】
将手工测试用例 Excel 放入 Manual_testcase/
/testcase-analyze --all
    ↓
查看红色行（需澄清），在 Excel 第 14 列填写用户答复
/testcase-analyze --module <有疑问的模块>
    ↓
重复直到所有行变为深绿色
    ↓
【脚本生成，未来 Skill】
/testcase-generate（待开发）
```

---

## 七、常见问题

**Q：调用后报"缺少必填参数"？**
A：确认 `--product` 和 `--url` 都已提供。示例：
```
/testcase_auto_init --product MyProduct --url https://192.168.1.100
```

**Q：目标目录已存在，会覆盖吗？**
A：会先询问确认（输出警告并等待 yes/no），回复 yes 后才继续覆盖。

**Q：smoke 测试全部 SKIPPED？**
A：使用了存根模式（未提供四个登录参数）。按「存根模式手动步骤」完成选择器配置后，
删除 `@pytest.mark.skip` 行并重新运行。

**Q：testcase-analyze 的文件没有被复制？**
A：确认在 Claude Code 中打开的是 gw_project_web_outo 目录（而非其他目录），
Skill 会从当前工作目录读取源文件。若仍有问题，手动复制以下文件：
```
auto_test_skills/testcase-analyze/（3 个文件）
.claude/commands/testcase-analyze.md
```

**Q：新项目能否使用不同的 Excel 列名？**
A：`excel_writer.py` 按固定列名匹配，若列名不同需修改 `excel_writer.py` 中的列名常量，
或统一使用与 gw_project_web_outo 相同的列名（推荐）。

---

## 八、相关文件

| 文件 | 说明 |
|------|------|
| `.claude/commands/testcase_auto_init.md` | Skill 调用入口（执行逻辑） |
| `auto_test_skills/testcase_auto_init/testcase_auto_init.md` | 存档副本 |
| `auto_test_skills/testcase-analyze/testcase-analyze.md` | testcase-analyze Skill（将被复制到新项目） |
| `product_structure_testcase_regulation/autotest_generativerule.md` | 自动化通用规则（将被复制到新项目） |
