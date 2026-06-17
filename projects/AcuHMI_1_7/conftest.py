import base64
import sys
from pathlib import Path

import pytest

from projects.AcuHMI_1_7.settings import BROWSER, HEADLESS, SLOW_MO, BASE_URL

# 强制标准输出/错误流使用 UTF-8，避免 Windows 控制台中文乱码
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

# 注：.env 已由 projects.AcuHMI_1_7.settings 适配层加载，无需在此重复 load_dotenv。


# ── 浏览器级 fixture ────────────────────────────────────────────────────────
# pytest-playwright 会自动识别以下两个固定名称的 fixture，并用它们覆盖插件默认值。
# 测试函数只需声明 `page` 参数，插件会按照下面的配置自动创建 browser → context → page，
# 无需在每个测试中手动调用 sync_playwright() / launch() / new_context() / new_page()。

@pytest.fixture(scope="session")
def browser_type_launch_args():
    # 控制浏览器的启动参数，scope="session" 表示整个测试会话只创建一次浏览器进程
    # headless=False 表示显示浏览器窗口（方便调试）；slow_mo 为每步操作增加毫秒延迟
    return {"headless": HEADLESS, "slow_mo": SLOW_MO}


@pytest.fixture(scope="session")
def browser_context_args(browser_context_args):
    # 控制浏览器上下文（相当于一个独立的"浏览器标签组"）的创建参数
    # **browser_context_args 保留插件自身传入的默认参数，再用下面的键值覆盖或追加
    return {
        **browser_context_args,
        "base_url": BASE_URL,           # 设置基础 URL，page.goto("/path") 会自动拼接
        "ignore_https_errors": True,    # 忽略自签名证书错误，适用于内网设备（如 192.168.x.x）
        "viewport": {"width": 1280, "height": 720},  # 统一窗口分辨率，保证截图/元素定位一致
    }


# ── Page Object fixture ─────────────────────────────────────────────────────

@pytest.fixture
def login_page(page):
    # 将 playwright 的 page 对象包装成 LoginPage，隐藏页面操作细节
    # 测试函数声明 `login_page` 参数即可直接调用封装好的登录方法
    from projects.AcuHMI_1_7.pages.login_page import LoginPage  ##导入类LoginPage
    return LoginPage(page)    ##page它是一个浏览器页面对象（Browser Page）


# ── 把 page 暴露给报告钩子（供失败时截图）──────────────────────────────────

@pytest.fixture(autouse=True)
def _expose_page_for_report(request, page):
    # autouse=True：对所有用例生效，让 pytest_runtest_makereport 能在用例失败时取到 page 截图。
    request.node.funcargs_page = page
    yield


# ── 报告增强：执行信息（用例说明）+ 失败错误信息（traceback 原生 + 截图内嵌）──

_CASE_DOC_MARKERS = ("用例编号", "用例标题", "测试步骤", "预期结果", "预置条件")


def _extract_case_doc(item) -> str:
    """取用例说明：优先函数 docstring；没有则在源文件里找含用例标记的 `#` 注释块。

    本项目用例把"用例编号/标题/测试步骤/预期结果"写成 `#` 注释块，但位置不固定：
    有的紧贴 def 上方，有的在文件顶部（def 与注释块之间隔着辅助函数）。故按"连续注释行"
    切分全文件，返回首个含上述标记的块，兼容两种结构。
    """
    func = item.function
    doc = (getattr(func, "__doc__", None) or "").strip()
    if doc:
        return doc
    try:
        lines = Path(func.__code__.co_filename).read_text(encoding="utf-8").splitlines()
    except OSError:
        return ""
    blocks: list[list[str]] = []
    current: list[str] = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("#"):
            current.append(stripped.lstrip("#").rstrip())
        elif current:
            blocks.append(current)
            current = []
    if current:
        blocks.append(current)
    for block in blocks:
        text = "\n".join(block).strip()
        if any(marker in text for marker in _CASE_DOC_MARKERS):
            return text
    return ""


@pytest.hookimpl(tryfirst=True, hookwrapper=True)
def pytest_runtest_makereport(item, call):
    # pytest hook：每个阶段（setup/call/teardown）生成报告对象后：
    #   1. 挂到 item 上（item.rep_call 等），供其他逻辑读取
    #   2. 把用例 docstring 写成报告“执行信息”extra（用例标题/测试步骤/预期结果）
    #   3. call 阶段失败时截图：既存盘到 screenshots/，又 base64 内嵌进报告对应行
    pytest_html = item.config.pluginmanager.getplugin("html")
    outcome = yield
    rep = outcome.get_result()
    setattr(item, f"rep_{rep.when}", rep)

    if pytest_html is None:
        return

    extras = list(getattr(rep, "extras", []))

    # 执行信息：用例说明（docstring 或函数上方 # 注释块），仅 call 阶段挂一次
    if rep.when == "call":
        doc = _extract_case_doc(item)
        if doc:
            extras.append(pytest_html.extras.text(doc, name="执行信息（用例步骤/预期）"))

    # 错误信息：call 阶段失败 → 截图存盘 + 内嵌报告
    if rep.when == "call" and rep.failed:
        page = getattr(item, "funcargs_page", None)
        if page is not None:
            try:
                from projects.AcuHMI_1_7.settings import get_screenshot_dir
                from projects.AcuHMI_1_7.helpers.web_helpers import timestamp
                png = page.screenshot()
                shot_path = get_screenshot_dir() / f"FAIL_{item.name}_{timestamp()}.png"
                shot_path.write_bytes(png)
                b64 = base64.b64encode(png).decode("ascii")
                extras.append(pytest_html.extras.image(b64, name="失败截图", mime_type="image/png"))
            except Exception:
                pass  # 截图失败不应影响报告生成

    rep.extras = extras
