import pytest
from pathlib import Path
from dotenv import load_dotenv

# 加载项目根目录下的 .env 文件（若存在），使环境变量在 settings.py 中生效
load_dotenv(Path(__file__).parent / ".env")

from config.settings import BROWSER, HEADLESS, SLOW_MO, BASE_URL


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
    from pages.login_page import LoginPage  ##导入类LoginPage
    return LoginPage(page)    ##page它是一个浏览器页面对象（Browser Page）


# ── 失败自动截图 ────────────────────────────────────────────────────────────

@pytest.fixture(autouse=True)
def screenshot_on_failure(request, page):
    # autouse=True 表示对所有测试自动生效，无需手动声明此 fixture
    yield  # 此处暂停，等待测试用例执行完毕
    # 测试执行完后检查是否失败，失败则截图保存到 screenshots 目录
    if request.node.rep_call.failed if hasattr(request.node, "rep_call") else False:
        from config.settings import SCREENSHOT_DIR
        from utils.helpers import timestamp
        name = f"FAIL_{request.node.name}_{timestamp()}"
        page.screenshot(path=str(SCREENSHOT_DIR / f"{name}.png"))


@pytest.hookimpl(tryfirst=True, hookwrapper=True)
def pytest_runtest_makereport(item, call):
    # pytest hook：在每个测试阶段（setup/call/teardown）生成报告对象后，
    # 将结果挂载到 item 上（如 item.rep_call），供 screenshot_on_failure 读取
    outcome = yield
    rep = outcome.get_result()
    setattr(item, f"rep_{rep.when}", rep)