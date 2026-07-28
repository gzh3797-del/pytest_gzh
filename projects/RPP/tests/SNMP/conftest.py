"""
SNMP 测试专用 conftest
- 将本目录加入 sys.path，使 test_snmp_*.py 可以直接 import snmp_utils / snmp_oid_map 等
- 覆盖上层 tests/conftest.py 的 screenshot_on_fail 和 app_page，
  避免 SNMP 测试被迫启动一个多余的浏览器 session
- 每次测试会话开始时自动从 HMI 页面下载 MIB 文件并生成 mib_mapping.json
"""
import sys
import os
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


# ── 覆盖上层 autouse fixture，让 SNMP 测试不依赖 app_page ─────────────────────

@pytest.fixture(scope="session")
def app_page():
    """SNMP 测试管理自己的浏览器，此处返回 None 以禁用上层 session 浏览器。"""
    return None


@pytest.fixture(autouse=True)
def screenshot_on_fail(request):
    """SNMP 测试不做截图（各用例自行管理 Playwright browser）。"""
    yield


# ── MIB 自动下载（每次 session 执行一次）────────────────────────────────────────

@pytest.fixture(scope="session", autouse=True)
def _mib_setup():
    """
    每次 SNMP 测试会话开始前自动执行：
      1. 从 HMI SNMP 页面下载 MIB 文件并解压到 mib/ 目录
      2. 读取每台设备 Settings > Connection 的 Template 字段
      3. 将 Template → MIB 文件映射写入 mib_mapping.json
    供 snmp_oid_map.py 在运行时动态加载对应 MIB。
    """
    import mib_manager
    try:
        mib_manager.build_and_save_mapping()
    except Exception as exc:
        print(f"\n[MIB] WARNING: MIB 自动下载失败 ({exc})，将使用已有 mib/ 文件（如存在）\n",
              flush=True)
    yield


# ── 共享浏览器 session（MIB 下载完成后登录一次，整个 session 复用）──────────────

@pytest.fixture(scope="session")
def snmp_browser_page(_mib_setup, playwright):
    """
    MIB 下载完成后打开浏览器并登录一次，整个测试 session 复用此 page。
    复用 pytest-playwright 提供的 playwright 实例，避免多个 sync_playwright() 冲突。
    helpers_data._select_and_walk 和 restore_after_test 通过
    goto_snmp + apply_snmp_v2c 直接在该 page 上操作，无需反复登录。
    """
    from configure_snmp import login_and_goto_snmp
    import helpers_data
    browser = playwright.chromium.launch(headless=False)
    ctx = browser.new_context(
        ignore_https_errors=True,
        viewport={"width": 1920, "height": 1080},
    )
    page = ctx.new_page()
    login_and_goto_snmp(page)
    helpers_data.set_shared_page(page)
    yield page
    helpers_data.set_shared_page(None)
    browser.close()


@pytest.fixture
def snmp_page(snmp_browser_page):
    """
    覆盖 tests/conftest.py 的同名 fixture（返回 BasePage 子类），
    改为返回 snmp_browser_page 的原生 Playwright Page 对象，
    供 helpers_ui.SNMPBase 的 UI 测试直接操作。
    """
    return snmp_browser_page
