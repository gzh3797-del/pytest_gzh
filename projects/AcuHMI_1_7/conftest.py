import base64
import sys
from pathlib import Path

import allure
import pytest
from playwright.sync_api import sync_playwright, Browser

from projects.AcuHMI_1_7.settings import HEADLESS, SLOW_MO, BASE_URL

# 强制标准输出/错误流使用 UTF-8，避免 Windows 控制台中文乱码
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

# 注：.env 已由 projects.AcuHMI_1_7.settings 适配层加载，无需在此重复 load_dotenv。


# ── 全项目唯一 Playwright 实例 ─────────────────────────────────────────────
# 在项目级集中管理 playwright + browser，覆盖 pytest-playwright 内置 fixture，让所有
# 子模块（BacnetIP / parameter_settings / ui / SNMP 等）共享同一实例，避免多处调用
# sync_playwright() 在同一线程内竞争事件循环。
#
# ⚠️ 覆盖的实例 fixture 必须叫 `playwright`：pytest-playwright ≥0.5 已把内置实例
# fixture 由旧名 `playwright_instance` 改名为 `playwright`。若仍用旧名覆盖，新版插件
# 根本不会用到它——覆盖失效，反而额外起了一个独立的 sync_playwright，与插件自带的
# `playwright` 同时存在两个实例；先建立者在主线程留下正在运行的事件循环，插件随后
# sync_playwright().start() 撞上它，抛 "Playwright Sync API inside the asyncio loop"。
#
# 注：browser 故意不依赖插件的 `browser_name` fixture，从而不会触发 pytest-playwright
# 的 [chromium] 参数化，保持各子模块用例 nodeid 与历史一致。

@pytest.fixture(scope="session")
def playwright():
    with sync_playwright() as pw:
        yield pw


@pytest.fixture(scope="session")
def browser(playwright) -> Browser:
    br = playwright.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)
    yield br
    br.close()


@pytest.fixture(scope="session")
def browser_context_args(browser_context_args):
    return {
        **browser_context_args,
        "base_url": BASE_URL,
        "ignore_https_errors": True,
        "viewport": {"width": 1280, "height": 720},
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
def _expose_page_for_report(request):
    # autouse=True：对所有用例生效，让 pytest_runtest_makereport 能在用例失败时取到 page 截图。
    # 仅当用例已通过其他 fixture 持有 page 时才暴露，不主动创建新 context（避免产生额外浏览器窗口）。
    # 各模块可在自己的 conftest 中覆盖此 fixture，提供对应的 page 对象。
    if "app_page" in request.fixturenames:
        request.node.funcargs_page = request.getfixturevalue("app_page")
    # 其他模块（BACnet/DataCollect/Device Mirror/Pass Through）在各自 conftest 覆盖本 fixture
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


# ── Allure 标签（parameter_settings 子设备） ────────────────────────────────
# 各设备的 feature/story 映射，key = 测试函数名（originalname）
_ALLURE_LABELS: dict[str, dict[str, tuple[str, str]]] = {
    "acurev4100": {
        "test_password":               ("基本配置",    "Password"),
        "test_backlight":              ("基本配置",    "Backlight"),
        "test_nominal_frequency":      ("基本配置",    "Nominal Frequency"),
        "test_pt_ratio":               ("基本配置",    "PT1 / PT2 Ratio"),
        "test_nominal_current":        ("基本配置",    "Nominal Current"),
        "test_phase_order":            ("基本配置",    "Phase Order"),
        "test_energy_reading_format":  ("能量设置",    "Energy Reading Format"),
        "test_energy_pulse_constant":  ("能量设置",    "Energy Pulse Constant"),
        "test_demand_method":          ("需量设置",    "Demand Method"),
        "test_demand_interval":        ("需量设置",    "Demand Interval"),
        "test_demand_update_rate":     ("需量设置",    "Demand Update Rate"),
        "test_var_pf":                 ("电能质量",    "VAR/PF Convention"),
        "test_reactive_method":        ("电能质量",    "Reactive Power Method"),
        "test_led_pulse_width":        ("LED 脉冲",   "LED Pulse Width"),
        "test_led_pulse_parameter":    ("LED 脉冲",   "LED Pulse Parameter"),
        "test_md_clear_mode":          ("最大需量复位", "Clear Mode"),
        "test_md_auto_reset_date":     ("最大需量复位", "Auto Reset Date"),
        "test_volt_sag_threshold":     ("电压跌落",    "Threshold"),
        "test_volt_sag_hysteresis":    ("电压跌落",    "Hysteresis"),
        "test_volt_sag_do_output":     ("电压跌落",    "DO Output"),
        "test_volt_sag_ro_output":     ("电压跌落",    "RO Output"),
        "test_volt_swell_threshold":   ("电压突升",    "Threshold"),
        "test_volt_swell_hysteresis":  ("电压突升",    "Hysteresis"),
        "test_volt_intr_threshold":    ("电压中断",    "Threshold"),
        "test_volt_intr_hysteresis":   ("电压中断",    "Hysteresis"),
        "test_curr_swell_threshold":   ("电流突升",    "Threshold"),
        "test_curr_swell_hysteresis":  ("电流突升",    "Hysteresis"),
        "test_waveform_sampling_rate": ("波形录制",    "Sample Rate"),
        "test_waveform_pre_cycles":    ("波形录制",    "Num of Cycles Before"),
        "test_waveform_post_cycles":   ("波形录制",    "Num of Cycles After"),
        "test_waveform_di1_trigger":   ("波形录制",    "DI1 Trigger"),
        "test_waveform_di2_trigger":   ("波形录制",    "DI2 Trigger"),
        "test_waveform_di3_trigger":   ("波形录制",    "DI3 Trigger"),
        "test_waveform_di4_trigger":   ("波形录制",    "DI4 Trigger"),
        "test_waveform_manual_trigger":("波形录制",    "Manually Trigger"),
    },
    "acuvim_iiw": {
        "test_nominal_frequency":    ("基本配置", "Nominal Frequency"),
        "test_password":             ("基本配置", "Password"),
        "test_pt_ratio":             ("基本配置", "PT1 / PT2 Ratio"),
        "test_ct1":                  ("基本配置", "CT1"),
        "test_backlight":            ("基本配置", "LCD Backlight"),
        "test_demand_interval":      ("需量设置", "Demand Interval"),
        "test_demand_method":        ("需量设置", "Demand Method"),
        "test_var_pf":               ("电能质量", "VAR/PF Convention"),
        "test_energy_calc_method":   ("电能质量", "Energy Calculation Method"),
        "test_reactive_method":      ("电能质量", "Reactive Power Method"),
        "test_kwh_pulse_constant":   ("脉冲设置", "kWh Pulse Constant"),
        "test_kvarh_pulse_constant": ("脉冲设置", "kvarh Pulse Constant"),
    },
}


@pytest.fixture(autouse=True)
def allure_labels(request):
    # 根据测试文件所在的设备子目录名选择对应的标签表
    test_path = Path(request.node.fspath)
    device_dir = test_path.parent.name          # e.g. "acurev4100"
    labels = _ALLURE_LABELS.get(device_dir, {})
    func = request.node.originalname
    feature, story = labels.get(func, (device_dir, func))
    allure.dynamic.feature(feature)
    allure.dynamic.story(story)
    allure.dynamic.title(request.node.name)
    yield
