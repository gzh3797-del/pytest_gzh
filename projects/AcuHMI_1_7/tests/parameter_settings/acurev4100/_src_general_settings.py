"""
test_general_settings.py  –  AcuRev4100 General Settings 自动化测试

页面路径：Physical Devices → AcuRev4100 → Settings → General（单页面，无标签页）

已验证字段（来自 inspect_acurev4100.py）：
  PT1 / PT2 / Nominal Current / Nominal Voltage
  VAR/PF Convention / Reactive Power Calculation Method
  Demand Method / Demand Interval / Demand Update Rate
  Nominal Frequency (Hz)

寄存器来源：AcuRev4100 Modbus Address Table v1.02 20260202.xlsx

运行：
  pytest tests/acurev4100/test_general_settings.py -v
"""

import allure
import pytest
import logging
from playwright.sync_api import Page
from pymodbus.client import ModbusTcpClient

logging.getLogger("pymodbus").setLevel(logging.DEBUG)

# ── 设备配置 ────────────────────────────────────────────────────────
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[5]))
import projects.AcuHMI_1_7.settings as _s

DEVICE_NAME = "AcuRev4100"
MODBUS_HOST = _s.METER_TCP_IP
MODBUS_PORT = _s.METER_TCP_PORT
SLAVE_ID    = _s.MODBUS_SLAVE

# ── 寄存器地址（Basic Setting 表）──────────────────────────────────
REG_PASSWORD        = 4096   # 0x1000  Password (0000-9999，必须 4 位)
REG_NOMINAL_CURRENT = 25856  # 0x6500  Nominal Current (1-50000 A，Waveform|PQ Event sheet)
REG_BACKLIGHT       = 4119   # 0x1017  LCD Backlight Time (0-120 min)
REG_MD_CLEAR_MODE   = 4124   # 0x101C  MaxDemand Clear Mode (0=Manual, 1=Auto, 2=Disable)
REG_MD_CLEAR_DATE   = 4125   # 0x101D  MaxDemand Clear Date Setting (high byte=day 1-31)
REG_FREQ_SEL        = 4161   # 0x1041  Frequency Selection (0=50Hz, 1=60Hz)
REG_SERVICE_CONFIG  = 4162   # 0x1042  Service Configuration / Wiring
REG_PT1             = 4163   # 0x1043  PT Primary Voltage (2 regs, uint32 big-endian)
REG_PT2             = 4165   # 0x1045  PT Secondary Voltage
REG_CH1_CT_TYPE     = 4167   # 0x1047  Channel 1 CT Type  (+5 per channel)
REG_CH1_PRIMARY     = 4168   # 0x1048  Channel 1 CT Primary (+5 per channel)
REG_DEMAND_METHOD   = 4310   # 0x10D6  Demand Method (0=Fixed, 1=Sliding)
REG_DEMAND_INTERVAL = 4311   # 0x10D7  Demand Interval (1-30 min)
REG_DEMAND_SUBINT   = 4312   # 0x10D8  Demand Update Rate (1-30 min)
REG_VAR_PF          = 4313   # 0x10D9  VAR/PF Convention (0=IEC, 1=IEEE)
REG_REACTIVE_METHOD = 4315   # 0x10DB  Reactive Power Calculation Method (0=Generic, 1=True)
REG_PHASE_ORDER     = 4316   # 0x10DC  Phase Order (0=ABC, 1=ACB)
REG_ENERGY_FMT      = 4317   # 0x10DD  Energy Reading Display Format (0=0.1kWh, 1=0.01kWh, 2=0.001kWh)
REG_ENERGY_PULSE    = 4319   # 0x10DF  Energy Pulse Constant (uint32, ×0.001, range 100–100000000)
REG_LED_PULSE_WIDTH = 4321   # 0x10E1  LED Pulse Width (20-100 ms)
REG_LED_PULSE_PARAM = 4322   # 0x10E2  LED Pulse Parameter (0=not in use, 1-360)

CH_COUNT = 18

# ── 寄存器描述表（用于日志输出）────────────────────────────────────
_REG_DESC: dict[int, str] = {
    4096:  "Password",
    4119:  "LCD Backlight Time",
    25856: "Nominal Current",
    4124: "MaxDemand Clear Mode",
    4125: "MaxDemand Clear Date Setting",
    4161: "Frequency Selection",
    4162: "Service Configuration",
    4163: "PT1 Primary Voltage (high)",
    4164: "PT1 Primary Voltage (low)",
    4165: "PT2 Secondary Voltage",
    4310: "Demand Method",
    4311: "Demand Interval",
    4312: "Demand Update Rate",
    4313: "VAR/PF Convention",
    4315: "Reactive Power Calculation Method",
    4316: "Phase Order",
    4317: "Energy Reading Display Format",
    4319: "Energy Pulse Constant (high)",
    4320: "Energy Pulse Constant (low)",
    4321: "Pulse LED Width",
    4322: "Pulse LED Parameter Selection",
}


def _reg_desc(address: int) -> str:
    """返回寄存器的字段描述，通道寄存器按步长动态计算。"""
    if address in _REG_DESC:
        return _REG_DESC[address]
    offset_type = address - REG_CH1_CT_TYPE
    if 0 <= offset_type < CH_COUNT * 5 and offset_type % 5 == 0:
        return f"Channel {offset_type // 5 + 1} CT Type"
    offset_pri = address - REG_CH1_PRIMARY
    if 0 <= offset_pri < CH_COUNT * 5 and offset_pri % 5 == 0:
        return f"Channel {offset_pri // 5 + 1} CT Primary"
    return ""


# ── 编码映射（下拉选项文字需与 HMI 实际保持一致）─────────────────
# TODO: 运行 inspect_acurev4100_dropdowns.py 确认后更新以下值
DEMAND_ENCODE = {
    "Fixed":   0,   # inspect 显示当前值为 "Fixed"
    "Sliding": 1,
}

VAR_PF_ENCODE = {
    "IEC":  0,
    "IEEE": 1,
}

REACTIVE_ENCODE = {
    "Generic": 0,   # inspect 显示当前值为 "True"，另一项推测为 "Generic"
    "True":    1,
}

FREQ_ENCODE = {
    "50Hz": 0,
    "60Hz": 1,
}

PHASE_ORDER_ENCODE = {
    "ABC": 0,
    "ACB": 1,
}

ENERGY_FMT_ENCODE = {
    "0.1 kWh":   0,
    "0.01 kWh":  1,
    "0.001 kWh": 2,
}

MD_MODE_ENCODE = {
    "Manual Reset": 0,
    "Auto Reset":   1,
    "Disable":      2,
}

_log = logging.getLogger("acurev4100_test")


# ════════════════════════════════════════════════════════════════════
# 工具函数
# ════════════════════════════════════════════════════════════════════

def step(msg: str) -> None:
    _log.info("[STEP] %s", msg)
    with allure.step(msg):
        pass


# ── 持久 Modbus 连接（整个进程共享，按需重连）────────────────────────
_modbus_client: ModbusTcpClient | None = None


def _get_modbus_client() -> ModbusTcpClient:
    import time as _time
    global _modbus_client
    for attempt in range(10):
        if _modbus_client is None:
            _modbus_client = ModbusTcpClient(MODBUS_HOST, port=MODBUS_PORT, timeout=5)
        if _modbus_client.connect():
            return _modbus_client
        _log.warning("[MODBUS] connect failed (attempt %d/10), wait 3s ...", attempt + 1)
        _modbus_client.close()
        _modbus_client = None
        _time.sleep(3)
    raise RuntimeError(f"Cannot connect to Modbus server {MODBUS_HOST}:{MODBUS_PORT} after 10 attempts")


@allure.step("Modbus 读取 addr={address}  count={count}")
def modbus_read(address: int, count: int = 1) -> list:
    import time as _time
    global _modbus_client
    desc = _reg_desc(address)
    _log.info("[MODBUS TX] FC=03  addr=%d(0x%04X)  count=%d%s",
              address, address, count,
              f"  [{desc}]" if desc else "")
    _RETRIES = 5
    _RETRY_DELAY = 3
    last_err = None
    for attempt in range(_RETRIES):
        if attempt > 0:
            _log.warning("[MODBUS RETRY] attempt %d/%d, wait %ds ...", attempt + 1, _RETRIES, _RETRY_DELAY)
            _time.sleep(_RETRY_DELAY)
            _modbus_client = None
        try:
            client = _get_modbus_client()
            rsp = client.read_holding_registers(address, count=count, device_id=SLAVE_ID)
            if rsp.isError():
                last_err = f"Modbus FC=03 error at addr={address}: {rsp}"
                _log.warning("[MODBUS ERROR] %s", last_err)
                _modbus_client = None
                continue
            regs = list(rsp.registers)
            _log.info("[MODBUS RX] registers=%s", regs)
            return regs
        except Exception as exc:
            last_err = f"Exception at addr={address}: {exc}"
            _log.warning("[MODBUS ERROR] %s", last_err)
            _modbus_client = None
            continue
    raise RuntimeError(
        f"Modbus read failed after {_RETRIES} attempts. addr={address}  Last error: {last_err}"
    )


def modbus_read_32(address: int) -> int:
    """读取 32-bit 无符号整数（高寄存器在前）。"""
    regs = modbus_read(address, count=2)
    return (regs[0] << 16) | regs[1]


@allure.step("验证寄存器 [{label}]  addr={address}  expected={expected}")
def verify_modbus(address: int, expected, label: str = "") -> None:
    regs = modbus_read(address, count=1)
    actual = regs[0]
    ok = actual == expected
    _log.info("[VERIFY] %-40s  expected=%-6s  actual=%-6s  → %s",
              label, expected, actual, "✓ PASS" if ok else "✗ FAIL")
    assert ok, f"{label}: expected {expected}, got {actual}"


@allure.step("验证寄存器32bit [{label}]  addr={address}  expected={expected}")
def verify_modbus_32(address: int, expected: int, label: str = "") -> None:
    actual = modbus_read_32(address)
    ok = actual == expected
    _log.info("[VERIFY32] %-40s  expected=%-10s  actual=%-10s  → %s",
              label, expected, actual, "✓ PASS" if ok else "✗ FAIL")
    assert ok, f"{label}: expected {expected}, got {actual}"


def verify_channels_primary(expected: int) -> None:
    """验证 Channel 1~CH_COUNT 的 CT Primary 寄存器。"""
    all_ok = True
    for ch in range(1, CH_COUNT + 1):
        reg = REG_CH1_PRIMARY + (ch - 1) * 5
        regs = modbus_read(reg, count=1)
        ok = regs[0] == expected
        _log.info("[VERIFY] Channel %2d Primary  expected=%-6s  actual=%-6s  → %s",
                  ch, expected, regs[0], "✓" if ok else "✗")
        if not ok:
            all_ok = False
    assert all_ok, f"CT Primary: 有通道值不匹配 (expected {expected})"


# ── 页面操作 ──────────────────────────────────────────────────────

@allure.step("导航到 General Settings 页面")
def nav_to_general(page: Page) -> None:
    """Physical Devices → AcuRev4100 → Settings → General。
    页面为单页面，无标签页，等待 'PT1' 字段出现即表示加载完毕。
    """
    step("Click Physical Devices")
    page.locator(
        "a:has-text('Physical Devices'), "
        "span:has-text('Physical Devices'), "
        "li:has-text('Physical Devices')"
    ).first.click()
    page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)

    step(f"Click device: {DEVICE_NAME}")
    page.locator(f"text={DEVICE_NAME}").first.click()
    page.wait_for_selector("text=Settings", timeout=10000)

    step("Click Settings menu")
    page.locator(
        "span:has-text('Settings'), a:has-text('Settings'), button:has-text('Settings')"
    ).first.click()
    page.wait_for_selector(
        "a:has-text('General'), li:has-text('General'), span:has-text('General')",
        timeout=5000,
    )

    step("Click General submenu")
    page.locator(
        "a:has-text('General'), li:has-text('General'), span:has-text('General')"
    ).first.click()
    page.wait_for_selector("button:has-text('Save')", timeout=12000)
    page.wait_for_timeout(800)
    step("Ready: General Settings")


@allure.step("输入字段 [{label}] = {value}")
def set_input(page: Page, label: str, value) -> None:
    step(f"Set input [{label}] = {value!r}")
    label_els = page.get_by_text(label, exact=False)
    n = label_els.count()
    if n == 0:
        raise RuntimeError(f"set_input: label {label!r} not found on page")
    inp = label_els.first.locator("xpath=following::input[1]")
    if inp.count() == 0:
        raise RuntimeError(f"set_input: no input found after label {label!r}")
    inp.first.scroll_into_view_if_needed()
    inp.first.click()
    inp.first.press("Control+a")
    inp.first.press("Delete")
    inp.first.type(str(value), delay=50)
    _log.info("  typed %r", value)


@allure.step("下拉选择 [{label}] → {option_text}")
def set_dropdown(page: Page, label: str, option_text: str) -> None:
    """点击下拉框并选择指定选项。

    Element UI 将所有下拉框的选项都渲染在 DOM 中，必须加 :visible
    过滤，只匹配当前打开的那个下拉框的可见选项。
    """
    step(f"Dropdown [{label}] → {option_text!r}")
    trigger = page.get_by_text(label, exact=False).first.locator(
        "xpath=following::div[contains(@class,'el-select')][1]"
    )
    trigger.click()
    page.wait_for_timeout(600)

    # 只在可见的下拉项里查找（避免命中其他下拉框的隐藏选项）
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(
        option_text, exact=True
    )
    if target.count() == 0:
        visible_items = page.locator(".el-select-dropdown__item:visible")
        opt_texts = [
            visible_items.nth(i).inner_text().strip()
            for i in range(min(visible_items.count(), 30))
        ]
        _log.info("  visible options: %s", opt_texts)
        page.keyboard.press("Escape")
        raise RuntimeError(
            f"set_dropdown: option {option_text!r} not found for [{label}]. "
            f"Visible options: {opt_texts}"
        )
    target.first.click()
    page.wait_for_timeout(400)


def get_visible_errors(page: Page) -> list[str]:
    errors = []
    locs = page.locator(".el-form-item__error")
    for i in range(locs.count()):
        loc = locs.nth(i)
        if loc.is_visible():
            txt = loc.inner_text().strip()
            if txt:
                errors.append(txt)
                _log.info("[FORM ERROR] %r", txt)
    return errors


def assert_field_error(page: Page, expected_text: str) -> None:
    errors = get_visible_errors(page)
    matched = [e for e in errors if expected_text.lower() in e.lower()]
    _log.info("[VERIFY ERROR] expected=%r", expected_text)
    _log.info("[VERIFY ERROR] visible_errors=%s", errors)
    if matched:
        _log.info("[VERIFY ERROR] ✓ PASS  matched=%r", matched[0])
    else:
        raise AssertionError(
            f"Expected error text {expected_text!r}\nVisible errors: {errors}"
        )


@allure.step("打开 LED Pulse Parameter Selector 弹窗")
def open_parameter_selector(page: Page) -> None:
    """点击 LED Pulse Parameter 字段，等待 Parameter Selector 弹窗出现。"""
    step("Open LED Pulse Parameter selector")
    form_item = page.locator(".el-form-item").filter(has_text="LED Pulse Parameter")
    form_item.first.scroll_into_view_if_needed()
    page.wait_for_timeout(300)
    content = form_item.first.locator(".el-form-item__content > div, .el-form-item__content")
    content.first.click()
    page.wait_for_selector("text=Parameter Selector", timeout=5000)
    page.wait_for_timeout(500)


@allure.step("选择 LED 参数 [{subcategory}] → {param_name}")
def select_led_param(page: Page, subcategory: str, param_name: str) -> None:
    """在 Parameter Selector 弹窗中选择参数，点击 Apply。

    Args:
        subcategory: "Input Channel Energy" / "Switch" / "System Energy" / "User Channel Energy"
        param_name:  参数全名，如 "Input Channel 1 active energy import"

    参数 ID 映射（寄存器 4322 存储的值，基于参数列表顺序）：
      0        : Switch → Disable（不使用）
      1–216    : Input Channel Energy（24 通道 × 9 种）
      217–252  : System Energy（Phase A/B/C 各 9 + System 9 = 36）
      253–360  : User Channel Energy（12 通道 × 9 种）
    """
    step(f"LED Param [{subcategory}] → {param_name!r}")

    # 确保 Energy 树已展开（子分类可见）
    subcat_el = page.get_by_text(subcategory, exact=True)
    if subcat_el.count() == 0 or not subcat_el.first.is_visible():
        step("Expand Energy tree")
        page.get_by_text("Energy", exact=True).first.click()
        page.wait_for_timeout(600)

    # 点击子分类，右侧加载参数列表
    page.get_by_text(subcategory, exact=True).first.click()
    page.wait_for_timeout(800)

    # 用搜索框过滤，避免在长列表中滚动查找
    search = page.locator(
        ".el-dialog input[placeholder*='search' i], "
        ".el-dialog input[placeholder*='keyword' i]"
    )
    if search.count() > 0 and search.first.is_visible():
        search.first.click()
        search.first.fill(param_name)
        page.wait_for_timeout(600)

    # 点击精确匹配的参数
    param_el = page.locator(".el-dialog").get_by_text(param_name, exact=True)
    if param_el.count() == 0:
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)
        raise RuntimeError(
            f"select_led_param: 参数 {param_name!r} 未在 [{subcategory}] 中找到"
        )
    param_el.first.click()
    page.wait_for_timeout(300)

    # 点击 Apply 关闭弹窗
    page.locator(".el-dialog button:has-text('Apply')").first.click()
    page.wait_for_timeout(500)
    step(f"Applied: {param_name!r}")


@allure.step("点击保存并检查结果")
def save_and_check(page: Page) -> bool:
    """点击 Save，先检测内联校验错误，再等待 toast。"""
    step("Click Save")
    page.locator("button:has-text('Save')").last.click()
    page.wait_for_timeout(1500)

    form_errors = get_visible_errors(page)
    if form_errors:
        _log.info("[SAVE] FORM VALIDATION FAILED  errors=%s", form_errors)
        return False

    page.wait_for_timeout(4500)
    for sel in [".el-message--success"]:
        loc = page.locator(sel)
        if loc.count() > 0 and loc.first.is_visible():
            _log.info("[SAVE] SUCCESS toast: %r", loc.first.inner_text()[:80])
            return True
    for sel in [".el-message--error", ".el-message--warning"]:
        loc = page.locator(sel)
        if loc.count() > 0 and loc.first.is_visible():
            _log.info("[SAVE] ERROR toast: %r", loc.first.inner_text()[:80])
            return False

    _log.info("[SAVE] No toast → assuming success")
    return True


# ════════════════════════════════════════════════════════════════════
# TC_Freq_001~002  Nominal Frequency（50Hz / 60Hz）
#   寄存器 4161：0=50Hz, 1=60Hz
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("50Hz", 0, id="TC_Freq_001"),
    pytest.param("60Hz", 1, id="TC_Freq_002"),
])
def test_nominal_frequency(app_page, option, encoded):
    nav_to_general(app_page)
    set_dropdown(app_page, "Nominal Frequency", option)
    saved = save_and_check(app_page)
    assert saved, f"Nominal Frequency={option}: 保存失败"
    verify_modbus(REG_FREQ_SEL, encoded, label=f"FrequencySelection({option})")


# ════════════════════════════════════════════════════════════════════
# TC_PT_001~012  PT1 / PT2 联合测试
#
#   约束：
#     PT1 ∈ [50, 1000000]（uint32，2 个寄存器）
#     PT2 ∈ [50, 830]（uint16）
#     PT2 < PT1（必须严格小于）
#
#   注：PT2 最小值为 50，因此 PT1 最小有效值为 51
#       （PT1=50 时无法找到满足 PT2 < 50 的合法 PT2）
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("pt1,pt2,ok,note", [
    # ── PT1 范围边界（PT2 固定为 50）────────────────────────────────
    pytest.param(51,      50,  True,  "PT1 有效最小值",      id="TC_PT_001"),
    pytest.param(480,     50,  True,  "PT1 典型值",          id="TC_PT_002"),
    pytest.param(1000000, 50,  True,  "PT1 最大值",          id="TC_PT_003"),
    pytest.param(49,      50,  False, "PT1 低于下限",        id="TC_PT_004"),
    pytest.param(1000001, 50,  False, "PT1 超出上限",        id="TC_PT_005"),
    # ── PT2 范围边界（PT1 固定为 1000）──────────────────────────────
    pytest.param(1000,    50,  True,  "PT2 最小值",          id="TC_PT_006"),
    pytest.param(1000,    830, True,  "PT2 最大值",          id="TC_PT_007"),
    pytest.param(1000,    49,  False, "PT2 低于下限",        id="TC_PT_008"),
    pytest.param(1000,    831, False, "PT2 超出上限",        id="TC_PT_009"),
    # ── PT2 < PT1 约束────────────────────────────────────────────────
    pytest.param(500,     499, True,  "PT2=PT1-1 刚好满足",  id="TC_PT_010"),
    pytest.param(500,     500, False, "PT2=PT1 不满足约束",  id="TC_PT_011"),
    pytest.param(500,     501, False, "PT2>PT1 不满足约束",  id="TC_PT_012"),
])
def test_pt_ratio(app_page, pt1, pt2, ok, note):
    nav_to_general(app_page)
    set_input(app_page, "PT1", pt1)
    set_input(app_page, "PT2", pt2)
    saved = save_and_check(app_page)
    _log.info("[TC] PT1=%-8s PT2=%-5s ok=%s  note=%s  result=%s",
              pt1, pt2, ok, note, saved)
    if ok:
        assert saved, f"PT1={pt1}, PT2={pt2} ({note}): 期望成功但出现错误"
        verify_modbus_32(REG_PT1, pt1, label="PT1 Primary Voltage")
        verify_modbus(REG_PT2, pt2, label="PT2 Secondary Voltage")
    else:
        assert not saved, f"PT1={pt1}, PT2={pt2} ({note}): 期望校验失败但保存成功"


# ════════════════════════════════════════════════════════════════════
# TC_NomCur_001~005  Nominal Current（1~50000 A）
#   HMI 字段：'Nominal Current'
#   Modbus 验证：寄存器 25856（0x6500），Waveform|PQ Event sheet
# ════════════════════════════════════════════════════════════════════

_ERR_NOM_CUR = "must be between 1 and 50000"

@pytest.mark.parametrize("value,ok", [
    pytest.param(1,     True,  id="TC_NomCur_001"),
    pytest.param(1000,  True,  id="TC_NomCur_002"),
    pytest.param(50000, True,  id="TC_NomCur_003"),
    pytest.param(0,     False, id="TC_NomCur_004"),
    pytest.param(50001, False, id="TC_NomCur_005"),
])
def test_nominal_current(app_page, value, ok):
    nav_to_general(app_page)
    set_input(app_page, "Nominal Current", value)
    saved = save_and_check(app_page)
    _log.info("[TC] Nominal Current=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Nominal Current={value}: 保存失败"
        verify_modbus(REG_NOMINAL_CURRENT, value, label="Nominal Current")
    else:
        assert not saved, f"Nominal Current={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, _ERR_NOM_CUR)


# ════════════════════════════════════════════════════════════════════
# TC_Demand_001~002  Demand Method
#   寄存器 4310：0=Fixed, 1=Sliding
#   下拉选项：inspect 显示当前值 "Fixed"，选项待 inspect_dropdowns 确认
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("Fixed",   0, id="TC_Demand_001"),
    pytest.param("Sliding", 1, id="TC_Demand_002"),
])
def test_demand_method(app_page, option, encoded):
    nav_to_general(app_page)
    set_dropdown(app_page, "Demand Method", option)
    saved = save_and_check(app_page)
    assert saved, f"Demand Method={option}: 保存失败"
    verify_modbus(REG_DEMAND_METHOD, encoded, label=f"DemandMethod({option})")


# ════════════════════════════════════════════════════════════════════
# TC_DmndInt_001~005  Demand Interval（1~30 min）
#   寄存器 4311
# ════════════════════════════════════════════════════════════════════

_ERR_DMND_INT = "must be between 1 and 30"

@pytest.mark.parametrize("value,ok", [
    pytest.param(1,  True,  id="TC_DmndInt_001"),
    pytest.param(15, True,  id="TC_DmndInt_002"),
    pytest.param(30, True,  id="TC_DmndInt_003"),
    pytest.param(0,  False, id="TC_DmndInt_004"),
    pytest.param(31, False, id="TC_DmndInt_005"),
])
def test_demand_interval(app_page, value, ok):
    nav_to_general(app_page)
    set_input(app_page, "Demand Interval", value)
    saved = save_and_check(app_page)
    _log.info("[TC] Demand Interval=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Demand Interval={value}: 保存失败"
        verify_modbus(REG_DEMAND_INTERVAL, value, label="Demand Interval")
    else:
        assert not saved, f"Demand Interval={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, _ERR_DMND_INT)


# ════════════════════════════════════════════════════════════════════
# TC_DmndRate_001~005  Demand Update Rate（1~30 min）
#   寄存器 4312
#   前置：Demand Method = Sliding（Fixed 时此字段可能置灰）
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("value,ok", [
    pytest.param(1,  True,  id="TC_DmndRate_001"),
    pytest.param(15, True,  id="TC_DmndRate_002"),
    pytest.param(30, True,  id="TC_DmndRate_003"),
    pytest.param(0,  False, id="TC_DmndRate_004"),
    pytest.param(31, False, id="TC_DmndRate_005"),
])
def test_demand_update_rate(app_page, value, ok):
    step("前置：Demand Method = Sliding（避免 Update Rate 置灰）")
    nav_to_general(app_page)
    set_dropdown(app_page, "Demand Method", "Sliding")
    app_page.wait_for_timeout(500)
    set_input(app_page, "Demand Update Rate", value)
    saved = save_and_check(app_page)
    _log.info("[TC] Demand Update Rate=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Demand Update Rate={value}: 保存失败"
        verify_modbus(REG_DEMAND_SUBINT, value, label="Demand Update Rate")
    else:
        assert not saved, f"Demand Update Rate={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, _ERR_DMND_INT)


# ════════════════════════════════════════════════════════════════════
# TC_VARPF_001~002  VAR/PF Convention（IEC / IEEE）
#   寄存器 4313：0=IEC, 1=IEEE
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("IEC",  0, id="TC_VARPF_001"),
    pytest.param("IEEE", 1, id="TC_VARPF_002"),
])
def test_var_pf(app_page, option, encoded):
    nav_to_general(app_page)
    set_dropdown(app_page, "VAR/PF Convention", option)
    saved = save_and_check(app_page)
    assert saved, f"VAR/PF={option}: 保存失败"
    verify_modbus(REG_VAR_PF, encoded, label=f"VAR/PF({option})")


# ════════════════════════════════════════════════════════════════════
# TC_RPM_001~002  Reactive Power Calculation Method
#   寄存器 4315：0=Generic, 1=True
#   下拉选项：inspect 显示当前值 "True"，另一选项推测为 "Generic"
#   待 inspect_dropdowns.py 运行结果确认后更新选项文字
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("Generic", 0, id="TC_RPM_001"),
    pytest.param("True",    1, id="TC_RPM_002"),
])
def test_reactive_method(app_page, option, encoded):
    nav_to_general(app_page)
    set_dropdown(app_page, "Reactive Power Calculation Method", option)
    saved = save_and_check(app_page)
    assert saved, f"Reactive Power Method={option}: 保存失败"
    verify_modbus(REG_REACTIVE_METHOD, encoded, label=f"ReactiveMethod({option})")


# ════════════════════════════════════════════════════════════════════
# TC_LedPulse_001~005  LED Pulse Parameter
#   寄存器 4322（0x10E2）：0=不使用，1-360=能量参数 ID
#
#   参数 ID 映射（实测 TC_LedPulse_003 = 43 推算，ID 按固件定义顺序）：
#     0        : Switch → Disable（不使用）
#     1–36     : System Energy（Phase A/B/C 各 9 + System 9 = 36）
#     37–252   : Input Channel Energy（24 通道 × 9 种 = 216）
#     253–360  : User Channel Energy（12 通道 × 9 种 = 108）
#
#   操作流程：
#     1. 点击 LED Pulse Parameter 字段 → Parameter Selector 弹窗
#     2. 展开 Energy 树 → 点击子分类 → 右侧显示参数列表
#     3. 搜索框过滤 → 点击目标参数 → Apply
#     4. Save → Modbus 验证寄存器 4322
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("subcategory,param_name,expected_reg", [
    pytest.param(
        "Switch", "Disable", 0,
        id="TC_LedPulse_001"
    ),
    pytest.param(
        "Input Channel Energy", "Input Channel 1 active energy import", 37,
        id="TC_LedPulse_002"
    ),
    pytest.param(
        "Input Channel Energy", "Input Channel 1 reactive energy net", 43,
        id="TC_LedPulse_003"
    ),
    pytest.param(
        "System Energy", "Phase A active energy import", 1,
        id="TC_LedPulse_004"
    ),
    pytest.param(
        "User Channel Energy", "User Channel 1 active energy import", 253,
        id="TC_LedPulse_005"
    ),
])
def test_led_pulse_parameter(app_page, subcategory, param_name, expected_reg):
    nav_to_general(app_page)
    open_parameter_selector(app_page)
    select_led_param(app_page, subcategory, param_name)
    saved = save_and_check(app_page)
    assert saved, f"LED Pulse Parameter={param_name!r}: 保存失败"
    verify_modbus(REG_LED_PULSE_PARAM, expected_reg, label=f"LED Pulse Param ({param_name})")


# ════════════════════════════════════════════════════════════════════
# TC_Password_001~006  Password（0000–9999，必须恰好 4 位数字）
#   寄存器 4096（0x1000），存储 uint16 数值
#   注：有效用例测试后自动恢复为 "0000"，避免影响后续登录
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("value,reg_val,ok", [
    pytest.param("0000", 0,    True,  id="TC_Password_001"),  # 最小值
    pytest.param("1234", 1234, True,  id="TC_Password_002"),  # 典型值
    pytest.param("9999", 9999, True,  id="TC_Password_003"),  # 最大值
    pytest.param("10000", None, False, id="TC_Password_004"), # 超出范围（5 位）
    pytest.param("123",   None, False, id="TC_Password_005"), # 不足 4 位
    pytest.param("0",     None, False, id="TC_Password_006"), # 不足 4 位
])
def test_password(app_page, value, reg_val, ok):
    nav_to_general(app_page)
    set_input(app_page, "Password", value)
    saved = save_and_check(app_page)
    _log.info("[TC] Password=%r  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Password={value!r}: 保存失败"
        verify_modbus(REG_PASSWORD, reg_val, label=f"Password({value})")
        # 恢复默认值 "0000"，避免影响后续测试登录
        nav_to_general(app_page)
        set_input(app_page, "Password", "0000")
        save_and_check(app_page)
    else:
        assert not saved, f"Password={value!r}: 期望校验失败但保存成功"


# ════════════════════════════════════════════════════════════════════
# TC_Backlight_001~004  LCD Backlight Time（0–120 min）
#   寄存器 4119（0x1017）
# ════════════════════════════════════════════════════════════════════

_ERR_BACKLIGHT = "must be between 0 and 120"

@pytest.mark.parametrize("value,ok", [
    pytest.param(0,   True,  id="TC_Backlight_001"),
    pytest.param(60,  True,  id="TC_Backlight_002"),
    pytest.param(120, True,  id="TC_Backlight_003"),
    pytest.param(121, False, id="TC_Backlight_004"),
])
def test_backlight(app_page, value, ok):
    nav_to_general(app_page)
    set_input(app_page, "Backlight", value)
    saved = save_and_check(app_page)
    _log.info("[TC] Backlight=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Backlight={value}: 保存失败"
        verify_modbus(REG_BACKLIGHT, value, label="LCD Backlight Time")
    else:
        assert not saved, f"Backlight={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, _ERR_BACKLIGHT)


# ════════════════════════════════════════════════════════════════════
# TC_PhaseOrder_001~002  Phase Order（ABC / ACB）
#   寄存器 4316（0x10DC）：0=ABC, 1=ACB
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("ABC", 0, id="TC_PhaseOrder_001"),
    pytest.param("ACB", 1, id="TC_PhaseOrder_002"),
])
def test_phase_order(app_page, option, encoded):
    nav_to_general(app_page)
    set_dropdown(app_page, "Phase Order", option)
    saved = save_and_check(app_page)
    assert saved, f"Phase Order={option}: 保存失败"
    verify_modbus(REG_PHASE_ORDER, encoded, label=f"PhaseOrder({option})")


# ════════════════════════════════════════════════════════════════════
# TC_EnergyFmt_001~003  Energy Reading Display Format
#   寄存器 4317（0x10DD）：0=0.1kWh, 1=0.01kWh, 2=0.001kWh
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("0.1 kWh",   0, id="TC_EnergyFmt_001"),
    pytest.param("0.01 kWh",  1, id="TC_EnergyFmt_002"),
    pytest.param("0.001 kWh", 2, id="TC_EnergyFmt_003"),
])
def test_energy_reading_format(app_page, option, encoded):
    nav_to_general(app_page)
    set_dropdown(app_page, "Energy Reading Display Format", option)
    saved = save_and_check(app_page)
    assert saved, f"Energy Reading Display Format={option}: 保存失败"
    verify_modbus(REG_ENERGY_FMT, encoded, label=f"EnergyFmt({option})")


# ════════════════════════════════════════════════════════════════════
# TC_EnergyPulse_001~005  Energy Pulse Constant（0.100–100000.000）
#   寄存器 4319（0x10DF，uint32，2 个寄存器）
#   HMI 输入值 × 1000 = 寄存器原始值
#   范围：HMI 0.100~100000.000 → 寄存器 100~100000000
# ════════════════════════════════════════════════════════════════════

_ERR_ENERGY_PULSE = "must be between"

@pytest.mark.parametrize("hmi_value,reg_value,ok", [
    pytest.param("0.100",      100,        True,  id="TC_EnergyPulse_001"),
    pytest.param("1000.000",   1000000,    True,  id="TC_EnergyPulse_002"),
    pytest.param("100000.000", 100000000,  True,  id="TC_EnergyPulse_003"),
    pytest.param("0.099",      99,         False, id="TC_EnergyPulse_004"),
    pytest.param("100000.001", 100000001,  False, id="TC_EnergyPulse_005"),
])
def test_energy_pulse_constant(app_page, hmi_value, reg_value, ok):
    nav_to_general(app_page)
    set_input(app_page, "Energy Pulse Constant", hmi_value)
    saved = save_and_check(app_page)
    _log.info("[TC] Energy Pulse Constant=%s  ok=%s  result=%s", hmi_value, ok, saved)
    if ok:
        assert saved, f"Energy Pulse Constant={hmi_value}: 保存失败"
        verify_modbus_32(REG_ENERGY_PULSE, reg_value, label=f"EnergyPulseConst({hmi_value})")
    else:
        assert not saved, f"Energy Pulse Constant={hmi_value}: 期望校验失败但保存成功"
        assert_field_error(app_page, _ERR_ENERGY_PULSE)


# ════════════════════════════════════════════════════════════════════
# TC_LedWidth_001~005  LED Pulse Width（20–100 ms）
#   寄存器 4321（0x10E1）
# ════════════════════════════════════════════════════════════════════

_ERR_LED_WIDTH = "must be between 20 and 100"

@pytest.mark.parametrize("value,ok", [
    pytest.param(20,  True,  id="TC_LedWidth_001"),
    pytest.param(60,  True,  id="TC_LedWidth_002"),
    pytest.param(100, True,  id="TC_LedWidth_003"),
    pytest.param(19,  False, id="TC_LedWidth_004"),
    pytest.param(101, False, id="TC_LedWidth_005"),
])
def test_led_pulse_width(app_page, value, ok):
    nav_to_general(app_page)
    set_input(app_page, "LED Pulse Width", value)
    saved = save_and_check(app_page)
    _log.info("[TC] LED Pulse Width=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"LED Pulse Width={value}: 保存失败"
        verify_modbus(REG_LED_PULSE_WIDTH, value, label="LED Pulse Width")
    else:
        assert not saved, f"LED Pulse Width={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, _ERR_LED_WIDTH)


# ════════════════════════════════════════════════════════════════════
# TC_MdMode_001~003  MaxDemand Clear Mode
#   寄存器 4124（0x101C）：0=Manual Reset, 1=Auto Reset, 2=Disable
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("Manual Reset", 0, id="TC_MdMode_001"),
    pytest.param("Auto Reset",   1, id="TC_MdMode_002"),
    pytest.param("Disable",      2, id="TC_MdMode_003"),
])
def test_md_clear_mode(app_page, option, encoded):
    nav_to_general(app_page)
    set_dropdown(app_page, "Mode", option)
    saved = save_and_check(app_page)
    assert saved, f"MaxDemand Clear Mode={option}: 保存失败"
    verify_modbus(REG_MD_CLEAR_MODE, encoded, label=f"MaxDemandClearMode({option})")


# ════════════════════════════════════════════════════════════════════
# TC_MdDate_001~002  MaxDemand Auto Reset Date（day 1–31）
#   寄存器 4125（0x101D）：高字节=day，低字节=hour
#   前置：MaxDemand Clear Mode = Auto Reset（否则日期字段置灰）
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("day_option,expected_day", [
    pytest.param("01", 1,  id="TC_MdDate_001"),
    pytest.param("31", 31, id="TC_MdDate_002"),
])
def test_md_auto_reset_date(app_page, day_option, expected_day):
    nav_to_general(app_page)
    step("前置：MaxDemand Clear Mode = Auto Reset")
    set_dropdown(app_page, "Mode", "Auto Reset")
    app_page.wait_for_timeout(300)
    set_dropdown(app_page, "Auto Reset Date", day_option)
    saved = save_and_check(app_page)
    assert saved, f"Auto Reset Date={day_option}: 保存失败"
    regs = modbus_read(REG_MD_CLEAR_DATE, count=1)
    actual_day = (regs[0] >> 8) & 0xFF
    _log.info("[VERIFY] AutoResetDate  expected_day=%-4s  actual_day=%-4s  reg=0x%04X  → %s",
              expected_day, actual_day, regs[0], "✓ PASS" if actual_day == expected_day else "✗ FAIL")
    assert actual_day == expected_day, (
        f"Auto Reset Date: expected day={expected_day}, got {actual_day} (reg={regs[0]})"
    )
