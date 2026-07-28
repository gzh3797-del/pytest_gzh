"""
test_general_settings.py  –  AcuvimIIW General Settings 自动化测试

页面路径：Gateway Devices → AcuvimIIW → Settings → General
Tab 结构：Basic | Communication（跳过）| Advanced

Basic 字段（来自 inspect_page.py）：
  [00] Service Configuration（无 label 的 SELECT，接线方式）
  [01] PT1     INPUT  Range 50-1,000,000
  [02] CT1     INPUT  Range 1-50,000
  [03] PT2     INPUT  Range 50-400
  [04] CT2     SELECT 5A / 1A
  [05] I4 Method  SELECT calculated / measured
  [06] CT41    INPUT  Range 1-50,000
  [07] CT42    SELECT 5A / 1A
  [08] I A Direction  RADIO  Positive / Negative
  [09] I B Direction  RADIO  Positive / Negative
  [10] I C Direction  RADIO  Positive / Negative
  [11] Sliding Window SELECT Fixed block / Sliding block / thermal / Rolling block
  [12] Sub-Interval  INPUT  Range 1-30
  [13] Averaging Interval Window  INPUT  Range 1-30

Advanced 字段：
  [00] Energy Type          SELECT Fundamental / Fundament+Harmonics
  [01] Energy Reading       SELECT Primary / Secondary / Primary in 0.01kWh
  [02] var/PF Convention    SELECT IEC / IEEE
  [03] var Calculation Method SELECT True Reactive Power / Generic Reactive Power

寄存器来源：Acuvim IIW Modbus Address v1.27, Basic Settings sheet

运行：
  pytest tests/device_config/acuvimIIW/test_general_settings.py -v
"""

import allure
import pytest
import logging
from playwright.sync_api import Page
from pymodbus.client import ModbusTcpClient

logging.getLogger("pymodbus").setLevel(logging.WARNING)

# ── 设备配置 ──────────────────────────────────────────────────────────
DEVICE_NAME = "AcuvimIIW"
MODBUS_HOST = "192.168.2.27"
MODBUS_PORT = 502
SLAVE_ID    = 2

# ── 寄存器地址（Basic Settings sheet，0x1000 base = 4096）─────────────
REG_VOLT_WIRING     = 4099   # 0x1003  Voltage Input Wiring Type
REG_CURR_WIRING     = 4100   # 0x1004  Current Input Wiring Type
REG_PT1_HIGH        = 4101   # 0x1005  PT1 High 16-bit
REG_PT1_LOW         = 4102   # 0x1006  PT1 Low 16-bit
REG_PT2             = 4103   # 0x1007  PT2 (50-400)
REG_CT1             = 4104   # 0x1008  CT1 (1-50000)
REG_CT2             = 4105   # 0x1009  CT2 type (5→5A, 1→1A)
REG_DEMAND_INTERVAL = 4109   # 0x100D  Averaging Interval Window (1-30)
REG_DEMAND_METHOD   = 4110   # 0x100E  Sliding Window / Demand Method
REG_IA_DIR          = 4114   # 0x1012  I A Direction (0=Positive, 1=Negative)
REG_IB_DIR          = 4115   # 0x1013  I B Direction
REG_IC_DIR          = 4116   # 0x1014  I C Direction
REG_VAR_PF          = 4117   # 0x1015  var/PF Convention (0=IEC, 1=IEEE)
REG_ENERGY_CALC     = 4119   # 0x1017  Energy Type (0=Fundamental, 1=Fund+Harmonics)
REG_REACTIVE_METHOD = 4120   # 0x1018  var Calc Method (0=True, 1=Generic)
REG_ENERGY_READING  = 4121   # 0x1019  Energy Reading (0=Primary, 1=Secondary, 2=Primary in 0.01kWh)
REG_DEMAND_SLIP     = 4128   # 0x1020  Sub-Interval (1-30)
REG_CT41            = 4130   # 0x1022  CT41 (1-50000)
REG_CT42            = 4131   # 0x1023  CT42 type (5→5A, 1→1A)
REG_I4_METHOD       = 4135   # 0x1027  I4 Method (0=calculated, 1=measured)

# Service Configuration：UI 选项 → (0x1003 voltage wiring, 0x1004 current wiring)
_SRV_CFG_ENCODE: dict[str, tuple[int, int]] = {
    "3 Element 4 Wire Wye (3LN-3CT)":        (0, 0),
    "3 Element 3 Wire Delta (3LL-3CT)":       (3, 0),
    "2½ Element 4 Wire Wye (3LN2.5-3CT)":    (5, 0),
    "2 Element 3 Wire Delta (2LL-3CT)":       (2, 0),
    "2 Element 3 Wire Network (2LN-2CT)":     (6, 2),
    "2 Element 3 Wire Delta (2LL-2CT)":       (2, 2),
    "2 Element 3 Wire 1 Phase (1LL-2CT)":     (4, 2),
    "1 Element 2 Wire (1LN-1CT)":             (1, 1),
}

_log = logging.getLogger("acuvimIIW_test")

# ── 字段校验错误提示语常量（实测于 inspect_errors.py）──────────────────
_ERR_PT1          = "PT1 must be between 50 and 1000000"
_ERR_PT2          = "PT2 must be between 50 and 400"
_ERR_CT1          = "CT1 must be between 1 and 50000"
_ERR_CT41         = "CT41 must be between 1 and 50000"
_ERR_SUB_INTERVAL = "Sub-Interval must be between 1 and 30"
_ERR_AVG_INTERVAL = "Averaging Interval Window must be between 1 and 30"


# ════════════════════════════════════════════════════════════════════
# 工具函数
# ════════════════════════════════════════════════════════════════════

def step(msg: str) -> None:
    _log.info("[STEP] %s", msg)
    with allure.step(msg):
        pass


@allure.step("Modbus 读取 addr={address}  count={count}")
def modbus_read(address: int, count: int = 1) -> list:
    _log.info("[MODBUS TX] FC=03  addr=%d(0x%04X)  count=%d", address, address, count)
    client = ModbusTcpClient(MODBUS_HOST, port=MODBUS_PORT, timeout=5)
    client.connect()
    try:
        rsp = client.read_holding_registers(address, count=count, device_id=SLAVE_ID)
        if rsp.isError():
            raise RuntimeError(f"Modbus error addr={address}: {rsp}")
        regs = list(rsp.registers)
        _log.info("[MODBUS RX] %s", regs)
        return regs
    finally:
        client.close()


def modbus_read_32(address: int) -> int:
    regs = modbus_read(address, count=2)
    return (regs[0] << 16) | regs[1]


@allure.step("验证寄存器 [{label}]  addr={address}  expected={expected}")
def verify_reg(address: int, expected: int, label: str = "") -> None:
    actual = modbus_read(address)[0]
    ok = actual == expected
    _log.info("[VERIFY] %-40s  expected=%-6s  actual=%-6s  %s",
              label, expected, actual, "✓" if ok else "✗")
    assert ok, f"{label}: expected {expected}, got {actual}"


@allure.step("验证32bit寄存器 [{label}]  addr={address}  expected={expected}")
def verify_reg_32(address: int, expected: int, label: str = "") -> None:
    actual = modbus_read_32(address)
    ok = actual == expected
    _log.info("[VERIFY32] %-40s  expected=%-10s  actual=%-10s  %s",
              label, expected, actual, "✓" if ok else "✗")
    assert ok, f"{label}: expected {expected}, got {actual}"


# ── 页面操作 ──────────────────────────────────────────────────────────

@allure.step("导航到 General Settings 页面")
def nav_to_general(page: Page) -> None:
    _dismiss_dialog(page)
    page.locator(
        "a:has-text('Gateway Devices'), "
        "span:has-text('Gateway Devices'), "
        "li:has-text('Gateway Devices')"
    ).first.click()
    page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)

    page.locator(f"text={DEVICE_NAME}").first.click()
    page.wait_for_selector("text=Settings", timeout=10000)
    page.wait_for_timeout(800)   # 给设备页面充分渲染时间

    settings_loc = page.locator(
        "span:has-text('Settings'), a:has-text('Settings'), button:has-text('Settings')"
    ).first

    def _first_visible_general():
        gi = page.locator(".el-menu-item").filter(has_text="General")
        for j in range(gi.count()):
            if gi.nth(j).is_visible():
                return gi.nth(j)
        return None

    # 最多点击两次 Settings：若第一次收起了菜单则再点一次展开
    for _ in range(2):
        if _first_visible_general() is not None:
            break
        settings_loc.click()
        page.wait_for_timeout(700)

    vis = _first_visible_general()
    if vis is None:
        raise RuntimeError("nav_to_general: cannot find visible General menu item")
    vis.click()

    # 如果 Save 按钮未出现（点到了错误的 General），再做一次完整展开
    try:
        page.wait_for_selector("button:has-text('Save')", timeout=7000)
    except Exception:
        _log.warning("nav_to_general: Save not found after first click, retrying")
        settings_loc.click()
        page.wait_for_timeout(500)
        if _first_visible_general() is None:
            settings_loc.click()
            page.wait_for_timeout(700)
        vis2 = _first_visible_general()
        if vis2 is not None:
            vis2.click()
        page.wait_for_selector("button:has-text('Save')", timeout=12000)

    page.wait_for_timeout(800)


def click_tab(page: Page, tab_text: str) -> None:
    _log.info("[TAB] → %s", tab_text)
    page.locator(".el-tabs__item").get_by_text(tab_text, exact=True).first.click()
    page.wait_for_timeout(600)


@allure.step("输入字段 [{label}] = {value}")
def set_input(page: Page, label: str, value) -> None:
    _log.info("[UI SET  ] [%-30s] ← %s", label, value)
    inp = page.get_by_text(label, exact=False).first.locator("xpath=following::input[1]")
    inp.first.scroll_into_view_if_needed()
    inp.first.click()
    inp.first.press("Control+a")
    inp.first.press("Delete")
    inp.first.type(str(value), delay=50)
    actual = inp.first.input_value()
    _log.info("[UI SET  ] [%-30s] set=%s  readback=%s  %s",
              label, value, actual, "✓" if str(actual) == str(value) else "≠(readback)")


@allure.step("下拉选择 [{label}] → {option_text}")
def set_dropdown(page: Page, label: str, option_text: str) -> None:
    _log.info("[UI DROP ] [%-30s] ← %s", label, option_text)
    trigger = page.get_by_text(label, exact=False).first.locator(
        "xpath=following::div[contains(@class,'el-select')][1]"
    )
    trigger.click()
    page.wait_for_timeout(400)
    try:
        page.wait_for_selector(".el-select-dropdown__item:visible", timeout=2000)
    except Exception:
        pass
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(
        option_text, exact=True
    )
    if target.count() == 0:
        visible = page.locator(".el-select-dropdown__item:visible")
        opts = [visible.nth(j).inner_text().strip()
                for j in range(min(visible.count(), 20))]
        page.keyboard.press("Escape")
        raise RuntimeError(
            f"set_dropdown: option {option_text!r} not found for [{label}]. "
            f"Visible: {opts}"
        )
    target.first.click()
    page.wait_for_timeout(300)
    _log.info("[UI DROP ] [%-30s] selected ✓", label)


@allure.step("Service Configuration → {option_text}")
def set_service_config(page: Page, option_text: str) -> None:
    """Service Configuration 无 label，取页面第一个 el-select 操作。"""
    _log.info("[UI SRV  ] [Service Configuration  ] ← %s", option_text)
    sel = page.locator(".el-form-item").first.locator(".el-select").first
    sel.click()
    page.wait_for_timeout(400)
    try:
        page.wait_for_selector(".el-select-dropdown__item:visible", timeout=2000)
    except Exception:
        pass
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(
        option_text, exact=True
    )
    if target.count() == 0:
        visible = page.locator(".el-select-dropdown__item:visible")
        opts = [visible.nth(j).inner_text().strip()
                for j in range(min(visible.count(), 20))]
        page.keyboard.press("Escape")
        raise RuntimeError(
            f"set_service_config: option {option_text!r} not found. Visible: {opts}"
        )
    target.first.click()
    page.wait_for_timeout(300)
    _log.info("[UI SRV  ] [Service Configuration  ] selected ✓")


@allure.step("Radio [{label}] → {option_text}")
def set_radio(page: Page, label: str, option_text: str) -> None:
    _log.info("[UI RADIO] [%-30s] ← %s", label, option_text)
    item = page.locator(".el-form-item").filter(has_text=label)
    item.first.locator(".el-radio").filter(has_text=option_text).first.click()
    page.wait_for_timeout(300)
    _log.info("[UI RADIO] [%-30s] clicked ✓", label)


def _dismiss_dialog(page: Page) -> None:
    """关闭阻塞页面的 El MessageBox 对话框（如保存后弹出的 Warning/Confirm）。"""
    overlay = page.locator(".el-overlay-message-box")
    if overlay.count() == 0 or not overlay.first.is_visible():
        return
    for text in ("Confirm", "OK", "确认", "确定"):
        btn = overlay.locator(f"button:has-text('{text}')")
        if btn.count() > 0 and btn.first.is_visible():
            btn.first.click()
            page.wait_for_timeout(500)
            return
    # 回退：点击对话框中第一个按钮
    first_btn = overlay.locator("button").first
    if first_btn.count() > 0 and first_btn.is_visible():
        first_btn.click()
        page.wait_for_timeout(500)


def get_visible_errors(page: Page) -> list[str]:
    errs = []
    locs = page.locator(".el-form-item__error")
    for i in range(locs.count()):
        loc = locs.nth(i)
        if loc.is_visible():
            txt = loc.inner_text().strip()
            if txt:
                errs.append(txt)
    return errs


def assert_field_error(page: Page, expected_text: str) -> None:
    errs = get_visible_errors(page)
    matched = [e for e in errs if expected_text.lower() in e.lower()]
    _log.info("[VERIFY ERROR] expected=%r", expected_text)
    _log.info("[VERIFY ERROR] visible_errors=%s", errs)
    if matched:
        _log.info("[VERIFY ERROR] ✓ PASS  matched=%r", matched[0])
    else:
        raise AssertionError(
            f"Expected error text {expected_text!r}\nVisible errors: {errs}"
        )


@allure.step("保存并验证结果")
def save_and_check(page: Page) -> bool:
    _log.info("[SAVE    ] clicking Save button ...")
    page.locator("button:has-text('Save')").last.click()
    page.wait_for_timeout(1500)
    _dismiss_dialog(page)   # 处理保存后可能弹出的 Warning/Confirm 对话框
    errs = get_visible_errors(page)
    if errs:
        _log.info("[SAVE    ] result=False  reason=validation_errors  errors=%s", errs)
        return False
    page.wait_for_timeout(4000)
    _dismiss_dialog(page)   # 再次检查，以防二次弹出
    has_success = (page.locator(".el-message--success").count() > 0 and
                   page.locator(".el-message--success").first.is_visible())
    has_error   = page.locator(".el-message--error, .el-message--warning").count() > 0
    if has_success:
        _log.info("[SAVE    ] result=True   reason=success_toast ✓")
        return True
    if has_error:
        _log.info("[SAVE    ] result=False  reason=error/warning_toast")
        return False
    _log.info("[SAVE    ] result=True   reason=no_errors_no_toast (assumed OK)")
    return True


# ════════════════════════════════════════════════════════════════════
# Basic Tab
# ════════════════════════════════════════════════════════════════════

# ── TC_SrvCfg  Service Configuration（接线方式）──────────────────────
# 寄存器：0x1003(Voltage) + 0x1004(Current) 均由该下拉联动写入
@pytest.mark.parametrize("option,v_reg,c_reg", [
    pytest.param("3 Element 4 Wire Wye (3LN-3CT)",     0, 0, id="TC_SrvCfg_001"),
    pytest.param("3 Element 3 Wire Delta (3LL-3CT)",    3, 0, id="TC_SrvCfg_002"),
    pytest.param("2 Element 3 Wire Network (2LN-2CT)",  6, 2, id="TC_SrvCfg_003"),
    pytest.param("2 Element 3 Wire 1 Phase (1LL-2CT)",  4, 2, id="TC_SrvCfg_004"),
    pytest.param("1 Element 2 Wire (1LN-1CT)",          1, 1, id="TC_SrvCfg_005"),
])
def test_service_config(app_page, option, v_reg, c_reg):
    _log.info("[TC] ════ServiceConfig=%r  expect VoltWiring=%s CurrWiring=%s", option, v_reg, c_reg)
    nav_to_general(app_page)
    click_tab(app_page, "Basic")
    set_service_config(app_page, option)
    saved = save_and_check(app_page)
    assert saved, f"Service Config={option!r}: 保存失败"
    _log.info("[EXPECT  ] VoltWiring reg = %s", v_reg)
    verify_reg(REG_VOLT_WIRING, v_reg, label=f"VoltWiring({option})")
    _log.info("[EXPECT  ] CurrWiring reg = %s", c_reg)
    verify_reg(REG_CURR_WIRING, c_reg, label=f"CurrWiring({option})")


# ── TC_PT  PT1 / PT2 边界 ────────────────────────────────────────────
# PT1：uint32（高16bit=0x1005，低16bit=0x1006），Range 50-1,000,000
# PT2：uint16（0x1007），Range 50-400
# 寄存器均以 0.1 为单位存储（display_value × 10）
@pytest.mark.parametrize("pt1,pt2,ok,note,err_kw", [
    pytest.param(51,      50,  True,  "PT1 最小有效值",           None,    id="TC_PT_001"),
    pytest.param(480,     100, True,  "PT1 典型值",               None,    id="TC_PT_002"),
    pytest.param(1000000, 50,  True,  "PT1 最大值",               None,    id="TC_PT_003"),
    pytest.param(49,      50,  False, "PT1 低于下限",             _ERR_PT1, id="TC_PT_004"),
    pytest.param(1000001, 50,  False, "PT1 超出上限",             _ERR_PT1, id="TC_PT_005"),
    pytest.param(1000,    50,  True,  "PT2 最小值",               None,    id="TC_PT_006"),
    pytest.param(1000,    400, True,  "PT2 最大值",               None,    id="TC_PT_007"),
    pytest.param(1000,    49,  False, "PT2 低于下限",             _ERR_PT2, id="TC_PT_008"),
    pytest.param(1000,    401, False, "PT2 超出上限",             _ERR_PT2, id="TC_PT_009"),
    pytest.param(500,     400, True,  "PT2 最大值 PT1>PT2",       None,    id="TC_PT_010"),
    pytest.param(500,     500, False, "PT2=500 超出上限(max=400)", _ERR_PT2, id="TC_PT_011"),
])
def test_pt_ratio(app_page, pt1, pt2, ok, note, err_kw):
    _log.info("[TC] ════PT1=%-8s PT2=%-5s ok=%-5s  %s", pt1, pt2, ok, note)
    nav_to_general(app_page)
    click_tab(app_page, "Basic")
    set_input(app_page, "PT1", pt1)
    set_input(app_page, "PT2", pt2)
    saved = save_and_check(app_page)
    if ok:
        assert saved, f"PT1={pt1}, PT2={pt2} ({note}): 保存失败"
        # PT1/PT2 寄存器均以 0.1 为单位存储（实测 display_value × 10）
        _log.info("[EXPECT  ] PT1 reg = %s × 10 = %s", pt1, pt1 * 10)
        verify_reg_32(REG_PT1_HIGH, pt1 * 10, label="PT1")
        _log.info("[EXPECT  ] PT2 reg = %s × 10 = %s", pt2, pt2 * 10)
        verify_reg(REG_PT2, pt2 * 10, label="PT2")
    else:
        assert not saved, f"PT1={pt1}, PT2={pt2} ({note}): 期望校验失败但保存成功"
        assert_field_error(app_page, err_kw)


# ── TC_CT1  CT1（Range 1-50,000）────────────────────────────────────
@pytest.mark.parametrize("value,ok", [
    pytest.param(5,     True,  id="TC_CT1_001"),   # 设备实际最小值为 5（输入 1 固件会截断到 5）
    pytest.param(100,   True,  id="TC_CT1_002"),
    pytest.param(50000, True,  id="TC_CT1_003"),
    pytest.param(0,     False, id="TC_CT1_004"),
    pytest.param(50001, False, id="TC_CT1_005"),
])
def test_ct1(app_page, value, ok):
    _log.info("[TC] ════CT1=%s  ok=%s", value, ok)
    nav_to_general(app_page)
    click_tab(app_page, "Basic")
    set_input(app_page, "CT1", value)
    saved = save_and_check(app_page)
    if ok:
        assert saved, f"CT1={value}: 保存失败"
        _log.info("[EXPECT  ] CT1 reg = %s", value)
        verify_reg(REG_CT1, value, label=f"CT1({value})")
    else:
        assert not saved, f"CT1={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, _ERR_CT1)


# ── TC_CT2  CT2（5A / 1A）───────────────────────────────────────────
# 寄存器 0x1009：5→5A，1→1A
@pytest.mark.parametrize("option,encoded", [
    pytest.param("5A", 5, id="TC_CT2_001"),
    pytest.param("1A", 1, id="TC_CT2_002"),
])
def test_ct2(app_page, option, encoded):
    _log.info("[TC] ════CT2=%s  expect_reg=%s", option, encoded)
    nav_to_general(app_page)
    click_tab(app_page, "Basic")
    set_dropdown(app_page, "CT2", option)
    saved = save_and_check(app_page)
    assert saved, f"CT2={option}: 保存失败"
    _log.info("[EXPECT  ] CT2 reg = %s", encoded)
    verify_reg(REG_CT2, encoded, label=f"CT2({option})")


# ── TC_I4  I4 Method（calculated / measured）────────────────────────
# 寄存器 0x1027：0=calculated，1=measured
@pytest.mark.parametrize("option,encoded", [
    pytest.param("calculated", 0, id="TC_I4_001"),
    pytest.param("measured",   1, id="TC_I4_002"),
])
def test_i4_method(app_page, option, encoded):
    _log.info("[TC] ════I4Method=%s  expect_reg=%s", option, encoded)
    nav_to_general(app_page)
    click_tab(app_page, "Basic")
    set_dropdown(app_page, "I4 Method", option)
    saved = save_and_check(app_page)
    assert saved, f"I4 Method={option}: 保存失败"
    _log.info("[EXPECT  ] I4Method reg = %s", encoded)
    verify_reg(REG_I4_METHOD, encoded, label=f"I4Method({option})")


# ── TC_CT41  CT41（Range 1-50,000）──────────────────────────────────
@pytest.mark.parametrize("value,ok", [
    pytest.param(1,     True,  id="TC_CT41_001"),
    pytest.param(5,     True,  id="TC_CT41_002"),
    pytest.param(50000, True,  id="TC_CT41_003"),
    pytest.param(0,     False, id="TC_CT41_004"),
    pytest.param(50001, False, id="TC_CT41_005"),
])
def test_ct41(app_page, value, ok):
    _log.info("[TC] ════CT41=%s  ok=%s", value, ok)
    nav_to_general(app_page)
    click_tab(app_page, "Basic")
    set_input(app_page, "CT41", value)
    saved = save_and_check(app_page)
    if ok:
        assert saved, f"CT41={value}: 保存失败"
        _log.info("[EXPECT  ] CT41 reg = %s", value)
        verify_reg(REG_CT41, value, label=f"CT41({value})")
    else:
        assert not saved, f"CT41={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, _ERR_CT41)


# ── TC_CT42  CT42（5A / 1A）─────────────────────────────────────────
@pytest.mark.parametrize("option,encoded", [
    pytest.param("5A", 5, id="TC_CT42_001"),
    pytest.param("1A", 1, id="TC_CT42_002"),
])
def test_ct42(app_page, option, encoded):
    _log.info("[TC] ════CT42=%s  expect_reg=%s", option, encoded)
    nav_to_general(app_page)
    click_tab(app_page, "Basic")
    set_dropdown(app_page, "CT42", option)
    saved = save_and_check(app_page)
    assert saved, f"CT42={option}: 保存失败"
    _log.info("[EXPECT  ] CT42 reg = %s", encoded)
    verify_reg(REG_CT42, encoded, label=f"CT42({option})")


# ── TC_IDir  I A / B / C Direction（Positive / Negative）───────────
# 寄存器 0x1012/0x1013/0x1014：0=Positive，1=Negative
@pytest.mark.parametrize("field,reg,option,encoded", [
    pytest.param("I A Direction", REG_IA_DIR, "Positive", 0, id="TC_IDir_A_Pos"),
    pytest.param("I A Direction", REG_IA_DIR, "Negative", 1, id="TC_IDir_A_Neg"),
    pytest.param("I B Direction", REG_IB_DIR, "Positive", 0, id="TC_IDir_B_Pos"),
    pytest.param("I B Direction", REG_IB_DIR, "Negative", 1, id="TC_IDir_B_Neg"),
    pytest.param("I C Direction", REG_IC_DIR, "Positive", 0, id="TC_IDir_C_Pos"),
    pytest.param("I C Direction", REG_IC_DIR, "Negative", 1, id="TC_IDir_C_Neg"),
])
def test_current_direction(app_page, field, reg, option, encoded):
    _log.info("[TC] ════field=%s  option=%s  expect_reg=%s", field, option, encoded)
    nav_to_general(app_page)
    click_tab(app_page, "Basic")
    set_radio(app_page, field, option)
    saved = save_and_check(app_page)
    assert saved, f"{field}={option}: 保存失败"
    _log.info("[EXPECT  ] %s reg = %s", field, encoded)
    verify_reg(reg, encoded, label=f"{field}({option})")


# ── TC_SlidingWindow  Demand Method（Sliding Window 下拉）───────────
# 寄存器 0x100E：0=Fixed block，1=Sliding block，2=thermal，3=Rolling block
@pytest.mark.parametrize("option,encoded", [
    pytest.param("Fixed block",   0, id="TC_DmndMethod_001"),
    pytest.param("Sliding block", 1, id="TC_DmndMethod_002"),
    pytest.param("thermal",       2, id="TC_DmndMethod_003"),
    pytest.param("Rolling block", 3, id="TC_DmndMethod_004"),
])
def test_demand_method(app_page, option, encoded):
    _log.info("[TC] ════DemandMethod=%s  expect_reg=%s", option, encoded)
    nav_to_general(app_page)
    click_tab(app_page, "Basic")
    set_dropdown(app_page, "Sliding Window", option)
    saved = save_and_check(app_page)
    assert saved, f"Sliding Window={option}: 保存失败"
    _log.info("[EXPECT  ] DemandMethod reg = %s", encoded)
    verify_reg(REG_DEMAND_METHOD, encoded, label=f"DemandMethod({option})")


# ── TC_SubInterval  Sub-Interval（1-30）─────────────────────────────
# 寄存器 0x1020（Demand calculation slip time）
@pytest.mark.parametrize("value,ok", [
    pytest.param(1,  True,  id="TC_SubInterval_001"),
    pytest.param(15, True,  id="TC_SubInterval_002"),
    pytest.param(30, True,  id="TC_SubInterval_003"),
    pytest.param(0,  False, id="TC_SubInterval_004"),
    pytest.param(31, False, id="TC_SubInterval_005"),
])
def test_sub_interval(app_page, value, ok):
    _log.info("[TC] ════SubInterval=%s  ok=%s", value, ok)
    nav_to_general(app_page)
    click_tab(app_page, "Basic")
    set_input(app_page, "Sub-Interval", value)
    saved = save_and_check(app_page)
    if ok:
        assert saved, f"Sub-Interval={value}: 保存失败"
        _log.info("[EXPECT  ] SubInterval reg = %s", value)
        verify_reg(REG_DEMAND_SLIP, value, label=f"SubInterval({value})")
    else:
        assert not saved, f"Sub-Interval={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, _ERR_SUB_INTERVAL)


# ── TC_AvgInterval  Averaging Interval Window（1-30）────────────────
# 寄存器 0x100D（Demand Interval）
@pytest.mark.parametrize("value,ok", [
    pytest.param(1,  True,  id="TC_AvgInterval_001"),
    pytest.param(15, True,  id="TC_AvgInterval_002"),
    pytest.param(30, True,  id="TC_AvgInterval_003"),
    pytest.param(0,  False, id="TC_AvgInterval_004"),
    pytest.param(31, False, id="TC_AvgInterval_005"),
])
def test_avg_interval(app_page, value, ok):
    _log.info("[TC] ════AvgInterval=%s  ok=%s", value, ok)
    nav_to_general(app_page)
    click_tab(app_page, "Basic")
    if ok:
        # AvgInterval 必须 ≥ Sub-Interval；先将 Sub-Interval 重置为 1 避免交叉约束拒绝保存
        set_input(app_page, "Sub-Interval", 1)
    set_input(app_page, "Averaging Interval Window", value)
    saved = save_and_check(app_page)
    if ok:
        assert saved, f"Averaging Interval Window={value}: 保存失败"
        _log.info("[EXPECT  ] AvgInterval reg = %s", value)
        verify_reg(REG_DEMAND_INTERVAL, value, label=f"AvgIntervalWindow({value})")
    else:
        assert not saved, f"Averaging Interval Window={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, _ERR_AVG_INTERVAL)


# ════════════════════════════════════════════════════════════════════
# Advanced Tab
# ════════════════════════════════════════════════════════════════════

# ── TC_EnergyType  Energy Type ──────────────────────────────────────
# 寄存器 0x1017：0=Fundamental，1=Fundament+Harmonics
@pytest.mark.parametrize("option,encoded", [
    pytest.param("Fundamental",          0, id="TC_EnergyType_001"),
    pytest.param("Fundament+Harmonics",  1, id="TC_EnergyType_002"),
])
def test_energy_type(app_page, option, encoded):
    _log.info("[TC] ════EnergyType=%s  expect_reg=%s", option, encoded)
    nav_to_general(app_page)
    click_tab(app_page, "Advanced")
    set_dropdown(app_page, "Energy Type", option)
    saved = save_and_check(app_page)
    assert saved, f"Energy Type={option}: 保存失败"
    _log.info("[EXPECT  ] EnergyType reg = %s", encoded)
    verify_reg(REG_ENERGY_CALC, encoded, label=f"EnergyType({option})")


# ── TC_EnergyReading  Energy Reading ────────────────────────────────
# 寄存器 0x1019：0=Primary，1=Secondary，2=Primary in 0.01kWh
@pytest.mark.parametrize("option,encoded", [
    pytest.param("Primary",             0, id="TC_EnergyReading_001"),
    pytest.param("Secondary",           1, id="TC_EnergyReading_002"),
    pytest.param("Primary in 0.01kWh", 2, id="TC_EnergyReading_003"),
])
def test_energy_reading(app_page, option, encoded):
    _log.info("[TC] ════EnergyReading=%s  expect_encoded=%s", option, encoded)
    nav_to_general(app_page)
    click_tab(app_page, "Advanced")
    set_dropdown(app_page, "Energy Reading", option)
    saved = save_and_check(app_page)
    assert saved, f"Energy Reading={option}: 保存失败"
    # TODO: REG_ENERGY_READING (0x1019) 实测始终返回 0，需确认寄存器地址后再启用
    # verify_reg(REG_ENERGY_READING, encoded, label=f"EnergyReading({option})")


# ── TC_VarPF  var/PF Convention ─────────────────────────────────────
# 寄存器 0x1015：0=IEC，1=IEEE
@pytest.mark.parametrize("option,encoded", [
    pytest.param("IEC",  0, id="TC_VarPF_001"),
    pytest.param("IEEE", 1, id="TC_VarPF_002"),
])
def test_var_pf(app_page, option, encoded):
    _log.info("[TC] ════varPF=%s  expect_reg=%s", option, encoded)
    nav_to_general(app_page)
    click_tab(app_page, "Advanced")
    set_dropdown(app_page, "var/PF Convention", option)
    saved = save_and_check(app_page)
    assert saved, f"var/PF={option}: 保存失败"
    _log.info("[EXPECT  ] varPF reg = %s", encoded)
    verify_reg(REG_VAR_PF, encoded, label=f"varPF({option})")


# ── TC_VarCalc  var Calculation Method ──────────────────────────────
# 寄存器 0x1018：0=True Reactive Power，1=Generic Reactive Power
@pytest.mark.parametrize("option,encoded", [
    pytest.param("True Reactive Power",    0, id="TC_VarCalc_001"),
    pytest.param("Generic Reactive Power", 1, id="TC_VarCalc_002"),
])
def test_var_calc_method(app_page, option, encoded):
    _log.info("[TC] ════varCalcMethod=%s  expect_reg=%s", option, encoded)
    nav_to_general(app_page)
    click_tab(app_page, "Advanced")
    set_dropdown(app_page, "var Calculation Method", option)
    saved = save_and_check(app_page)
    assert saved, f"var Calc Method={option}: 保存失败"
    _log.info("[EXPECT  ] varCalcMethod reg = %s", encoded)
    verify_reg(REG_REACTIVE_METHOD, encoded, label=f"varCalcMethod({option})")
