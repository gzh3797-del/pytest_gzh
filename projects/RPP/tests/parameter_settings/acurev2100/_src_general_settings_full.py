"""
test_general_settings_full.py  –  AcuRev2100 General Settings 自动化测试

34 个用例：TC_Basic_001~030 + TC_Advanced_001~004

运行命令：pytest test_general_settings_full.py -v --tb=short

v4 改动：
  - save_and_check(): 先 1.5s 检测内联校验错误，无错误再等 4.5s 检测 toast
  - get_visible_errors(): 读取 .el-form-item__error 的实际错误文字
  - assert_field_error(): 断言错误文字，同时打印预期/实际对比日志
  - 所有异常边界用例增加 expected_error 参数，verify 错误文本内容
  - TC_Basic_023/026 根据设备约束（Avg ≥ 2×Sub）修正测试值
  - nav_to_general Advanced 标签等待正确元素
"""

import pytest
import logging
from playwright.sync_api import Page
from pymodbus.client import ModbusTcpClient

logging.getLogger("pymodbus").setLevel(logging.DEBUG)

# ── 配置 ───────────────────────────────────────────────────────────
DEVICE_NAME = "AcuRev2100"
MODBUS_HOST = "192.168.2.64"
MODBUS_PORT = 502
SLAVE_ID    = 101

# ── 寄存器地址 ─────────────────────────────────────────────────────
REG_DEMAND_METHOD   = 2058
REG_DEMAND_INTERVAL = 2059
REG_DEMAND_SUB      = 2060
REG_CT_TYPE         = 2061
REG_VAR_METHOD      = 2101
REG_VAR_PF          = 2102
REG_RATED_VOLTAGE   = 2183
REG_WIRING          = 2184
REG_CT_FS_BASE      = 2185   # Ch1~18 at 2185~2202

# ── 编码映射 ───────────────────────────────────────────────────────
WIRING_ENCODE = {
    "1 Element 2 Wire (1LN)":                  0,
    "3 Element 4 Wire Y (3LN)":                1,
    "2 Element 3 Wire 1 Phase (2LN)":          2,
    "3 Element 4 Wire Delta (High-leg delta)":  3,
    "2 Element 3 Wire Delta (2LL)":            4,
    "2 Element 3 Wire Network (2LN/Network)":  5,
}
CT2_ENCODE        = {"RCT": 120, "333mV": 333}
DEMAND_ENCODE     = {"Rolling Window Demand": 1, "Fixed Window Demand": 2}
VAR_METHOD_ENCODE = {"Method 1 (True)": 0, "Method 2 (Generalized)": 1}
VAR_PF_ENCODE     = {"IEC": 0, "IEEE": 1}

CT_LABELS = [f"Input Channel {i}" for i in range(1, 19)]

_log = logging.getLogger("acurev_test")

# ── 寄存器描述表（用于日志输出）────────────────────────────────────
_REG_DESC: dict[int, str] = {
    2058: "Demand Calculation Method",
    2059: "Demand Interval",
    2060: "Demand Sub-Interval",
    2061: "CT Type",
    2101: "VAR Calculation Method",
    2102: "VAR/PF Convention",
    2183: "Rated Voltage",
    2184: "Wiring Type",
    **{2185 + i: f"CT {i+1} Full-scale" for i in range(18)},
}


def _reg_desc(address: int) -> str:
    return _REG_DESC.get(address, "")


# ════════════════════════════════════════════════════════════════════
# 工具函数
# ════════════════════════════════════════════════════════════════════

def step(msg: str) -> None:
    _log.info("[STEP] %s", msg)


# ── Modbus ────────────────────────────────────────────────────────

def modbus_read(address: int, count: int = 1) -> list:
    """FC=03 读寄存器，打印完整 TX/RX 帧。"""
    tx_pdu  = bytes([0x03, address >> 8, address & 0xFF, 0x00, count])
    tx_mbap = bytes([0x00, 0x01, 0x00, 0x00, 0x00, 0x06, SLAVE_ID])
    desc = _reg_desc(address)
    _log.info("[MODBUS TX] FC=03  addr=%d(0x%04X)  count=%d%s",
              address, address, count,
              f"  [{desc}]" if desc else "")
    _log.info("            %s", " ".join(f"{b:02X}" for b in tx_mbap + tx_pdu))

    client = ModbusTcpClient(MODBUS_HOST, port=MODBUS_PORT, timeout=5)
    client.connect()
    try:
        rsp = client.read_holding_registers(address, count=count, device_id=SLAVE_ID)
        if rsp.isError():
            _log.error("[MODBUS RX] ERROR: %s", rsp)
            raise RuntimeError(f"Modbus FC=03 error at addr={address}: {rsp}")
        regs = list(rsp.registers)
        data_b = [b for r in regs for b in (r >> 8, r & 0xFF)]
        rx_pdu  = bytes([0x03, count * 2] + data_b)
        rx_mbap = bytes([0x00, 0x01, 0x00, 0x00, 0x00, len(rx_pdu) + 1, SLAVE_ID])
        _log.info("[MODBUS RX] registers=%s", regs)
        _log.info("            %s", " ".join(f"{b:02X}" for b in rx_mbap + rx_pdu))
        return regs
    finally:
        client.close()


def verify_modbus(address: int, expected, label: str = "", count: int = 1) -> None:
    """读寄存器并断言，打印 expected / actual 对比日志。"""
    regs = modbus_read(address, count)
    if count == 1:
        actual = regs[0]
        ok = actual == expected
        _log.info("[VERIFY] %-35s  expected=%-6s  actual=%-6s  → %s",
                  label, expected, actual, "✓ PASS" if ok else "✗ FAIL")
        assert ok, f"{label}: Modbus expected {expected}, got {actual}"
    else:
        all_ok = True
        for i, r in enumerate(regs):
            ok = r == expected
            _log.info("[VERIFY] %s Ch%-2d  expected=%-6s  actual=%-6s  → %s",
                      label, i + 1, expected, r, "✓" if ok else "✗")
            if not ok:
                all_ok = False
        assert all_ok, f"{label}: one or more channels mismatch (expected {expected})"


# ── 页面操作 ──────────────────────────────────────────────────────

def nav_to_general(page: Page, tab: str = "Basic") -> None:
    """Gateway Devices → AcuRev2100 → Settings → General → tab。
    Basic 标签等待 'Rated Voltage'，Advanced 标签等待 'VAR Calculation Method'。
    """
    step("Click Gateway Devices")
    page.locator(
        "a:has-text('Gateway Devices'), "
        "span:has-text('Gateway Devices'), "
        "li:has-text('Gateway Devices')"
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
    page.wait_for_selector(
        f"button:has-text('{tab}'), a:has-text('{tab}'), [role='tab']:has-text('{tab}')",
        timeout=10000,
    )
    step(f"Switch to tab: {tab}")
    page.locator(
        f"button:has-text('{tab}'), a:has-text('{tab}'), [role='tab']:has-text('{tab}')"
    ).first.click()
    # 等待标签页对应的首个 section 出现
    sentinel = "VAR Calculation Method" if tab == "Advanced" else "Rated Voltage"
    page.wait_for_selector(f"text={sentinel}", timeout=8000)
    step(f"Ready: General / {tab}")


def set_input(page: Page, label: str, value) -> None:
    step(f"Set input [{label}] = {value!r}")
    label_els = page.get_by_text(label, exact=False)
    n = label_els.count()
    _log.info("  label matches: %d", n)
    if n == 0:
        raise RuntimeError(f"set_input: label {label!r} not found")
    inp = label_els.first.locator("xpath=following::input[1]")
    if inp.count() == 0:
        raise RuntimeError(f"set_input: no input found after {label!r}")
    inp.first.scroll_into_view_if_needed()
    inp.first.click()
    inp.first.press("Control+a")
    inp.first.press("Delete")
    inp.first.type(str(value), delay=50)
    _log.info("  typed %r", value)


def set_dropdown(page: Page, label: str, option_text: str) -> None:
    step(f"Dropdown [{label}] → {option_text!r}")
    trigger = page.get_by_text(label, exact=False).first.locator(
        "xpath=following::div[contains(@class,'el-select')][1]"
    )
    trigger.click()
    page.wait_for_timeout(600)
    page.locator(".el-select-dropdown__item").get_by_text(
        option_text, exact=True
    ).click()
    page.wait_for_timeout(400)


def get_visible_errors(page: Page) -> list[str]:
    """返回当前页面所有可见的内联校验错误文字列表。

    截图确认：错误文字显示在 .el-form-item__error 中，
    格式为 '{Field} must be between {min} and {max}'。
    """
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
    """断言指定的错误提示文字可见，并打印预期/实际对比日志。

    expected_text: 错误文字的完整或部分内容（不区分大小写子串匹配）
    例如: "Rated Voltage must be between 10 and 400"
    """
    errors = get_visible_errors(page)
    matched = [e for e in errors if expected_text.lower() in e.lower()]
    _log.info("[VERIFY ERROR] expected=%r", expected_text)
    _log.info("[VERIFY ERROR] visible_errors=%s", errors)
    if matched:
        _log.info("[VERIFY ERROR] ✓ PASS  matched=%r", matched[0])
    else:
        _log.error("[VERIFY ERROR] ✗ FAIL  expected text not found in any error message")
        raise AssertionError(
            f"Expected error text {expected_text!r}\n"
            f"Visible errors on page: {errors}"
        )


def save_and_check(page: Page) -> bool:
    """点击 Save，检测保存结果。

    检测策略（顺序）：
    1. 点击后等待 1.5s → 读取内联校验错误（.el-form-item__error）
       如有错误 → 立刻返回 False（表单校验失败，数据未写入设备）
    2. 无内联错误 → 再等 4.5s 让设备处理 → 检测 toast 通知
       - .el-message--success 出现 → True
       - .el-message--error / warning 出现 → False
    3. 无 toast → 视为静默保存成功 → True
    """
    step("Click Save")
    page.locator("button:has-text('Save')").last.click()
    page.wait_for_timeout(1500)

    # 1. 立刻检查内联表单校验错误
    form_errors = get_visible_errors(page)
    if form_errors:
        _log.info("[SAVE] FORM VALIDATION FAILED  errors=%s", form_errors)
        return False

    # 2. 等待设备处理并检测 toast
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

    _log.info("[SAVE] No toast detected → assuming success")
    return True


def sub_interval_is_disabled(page: Page) -> bool:
    inp = page.get_by_text("Sub-Interval", exact=False).first.locator(
        "xpath=following::input[1]"
    )
    cls = inp.evaluate(
        "el => el.closest('.el-input') ? el.closest('.el-input').className : ''"
    )
    disabled = "is-disabled" in str(cls)
    _log.info("[LINKAGE] Sub-Interval el-input class=%r  disabled=%s", cls, disabled)
    return disabled


# ════════════════════════════════════════════════════════════════════
# TC_Basic_001 ~ TC_Basic_005  Rated Voltage  (范围 10-400)
#
# 异常断言：
#   点击 Save 后 .el-form-item__error 出现并显示：
#   "Rated Voltage must be between 10 and 400"
# ════════════════════════════════════════════════════════════════════

_ERR_RATED_VOLTAGE = "Rated Voltage must be between 10 and 400"

@pytest.mark.parametrize("value,ok,err_msg", [
    pytest.param(10,  True,  None,              id="TC_Basic_001"),
    pytest.param(200, True,  None,              id="TC_Basic_002"),
    pytest.param(400, True,  None,              id="TC_Basic_003"),
    pytest.param(9,   False, _ERR_RATED_VOLTAGE, id="TC_Basic_004"),
    pytest.param(401, False, _ERR_RATED_VOLTAGE, id="TC_Basic_005"),
])
def test_rated_voltage(app_page, value, ok, err_msg):
    nav_to_general(app_page, "Basic")
    set_input(app_page, "Rated Voltage", value)
    saved = save_and_check(app_page)
    _log.info("[TC] Rated Voltage=%s  ok_expected=%s  save_result=%s", value, ok, saved)
    if ok:
        assert saved, f"Rated Voltage={value}: 期望成功但出现错误"
        verify_modbus(REG_RATED_VOLTAGE, value, label="Rated Voltage")
    else:
        assert not saved, f"Rated Voltage={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, err_msg)


# ════════════════════════════════════════════════════════════════════
# TC_Basic_006 ~ TC_Basic_011  Wiring of Three Phase User
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option", [
    pytest.param("1 Element 2 Wire (1LN)",                  id="TC_Basic_006"),
    pytest.param("3 Element 4 Wire Y (3LN)",                id="TC_Basic_007"),
    pytest.param("2 Element 3 Wire 1 Phase (2LN)",          id="TC_Basic_008"),
    pytest.param("3 Element 4 Wire Delta (High-leg delta)",  id="TC_Basic_009"),
    pytest.param("2 Element 3 Wire Delta (2LL)",            id="TC_Basic_010"),
    pytest.param("2 Element 3 Wire Network (2LN/Network)",  id="TC_Basic_011"),
])
def test_wiring(app_page, option):
    nav_to_general(app_page, "Basic")
    set_dropdown(app_page, "Wiring of Three Phase User", option)
    saved = save_and_check(app_page)
    assert saved, f"Wiring={option}: 保存失败"
    verify_modbus(REG_WIRING, WIRING_ENCODE[option], label=f"Wiring({option})")


# ════════════════════════════════════════════════════════════════════
# TC_Basic_012 ~ TC_Basic_013  CT2 Type
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("RCT",   120, id="TC_Basic_012"),
    pytest.param("333mV", 333, id="TC_Basic_013"),
])
def test_ct2_type(app_page, option, encoded):
    nav_to_general(app_page, "Basic")
    set_dropdown(app_page, "CT2", option)
    saved = save_and_check(app_page)
    assert saved, f"CT2={option}: 保存失败"
    verify_modbus(REG_CT_TYPE, encoded, label=f"CT2({option})")


# ════════════════════════════════════════════════════════════════════
# TC_Basic_014 ~ TC_Basic_018  CT Model Input Channel 1~18
#
# 异常断言：
#   "Input Channel {N} must be between 5 and 50000"
#   只校验第一个 Channel 的错误文字，其余 channel 同理
# ════════════════════════════════════════════════════════════════════

_ERR_CT_CH = "must be between 5 and 50000"   # 部分匹配，适用所有通道

@pytest.mark.parametrize("value,ok", [
    pytest.param(5,     True,  id="TC_Basic_014"),
    pytest.param(25000, True,  id="TC_Basic_015"),
    pytest.param(50000, True,  id="TC_Basic_016"),
    pytest.param(4,     False, id="TC_Basic_017"),
    pytest.param(50001, False, id="TC_Basic_018"),
])
def test_ct_channels(app_page, value, ok):
    nav_to_general(app_page, "Basic")
    for lbl in CT_LABELS:
        set_input(app_page, lbl, value)
    saved = save_and_check(app_page)
    _log.info("[TC] CT channels=%s  ok_expected=%s  save_result=%s", value, ok, saved)
    if ok:
        assert saved, f"CT channels={value}: 保存失败"
        verify_modbus(REG_CT_FS_BASE, value, label="CT Ch1~18", count=18)
    else:
        assert not saved, f"CT channels={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, _ERR_CT_CH)


# ════════════════════════════════════════════════════════════════════
# TC_Basic_019 ~ TC_Basic_020  Demand Method + Sub-Interval 联动
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded,sub_disabled", [
    pytest.param("Rolling Window Demand", 1, False, id="TC_Basic_019"),
    pytest.param("Fixed Window Demand",   2, True,  id="TC_Basic_020"),
])
def test_demand_method(app_page, option, encoded, sub_disabled):
    nav_to_general(app_page, "Basic")
    set_dropdown(app_page, "Sliding Window", option)
    app_page.wait_for_timeout(600)
    actual_disabled = sub_interval_is_disabled(app_page)
    assert actual_disabled == sub_disabled, (
        f"Demand={option}: Sub-Interval 置灰期望={sub_disabled}, 实际={actual_disabled}"
    )
    saved = save_and_check(app_page)
    assert saved, f"Demand Method={option}: 保存失败"
    verify_modbus(REG_DEMAND_METHOD, encoded, label=f"DemandMethod({option})")


# ════════════════════════════════════════════════════════════════════
# TC_Basic_021 ~ TC_Basic_025  Sub-Interval  (范围 1-30)
#
# 设备约束：Avg ≥ 2×Sub → Avg=30 时 Sub 最大可设 15
# TC_Basic_023 原设计 Sub=30 超出设备约束，改为 Sub=15（真实可达最大值）
#
# 异常断言：
#   "Sub-Interval must be between 1 and 30"
# ════════════════════════════════════════════════════════════════════

_ERR_SUB_INTERVAL = "Sub-Interval must be between 1 and 30"

@pytest.mark.parametrize("value,ok,err_msg,avg_pre", [
    pytest.param(1,  True,  None,              30, id="TC_Basic_021"),
    pytest.param(15, True,  None,              30, id="TC_Basic_022"),
    pytest.param(15, True,  None,              30, id="TC_Basic_023"),  # Sub=30 不可达，改为15
    pytest.param(0,  False, _ERR_SUB_INTERVAL, 30, id="TC_Basic_024"),
    pytest.param(31, False, _ERR_SUB_INTERVAL, 30, id="TC_Basic_025"),
])
def test_sub_interval(app_page, value, ok, err_msg, avg_pre):
    step(f"前置：设 Averaging Interval Window = {avg_pre}")
    nav_to_general(app_page, "Basic")
    set_dropdown(app_page, "Sliding Window", "Rolling Window Demand")
    app_page.wait_for_timeout(500)
    set_input(app_page, "Averaging Interval Window", avg_pre)
    save_and_check(app_page)

    step(f"测试：设 Sub-Interval = {value}")
    nav_to_general(app_page, "Basic")
    set_dropdown(app_page, "Sliding Window", "Rolling Window Demand")
    app_page.wait_for_timeout(500)
    set_input(app_page, "Sub-Interval", value)
    saved = save_and_check(app_page)
    _log.info("[TC] Sub-Interval=%s  ok_expected=%s  save_result=%s", value, ok, saved)
    if ok:
        assert saved, f"Sub-Interval={value}: 保存失败"
        verify_modbus(REG_DEMAND_SUB, value, label="Sub-Interval")
    else:
        assert not saved, f"Sub-Interval={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, err_msg)


# ════════════════════════════════════════════════════════════════════
# TC_Basic_026 ~ TC_Basic_030  Averaging Interval Window  (范围 1-30)
#
# 设备约束：Avg ≥ 2×Sub → Sub=1 时 Avg 最小可设 2
# TC_Basic_026 原设计 Avg=1 超出设备约束，改为 Avg=2（真实可达最小值）
#
# 异常断言：
#   "Averaging Interval Window must be between 1 and 30"
# ════════════════════════════════════════════════════════════════════

_ERR_AVG_INTERVAL = "Averaging Interval Window must be between 1 and 30"

@pytest.mark.parametrize("value,ok,err_msg", [
    pytest.param(2,  True,  None,              id="TC_Basic_026"),  # Avg=1 不可达，改为2
    pytest.param(15, True,  None,              id="TC_Basic_027"),
    pytest.param(30, True,  None,              id="TC_Basic_028"),
    pytest.param(0,  False, _ERR_AVG_INTERVAL, id="TC_Basic_029"),
    pytest.param(31, False, _ERR_AVG_INTERVAL, id="TC_Basic_030"),
])
def test_averaging_interval(app_page, value, ok, err_msg):
    step("前置：设 Sub-Interval = 1")
    nav_to_general(app_page, "Basic")
    set_dropdown(app_page, "Sliding Window", "Rolling Window Demand")
    app_page.wait_for_timeout(500)
    set_input(app_page, "Sub-Interval", 1)
    save_and_check(app_page)

    step(f"测试：设 Averaging Interval Window = {value}")
    nav_to_general(app_page, "Basic")
    set_input(app_page, "Averaging Interval Window", value)
    saved = save_and_check(app_page)
    _log.info("[TC] Averaging Interval=%s  ok_expected=%s  save_result=%s", value, ok, saved)
    if ok:
        assert saved, f"Averaging Interval={value}: 保存失败"
        verify_modbus(REG_DEMAND_INTERVAL, value, label="Averaging Interval Window")
    else:
        assert not saved, f"Averaging Interval={value}: 期望校验失败但保存成功"
        assert_field_error(app_page, err_msg)


# ════════════════════════════════════════════════════════════════════
# TC_Advanced_001 ~ TC_Advanced_002  VAR Calculation Method
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("Method 1 (True)",        0, id="TC_Advanced_001"),
    pytest.param("Method 2 (Generalized)", 1, id="TC_Advanced_002"),
])
def test_var_method(app_page, option, encoded):
    nav_to_general(app_page, "Advanced")
    set_dropdown(app_page, "VAR Calculation Method", option)
    saved = save_and_check(app_page)
    assert saved, f"VAR Method={option}: 保存失败"
    verify_modbus(REG_VAR_METHOD, encoded, label=f"VAR Method({option})")


# ════════════════════════════════════════════════════════════════════
# TC_Advanced_003 ~ TC_Advanced_004  VAR/PF Convention
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("IEC",  0, id="TC_Advanced_003"),
    pytest.param("IEEE", 1, id="TC_Advanced_004"),
])
def test_var_pf(app_page, option, encoded):
    nav_to_general(app_page, "Advanced")
    set_dropdown(app_page, "VAR/PF Convention", option)
    saved = save_and_check(app_page)
    assert saved, f"VAR/PF={option}: 保存失败"
    verify_modbus(REG_VAR_PF, encoded, label=f"VAR/PF({option})")
