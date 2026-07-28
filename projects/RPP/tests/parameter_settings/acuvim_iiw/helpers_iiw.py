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

    # 最多点击 4 次 Settings 展开子菜单；菜单在 click 前收起则快速重试
    clicked = False
    for _ in range(4):
        vis = _first_visible_general()
        if vis is None:
            settings_loc.click()
            page.wait_for_timeout(700)
            vis = _first_visible_general()
        if vis is not None:
            try:
                vis.click(timeout=3000)
                clicked = True
                break
            except Exception:
                pass  # 菜单在 click 前收起，继续重试

    if not clicked:
        _log.warning("nav: Settings 4次未展开 → reload 重置页面")
        page.reload()
        page.wait_for_timeout(1000)
        page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)
        page.locator(f"text={DEVICE_NAME}").first.click()
        page.wait_for_selector("text=Settings", timeout=10000)
        page.wait_for_timeout(500)
        settings_loc.click()
        page.wait_for_timeout(700)
        vis = _first_visible_general()
        if vis is not None:
            vis.click(timeout=5000)
        else:
            raise RuntimeError("nav_to_general: reload 后仍无法找到 General 菜单项")

    try:
        page.wait_for_selector("button:has-text('Save')", timeout=12000)
    except Exception:
        _log.warning("nav: General 页面 Save 按钮未出现 → reload 完整重试")
        page.reload()
        page.wait_for_timeout(1500)
        page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)
        page.locator(f"text={DEVICE_NAME}").first.click()
        page.wait_for_selector("text=Settings", timeout=10000)
        page.wait_for_timeout(500)
        settings_loc.click()
        page.wait_for_timeout(700)
        vis = _first_visible_general()
        if vis is not None:
            vis.click(timeout=5000)
        page.wait_for_selector("button:has-text('Save')", timeout=15000)

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



# ── 补充：被修复脚本误删的寄存器地址 & 校验错误常量 ─────────────────
REG_IA_DIR          = 4114   # 0x1012  I A Direction (0=Positive, 1=Negative)
REG_IB_DIR          = 4115   # 0x1013  I B Direction
REG_IC_DIR          = 4116   # 0x1014  I C Direction
_ERR_PT1          = "PT1 must be between 50 and 1000000"
_ERR_PT2          = "PT2 must be between 50 and 400"
_ERR_CT1          = "CT1 must be between 1 and 50000"
_ERR_CT41         = "CT41 must be between 1 and 50000"
_ERR_SUB_INTERVAL = "Sub-Interval must be between 1 and 30"
_ERR_AVG_INTERVAL = "Averaging Interval Window must be between 1 and 30"

# 让 from helpers import * 包含 _log 等下划线名称（必须放文件最末）

# ════════════════════════════════════════════════════════════════════
# Class S Event (PQ Event) – 寄存器地址 & 编码映射
# Settings → Class S Event
# 来源：Acuvim IIW Modbus Address v1.27, sheet "4-30 PQ Event"
# ════════════════════════════════════════════════════════════════════

# ── 全局配置寄存器 ──────────────────────────────────────────────────
REG_CS_VOLT_RATED   = 44902  # 0xAF66  Voltage Rated Value (10-690 V)
REG_CS_CURR_RATED   = 44903  # 0xAF67  Current Rated Value (0.1-5 A, ×1000 in reg)
REG_CS_SAMPLE_RATE  = 44904  # 0xAF68  Waveform Sample Rate (0=16 … 5=512 pts)

# ── 各事件类型基地址（UI 行顺序） ──────────────────────────────────
# 每个事件块：[base+0]=Enable [base+1]=Threshold [base+2]=Hysteresis
#             [base+3]=WF Trigger  [base+4]=DO Output  [base+5]=RO Output
_CS_BASE = {
    "VoltDip"  : 44905,  # 0xAF69  Voltage Dip
    "VoltSwl"  : 44911,  # 0xAF6F  Voltage Swell
    "VoltIntr" : 44917,  # 0xAF75  Voltage Interruption
    "UnbalVolt": 44923,  # 0xAF7B  Unbalance Voltage
    "CurrSwl"  : 44935,  # 0xAF87  Current Swell（跳过 Current Dip 44929-44934，UI不显示）
    "UnbalCurr": 44941,  # 0xAF8D  Unbalance Current
}
# UI 行索引（0-based）→ 事件名称
CS_ROW_NAMES = ["VoltDip", "VoltSwl", "VoltIntr", "UnbalVolt", "CurrSwl", "UnbalCurr"]


def cs_reg(event_name: str, offset: int) -> int:
    """返回指定事件、偏移量对应的寄存器地址。"""
    return _CS_BASE[event_name] + offset


CS_OFFSET_ENABLE = 0
CS_OFFSET_THR    = 1
CS_OFFSET_HYS    = 2
CS_OFFSET_WF     = 3
CS_OFFSET_DO     = 4
CS_OFFSET_RO     = 5

# ── 编码映射 ─────────────────────────────────────────────────────────
CS_SAMPLE_RATE_ENCODE = {
    "16 sample/cycle" : 0,
    "32 sample/cycle" : 1,
    "64 sample/cycle" : 2,
    "128 sample/cycle": 3,
    "256 sample/cycle": 4,
    "512 sample/cycle": 5,
}
CS_DO_ENCODE = {
    "No Output": 0,
    "2-1 DO1"  : 1,
    "2-1 DO2"  : 2,
    "2-2 DO1"  : 3,
    "2-2 DO2"  : 4,
}
CS_RO_ENCODE = {
    "No Output": 0,
    "1-1 RO1"  : 1,
    "1-1 RO2"  : 2,
    "3-1 RO1"  : 4,
    "3-1 RO2"  : 8,
    "1-2 RO1"  : 16,
    "1-2 RO2"  : 32,
    "3-2 RO1"  : 64,
    "3-2 RO2"  : 128,
}

# ── 页面导航 ─────────────────────────────────────────────────────────

@allure.step("导航到 Class S Event 页面")
def nav_to_class_s_event(page: Page) -> None:
    """Gateway Devices → AcuvimIIW → Settings ▾ → Class S Event"""
    _dismiss_dialog(page)
    page.locator(
        "a:has-text('Gateway Devices'), "
        "span:has-text('Gateway Devices'), "
        "li:has-text('Gateway Devices')"
    ).first.click()
    page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)

    page.locator(f"text={DEVICE_NAME}").first.click()
    page.wait_for_selector("text=Settings", timeout=10000)
    page.wait_for_timeout(800)

    settings_loc = page.locator(
        "span:has-text('Settings'), a:has-text('Settings'), button:has-text('Settings')"
    ).first

    def _find_cs():
        gi = page.locator(".el-menu-item").filter(has_text="Class S Event")
        for j in range(gi.count()):
            if gi.nth(j).is_visible():
                return gi.nth(j)
        return None

    # 最多点击 4 次 Settings 展开子菜单；菜单在 click 前收起则快速重试
    clicked = False
    for _ in range(4):
        vis = _find_cs()
        if vis is None:
            settings_loc.click()
            page.wait_for_timeout(700)
            vis = _find_cs()
        if vis is not None:
            try:
                vis.click(timeout=3000)
                clicked = True
                break
            except Exception:
                pass  # 菜单在 click 前收起，继续重试

    if not clicked:
        _log.warning("nav: Settings 4次未展开 → reload 重置页面")
        page.reload()
        page.wait_for_timeout(1000)
        page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)
        page.locator(f"text={DEVICE_NAME}").first.click()
        page.wait_for_selector("text=Settings", timeout=10000)
        page.wait_for_timeout(500)
        settings_loc.click()
        page.wait_for_timeout(700)
        vis = _find_cs()
        if vis is not None:
            vis.click(timeout=5000)
        else:
            raise RuntimeError("nav_to_class_s_event: reload 后仍无法找到 Class S Event 菜单项")

    try:
        page.wait_for_selector("button:has-text('Save')", timeout=12000)
    except Exception:
        _log.warning("nav: Class S Event 页面 Save 按钮未出现 → reload 完整重试")
        page.reload()
        page.wait_for_timeout(1500)
        page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)
        page.locator(f"text={DEVICE_NAME}").first.click()
        page.wait_for_selector("text=Settings", timeout=10000)
        page.wait_for_timeout(500)
        settings_loc.click()
        page.wait_for_timeout(700)
        vis = _find_cs()
        if vis is not None:
            vis.click(timeout=5000)
        page.wait_for_selector("button:has-text('Save')", timeout=15000)

    page.wait_for_timeout(800)


# ── 事件表格行操作 ────────────────────────────────────────────────────

@allure.step("设置事件行 [{row_idx}] Enable = {enable}")
def set_event_enable(page: Page, row_idx: int, enable: bool) -> None:
    """Enable 开关（el-switch）；Enable 位于 el-switch 奇数索引：0,2,4,6,8,10。"""
    sw_idx = row_idx * 2
    sw = page.locator(".el-switch").nth(sw_idx)
    sw.scroll_into_view_if_needed()
    is_on = "is-checked" in (sw.get_attribute("class") or "")
    _log.info("[CS ENABLE] row=%d is_on=%s target=%s", row_idx, is_on, enable)
    if is_on != enable:
        sw.click()
        page.wait_for_timeout(300)


@allure.step("设置事件行 [{row_idx}] WF Trigger = {enable}")
def set_event_wf(page: Page, row_idx: int, enable: bool) -> None:
    """Waveform Trigger 开关；位于 el-switch 偶数索引：1,3,5,7,9,11。"""
    sw_idx = row_idx * 2 + 1
    sw = page.locator(".el-switch").nth(sw_idx)
    sw.scroll_into_view_if_needed()
    is_on = "is-checked" in (sw.get_attribute("class") or "")
    _log.info("[CS WF    ] row=%d is_on=%s target=%s", row_idx, is_on, enable)
    if is_on != enable:
        sw.click()
        page.wait_for_timeout(300)


@allure.step("设置事件行 [{row_idx}] Threshold = {value}")
def set_event_thr(page: Page, row_idx: int, value) -> None:
    inp = page.locator("input[placeholder='Enter Threshold']").nth(row_idx)
    inp.scroll_into_view_if_needed()
    inp.click()
    inp.press("Control+a")
    inp.press("Delete")
    inp.type(str(value), delay=50)
    _log.info("[CS THR   ] row=%d value=%s", row_idx, value)


@allure.step("设置事件行 [{row_idx}] Hysteresis = {value}")
def set_event_hys(page: Page, row_idx: int, value) -> None:
    inp = page.locator("input[placeholder='Enter Hysteresis']").nth(row_idx)
    inp.scroll_into_view_if_needed()
    inp.click()
    inp.press("Control+a")
    inp.press("Delete")
    inp.type(str(value), delay=50)
    _log.info("[CS HYS   ] row=%d value=%s", row_idx, value)


@allure.step("设置事件行 [{row_idx}] DO = {option}")
def set_event_do(page: Page, row_idx: int, option: str) -> None:
    """DO SELECT；位于 el-select 奇数索引：1,3,5,7,9,11（跳过第0个 Sample Rate）。"""
    sel_idx = 1 + row_idx * 2
    sel = page.locator(".el-select").nth(sel_idx)
    sel.scroll_into_view_if_needed()
    sel.click()
    page.wait_for_selector(".el-select-dropdown__item:visible", timeout=3000)
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(option, exact=True)
    if target.count() == 0:
        page.keyboard.press("Escape")
        raise RuntimeError(f"set_event_do: option {option!r} not found")
    target.first.click()
    page.wait_for_timeout(300)
    _log.info("[CS DO    ] row=%d option=%s", row_idx, option)


@allure.step("设置事件行 [{row_idx}] RO = {option}")
def set_event_ro(page: Page, row_idx: int, option: str) -> None:
    """RO SELECT；位于 el-select 偶数索引：2,4,6,8,10,12。"""
    sel_idx = 2 + row_idx * 2
    sel = page.locator(".el-select").nth(sel_idx)
    sel.scroll_into_view_if_needed()
    sel.click()
    page.wait_for_selector(".el-select-dropdown__item:visible", timeout=3000)
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(option, exact=True)
    if target.count() == 0:
        page.keyboard.press("Escape")
        raise RuntimeError(f"set_event_ro: option {option!r} not found")
    target.first.click()
    page.wait_for_timeout(300)
    _log.info("[CS RO    ] row=%d option=%s", row_idx, option)


@allure.step("验证 CS 寄存器 [{label}] addr={address} expected={expected}")
def cs_verify(address: int, expected: int, label: str = "") -> None:
    """读 Class S Event 寄存器并断言。"""
    actual = modbus_read(address)[0]
    ok = actual == expected
    _log.info("[CS VERIFY] %-40s expected=%-6s actual=%-6s %s",
              label, expected, actual, "✓" if ok else "✗")
    assert ok, f"{label}: expected {expected}, got {actual}"


# 让 from helpers_iiw import * 包含 _log 等下划线名称（必须放文件最末）

# ════════════════════════════════════════════════════════════════════
# Class S Event (PQ Event) – 寄存器地址 & 编码映射
# Settings -> Class S Event
# 来源：Acuvim IIW Modbus Address v1.27, sheet "4-30 PQ Event"
# ════════════════════════════════════════════════════════════════════

REG_CS_VOLT_RATED   = 44902  # Voltage Rated Value (10-690 V)
REG_CS_CURR_RATED   = 44903  # Current Rated Value (0.1-5 A, x1000 in reg)
REG_CS_SAMPLE_RATE  = 44904  # Waveform Sample Rate (0=16 ... 5=512 pts)

_CS_BASE = {
    "VoltDip"  : 44905,
    "VoltSwl"  : 44911,
    "VoltIntr" : 44917,
    "UnbalVolt": 44923,
    "CurrSwl"  : 44935,
    "UnbalCurr": 44941,
}
CS_ROW_NAMES = ["VoltDip", "VoltSwl", "VoltIntr", "UnbalVolt", "CurrSwl", "UnbalCurr"]

CS_OFFSET_ENABLE = 0
CS_OFFSET_THR    = 1
CS_OFFSET_HYS    = 2
CS_OFFSET_WF     = 3
CS_OFFSET_DO     = 4
CS_OFFSET_RO     = 5


def cs_reg(event_name, offset):
    return _CS_BASE[event_name] + offset


CS_SAMPLE_RATE_ENCODE = {
    "16 sample/cycle" : 0,
    "32 sample/cycle" : 1,
    "64 sample/cycle" : 2,
    "128 sample/cycle": 3,
    "256 sample/cycle": 4,
    "512 sample/cycle": 5,
}
CS_DO_ENCODE = {
    "No Output": 0,
    "2-1 DO1"  : 1,
    "2-1 DO2"  : 2,
    "2-2 DO1"  : 3,
    "2-2 DO2"  : 4,
}
CS_RO_ENCODE = {
    "No Output": 0,
    "1-1 RO1"  : 1,
    "1-1 RO2"  : 2,
    "3-1 RO1"  : 4,
    "3-1 RO2"  : 8,
    "1-2 RO1"  : 16,
    "1-2 RO2"  : 32,
    "3-2 RO1"  : 64,
    "3-2 RO2"  : 128,
}


@allure.step("navigate to Class S Event page")
def nav_to_class_s_event(page: Page) -> None:
    """Gateway Devices -> AcuvimIIW -> Settings -> Class S Event"""
    _dismiss_dialog(page)
    page.locator(
        "a:has-text('Gateway Devices'), "
        "span:has-text('Gateway Devices'), "
        "li:has-text('Gateway Devices')"
    ).first.click()
    page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)
    page.locator(f"text={DEVICE_NAME}").first.click()
    page.wait_for_selector("text=Settings", timeout=10000)
    page.wait_for_timeout(800)

    settings_loc = page.locator(
        "span:has-text('Settings'), a:has-text('Settings'), button:has-text('Settings')"
    ).first

    def _find_cs():
        gi = page.locator(".el-menu-item").filter(has_text="Class S Event")
        for j in range(gi.count()):
            if gi.nth(j).is_visible():
                return gi.nth(j)
        return None

    # 最多点击 4 次 Settings 展开子菜单；菜单在 click 前收起则快速重试
    clicked = False
    for _ in range(4):
        vis = _find_cs()
        if vis is None:
            settings_loc.click()
            page.wait_for_timeout(700)
            vis = _find_cs()
        if vis is not None:
            try:
                vis.click(timeout=3000)
                clicked = True
                break
            except Exception:
                pass  # 菜单在 click 前收起，继续重试

    if not clicked:
        _log.warning("nav: Settings 4次未展开 → reload 重置页面")
        page.reload()
        page.wait_for_timeout(1000)
        page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)
        page.locator(f"text={DEVICE_NAME}").first.click()
        page.wait_for_selector("text=Settings", timeout=10000)
        page.wait_for_timeout(500)
        settings_loc.click()
        page.wait_for_timeout(700)
        vis = _find_cs()
        if vis is not None:
            vis.click(timeout=5000)
        else:
            raise RuntimeError("nav_to_class_s_event: reload 后仍无法找到 Class S Event 菜单项")

    try:
        page.wait_for_selector("button:has-text('Save')", timeout=12000)
    except Exception:
        _log.warning("nav: Class S Event 页面 Save 按钮未出现 → reload 完整重试")
        page.reload()
        page.wait_for_timeout(1500)
        page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)
        page.locator(f"text={DEVICE_NAME}").first.click()
        page.wait_for_selector("text=Settings", timeout=10000)
        page.wait_for_timeout(500)
        settings_loc.click()
        page.wait_for_timeout(700)
        vis = _find_cs()
        if vis is not None:
            vis.click(timeout=5000)
        page.wait_for_selector("button:has-text('Save')", timeout=15000)

    page.wait_for_timeout(800)


@allure.step("set event row [{row_idx}] Enable={enable}")
def set_event_enable(page: Page, row_idx: int, enable: bool) -> None:
    """Enable toggle (el-switch); Enable switches at indices 0,2,4,6,8,10."""
    sw = page.locator(".el-switch").nth(row_idx * 2)
    sw.scroll_into_view_if_needed()
    is_on = "is-checked" in (sw.get_attribute("class") or "")
    _log.info("[CS ENABLE] row=%d is_on=%s target=%s", row_idx, is_on, enable)
    if is_on != enable:
        sw.click()
        page.wait_for_timeout(300)


@allure.step("set event row [{row_idx}] WF Trigger={enable}")
def set_event_wf(page: Page, row_idx: int, enable: bool) -> None:
    """Waveform Trigger toggle; WF switches at indices 1,3,5,7,9,11."""
    sw = page.locator(".el-switch").nth(row_idx * 2 + 1)
    sw.scroll_into_view_if_needed()
    is_on = "is-checked" in (sw.get_attribute("class") or "")
    _log.info("[CS WF    ] row=%d is_on=%s target=%s", row_idx, is_on, enable)
    if is_on != enable:
        sw.click()
        page.wait_for_timeout(300)


@allure.step("set event row [{row_idx}] Threshold={value}")
def set_event_thr(page: Page, row_idx: int, value) -> None:
    inp = page.locator("input[placeholder='Enter Threshold']").nth(row_idx)
    inp.scroll_into_view_if_needed()
    inp.click()
    inp.press("Control+a")
    inp.press("Delete")
    inp.type(str(value), delay=50)
    _log.info("[CS THR   ] row=%d value=%s", row_idx, value)


@allure.step("set event row [{row_idx}] Hysteresis={value}")
def set_event_hys(page: Page, row_idx: int, value) -> None:
    inp = page.locator("input[placeholder='Enter Hysteresis']").nth(row_idx)
    inp.scroll_into_view_if_needed()
    inp.click()
    inp.press("Control+a")
    inp.press("Delete")
    inp.type(str(value), delay=50)
    _log.info("[CS HYS   ] row=%d value=%s", row_idx, value)


@allure.step("set event row [{row_idx}] DO={option}")
def set_event_do(page: Page, row_idx: int, option: str) -> None:
    """DO SELECT; at el-select indices 1,3,5,7,9,11 (skip index 0 = Sample Rate)."""
    sel = page.locator(".el-select").nth(1 + row_idx * 2)
    sel.scroll_into_view_if_needed()
    sel.click()
    page.wait_for_selector(".el-select-dropdown__item:visible", timeout=3000)
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(option, exact=True)
    if target.count() == 0:
        page.keyboard.press("Escape")
        raise RuntimeError(f"set_event_do: option {option!r} not found")
    target.first.click()
    page.wait_for_timeout(300)
    _log.info("[CS DO    ] row=%d option=%s", row_idx, option)


@allure.step("set event row [{row_idx}] RO={option}")
def set_event_ro(page: Page, row_idx: int, option: str) -> None:
    """RO SELECT; at el-select indices 2,4,6,8,10,12."""
    sel = page.locator(".el-select").nth(2 + row_idx * 2)
    sel.scroll_into_view_if_needed()
    sel.click()
    page.wait_for_selector(".el-select-dropdown__item:visible", timeout=3000)
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(option, exact=True)
    if target.count() == 0:
        page.keyboard.press("Escape")
        raise RuntimeError(f"set_event_ro: option {option!r} not found")
    target.first.click()
    page.wait_for_timeout(300)
    _log.info("[CS RO    ] row=%d option=%s", row_idx, option)


@allure.step("verify CS reg [{label}] addr={address} expected={expected}")
def cs_verify(address: int, expected: int, label: str = "") -> None:
    actual = modbus_read(address)[0]
    ok = actual == expected
    _log.info("[CS VERIFY] %-40s expected=%-6s actual=%-6s %s",
              label, expected, actual, "\u2713" if ok else "\u2717")
    assert ok, f"{label}: expected {expected}, got {actual}"

# 让 from helpers_iiw import * 包含 _log 等下划线名称（必须放文件最末）
__all__ = [_n for _n in globals() if not _n.startswith('__')]
