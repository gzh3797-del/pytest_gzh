"""
test_event_waveform.py – AcuRev4100  Settings → Event & Waveform 自动化测试

页面路径：Physical Devices → AcuRev4100 → Settings → Event & Waveform

页面实际结构（来自 inspect_event_waveform.py）：
  ┌─ Power Quality Events Settings ──────────────────────────────┐
  │  Voltage Sag       : Threshold(input) / Hysteresis(input)    │
  │                      DO output(select) / RO output(select)   │
  │  Voltage Swell     : 同上，Threshold 范围 110-150            │
  │  Voltage Interruption: 同上，Threshold 范围 5-20             │
  │  Current Swell     : 同上，Threshold 范围 110-150            │
  └──────────────────────────────────────────────────────────────┘
  ┌─ Waveform ───────────────────────────────────────────────────┐
  │  DI1/DI2/DI3/DI4 Trigger (select: Disable/User1-12/Only Voltage)│
  │  Manually Trigger  (select: 同上)                             │
  │  Sample Rate       (select: 16/32/64/128 points)             │
  │  Num of Cycles Before (input, 1-159)                         │
  │  Num of Cycles After  (input, 1-159)                         │
  └──────────────────────────────────────────────────────────────┘

PQ Event 区字段无 el-form-item__label，通过 placeholder 定位：
  - 'Enter Threshold'   → nth(0)=Sag, nth(1)=Swell, nth(2)=Intr, nth(3)=CurrSwell
  - 'Enter Hysteresis'  → 同上顺序

DO/RO 选项（实测）：
  - DO: 'None', 'DO1'~'DO8'
  - RO: 'None', 'RO1', 'RO2'

寄存器来源：AcuRev4100 Modbus Address Table v1.02，Sheet: Waveform|PQ Event

运行：
  pytest tests/acurev4100/test_event_waveform.py -v
"""

import allure
import pytest
import logging
from playwright.sync_api import Page
from pymodbus.client import ModbusTcpClient

logging.getLogger("pymodbus").setLevel(logging.WARNING)

# ── 设备配置：由 conftest.py 的 _bind_acurev4100_device（autouse，会话级）
# 在测试开始前通过网关 API 动态发现当前在线的 AcuRev4100 设备后写入真实值，
# 替代原来从 config.yaml device_modbus 段静态解析——避免物理设备切换/离线后
# 仍连着一个旧地址。以下仅为占位默认值，实际连接前会被覆盖。
DEVICE_NAME = ""
MODBUS_HOST = ""
MODBUS_PORT = 502
SLAVE_ID    = 1

# ── 寄存器地址（Waveform|PQ Event sheet）─────────────────────────────

# Voltage Sag  (0x6502–0x6507)
REG_VOLT_SAG_THR = 25859   # 0x6503  Threshold    10–90 %
REG_VOLT_SAG_HYS = 25860   # 0x6504  Hysteresis   1–10 %
REG_VOLT_SAG_DO  = 25862   # 0x6506  DO Output    0–8
REG_VOLT_SAG_RO  = 25863   # 0x6507  RO Output    0–2

# Voltage Swell  (0x6508–0x650D)
REG_VOLT_SWL_THR = 25865   # 0x6509  Threshold   110–150 %
REG_VOLT_SWL_HYS = 25866   # 0x650A  Hysteresis   1–10 %
REG_VOLT_SWL_DO  = 25868   # 0x650C  DO Output
REG_VOLT_SWL_RO  = 25869   # 0x650D  RO Output

# Voltage Interruption  (0x650E–0x6513)
REG_VOLT_INT_THR = 25871   # 0x650F  Threshold    5–20 %
REG_VOLT_INT_HYS = 25872   # 0x6510  Hysteresis   1–10 %
REG_VOLT_INT_DO  = 25874   # 0x6512  DO Output
REG_VOLT_INT_RO  = 25875   # 0x6513  RO Output

# Current Swell  (0x6514–0x6519)
REG_CURR_SWL_THR = 25877   # 0x6515  Threshold   110–150 %
REG_CURR_SWL_HYS = 25878   # 0x6516  Hysteresis   1–10 %
REG_CURR_SWL_DO  = 25880   # 0x6518  DO Output
REG_CURR_SWL_RO  = 25881   # 0x6519  RO Output

# Waveform Record Parameter Setting  (0x6570–0x6577)
REG_WF_SAMPLE_RATE = 25968  # 0x6570  0=16pts 1=32pts 2=64pts 3=128pts
REG_WF_PRE_CYCLES  = 25969  # 0x6571  Pre-event cycles  1–159
REG_WF_POST_CYCLES = 25970  # 0x6572  Post-event cycles 1–(160-pre)
REG_WF_DI1_TRIG    = 25971  # 0x6573  0=Disable 1-12=User1-12 255=OnlyVoltage
REG_WF_DI2_TRIG    = 25972  # 0x6574
REG_WF_DI3_TRIG    = 25973  # 0x6575
REG_WF_DI4_TRIG    = 25974  # 0x6576
REG_WF_MAN_TRIG    = 25975  # 0x6577

# ── PQ 事件类型索引（对应 Threshold/Hysteresis 的 nth 顺序）─────────
SAG_IDX = 0   # Voltage Sag
SWL_IDX = 1   # Voltage Swell
INT_IDX = 2   # Voltage Interruption
CSW_IDX = 3   # Current Swell

# DO selects 整体位置：页面 .el-select 中第 6+idx*2 个
# RO selects：第 7+idx*2 个（基于 inspect 输出的顺序）
_DO_SELECT_BASE = 6
_RO_SELECT_BASE = 7

# ── DI Trigger 选项编码 ────────────────────────────────────────────
DI_TRIG_ENCODE = {
    "Disable":             0,
    "User 1":              1,
    "User 12":            12,
    "Only Record Voltage": 255,
}

# ── Sample Rate 编码 ──────────────────────────────────────────────
SAMPLE_RATE_ENCODE = {
    "16 points":  0,
    "32 points":  1,
    "64 points":  2,
    "128 points": 3,
}

# ── DO/RO 编码 ────────────────────────────────────────────────────
DO_ENCODE = {"None": 0, "DO1": 1, "DO2": 2, "DO8": 8}
RO_ENCODE = {"None": 0, "RO1": 1, "RO2": 2}

_log = logging.getLogger("acurev4100_ew")


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
    _log.info("[MODBUS TX] FC=03  addr=%d(0x%04X)  count=%d", address, address, count)
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


@allure.step("验证寄存器 [{label}]  addr={address}  expected={expected}")
def verify_modbus(address: int, expected, label: str = "") -> None:
    regs = modbus_read(address, count=1)
    actual = regs[0]
    ok = actual == expected
    _log.info("[VERIFY] %-50s  expected=%-6s  actual=%-6s  → %s",
              label, expected, actual, "✓ PASS" if ok else "✗ FAIL")
    assert ok, f"{label}: expected {expected}, got {actual}"


def get_visible_errors(page: Page) -> list[str]:
    errors = []
    locs = page.locator(".el-form-item__error")
    for i in range(locs.count()):
        loc = locs.nth(i)
        if loc.is_visible():
            txt = loc.inner_text().strip()
            if txt:
                errors.append(txt)
    return errors


@allure.step("导航到 Event & Waveform 页面")
def nav_to_event_waveform(page: Page) -> None:
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
    _settings_loc = page.locator(
        "span:has-text('Settings'), a:has-text('Settings'), button:has-text('Settings')"
    )
    _ew_loc = page.locator(
        "a:has-text('Event & Waveform'), li:has-text('Event & Waveform'), "
        "span:has-text('Event & Waveform'), "
        "a:has-text('Event Waveform'), li:has-text('Event Waveform')"
    )
    step("Click Event & Waveform submenu")
    clicked = False
    for _ in range(4):
        _settings_loc.first.click()
        page.wait_for_timeout(700)
        if _ew_loc.first.is_visible():
            try:
                _ew_loc.first.click(timeout=3000)
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
        _settings_loc.first.click()
        page.wait_for_timeout(700)
        if _ew_loc.first.is_visible():
            _ew_loc.first.click()
        else:
            _ew_loc.first.click(force=True, timeout=5000)

    try:
        page.wait_for_selector("button:has-text('Save')", timeout=12000)
    except Exception:
        _log.warning("nav: E&W 页面 Save 按钮未出现 → reload 完整重试")
        page.reload()
        page.wait_for_timeout(1500)
        page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)
        page.locator(f"text={DEVICE_NAME}").first.click()
        page.wait_for_selector("text=Settings", timeout=10000)
        page.wait_for_timeout(500)
        _settings_loc.first.click()
        page.wait_for_timeout(700)
        if _ew_loc.first.is_visible():
            _ew_loc.first.click()
        else:
            _ew_loc.first.click(force=True, timeout=5000)
        page.wait_for_selector("button:has-text('Save')", timeout=15000)

    page.wait_for_timeout(1000)
    step("Ready: Event & Waveform")


@allure.step("设置 PQ Threshold  event_idx={event_idx}  value={value}")
def set_pq_threshold(page: Page, event_idx: int, value) -> None:
    """通过 placeholder 定位 Threshold 输入框（顺序：Sag=0, Swell=1, Intr=2, CurrSwell=3）。"""
    inp = page.locator("input[placeholder='Enter Threshold']").nth(event_idx)
    inp.scroll_into_view_if_needed()
    inp.click()
    inp.press("Control+a")
    inp.press("Delete")
    inp.type(str(value), delay=50)
    _log.info("  set Threshold[%d] = %r", event_idx, value)


@allure.step("设置 PQ Hysteresis  event_idx={event_idx}  value={value}")
def set_pq_hysteresis(page: Page, event_idx: int, value) -> None:
    """通过 placeholder 定位 Hysteresis 输入框。"""
    inp = page.locator("input[placeholder='Enter Hysteresis']").nth(event_idx)
    inp.scroll_into_view_if_needed()
    inp.click()
    inp.press("Control+a")
    inp.press("Delete")
    inp.type(str(value), delay=50)
    _log.info("  set Hysteresis[%d] = %r", event_idx, value)


@allure.step("设置 PQ DO Output  event_idx={event_idx}  option={option}")
def set_pq_do(page: Page, event_idx: int, option: str) -> None:
    """DO 输出下拉框：按页面内 .el-select 顺序索引（6+idx*2）。"""
    sel_idx = _DO_SELECT_BASE + event_idx * 2
    sel = page.locator(".el-select").nth(sel_idx)
    sel.scroll_into_view_if_needed()
    sel.click()
    page.wait_for_timeout(400)
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(option, exact=True)
    if target.count() == 0:
        page.keyboard.press("Escape")
        raise RuntimeError(f"set_pq_do: option {option!r} not found")
    target.first.click()
    page.wait_for_timeout(300)


@allure.step("设置 PQ RO Output  event_idx={event_idx}  option={option}")
def set_pq_ro(page: Page, event_idx: int, option: str) -> None:
    """RO 输出下拉框：按页面内 .el-select 顺序索引（7+idx*2）。"""
    sel_idx = _RO_SELECT_BASE + event_idx * 2
    sel = page.locator(".el-select").nth(sel_idx)
    sel.scroll_into_view_if_needed()
    sel.click()
    page.wait_for_timeout(400)
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(option, exact=True)
    if target.count() == 0:
        page.keyboard.press("Escape")
        raise RuntimeError(f"set_pq_ro: option {option!r} not found")
    target.first.click()
    page.wait_for_timeout(300)


@allure.step("下拉选择 [{label}] → {option_text}")
def set_dropdown(page: Page, label: str, option_text: str) -> None:
    step(f"Dropdown [{label}] → {option_text!r}")
    trigger = page.get_by_text(label, exact=False).first.locator(
        "xpath=following::div[contains(@class,'el-select')][1]"
    )
    trigger.click()
    page.wait_for_timeout(500)
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(
        option_text, exact=True
    )
    if target.count() == 0:
        visible_items = page.locator(".el-select-dropdown__item:visible")
        opt_texts = [
            visible_items.nth(i).inner_text().strip()
            for i in range(min(visible_items.count(), 30))
        ]
        page.keyboard.press("Escape")
        raise RuntimeError(
            f"set_dropdown: option {option_text!r} not found for [{label}]. "
            f"Visible options: {opt_texts}"
        )
    target.first.click()
    page.wait_for_timeout(300)


@allure.step("输入字段 [{label}] = {value}")
def set_input(page: Page, label: str, value) -> None:
    step(f"Set input [{label}] = {value!r}")
    label_els = page.get_by_text(label, exact=False)
    if label_els.count() == 0:
        raise RuntimeError(f"set_input: label {label!r} not found on page")
    inp = label_els.first.locator("xpath=following::input[1]")
    if inp.count() == 0:
        raise RuntimeError(f"set_input: no input found after label {label!r}")
    inp.first.scroll_into_view_if_needed()
    inp.first.click()
    inp.first.press("Control+a")
    inp.first.press("Delete")
    inp.first.type(str(value), delay=50)


@allure.step("点击保存并检查结果")
def save_and_check(page: Page) -> bool:
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
            _log.info("[SAVE] SUCCESS toast")
            return True
    for sel in [".el-message--error", ".el-message--warning"]:
        loc = page.locator(sel)
        if loc.count() > 0 and loc.first.is_visible():
            _log.info("[SAVE] ERROR toast: %r", loc.first.inner_text()[:80])
            return False
    _log.info("[SAVE] No toast → assuming success")
    return True


# ════════════════════════════════════════════════════════════════════
# TC_VoltSag_Thr_001~005  Voltage Sag Threshold（10–90 %）
#   寄存器 0x6503 (25859)
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("value,ok", [
    pytest.param(10, True,  id="TC_VoltSag_Thr_001"),
    pytest.param(50, True,  id="TC_VoltSag_Thr_002"),
    pytest.param(90, True,  id="TC_VoltSag_Thr_003"),
    pytest.param(9,  False, id="TC_VoltSag_Thr_004"),
    pytest.param(91, False, id="TC_VoltSag_Thr_005"),
])
def test_volt_sag_threshold(app_page, value, ok):
    nav_to_event_waveform(app_page)
    set_pq_threshold(app_page, SAG_IDX, value)
    saved = save_and_check(app_page)
    _log.info("[TC] VoltSagThreshold=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Voltage Sag Threshold={value}: 保存失败"
        verify_modbus(REG_VOLT_SAG_THR, value, label=f"VoltSagThreshold({value})")
    else:
        assert not saved, f"Voltage Sag Threshold={value}: 期望校验失败但保存成功"


# ════════════════════════════════════════════════════════════════════
# TC_VoltSag_Hys_001~005  Voltage Sag Hysteresis（1–10 %）
#   寄存器 0x6504 (25860)
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("value,ok", [
    pytest.param(1,  True,  id="TC_VoltSag_Hys_001"),
    pytest.param(5,  True,  id="TC_VoltSag_Hys_002"),
    pytest.param(10, True,  id="TC_VoltSag_Hys_003"),
    pytest.param(0,  False, id="TC_VoltSag_Hys_004"),
    pytest.param(11, False, id="TC_VoltSag_Hys_005"),
])
def test_volt_sag_hysteresis(app_page, value, ok):
    nav_to_event_waveform(app_page)
    set_pq_hysteresis(app_page, SAG_IDX, value)
    saved = save_and_check(app_page)
    _log.info("[TC] VoltSagHysteresis=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Voltage Sag Hysteresis={value}: 保存失败"
        verify_modbus(REG_VOLT_SAG_HYS, value, label=f"VoltSagHysteresis({value})")
    else:
        assert not saved, f"Voltage Sag Hysteresis={value}: 期望校验失败但保存成功"


# ════════════════════════════════════════════════════════════════════
# TC_VoltSag_DO_001~003  Voltage Sag DO Output
#   寄存器 0x6506 (25862)：0=None, 1=DO1 … 8=DO8
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("None", 0, id="TC_VoltSag_DO_001"),
    pytest.param("DO1",  1, id="TC_VoltSag_DO_002"),
    pytest.param("DO8",  8, id="TC_VoltSag_DO_003"),
])
def test_volt_sag_do_output(app_page, option, encoded):
    nav_to_event_waveform(app_page)
    set_pq_do(app_page, SAG_IDX, option)
    saved = save_and_check(app_page)
    assert saved, f"Volt Sag DO={option}: 保存失败"
    verify_modbus(REG_VOLT_SAG_DO, encoded, label=f"VoltSagDO({option})")


# ════════════════════════════════════════════════════════════════════
# TC_VoltSag_RO_001~003  Voltage Sag RO Output
#   寄存器 0x6507 (25863)：0=None, 1=RO1, 2=RO2
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("None", 0, id="TC_VoltSag_RO_001"),
    pytest.param("RO1",  1, id="TC_VoltSag_RO_002"),
    pytest.param("RO2",  2, id="TC_VoltSag_RO_003"),
])
def test_volt_sag_ro_output(app_page, option, encoded):
    nav_to_event_waveform(app_page)
    set_pq_ro(app_page, SAG_IDX, option)
    saved = save_and_check(app_page)
    assert saved, f"Volt Sag RO={option}: 保存失败"
    verify_modbus(REG_VOLT_SAG_RO, encoded, label=f"VoltSagRO({option})")


# ════════════════════════════════════════════════════════════════════
# TC_VoltSwl_Thr_001~005  Voltage Swell Threshold（110–150 %）
#   寄存器 0x6509 (25865)
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("value,ok", [
    pytest.param(110, True,  id="TC_VoltSwl_Thr_001"),
    pytest.param(130, True,  id="TC_VoltSwl_Thr_002"),
    pytest.param(150, True,  id="TC_VoltSwl_Thr_003"),
    pytest.param(109, False, id="TC_VoltSwl_Thr_004"),
    pytest.param(151, False, id="TC_VoltSwl_Thr_005"),
])
def test_volt_swell_threshold(app_page, value, ok):
    nav_to_event_waveform(app_page)
    set_pq_threshold(app_page, SWL_IDX, value)
    saved = save_and_check(app_page)
    _log.info("[TC] VoltSwellThreshold=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Voltage Swell Threshold={value}: 保存失败"
        verify_modbus(REG_VOLT_SWL_THR, value, label=f"VoltSwellThreshold({value})")
    else:
        assert not saved, f"Voltage Swell Threshold={value}: 期望校验失败但保存成功"


# ════════════════════════════════════════════════════════════════════
# TC_VoltSwl_Hys_001~005  Voltage Swell Hysteresis（1–10 %）
#   寄存器 0x650A (25866)
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("value,ok", [
    pytest.param(1,  True,  id="TC_VoltSwl_Hys_001"),
    pytest.param(5,  True,  id="TC_VoltSwl_Hys_002"),
    pytest.param(10, True,  id="TC_VoltSwl_Hys_003"),
    pytest.param(0,  False, id="TC_VoltSwl_Hys_004"),
    pytest.param(11, False, id="TC_VoltSwl_Hys_005"),
])
def test_volt_swell_hysteresis(app_page, value, ok):
    nav_to_event_waveform(app_page)
    set_pq_hysteresis(app_page, SWL_IDX, value)
    saved = save_and_check(app_page)
    _log.info("[TC] VoltSwellHysteresis=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Voltage Swell Hysteresis={value}: 保存失败"
        verify_modbus(REG_VOLT_SWL_HYS, value, label=f"VoltSwellHysteresis({value})")
    else:
        assert not saved, f"Voltage Swell Hysteresis={value}: 期望校验失败但保存成功"


# ════════════════════════════════════════════════════════════════════
# TC_VoltInt_Thr_001~005  Voltage Interruption Threshold（5–20 %）
#   寄存器 0x650F (25871)
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("value,ok", [
    pytest.param(5,  True,  id="TC_VoltInt_Thr_001"),
    pytest.param(10, True,  id="TC_VoltInt_Thr_002"),
    pytest.param(20, True,  id="TC_VoltInt_Thr_003"),
    pytest.param(4,  False, id="TC_VoltInt_Thr_004"),
    pytest.param(21, False, id="TC_VoltInt_Thr_005"),
])
def test_volt_intr_threshold(app_page, value, ok):
    nav_to_event_waveform(app_page)
    set_pq_threshold(app_page, INT_IDX, value)
    saved = save_and_check(app_page)
    _log.info("[TC] VoltIntrThreshold=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Voltage Interruption Threshold={value}: 保存失败"
        verify_modbus(REG_VOLT_INT_THR, value, label=f"VoltIntrThreshold({value})")
    else:
        assert not saved, f"Voltage Interruption Threshold={value}: 期望校验失败但保存成功"


# ════════════════════════════════════════════════════════════════════
# TC_VoltInt_Hys_001~005  Voltage Interruption Hysteresis（1–10 %）
#   寄存器 0x6510 (25872)
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("value,ok", [
    pytest.param(1,  True,  id="TC_VoltInt_Hys_001"),
    pytest.param(5,  True,  id="TC_VoltInt_Hys_002"),
    pytest.param(10, True,  id="TC_VoltInt_Hys_003"),
    pytest.param(0,  False, id="TC_VoltInt_Hys_004"),
    pytest.param(11, False, id="TC_VoltInt_Hys_005"),
])
def test_volt_intr_hysteresis(app_page, value, ok):
    nav_to_event_waveform(app_page)
    set_pq_hysteresis(app_page, INT_IDX, value)
    saved = save_and_check(app_page)
    _log.info("[TC] VoltIntrHysteresis=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Voltage Interruption Hysteresis={value}: 保存失败"
        verify_modbus(REG_VOLT_INT_HYS, value, label=f"VoltIntrHysteresis({value})")
    else:
        assert not saved, f"Voltage Interruption Hysteresis={value}: 期望校验失败但保存成功"


# ════════════════════════════════════════════════════════════════════
# TC_CurrSwl_Thr_001~005  Current Swell Threshold（110–150 %）
#   寄存器 0x6515 (25877)
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("value,ok", [
    pytest.param(110, True,  id="TC_CurrSwl_Thr_001"),
    pytest.param(130, True,  id="TC_CurrSwl_Thr_002"),
    pytest.param(150, True,  id="TC_CurrSwl_Thr_003"),
    pytest.param(109, False, id="TC_CurrSwl_Thr_004"),
    pytest.param(151, False, id="TC_CurrSwl_Thr_005"),
])
def test_curr_swell_threshold(app_page, value, ok):
    nav_to_event_waveform(app_page)
    set_pq_threshold(app_page, CSW_IDX, value)
    saved = save_and_check(app_page)
    _log.info("[TC] CurrSwellThreshold=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Current Swell Threshold={value}: 保存失败"
        verify_modbus(REG_CURR_SWL_THR, value, label=f"CurrSwellThreshold({value})")
    else:
        assert not saved, f"Current Swell Threshold={value}: 期望校验失败但保存成功"


# ════════════════════════════════════════════════════════════════════
# TC_CurrSwl_Hys_001~005  Current Swell Hysteresis（1–10 %）
#   寄存器 0x6516 (25878)
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("value,ok", [
    pytest.param(1,  True,  id="TC_CurrSwl_Hys_001"),
    pytest.param(5,  True,  id="TC_CurrSwl_Hys_002"),
    pytest.param(10, True,  id="TC_CurrSwl_Hys_003"),
    pytest.param(0,  False, id="TC_CurrSwl_Hys_004"),
    pytest.param(11, False, id="TC_CurrSwl_Hys_005"),
])
def test_curr_swell_hysteresis(app_page, value, ok):
    nav_to_event_waveform(app_page)
    set_pq_hysteresis(app_page, CSW_IDX, value)
    saved = save_and_check(app_page)
    _log.info("[TC] CurrSwellHysteresis=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Current Swell Hysteresis={value}: 保存失败"
        verify_modbus(REG_CURR_SWL_HYS, value, label=f"CurrSwellHysteresis({value})")
    else:
        assert not saved, f"Current Swell Hysteresis={value}: 期望校验失败但保存成功"


# ════════════════════════════════════════════════════════════════════
# TC_WfSmpl_001~004  Waveform Sampling Rate
#   寄存器 0x6570 (25968)，HMI 标签：'Sample Rate'
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("16 points",  0, id="TC_WfSmpl_001"),
    pytest.param("32 points",  1, id="TC_WfSmpl_002"),
    pytest.param("64 points",  2, id="TC_WfSmpl_003"),
    pytest.param("128 points", 3, id="TC_WfSmpl_004"),
])
def test_waveform_sampling_rate(app_page, option, encoded):
    nav_to_event_waveform(app_page)
    set_dropdown(app_page, "Sample Rate", option)
    saved = save_and_check(app_page)
    assert saved, f"Sampling Rate={option}: 保存失败"
    verify_modbus(REG_WF_SAMPLE_RATE, encoded, label=f"SamplingRate({option})")


# ════════════════════════════════════════════════════════════════════
# TC_WfPre_001~005  Num of Cycles Before（1~159）
#   寄存器 0x6571 (25969)，HMI 标签：'Num of Cycles Before'
#
#   约束：Sample Rate(pts) × (Pre + Post) < 2560
#   前置：Sample Rate = 16 points，Post 固定为 1
#     → 合法上限：Pre < (2560/16 - 1) = 159，即最大有效值 = 158
#       16 × (158+1) = 2544 < 2560 ✓
#       16 × (159+1) = 2560 不满足严格小于 2560 ✗
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("value,ok", [
    pytest.param(1,   True,  id="TC_WfPre_001"),
    pytest.param(50,  True,  id="TC_WfPre_002"),
    pytest.param(158, True,  id="TC_WfPre_003"),   # 16×(158+1)=2544 < 2560
    pytest.param(0,   False, id="TC_WfPre_004"),   # 超出下限
    pytest.param(160, False, id="TC_WfPre_005"),   # 超出上限 1-159
])
def test_waveform_pre_cycles(app_page, value, ok):
    nav_to_event_waveform(app_page)
    set_dropdown(app_page, "Sample Rate", "16 points")   # 扩大总容量上限
    set_input(app_page, "Num of Cycles After", 1)        # Post=1，最小化占用
    set_input(app_page, "Num of Cycles Before", value)
    saved = save_and_check(app_page)
    _log.info("[TC] WfPreCycles=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Num of Cycles Before={value}: 保存失败"
        verify_modbus(REG_WF_PRE_CYCLES, value, label=f"PreCycles({value})")
    else:
        assert not saved, f"Num of Cycles Before={value}: 期望校验失败但保存成功"


# ════════════════════════════════════════════════════════════════════
# TC_WfPost_001~005  Num of Cycles After（1~159，Pre 固定为 1）
#   寄存器 0x6572 (25970)，HMI 标签：'Num of Cycles After'
#
#   前置：Sample Rate = 16 points，Pre 固定为 1
#     → 合法上限：Post ≤ 158（16×(1+158)=2544 < 2560）
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("value,ok", [
    pytest.param(1,   True,  id="TC_WfPost_001"),
    pytest.param(50,  True,  id="TC_WfPost_002"),
    pytest.param(158, True,  id="TC_WfPost_003"),
    pytest.param(0,   False, id="TC_WfPost_004"),
    pytest.param(160, False, id="TC_WfPost_005"),
])
def test_waveform_post_cycles(app_page, value, ok):
    nav_to_event_waveform(app_page)
    set_dropdown(app_page, "Sample Rate", "16 points")
    set_input(app_page, "Num of Cycles Before", 1)       # Pre=1，最小化占用
    set_input(app_page, "Num of Cycles After", value)
    saved = save_and_check(app_page)
    _log.info("[TC] WfPostCycles=%s  ok=%s  result=%s", value, ok, saved)
    if ok:
        assert saved, f"Num of Cycles After={value}: 保存失败"
        verify_modbus(REG_WF_POST_CYCLES, value, label=f"PostCycles({value})")
    else:
        assert not saved, f"Num of Cycles After={value}: 期望校验失败但保存成功"


# ════════════════════════════════════════════════════════════════════
# TC_DI1_001~003  DI1 Trigger
#   寄存器 0x6573 (25971)，HMI 标签：'DI1 Trigger'
#   选项：Disable(0) / User 1(1) / User 12(12) / Only Record Voltage(255)
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("Disable",             0,   id="TC_DI1_001"),
    pytest.param("User 1",              1,   id="TC_DI1_002"),
    pytest.param("Only Record Voltage", 255, id="TC_DI1_003"),
])
def test_waveform_di1_trigger(app_page, option, encoded):
    nav_to_event_waveform(app_page)
    set_dropdown(app_page, "DI1 Trigger", option)
    saved = save_and_check(app_page)
    assert saved, f"DI1 Trigger={option}: 保存失败"
    verify_modbus(REG_WF_DI1_TRIG, encoded, label=f"DI1Trigger({option})")


# ════════════════════════════════════════════════════════════════════
# TC_DI2_001~003  DI2 Trigger  (寄存器 0x6574)
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("Disable",             0,   id="TC_DI2_001"),
    pytest.param("User 1",              1,   id="TC_DI2_002"),
    pytest.param("Only Record Voltage", 255, id="TC_DI2_003"),
])
def test_waveform_di2_trigger(app_page, option, encoded):
    nav_to_event_waveform(app_page)
    set_dropdown(app_page, "DI2 Trigger", option)
    saved = save_and_check(app_page)
    assert saved, f"DI2 Trigger={option}: 保存失败"
    verify_modbus(REG_WF_DI2_TRIG, encoded, label=f"DI2Trigger({option})")


# ════════════════════════════════════════════════════════════════════
# TC_DI3_001~003  DI3 Trigger  (寄存器 0x6575)
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("Disable",             0,   id="TC_DI3_001"),
    pytest.param("User 1",              1,   id="TC_DI3_002"),
    pytest.param("Only Record Voltage", 255, id="TC_DI3_003"),
])
def test_waveform_di3_trigger(app_page, option, encoded):
    nav_to_event_waveform(app_page)
    set_dropdown(app_page, "DI3 Trigger", option)
    saved = save_and_check(app_page)
    assert saved, f"DI3 Trigger={option}: 保存失败"
    verify_modbus(REG_WF_DI3_TRIG, encoded, label=f"DI3Trigger({option})")


# ════════════════════════════════════════════════════════════════════
# TC_DI4_001~003  DI4 Trigger  (寄存器 0x6576)
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("Disable",             0,   id="TC_DI4_001"),
    pytest.param("User 1",              1,   id="TC_DI4_002"),
    pytest.param("Only Record Voltage", 255, id="TC_DI4_003"),
])
def test_waveform_di4_trigger(app_page, option, encoded):
    nav_to_event_waveform(app_page)
    set_dropdown(app_page, "DI4 Trigger", option)
    saved = save_and_check(app_page)
    assert saved, f"DI4 Trigger={option}: 保存失败"
    verify_modbus(REG_WF_DI4_TRIG, encoded, label=f"DI4Trigger({option})")


# ════════════════════════════════════════════════════════════════════
# TC_ManTrig_001~003  Manually Trigger  (寄存器 0x6577)
#   HMI 标签：'Manually Trigger'
# ════════════════════════════════════════════════════════════════════

@pytest.mark.parametrize("option,encoded", [
    pytest.param("Disable",             0,   id="TC_ManTrig_001"),
    pytest.param("User 1",              1,   id="TC_ManTrig_002"),
    pytest.param("Only Record Voltage", 255, id="TC_ManTrig_003"),
])
def test_waveform_manual_trigger(app_page, option, encoded):
    nav_to_event_waveform(app_page)
    set_dropdown(app_page, "Manually Trigger", option)
    saved = save_and_check(app_page)
    assert saved, f"Manually Trigger={option}: 保存失败"
    verify_modbus(REG_WF_MAN_TRIG, encoded, label=f"ManuallyTrigger({option})")
