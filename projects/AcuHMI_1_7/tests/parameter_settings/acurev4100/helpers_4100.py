# ── test_general_settings.py ──
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
from _src_event_waveform import (  # noqa: F401
    nav_to_event_waveform,
    set_pq_threshold,
    set_pq_hysteresis,
    set_pq_do,
    set_pq_ro,
    SAG_IDX,
    SWL_IDX,
    INT_IDX,
    CSW_IDX,
)

logging.getLogger("pymodbus").setLevel(logging.DEBUG)

# ── 设备配置：由 conftest.py 的 _bind_acurev4100_device（autouse，会话级）
# 在测试开始前通过网关 API 动态发现当前在线的 AcuRev4100 设备后写入真实值，
# 替代原来从 config.yaml device_modbus 段静态解析——避免物理设备切换/离线后
# 仍连着一个旧地址。以下仅为占位默认值，实际连接前会被覆盖。
DEVICE_NAME = ""
MODBUS_HOST = ""
MODBUS_PORT = 502
SLAVE_ID    = 1

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
    """返回已连接的 Modbus 客户端，断线时自动重连。"""
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
            _log.warning("[MODBUS RETRY] attempt %d/%d, wait %ds ...",
                         attempt + 1, _RETRIES, _RETRY_DELAY)
            _time.sleep(_RETRY_DELAY)
            _modbus_client = None  # 强制重建连接
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
        f"Modbus read failed after {_RETRIES} attempts. "
        f"addr={address}  Last error: {last_err}"
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
    _settings_loc = page.locator(
        "span:has-text('Settings'), a:has-text('Settings'), button:has-text('Settings')"
    )
    _general_loc = page.locator(
        "a:has-text('General'), li:has-text('General'), span:has-text('General')"
    )
    # 点击 Settings 展开子菜单；确认 General 可见后立即点击（避免菜单动画再次收起）
    step("Click General submenu")
    clicked = False
    for _ in range(4):
        _settings_loc.first.click()
        page.wait_for_timeout(700)
        if _general_loc.first.is_visible():
            try:
                _general_loc.first.click(timeout=3000)
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
        if _general_loc.first.is_visible():
            _general_loc.first.click()
        else:
            _general_loc.first.click(force=True, timeout=5000)
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
        _settings_loc.first.click()
        page.wait_for_timeout(700)
        if _general_loc.first.is_visible():
            _general_loc.first.click()
        else:
            _general_loc.first.click(force=True, timeout=5000)
        page.wait_for_selector("button:has-text('Save')", timeout=15000)
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

    1. DI Trigger 系列（E&W 页面无 el-form-item 标签）用位置索引定位 el-select。
    2. 其余字段：优先通过 .el-form-item 容器找 .el-select，点击其内部 .el-input__inner。
    3. 重试最多 3 次；3 次全失败后 for…else 做一次干净重开。
    4. 找到 target 后等 300ms 动画稳定再 scroll_into_view + click。
    """
    step(f"Dropdown [{label}] → {option_text!r}")

    # ── DI Trigger 特殊处理：E&W 页面 DI 区无 el-form-item，按位置索引 ──
    _DI_IDX = {"DI1 Trigger": 0, "DI2 Trigger": 1, "DI3 Trigger": 2, "DI4 Trigger": 3}
    if label in _DI_IDX:
        sel = page.locator(".el-select").nth(_DI_IDX[label])
        sel.scroll_into_view_if_needed()
        _inner = sel.locator(".el-input__inner")
        trigger = _inner.first if _inner.count() > 0 else sel
    else:
        # 优先：在包含 label 的 el-form-item 中找 el-select，点击内部 input
        form_item = page.locator(".el-form-item").filter(has_text=label)
        if form_item.count() > 0:
            sel = form_item.first.locator(".el-select").first
            sel.scroll_into_view_if_needed()
            _inner = sel.locator(".el-input__inner")
            trigger = _inner.first if _inner.count() > 0 else sel
        else:
            # 回退：XPath following 定位
            trigger = page.get_by_text(label, exact=False).first.locator(
                "xpath=following::div[contains(@class,'el-select')][1]"
            )
            trigger.scroll_into_view_if_needed()

    # 点击触发器并等待选项出现，重试最多 3 次
    for _attempt in range(3):
        trigger.click()
        try:
            page.wait_for_selector(".el-select-dropdown__item:visible", timeout=3000)
            if page.locator(".el-select-dropdown__item:visible").count() > 0:
                break
        except Exception:
            pass
        _log.warning("  set_dropdown attempt %d: no visible items, retrying…", _attempt + 1)
        page.keyboard.press("Escape")
        page.wait_for_timeout(600)
    else:
        # 3 次重试均未打开 → 关闭动画结束后再做一次干净点击
        page.keyboard.press("Escape")
        page.wait_for_timeout(400)
        trigger.click()
        page.wait_for_timeout(1000)

    # 等待动画稳定后再查找选项
    page.wait_for_timeout(300)

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
    target.first.scroll_into_view_if_needed()
    page.wait_for_timeout(200)
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



# ── 补充缺失常量（必须在 __all__ 之前）──────────────────────────────
_ERR_NOM_CUR = "Nominal Current must"
_ERR_DMND_INT = "must be between 1 and 30"
_ERR_BACKLIGHT = "must be between 0 and 120"
_ERR_ENERGY_PULSE = "must be between"
_ERR_LED_WIDTH = "must be between 20 and 100"
REG_VOLT_SAG_THR = 25859   # 0x6503  Threshold    10–90 %
REG_VOLT_SAG_HYS = 25860   # 0x6504  Hysteresis   1–10 %
REG_VOLT_SAG_DO  = 25862   # 0x6506  DO Output    0–8
REG_VOLT_SAG_RO  = 25863   # 0x6507  RO Output    0–2
REG_VOLT_SWL_THR = 25865   # 0x6509  Threshold   110–150 %
REG_VOLT_SWL_HYS = 25866   # 0x650A  Hysteresis   1–10 %
REG_VOLT_SWL_DO  = 25868   # 0x650C  DO Output
REG_VOLT_SWL_RO  = 25869   # 0x650D  RO Output
REG_VOLT_INT_THR = 25871   # 0x650F  Threshold    5–20 %
REG_VOLT_INT_HYS = 25872   # 0x6510  Hysteresis   1–10 %
REG_VOLT_INT_DO  = 25874   # 0x6512  DO Output
REG_VOLT_INT_RO  = 25875   # 0x6513  RO Output
REG_CURR_SWL_THR = 25877   # 0x6515  Threshold   110–150 %
REG_CURR_SWL_HYS = 25878   # 0x6516  Hysteresis   1–10 %
REG_CURR_SWL_DO  = 25880   # 0x6518  DO Output
REG_CURR_SWL_RO  = 25881   # 0x6519  RO Output
REG_WF_SAMPLE_RATE = 25968  # 0x6570  0=16pts 1=32pts 2=64pts 3=128pts
REG_WF_PRE_CYCLES  = 25969  # 0x6571  Pre-event cycles  1–159
REG_WF_POST_CYCLES = 25970  # 0x6572  Post-event cycles 1–(160-pre)
REG_WF_DI1_TRIG    = 25971  # 0x6573  0=Disable 1-12=User1-12 255=OnlyVoltage
REG_WF_DI2_TRIG    = 25972  # 0x6574
REG_WF_DI3_TRIG    = 25973  # 0x6575
REG_WF_DI4_TRIG    = 25974  # 0x6576
REG_WF_MAN_TRIG    = 25975  # 0x6577

# 让 from helpers import * 包含 _log 等下划线名称（必须放文件最末）
__all__ = [_n for _n in globals() if not _n.startswith('__')]
