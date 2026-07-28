# -*- coding: utf-8 -*-
"""
helpers_iom04.py — AcuIOM-4（28 路 DI + 4 路 DO + 2 路 RO）页面操作 + Modbus 校验公共函数

参照 acurev4100/helpers_4100.py 的结构与约定（team conventions #9/#12）。

页面路径：Settings → Devices → <设备>（span.link-url） → 详情页顶部 Meter|IO 切换
→ IO → DI / DO / RO 二级标签
探查环境：https://192.168.3.47（ACM-41-WEB2 Gateway Mode，账号 l/1，探查用只读，
未点击任何真实 Save）

寄存器来源：knowledge/shared/templates/raw/AcuIOM-04_v1.01_20250205.xlsx（blockParams sheet）
  DI 配置块基址 8192，每路跨度 3 regs：
    +0  DIn Type            uint16   (Function 编码：0=Status Monitor / 1=Pulse Counter，推测)
    +1  DIn pulse constant  uint32×0.001  (Pulse Constant，2 regs)
  DI Unit 块基址 8448，每路跨度 2 regs（uint32，ASCII 打包推测，见 read_unit_ascii）
  DO 配置块基址 16384，每路跨度 2 regs：
    +0  DOn Type         uint16   (Control Mode 编码：0/1=Manual/Pulse，TODO 顺序未确认)
    +1  DOn Pulse width  uint16   (Pulse Width，无缩放，单位 ms)
  RO 配置块基址 24576，每路跨度 2 regs（结构同 DO）：
    +0  ROn Type          uint16
    +1  ROn Pulse Width   uint16
"""
from __future__ import annotations

import logging
import time as _time

import allure
from playwright.sync_api import Page
from pymodbus.client import ModbusTcpClient

# ── 设备连接常量：由 conftest.py 的 _bind_acuiom04_device（autouse，会话级）
# 在测试开始前通过网关 API 动态发现当前在线的 AcuIOM-4 设备后写入真实值，
# 替代静态配置——避免物理设备切换/离线后仍连着旧地址。以下仅为占位默认值。
DEVICE_NAME = ""
MODBUS_HOST = ""
MODBUS_PORT = 502
SLAVE_ID = 1

# ── 通道数量（AcuIOM-4 = 28 DI + 4 DO + 2 RO） ───────────────────────
DI_COUNT = 28
DO_COUNT = 4
RO_COUNT = 2

# ── DI 配置寄存器 ────────────────────────────────────────────────────
REG_DI_BASE = 8192
DI_STRIDE = 3
REG_DI_UNIT_BASE = 8448
DI_UNIT_STRIDE = 2


def di_type_reg(n: int) -> int:
    """DI n(1-based) Type 寄存器（Function 编码，uint16）。"""
    return REG_DI_BASE + (n - 1) * DI_STRIDE


def di_pulse_constant_reg(n: int) -> int:
    """DI n Pulse Constant（uint32×0.001，2 regs）。"""
    return di_type_reg(n) + 1


def di_unit_reg(n: int) -> int:
    """DI n Unit（Eng. Unit，uint32，2 regs）。"""
    return REG_DI_UNIT_BASE + (n - 1) * DI_UNIT_STRIDE


# ── DO 配置寄存器 ────────────────────────────────────────────────────
REG_DO_BASE = 16384
DO_STRIDE = 2


def do_type_reg(n: int) -> int:
    """DO n(1-based) Type 寄存器（Control Mode 编码，uint16）。"""
    return REG_DO_BASE + (n - 1) * DO_STRIDE


def do_pulse_width_reg(n: int) -> int:
    """DO n Pulse Width（uint16，无缩放，单位 ms）。"""
    return do_type_reg(n) + 1


# ── RO 配置寄存器 ────────────────────────────────────────────────────
REG_RO_BASE = 24576
RO_STRIDE = 2


def ro_type_reg(n: int) -> int:
    """RO n(1-based) Type 寄存器（Control Mode 编码，uint16）。"""
    return REG_RO_BASE + (n - 1) * RO_STRIDE


def ro_pulse_width_reg(n: int) -> int:
    """RO n Pulse Width(ms)（uint16，无缩放）。"""
    return ro_type_reg(n) + 1


# ── 下拉编码映射 ─────────────────────────────────────────────────────
# TODO(需真机确认)：探查阶段只做了只读页面探查（未做 Modbus 回读联动验证），
# 以下编码顺序均为**假定值**，建议真机跑通各 _001/_002 冒烟用例时用实际回读值
# 反推校正。
FUNCTION_ENCODE = {
    "Status Monitor": 0,
    "Pulse Counter":  1,
}

CONTROL_MODE_ENCODE = {
    "Manual": 0,
    "Pulse":  1,
}

_log = logging.getLogger("acuiom04_test")


# ════════════════════════════════════════════════════════════════════
# 工具函数
# ════════════════════════════════════════════════════════════════════

def step(msg: str) -> None:
    _log.info("[STEP] %s", msg)
    with allure.step(msg):
        pass


# ── 持久 Modbus 连接（整个进程共享，按需重连） ───────────────────────
_modbus_client: ModbusTcpClient | None = None


def _get_modbus_client() -> ModbusTcpClient:
    """返回已连接的 Modbus 客户端，断线时自动重连。"""
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
    global _modbus_client
    _log.info("[MODBUS TX] FC=03  addr=%d(0x%04X)  count=%d", address, address, count)
    _retries = 5
    _retry_delay = 3
    last_err = None
    for attempt in range(_retries):
        if attempt > 0:
            _log.warning("[MODBUS RETRY] attempt %d/%d, wait %ds ...",
                         attempt + 1, _retries, _retry_delay)
            _time.sleep(_retry_delay)
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
        except Exception as exc:  # noqa: BLE001  重试兜底
            last_err = f"Exception at addr={address}: {exc}"
            _log.warning("[MODBUS ERROR] %s", last_err)
            _modbus_client = None
            continue
    raise RuntimeError(
        f"Modbus read failed after {_retries} attempts. addr={address}  Last error: {last_err}"
    )


def modbus_read_u32(address: int) -> int:
    """读取 32-bit 无符号整数（高寄存器在前，big-endian）。"""
    regs = modbus_read(address, count=2)
    return (regs[0] << 16) | regs[1]


def modbus_read_32(address: int) -> int:
    """向后兼容别名：无符号 32-bit 读取（同 modbus_read_u32）。"""
    return modbus_read_u32(address)


@allure.step("验证寄存器 [{label}]  addr={address}  expected={expected}")
def verify_modbus(address: int, expected, label: str = "") -> None:
    """验证 16-bit 寄存器整型值。"""
    regs = modbus_read(address, count=1)
    actual = regs[0]
    ok = actual == expected
    _log.info("[VERIFY] %-40s  expected=%-6s  actual=%-6s  -> %s",
              label, expected, actual, "PASS" if ok else "FAIL")
    assert ok, f"{label}: expected {expected}, got {actual}"


@allure.step("验证缩放寄存器 [{label}]  addr={address}  expected={expected_value}")
def verify_modbus_scaled(
    address: int,
    expected_value: float,
    scale: float = 0.001,
    label: str = "",
    tol: float = 6e-4,
) -> None:
    """验证 ×scale 缩放的 32-bit 无符号寄存器（如 Pulse Constant）。"""
    raw = modbus_read_u32(address)
    actual = raw * scale
    ok = abs(actual - expected_value) < tol
    _log.info("[VERIFY SCALED] %-40s  expected=%-10s  actual=%-10s  raw=%-10s -> %s",
              label, expected_value, actual, raw, "PASS" if ok else "FAIL")
    assert ok, f"{label}: expected {expected_value}, got {actual} (raw={raw})"


def read_unit_ascii(address: int, num_regs: int = 2) -> str:
    """读取 Eng. Unit 寄存器并按 ASCII 解包（每寄存器 2 字节，高字节在前）。

    TODO(需真机确认)：模板列出该字段 dataType=uint32（本设备 num_regs=2，
    即最多 4 个 ASCII 字符），此处按 AcuRev4100 User Channel Description 同款
    ASCII 打包模式**推测**实现，未做 Modbus 回读比对验证，请真机实测确认。
    """
    regs = modbus_read(address, count=num_regs)
    chars = []
    for reg in regs:
        chars.append((reg >> 8) & 0xFF)
        chars.append(reg & 0xFF)
    raw = bytes(b for b in chars if b != 0)
    return raw.decode("ascii", errors="replace").rstrip()


# ── 页面操作：通用（label-based，若后续需要 Edit 弹窗时复用） ────────

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
    inp.first.type(str(value), delay=40)
    _log.info("  typed %r", value)


@allure.step("下拉选择 [{label}] -> {option_text}")
def set_dropdown(page: Page, label: str, option_text: str) -> None:
    """点击下拉框并选择指定选项（Element Plus v2，aria-controls 精确定位 popper）。"""
    step(f"Dropdown [{label}] -> {option_text!r}")
    form_item = page.locator(".el-form-item").filter(has_text=label)
    if form_item.count() > 0:
        sel = form_item.first.locator(".el-select").first
    else:
        sel = page.get_by_text(label, exact=False).first.locator(
            "xpath=following::div[contains(@class,'el-select')][1]"
        )
    sel.scroll_into_view_if_needed()
    wrapper = sel.locator(".el-select__wrapper")
    trigger = wrapper.first if wrapper.count() > 0 else sel
    trigger.click()
    page.wait_for_timeout(400)

    aria_input = sel.locator(".el-select__input")
    aria = aria_input.get_attribute("aria-controls") if aria_input.count() > 0 else None
    listbox = page.locator(f"[id='{aria}']") if aria else page

    target = listbox.locator(".el-select-dropdown__item:visible").get_by_text(
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
    target.first.scroll_into_view_if_needed()
    target.first.click()
    page.wait_for_timeout(300)


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
    _log.info("[VERIFY ERROR] expected=%r  visible_errors=%s", expected_text, errors)
    if not matched:
        raise AssertionError(f"Expected error text {expected_text!r}\nVisible errors: {errors}")


@allure.step("点击 Save 并检查结果")
def save_and_check(page: Page) -> bool:
    """点击底部固定 Save 按钮，先看内联校验错误，再看顶部 toast。"""
    step("Click Save")
    page.locator("button:has-text('Save')").last.click()
    page.wait_for_timeout(1500)

    form_errors = get_visible_errors(page)
    if form_errors:
        _log.info("[SAVE] FORM VALIDATION FAILED  errors=%s", form_errors)
        return False

    page.wait_for_timeout(2500)
    loc = page.locator(".el-message--success")
    if loc.count() > 0 and loc.first.is_visible():
        _log.info("[SAVE] SUCCESS toast: %r", loc.first.inner_text()[:80])
        return True
    for sel in [".el-message--error", ".el-message--warning"]:
        loc = page.locator(sel)
        if loc.count() > 0 and loc.first.is_visible():
            _log.info("[SAVE] ERROR/WARNING toast: %r", loc.first.inner_text()[:80])
            return False

    _log.info("[SAVE] No toast -> assuming success")
    return True


# 让 from helpers_iom04 import * 包含 _log 等下划线名称（须放文件最末）
__all__ = [_n for _n in globals() if not _n.startswith('__')]
