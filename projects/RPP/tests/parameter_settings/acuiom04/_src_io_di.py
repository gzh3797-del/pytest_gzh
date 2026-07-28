# -*- coding: utf-8 -*-
"""
_src_io_di.py — AcuIOM-4 "IO → DI" 页面操作

页面路径：Settings → Devices → <DEVICE_NAME>（span.link-url）→ 详情页顶部
Meter|IO 切换按钮 → 点 IO → 二级标签 DI
URL hash（探查实测样例，随设备 hash 变化）：
  .../#/device/connection/deviceSetting/io/<设备hash>/<idx>/0/0

页面结构（探查实测，https://192.168.3.47，Element Plus v2）：
  单张 <table>，每行：checkbox(0) / ID(1,"DI n") / Function(2,el-select，选项
  "Status Monitor"/"Pulse Counter") / Pulse Constant(3,input，placeholder
  "Enter Pulse Constant") / Unit(4,input，placeholder "Enter Unit") / Action(5，
  仅 Copy 按钮，**无 Edit** —— 与 AI/AO 不同，本表所有字段均为行内直接编辑)。
  Function=Status Monitor 时 Pulse Constant/Unit 均带 disabled；切 Pulse Counter
  后 disabled 属性消失、可编辑（探查阶段仅验证了行为，未点 Save 真正下发）。
  顶部 Apply to Selected / Reset Selected；每行 Copy 点击后顶部提示变为
  "DI n Settings Copied!"；底部固定 Save 按钮。
"""
from __future__ import annotations

from playwright.sync_api import Page

from helpers_iom04 import step  # noqa: F401

# ── 导航 ──────────────────────────────────────────────────────────────


def nav_to_io(page: Page) -> None:
    """Settings → Devices → <设备> → 详情页默认 Meter → 点击顶部 IO 按钮。"""
    from helpers_iom04 import DEVICE_NAME

    # 复位到设备列表（幂等）：会话级 app_page 跨用例复用，上一条常停在设备详情 IO 页，
    # 此时点 Settings→Devices 回不到列表、span.link-url 取不到；改用 hash 路由复位，
    # 只切 hash 不整页 reload，登录态保留，从任意页面状态都能重新进设备。
    base = page.url.split("/#/")[0]
    page.goto(f"{base}/#/device/connection", wait_until="domcontentloaded")
    page.wait_for_timeout(1500)

    step(f"Click device: {DEVICE_NAME}")
    page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)
    page.locator("span.link-url", has_text=DEVICE_NAME).first.click()
    page.wait_for_selector("text=IO", timeout=10000)

    step("Click IO toggle")
    page.get_by_text("IO", exact=True).first.click()
    page.wait_for_timeout(800)


def nav_to_io_di(page: Page) -> None:
    """导航到 IO → DI 标签，等待表格行出现。"""
    nav_to_io(page)
    step("Click DI sub-tab")
    page.get_by_text("DI", exact=True).first.click()
    page.wait_for_selector("table tbody tr", timeout=10000)
    page.wait_for_timeout(500)


# ── 表格行定位 ────────────────────────────────────────────────────────

def _di_table(page: Page):
    return page.locator("table.el-table__body")


def di_row_count(page: Page) -> int:
    return _di_table(page).locator("tbody tr").count()


def _di_row(page: Page, n: int):
    """DI n(1-based) 行。"""
    return _di_table(page).locator("tbody tr").nth(n - 1)


def _function_wrapper(page: Page, n: int):
    return _di_row(page, n).locator("td").nth(2).locator(".el-select__wrapper")


def get_function_text(page: Page, n: int) -> str:
    return _function_wrapper(page, n).inner_text().strip()


def set_function(page: Page, n: int, option_text: str) -> None:
    """设置 DI n 的 Function（Status Monitor / Pulse Counter）。"""
    step(f"Set DI{n} Function -> {option_text}")
    wrapper = _function_wrapper(page, n)
    wrapper.scroll_into_view_if_needed()
    wrapper.click()
    page.wait_for_timeout(400)
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(
        option_text, exact=True
    )
    if target.count() == 0:
        page.keyboard.press("Escape")
        raise RuntimeError(f"set_function: option {option_text!r} not found for DI{n}")
    target.first.click()
    page.wait_for_timeout(300)


def _pulse_constant_input(page: Page, n: int):
    return _di_row(page, n).locator("td").nth(3).locator("input")


def _unit_input(page: Page, n: int):
    return _di_row(page, n).locator("td").nth(4).locator("input")


def pulse_constant_disabled(page: Page, n: int) -> bool:
    return _pulse_constant_input(page, n).get_attribute("disabled") is not None


def unit_disabled(page: Page, n: int) -> bool:
    return _unit_input(page, n).get_attribute("disabled") is not None


def set_pulse_constant(page: Page, n: int, value) -> None:
    step(f"Set DI{n} Pulse Constant = {value!r}")
    inp = _pulse_constant_input(page, n)
    inp.scroll_into_view_if_needed()
    inp.click()
    inp.press("Control+a")
    inp.press("Delete")
    if value not in (None, ""):
        inp.type(str(value), delay=40)


def set_unit(page: Page, n: int, value) -> None:
    step(f"Set DI{n} Unit = {value!r}")
    inp = _unit_input(page, n)
    inp.scroll_into_view_if_needed()
    inp.click()
    inp.press("Control+a")
    inp.press("Delete")
    if value not in (None, ""):
        inp.type(str(value), delay=40)


def unit_value(page: Page, n: int) -> str:
    return _unit_input(page, n).input_value()


# ── Copy / Apply to Selected / Reset Selected ────────────────────────

def click_copy(page: Page, n: int) -> None:
    step(f"Click Copy on DI{n}")
    _di_row(page, n).get_by_text("Copy", exact=True).click()
    page.wait_for_timeout(500)


def click_apply_to_selected(page: Page) -> None:
    step("Click Apply to Selected")
    page.get_by_text("Apply to Selected", exact=True).click()
    page.wait_for_timeout(500)


def click_reset_selected(page: Page) -> None:
    step("Click Reset Selected")
    page.get_by_text("Reset Selected", exact=True).click()
    page.wait_for_timeout(500)


def select_row_checkbox(page: Page, n: int) -> None:
    step(f"Select checkbox DI{n}")
    row = _di_row(page, n)
    cb = row.locator(".el-checkbox__inner").first
    cb.scroll_into_view_if_needed()
    cb.click()
    page.wait_for_timeout(300)


def copied_message_text(page: Page) -> str:
    msg = page.locator("text=Settings Copied")
    return msg.first.inner_text().strip() if msg.count() > 0 else ""


# 让 from _src_io_di import * 包含工具名（须放文件最末）
__all__ = [_n for _n in globals() if not _n.startswith('__')]
