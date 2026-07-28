# -*- coding: utf-8 -*-
"""
_src_io_ro.py — AcuIOM-4 "IO → RO" 页面操作

页面结构与 _src_io_do.py 的 DO 标签同构（同一 <table> 组件），列头为
"Pulse Width (ms)"。checkbox(0) / ID(1,"RO n") / Control Mode(2,el-select) /
Pulse Width(ms)(3,input) / Action(4，仅 Copy，无 Edit)。
"""
from __future__ import annotations

from playwright.sync_api import Page

from helpers_iom04 import step  # noqa: F401
from _src_io_di import nav_to_io  # noqa: F401


def nav_to_io_ro(page: Page) -> None:
    """导航到 IO → RO 标签，等待表格行出现。"""
    nav_to_io(page)
    step("Click RO sub-tab")
    page.get_by_text("RO", exact=True).first.click()
    page.wait_for_selector("table tbody tr", timeout=10000)
    page.wait_for_timeout(500)


def _ro_table(page: Page):
    return page.locator("table.el-table__body")


def ro_row_count(page: Page) -> int:
    return _ro_table(page).locator("tbody tr").count()


def _ro_row(page: Page, n: int):
    """RO n(1-based) 行。"""
    return _ro_table(page).locator("tbody tr").nth(n - 1)


def _control_mode_wrapper(page: Page, n: int):
    return _ro_row(page, n).locator("td").nth(2).locator(".el-select__wrapper")


def get_control_mode_text(page: Page, n: int) -> str:
    return _control_mode_wrapper(page, n).inner_text().strip()


def set_control_mode(page: Page, n: int, option_text: str) -> None:
    step(f"Set RO{n} Control Mode -> {option_text}")
    wrapper = _control_mode_wrapper(page, n)
    wrapper.scroll_into_view_if_needed()
    wrapper.click()
    page.wait_for_timeout(400)
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(
        option_text, exact=True
    )
    if target.count() == 0:
        page.keyboard.press("Escape")
        raise RuntimeError(f"set_control_mode: option {option_text!r} not found for RO{n}")
    target.first.click()
    page.wait_for_timeout(300)


def _pulse_width_input(page: Page, n: int):
    return _ro_row(page, n).locator("td").nth(3).locator("input")


def set_pulse_width(page: Page, n: int, value) -> None:
    step(f"Set RO{n} Pulse Width = {value!r}")
    inp = _pulse_width_input(page, n)
    inp.scroll_into_view_if_needed()
    inp.click()
    inp.press("Control+a")
    inp.press("Delete")
    if value not in (None, ""):
        inp.type(str(value), delay=40)


def pulse_width_value(page: Page, n: int) -> str:
    return _pulse_width_input(page, n).input_value()


def click_copy(page: Page, n: int) -> None:
    step(f"Click Copy on RO{n}")
    _ro_row(page, n).get_by_text("Copy", exact=True).click()
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
    step(f"Select checkbox RO{n}")
    row = _ro_row(page, n)
    cb = row.locator(".el-checkbox__inner").first
    cb.scroll_into_view_if_needed()
    cb.click()
    page.wait_for_timeout(300)


def copied_message_text(page: Page) -> str:
    msg = page.locator("text=Settings Copied")
    return msg.first.inner_text().strip() if msg.count() > 0 else ""


# 让 from _src_io_ro import * 包含工具名（须放文件最末）
__all__ = [_n for _n in globals() if not _n.startswith('__')]
