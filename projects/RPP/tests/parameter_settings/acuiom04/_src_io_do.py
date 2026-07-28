# -*- coding: utf-8 -*-
"""
_src_io_do.py — AcuIOM-4 "IO → DO" 页面操作

页面结构（探查实测，https://192.168.3.47，Element Plus v2）：
  单张 <table>，每行：checkbox(0) / ID(1,"DO n") / Control Mode(2,el-select，
  实测样例 "Pulse"/"Manual"，完整选项集未在探查中抓全) / Pulse Width(3,input，
  无单位标注) / Action(4，仅 Copy，无 Edit)。
  顶部 Apply to Selected / Reset Selected；Copy 点击后顶部提示变为
  "DO n Settings Copied!"；底部固定 Save 按钮。
"""
from __future__ import annotations

from playwright.sync_api import Page

from helpers_iom04 import step  # noqa: F401
from _src_io_di import nav_to_io  # noqa: F401  复用与 DI 标签共通的导航


def nav_to_io_do(page: Page) -> None:
    """导航到 IO → DO 标签，等待表格行出现。"""
    nav_to_io(page)
    step("Click DO sub-tab")
    page.get_by_text("DO", exact=True).first.click()
    page.wait_for_selector("table tbody tr", timeout=10000)
    page.wait_for_timeout(500)


def _do_table(page: Page):
    return page.locator("table.el-table__body")


def do_row_count(page: Page) -> int:
    return _do_table(page).locator("tbody tr").count()


def _do_row(page: Page, n: int):
    """DO n(1-based) 行。"""
    return _do_table(page).locator("tbody tr").nth(n - 1)


def _control_mode_wrapper(page: Page, n: int):
    return _do_row(page, n).locator("td").nth(2).locator(".el-select__wrapper")


def get_control_mode_text(page: Page, n: int) -> str:
    return _control_mode_wrapper(page, n).inner_text().strip()


def set_control_mode(page: Page, n: int, option_text: str) -> None:
    """设置 DO n 的 Control Mode。"""
    step(f"Set DO{n} Control Mode -> {option_text}")
    wrapper = _control_mode_wrapper(page, n)
    wrapper.scroll_into_view_if_needed()
    wrapper.click()
    page.wait_for_timeout(400)
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(
        option_text, exact=True
    )
    if target.count() == 0:
        page.keyboard.press("Escape")
        raise RuntimeError(f"set_control_mode: option {option_text!r} not found for DO{n}")
    target.first.click()
    page.wait_for_timeout(300)


def _pulse_width_input(page: Page, n: int):
    return _do_row(page, n).locator("td").nth(3).locator("input")


def set_pulse_width(page: Page, n: int, value) -> None:
    step(f"Set DO{n} Pulse Width = {value!r}")
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
    step(f"Click Copy on DO{n}")
    _do_row(page, n).get_by_text("Copy", exact=True).click()
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
    step(f"Select checkbox DO{n}")
    row = _do_row(page, n)
    cb = row.locator(".el-checkbox__inner").first
    cb.scroll_into_view_if_needed()
    cb.click()
    page.wait_for_timeout(300)


def copied_message_text(page: Page) -> str:
    msg = page.locator("text=Settings Copied")
    return msg.first.inner_text().strip() if msg.count() > 0 else ""


# 让 from _src_io_do import * 包含工具名（须放文件最末）
__all__ = [_n for _n in globals() if not _n.startswith('__')]
