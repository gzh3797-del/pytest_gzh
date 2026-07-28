# -*- coding: utf-8 -*-
"""
_src_io_ao.py — AcuIOM-2 "IO → AO" 页面操作

页面结构与 acuiom01/_src_io_ao.py 完全同构（同一前端组件，仅通道数不同：4 AO）。
Edit 弹窗 "AO n Details" 内 Breakpoints 列序与 AI **相反**："Eng. Value (x)" ->
"Output (V)"。
"""
from __future__ import annotations

from playwright.sync_api import Page

from helpers_iom02 import step  # noqa: F401
from _src_io_ai import (  # noqa: F401  复用与 AI 标签共通的导航/弹窗/Copy 操作
    nav_to_io,
    _dialog,
    dialog_set_input,
    dialog_input_value,
    dialog_set_dropdown,
    dialog_dropdown_text,
    set_segments,
    dialog_confirm,
    dialog_cancel,
    click_apply_to_selected,
    click_reset_selected,
)


def nav_to_io_ao(page: Page) -> None:
    """导航到 IO → AO 标签，等待表格行出现。"""
    nav_to_io(page)
    step("Click AO sub-tab")
    page.get_by_text("AO", exact=True).first.click()
    page.wait_for_selector("table tbody tr", timeout=10000)
    page.wait_for_timeout(500)


def _ao_table(page: Page):
    return page.locator("table.el-table__body")


def ao_row_count(page: Page) -> int:
    return _ao_table(page).locator("tbody tr").count()


def _ao_row(page: Page, n: int):
    """AO n(1-based) 行。"""
    return _ao_table(page).locator("tbody tr").nth(n - 1)


def _signal_type_wrapper(page: Page, n: int):
    return _ao_row(page, n).locator("td").nth(2).locator(".el-select__wrapper")


def get_signal_type_text(page: Page, n: int) -> str:
    return _signal_type_wrapper(page, n).inner_text().strip()


def set_signal_type(page: Page, n: int, option_text: str) -> None:
    """表格行内直接设置 AO n 的 Signal Type（不经 Edit 弹窗）。"""
    step(f"Set AO{n} Signal Type -> {option_text}")
    wrapper = _signal_type_wrapper(page, n)
    wrapper.scroll_into_view_if_needed()
    wrapper.click()
    page.wait_for_timeout(400)
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(
        option_text, exact=True
    )
    if target.count() == 0:
        page.keyboard.press("Escape")
        raise RuntimeError(f"set_signal_type: option {option_text!r} not found for AO{n}")
    target.first.click()
    page.wait_for_timeout(300)


def open_edit(page: Page, n: int) -> None:
    """点击 AO n 行的 Edit 按钮，等待 "AO n Details" 弹窗出现。"""
    step(f"Open AO{n} Edit dialog")
    _ao_row(page, n).get_by_text("Edit", exact=True).click()
    page.wait_for_selector(".el-dialog", timeout=5000)
    page.wait_for_timeout(500)


def _breakpoints_table(page: Page):
    return _dialog(page).locator("table.el-table__body")


def breakpoint_input(page: Page, point: int, col: int):
    """断点 Point(1~4) 行的第 col 列输入框。

    AO 列序与 AI 相反：col=1 -> "Eng. Value (x)" 列；col=2 -> "Output (V)" 列。
    """
    row = _breakpoints_table(page).locator("tbody tr").nth(point - 1)
    return row.locator("td").nth(col).locator("input")


def breakpoint_value(page: Page, point: int, col: int) -> str:
    """读取断点 Point(1~4) 第 col 列输入框当前值。"""
    return breakpoint_input(page, point, col).input_value()


def set_breakpoint(page: Page, point: int, col: int, value) -> None:
    step(f"[Dialog] Set AO breakpoint Point{point} col{col} = {value!r}")
    inp = breakpoint_input(page, point, col)
    inp.scroll_into_view_if_needed()
    inp.click()
    inp.press("Control+a")
    inp.press("Delete")
    inp.type(str(value), delay=30)


def breakpoint_disabled(page: Page, point: int, col: int) -> bool:
    return breakpoint_input(page, point, col).get_attribute("disabled") is not None


def click_copy(page: Page, n: int) -> None:
    step(f"Click Copy on AO{n}")
    _ao_row(page, n).get_by_text("Copy", exact=True).click()
    page.wait_for_timeout(500)


def select_row_checkbox(page: Page, n: int) -> None:
    step(f"Select checkbox AO{n}")
    row = _ao_row(page, n)
    cb = row.locator(".el-checkbox__inner").first
    cb.scroll_into_view_if_needed()
    cb.click()
    page.wait_for_timeout(300)


def copied_message_text(page: Page) -> str:
    msg = page.locator("text=Settings Copied")
    return msg.first.inner_text().strip() if msg.count() > 0 else ""


# 让 from _src_io_ao import * 包含工具名（须放文件最末）
__all__ = [_n for _n in globals() if not _n.startswith('__')]
