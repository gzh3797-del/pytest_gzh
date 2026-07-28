# -*- coding: utf-8 -*-
"""
_src_io_ai.py — AcuIOM-2 "IO → AI" 页面操作

页面结构与 acuiom01/_src_io_ai.py 完全同构（同一前端组件，仅通道数不同：
16 AI）。详见 acuiom01/_src_io_ai.py 顶部说明。
"""
from __future__ import annotations

from playwright.sync_api import Page

from helpers_iom02 import step  # noqa: F401

# ── 导航 ──────────────────────────────────────────────────────────────


def nav_to_io(page: Page) -> None:
    """Settings → Devices → <设备> → 详情页默认 Meter → 点击顶部 IO 按钮。"""
    from helpers_iom02 import DEVICE_NAME

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


def nav_to_io_ai(page: Page) -> None:
    """导航到 IO → AI 标签，等待表格行出现。"""
    nav_to_io(page)
    step("Click AI sub-tab")
    page.get_by_text("AI", exact=True).first.click()
    page.wait_for_selector("table tbody tr", timeout=10000)
    page.wait_for_timeout(500)


# ── 表格行定位 ────────────────────────────────────────────────────────

def _ai_table(page: Page):
    return page.locator("table.el-table__body")


def ai_row_count(page: Page) -> int:
    return _ai_table(page).locator("tbody tr").count()


def _ai_row(page: Page, n: int):
    """AI n(1-based) 行。"""
    return _ai_table(page).locator("tbody tr").nth(n - 1)


def _signal_type_wrapper(page: Page, n: int):
    return _ai_row(page, n).locator("td").nth(2).locator(".el-select__wrapper")


def get_signal_type_text(page: Page, n: int) -> str:
    """读取 AI n 当前 Signal Type 显示文本。"""
    return _signal_type_wrapper(page, n).inner_text().strip()


def set_signal_type(page: Page, n: int, option_text: str) -> None:
    """表格行内直接设置 AI n 的 Signal Type（不经 Edit 弹窗）。"""
    step(f"Set AI{n} Signal Type -> {option_text}")
    wrapper = _signal_type_wrapper(page, n)
    wrapper.scroll_into_view_if_needed()
    wrapper.click()
    page.wait_for_timeout(400)
    target = page.locator(".el-select-dropdown__item:visible").get_by_text(
        option_text, exact=True
    )
    if target.count() == 0:
        page.keyboard.press("Escape")
        raise RuntimeError(f"set_signal_type: option {option_text!r} not found for AI{n}")
    target.first.click()
    page.wait_for_timeout(300)


# ── Edit 弹窗 ─────────────────────────────────────────────────────────

def open_edit(page: Page, n: int) -> None:
    """点击 AI n 行的 Edit 按钮，等待 "AI n Details" 弹窗出现。"""
    step(f"Open AI{n} Edit dialog")
    _ai_row(page, n).get_by_text("Edit", exact=True).click()
    page.wait_for_selector(".el-dialog", timeout=5000)
    page.wait_for_timeout(500)


def _dialog(page: Page):
    return page.locator(".el-dialog").first


def dialog_set_input(page: Page, label: str, value) -> None:
    """在 Edit 弹窗内按 label 设置输入框（Input Lower/Upper Limit / Eng. Unit）。"""
    step(f"[Dialog] Set [{label}] = {value!r}")
    dlg = _dialog(page)
    label_el = dlg.get_by_text(label, exact=False)
    if label_el.count() == 0:
        raise RuntimeError(f"dialog_set_input: label {label!r} not found in dialog")
    inp = label_el.first.locator("xpath=following::input[1]")
    inp.first.scroll_into_view_if_needed()
    inp.first.click()
    inp.first.press("Control+a")
    inp.first.press("Delete")
    if value not in (None, ""):
        inp.first.type(str(value), delay=40)


def dialog_input_value(page: Page, label: str) -> str:
    dlg = _dialog(page)
    label_el = dlg.get_by_text(label, exact=False)
    inp = label_el.first.locator("xpath=following::input[1]")
    return inp.first.input_value()


def dialog_set_dropdown(page: Page, label: str, option_text: str) -> None:
    """在 Edit 弹窗内按 label 设置下拉（Signal Type / Number of Segments）。"""
    step(f"[Dialog] Dropdown [{label}] -> {option_text!r}")
    dlg = _dialog(page)
    form_item = dlg.locator(".el-form-item").filter(has_text=label)
    sel = form_item.first.locator(".el-select").first
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
        page.keyboard.press("Escape")
        raise RuntimeError(f"dialog_set_dropdown: option {option_text!r} not found for [{label}]")
    target.first.click()
    page.wait_for_timeout(300)


def set_segments(page: Page, n_segments: int) -> None:
    """设置 Edit 弹窗内 "Number of Segments" 下拉（选项 1/2/3）。"""
    dialog_set_dropdown(page, "Number of Segments", str(n_segments))


def dialog_dropdown_text(page: Page, label: str) -> str:
    """读取 Edit 弹窗内某下拉（如 "Number of Segments"）当前显示文本。"""
    dlg = _dialog(page)
    form_item = dlg.locator(".el-form-item").filter(has_text=label)
    sel = form_item.first.locator(".el-select").first
    wrapper = sel.locator(".el-select__wrapper")
    return (wrapper.first if wrapper.count() > 0 else sel).inner_text().strip()


def _breakpoints_table(page: Page):
    return _dialog(page).locator("table.el-table__body")


def breakpoint_input(page: Page, point: int, col: int):
    """断点 Point(1~4) 行的第 col 列输入框。

    col=1 -> "Input (V)" 列；col=2 -> "Eng. Value (x)" 列（AI 方向的列序）。
    """
    row = _breakpoints_table(page).locator("tbody tr").nth(point - 1)
    return row.locator("td").nth(col).locator("input")


def breakpoint_value(page: Page, point: int, col: int) -> str:
    """读取断点 Point(1~4) 第 col 列输入框当前值。"""
    return breakpoint_input(page, point, col).input_value()


def set_breakpoint(page: Page, point: int, col: int, value) -> None:
    step(f"[Dialog] Set breakpoint Point{point} col{col} = {value!r}")
    inp = breakpoint_input(page, point, col)
    inp.scroll_into_view_if_needed()
    inp.click()
    inp.press("Control+a")
    inp.press("Delete")
    inp.type(str(value), delay=30)


def breakpoint_disabled(page: Page, point: int, col: int) -> bool:
    """断点 Point(1~4) 第 col 列输入框是否 disabled（灰态不可编辑）。"""
    return breakpoint_input(page, point, col).get_attribute("disabled") is not None


def dialog_confirm(page: Page) -> None:
    step("[Dialog] Confirm")
    _dialog(page).get_by_text("Confirm", exact=True).click()
    page.wait_for_timeout(500)


def dialog_cancel(page: Page) -> None:
    step("[Dialog] Cancel")
    _dialog(page).get_by_text("Cancel", exact=True).click()
    page.wait_for_timeout(400)


# ── Copy / Apply to Selected / Reset Selected ────────────────────────

def click_copy(page: Page, n: int) -> None:
    step(f"Click Copy on AI{n}")
    _ai_row(page, n).get_by_text("Copy", exact=True).click()
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
    step(f"Select checkbox AI{n}")
    row = _ai_row(page, n)
    cb = row.locator(".el-checkbox__inner").first
    cb.scroll_into_view_if_needed()
    cb.click()
    page.wait_for_timeout(300)


def copied_message_text(page: Page) -> str:
    """顶部 "<ID> Settings Copied!" 提示文案（点 Copy 后出现）。"""
    msg = page.locator("text=Settings Copied")
    return msg.first.inner_text().strip() if msg.count() > 0 else ""


# 让 from _src_io_ai import * 包含工具名（须放文件最末）
__all__ = [_n for _n in globals() if not _n.startswith('__')]
