# -*- coding: utf-8 -*-
"""
test_bacnet_ui_config.py — BACnet/IP 配置页面功能验证用例（P1/P2）

用例覆盖：
  TestCase_AcuHMI-1-7_033_001_001: 页面入口与默认状态
  TestCase_AcuHMI-1-7_033_001_004: Device Object Name 配置保存展示正确
  TestCase_AcuHMI-1-7_033_001_005: Device Instance 唯一编号配置正常
  TestCase_AcuHMI-1-7_033_001_006: Advertised APDU Timeout 合法边界值配置正常
  TestCase_AcuHMI-1-7_033_001_007: Advertised APDU Retries 合法边界值配置正常
  TestCase_AcuHMI-1-7_033_001_008: Enable Foreign Device 开关联动
  TestCase_AcuHMI-1-7_033_001_009: Time To Live 合法边界值保存正常
  TestCase_AcuHMI-1-7_033_001_021: EPICS 与 COV 联动规则
  TestCase_AcuHMI-1-7_033_001_022: COV Increment 合法范围值保存
  TestCase_AcuHMI-1-7_033_001_033: COV Batch Update 配置 AcuRev-4100 参数保存正常（Select All + Polling Enable 前置）
  TestCase_AcuHMI-1-7_033_001_036: COV Batch Update 支持覆盖已有配置且未修改配置保持不变
  TestCase_AcuHMI-1-7_033_001_029: COV Increment 为空时切换 Parameter Type 被阻止
  TestCase_AcuHMI-1-7_033_001_030: COV Increment 为空时切换分页被阻止
  TestCase_AcuHMI-1-7_033_001_040: BACnet Port 非法值被拦截
  TestCase_AcuHMI-1-7_033_001_041: Network Number 非法值被拦截
  TestCase_AcuHMI-1-7_033_001_045: COV Increment 非法值被拦截
  TestCase_AcuHMI-1-7_033_001_042: Advertised APDU Timeout 下拉选项集合校验（El Plus v2 Select）
  TestCase_AcuHMI-1-7_033_001_043: Advertised APDU Retries 下拉选项集合校验（El Plus v2 Select）
  TestCase_AcuHMI-1-7_033_001_044: Time To Live 非法边界外值被拦截
  TestCase_AcuHMI-1-7_033_001_011: 非法 BBMD 配置保存被阻止
  TestCase_AcuHMI-1-7_033_001_032: COV Batch Update 配置 AcuRev2100 参数保存正常

运行：
  pytest projects/PX_EMD_G/tests/BacnetIP/test_bacnet_ui_config.py -v

页面结构说明（来自 v9_bacnet_full.html / v10_row0_dialog.html 快照）：
  - BACnet Enable / Foreign Device Function: el-radio-group，选项值 "true"/"false"
  - BACnet Port: placeholder="Enter BACnet Port"，范围 47808-49000，input[type=text]
  - Network Number: placeholder="Enter Network Number"，范围 1-65534，input[type=text]
  - Advertised APDU Timeout: El Plus v2 Select（非自由输入），选项=3/6/10/20/30/45/60 seconds
  - Advertised APDU Retries: El Plus v2 Select（非自由输入），选项=0/1/2/3/5/10
  - Foreign Device BBMD 字段: v-if 条件渲染，仅 Foreign Device = Enable 时出现
  - Parameter Config 弹窗: aria-label="Parameter Config"
    - Parameter Type 下拉: El Plus v2 select（.el-select__input）
    - COV Batch Update 按钮: button:has-text("COV Batch Update")
    - 参数表格列: Parameter | Polling Enable | COV Enable | COV Increment
    - COV Enable: 每行第 3 列（index 2）的 el-switch，aria-checked
    - COV Increment: 每行第 4 列（index 3）的 el-input，COV Enable=off 时 disabled
  - 验证错误: .el-form-item__error 或 .el-message / .el-notification
  - Save 按钮: button.el-button:has-text("Save")（主配置区固定按钮）
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional

import pytest
from playwright.sync_api import Page

# ── 路径 ─────────────────────────────────────────────────────────────────────
_REPO_ROOT = str(Path(__file__).parents[4])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

# ── 从 test_bacnet_ui_basic 复用辅助函数 ─────────────────────────────────────
from projects.PX_EMD_G.tests.BacnetIP.test_bacnet_ui_basic import (  # noqa: E402
    _get_all_device_rows,
    _ensure_device_mapped,
    _resolve_device_by_keyword,
    _set_device_checked,
    _open_param_dialog,
    _close_batch_update_dialog,
    _close_param_config_dialog,
)

# ── 常量 ─────────────────────────────────────────────────────────────────────
_DEFAULT_PORT = "47808"
_PORT_RANGE_HINT = "Range: 47808 - 49000"
_DEFAULT_NETWORK_NUMBER = "1"
_SAVE_WAIT_MS = 2000


# ═════════════════════════════════════════════════════════════════════════════
# 本文件专用辅助函数
# ═════════════════════════════════════════════════════════════════════════════

def _get_foreign_device_state(page: Page) -> bool:
    """读取 Foreign Device Function 当前是否为 Enable（True）。"""
    return page.evaluate(
        """() => {
            const radios = document.querySelectorAll(
                '.el-radio__original[value="true"]'
            );
            for (const r of radios) {
                // 找 Foreign Device Function 所在的 radio-group
                const group = r.closest('.el-radio-group');
                if (!group) continue;
                const formItem = group.closest('.el-form-item');
                const label = formItem && formItem.querySelector('.el-form-item__label');
                if (label && label.textContent.trim() === 'Foreign Device Function') {
                    return r.checked;
                }
            }
            return false;
        }"""
    )


def _set_foreign_device_enable(page: Page, enable: bool) -> None:
    """设置 Foreign Device Function 为 Enable 或 Disable。"""
    value = "true" if enable else "false"
    page.evaluate(
        """(val) => {
            const radios = document.querySelectorAll('.el-radio__original');
            for (const r of radios) {
                if (r.getAttribute('value') !== val) continue;
                const group = r.closest('.el-radio-group');
                if (!group) continue;
                const formItem = group.closest('.el-form-item');
                const label = formItem && formItem.querySelector('.el-form-item__label');
                if (label && label.textContent.trim() === 'Foreign Device Function') {
                    r.click();
                    return;
                }
            }
        }""",
        value,
    )
    page.wait_for_timeout(800)


def _get_field_value(page: Page, placeholder: str) -> str:
    """按 placeholder 读取输入框的当前值。"""
    return page.evaluate(
        """(ph) => {
            const inp = document.querySelector('input[placeholder="' + ph + '"]');
            return inp ? inp.value : '';
        }""",
        placeholder,
    )


def _set_field_value(page: Page, placeholder: str, value: str) -> None:
    """按 placeholder 清空并填写输入框值（模拟用户输入）。"""
    locator = page.locator(f'input[placeholder="{placeholder}"]')
    locator.first.click()
    locator.first.fill(value)


def _click_save(page: Page) -> None:
    """点击主配置区 Save 按钮并等待响应。

    必须跳过不可见按钮：已关闭的 Parameter Config 弹窗仍留在 DOM 中，
    其隐藏 Save 按钮在 DOM 顺序上先于主配置区 Save，误点会导致配置未保存。
    """
    page.evaluate(
        """() => {
            for (const btn of document.querySelectorAll('button')) {
                if (btn.textContent.trim() === 'Save' && !btn.disabled
                        && btn.offsetParent !== null) {
                    btn.click();
                    return;
                }
            }
        }"""
    )
    page.wait_for_timeout(_SAVE_WAIT_MS)


def _has_validation_error(page: Page) -> bool:
    """检查页面是否出现表单验证错误提示。

    覆盖四种错误呈现形式：
    1. El Plus 标准表单验证：.el-form-item__error
    2. Parameter Config 弹窗 COV Increment 自定义错误：.cov-error-tip（span 元素，
       紧随 .cov-input-error wrapper 后，内容如 "Invalid: must be >= 0, max 3 decimal places"）
    3. El Plus Message 全局提示：.el-message--error
    4. El Plus Notification：.el-notification（含 error/warning class）

    注意：.cov-error-tip 在 el-overlay 内，其 offsetParent 可能因 Teleport 返回 null，
    因此额外用 getBoundingClientRect().width > 0 判断实际可见性。
    """
    return page.evaluate(
        """() => {
            // 方式1: el-form-item__error（El Plus 标准表单验证）
            const formErr = document.querySelector('.el-form-item__error');
            if (formErr && formErr.offsetParent !== null) return true;

            // 方式2: COV Increment 自定义错误（.cov-error-tip）
            // 该元素在 el-overlay/Teleport 内，offsetParent 可能为 null，
            // 用 getBoundingClientRect 判断是否真实渲染在视口中
            for (const tip of document.querySelectorAll('.cov-error-tip')) {
                const rect = tip.getBoundingClientRect();
                if (rect.width > 0 || rect.height > 0) return true;
                // 备用：文字非空即视为可见（部分浏览器 rect 为 0 但文字存在）
                if (tip.textContent.trim().length > 0 && tip.offsetParent !== null) return true;
            }
            // 同时检查 .cov-input-error（输入框 wrapper 红框标记）
            for (const el of document.querySelectorAll('.cov-input-error')) {
                const rect = el.getBoundingClientRect();
                if (rect.width > 0 || rect.height > 0) return true;
            }

            // 方式3: el-message--error
            const msgErr = document.querySelector('.el-message--error');
            if (msgErr && msgErr.offsetParent !== null) return true;

            // 方式4: el-notification（含 error/warning class）
            const notif = document.querySelector('.el-notification');
            if (notif && notif.offsetParent !== null) {
                const cls = notif.className;
                if (cls.includes('error') || cls.includes('warning')) return true;
            }
            return false;
        }"""
    )


def _dismiss_toast(page: Page) -> None:
    """关闭可能出现的 toast/message 提示，避免覆盖后续操作。"""
    page.evaluate(
        """() => {
            const close = document.querySelector(
                '.el-message__closeBtn, .el-notification__closeBtn'
            );
            if (close) close.click();
        }"""
    )
    page.wait_for_timeout(300)


def _get_cov_increment_value_in_dialog(page: Page, param_name: str) -> Optional[str]:
    """
    在已打开的 Parameter Config 弹窗中，找到指定参数名对应行的 COV Increment 值。
    找不到则返回 None。
    """
    return page.evaluate(
        """(name) => {
            const rows = document.querySelectorAll(
                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
            );
            for (const row of rows) {
                const cells = row.querySelectorAll('td');
                if (cells.length < 4) continue;
                const paramName = cells[0].textContent.trim();
                if (paramName === name) {
                    const inp = cells[3].querySelector('input.el-input__inner');
                    return inp ? inp.value : null;
                }
            }
            return null;
        }""",
        param_name,
    )


def _set_cov_increment_first_row_in_dialog(page: Page, value: str) -> bool:
    """
    在已打开的 Parameter Config 弹窗中，设置第一行的 COV Increment 值。
    - 若 COV Enable 处于关闭状态，先打开它。
    - 返回是否成功找到并设置。
    """
    return page.evaluate(
        """(val) => {
            const rows = document.querySelectorAll(
                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
            );
            if (rows.length === 0) return false;
            const row = rows[0];
            const cells = row.querySelectorAll('td');
            if (cells.length < 4) return false;

            // 第 3 列: COV Enable switch
            const covSwitch = cells[2].querySelector('.el-switch__input');
            if (covSwitch && covSwitch.getAttribute('aria-checked') === 'false') {
                covSwitch.click();
            }
            // 第 4 列: COV Increment input
            const inp = cells[3].querySelector('input.el-input__inner');
            if (!inp) return false;
            inp.removeAttribute('disabled');
            inp.focus();
            inp.value = val;
            inp.dispatchEvent(new Event('input', {bubbles: true}));
            inp.dispatchEvent(new Event('change', {bubbles: true}));
            return true;
        }""",
        value,
    )


def _open_cov_batch_update(page: Page) -> bool:
    """
    在已打开的 Parameter Config 弹窗中点击 COV Batch Update 按钮。
    返回 Batch Update 弹窗是否出现。
    """
    clicked: bool = page.evaluate(
        """() => {
            for (const btn of document.querySelectorAll('button')) {
                if (btn.textContent.trim().includes('COV Batch')) {
                    btn.click();
                    return true;
                }
            }
            return false;
        }"""
    )
    if not clicked:
        return False
    page.wait_for_timeout(1500)
    return page.evaluate(
        "() => !!document.querySelector('[aria-label=\"Batch Update\"]')"
    )


def _set_batch_update_cov_increment(page: Page, value: str) -> None:
    """在 Batch Update 弹窗中设置 COV Increment 输入框的值。

    Batch Update 弹窗内有两个 input：
      1. .el-select__input（el-select 内部 input，placeholder 为空，用于参数多选）
      2. input.el-input__inner（COV Increment 文本输入框，placeholder="e.g. 0.001"）

    必须精确定位 .el-input__inner，否则会误写入 .el-select__input，
    导致 COV Increment 字段为空，触发 "COV Increment is required" 校验错误。
    """
    page.evaluate(
        """(val) => {
            const dlg = document.querySelector('[aria-label="Batch Update"]');
            if (!dlg) return;
            // 精确定位 COV Increment 输入框：class=el-input__inner（非 el-select__input）
            const inp = dlg.querySelector('input.el-input__inner');
            if (inp && !inp.disabled && !inp.readOnly) {
                inp.focus();
                inp.value = val;
                inp.dispatchEvent(new Event('input', {bubbles: true}));
                inp.dispatchEvent(new Event('change', {bubbles: true}));
                inp.dispatchEvent(new Event('blur', {bubbles: true}));
                return;
            }
            // 回退：找 placeholder 含 '0.001' 或 'cov' 或 'increment' 的 input
            for (const candidate of dlg.querySelectorAll('input')) {
                const ph = (candidate.getAttribute('placeholder') || '').toLowerCase();
                if (ph.includes('0.001') || ph.includes('cov') || ph.includes('increment')) {
                    candidate.focus();
                    candidate.value = val;
                    candidate.dispatchEvent(new Event('input', {bubbles: true}));
                    candidate.dispatchEvent(new Event('change', {bubbles: true}));
                    candidate.dispatchEvent(new Event('blur', {bubbles: true}));
                    return;
                }
            }
        }""",
        value,
    )
    page.wait_for_timeout(300)


def _playwright_batch_select_all(page: Page) -> bool:
    """
    Playwright 原生：在 Batch Update 弹窗中选中所有参数。

    JS click 不触发 El Plus v2 multiselect 的 Vue 3 选择事件，必须用 locator.click()。
    查找顺序：
    1. Select All / 全选 按钮
    2. el-checkbox 全选勾选框（顶部全选 header）
    3. 展开下拉后逐一 locator.click() 每个选项
    返回是否至少选中一项。
    """
    batch = page.locator('[aria-label="Batch Update"]')

    # 1. Select All 按钮
    for txt in ("Select All", "全选"):
        btn = batch.locator("button").filter(has_text=txt)
        if btn.count() > 0:
            btn.first.click()
            page.wait_for_timeout(500)
            return True

    # 2. el-checkbox 全选
    chk = batch.locator('.el-checkbox__input[aria-label="select all"], .check-all .el-checkbox__input')
    if chk.count() > 0:
        chk.first.click()
        page.wait_for_timeout(500)
        return True

    # 3. 展开 multiselect 下拉，用 Playwright locator.click() 逐一选项
    select_input = batch.locator(".el-select__input")
    if select_input.count() == 0:
        return False
    select_input.first.click()
    page.wait_for_timeout(600)

    # 等待展开
    for _ in range(10):
        expanded: bool = page.evaluate(
            """() => {
                const dlg = document.querySelector('[aria-label="Batch Update"]');
                const inp = dlg && dlg.querySelector('.el-select__input');
                return inp ? inp.getAttribute('aria-expanded') === 'true' : false;
            }"""
        )
        if expanded:
            break
        page.wait_for_timeout(300)

    # 读取 dropdown list ID，再用 Playwright locator 点击每一项
    list_id: str = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Batch Update"]');
            const inp = dlg && dlg.querySelector('.el-select__input');
            return (inp && inp.getAttribute('aria-controls')) || '';
        }"""
    )
    if not list_id:
        return False

    items = page.locator(f"#{list_id} .el-select-dropdown__item:not(.is-disabled)")
    n = items.count()
    for i in range(n):
        items.nth(i).click()
        page.wait_for_timeout(150)
    return n > 0


def _playwright_batch_fill_cov(page: Page, value: str) -> None:
    """
    Playwright 原生：在 Batch Update 弹窗中填写 COV Increment。

    JS inp.value + dispatchEvent 不更新 El Plus v2 el-input-number 的 Vue reactive state，
    必须用 locator.fill() 触发完整输入事件链。
    """
    batch = page.locator('[aria-label="Batch Update"]')
    # 优先找 type=number 输入框，其次找非 disabled 的 el-input__inner
    inp = batch.locator('input[type="number"]')
    if inp.count() == 0:
        inp = batch.locator('input.el-input__inner:not([disabled]):not([readonly])')
    inp.last.fill(value)
    page.keyboard.press("Tab")
    page.wait_for_timeout(300)


def _click_batch_update_apply(page: Page) -> bool:
    """在 Batch Update 弹窗中点击 Apply / Save / OK 按钮。"""
    return page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Batch Update"]');
            if (!dlg) return false;
            for (const btn of dlg.querySelectorAll('button')) {
                const t = btn.textContent.trim();
                if (['Apply', 'Save', 'OK', '确定', 'Confirm'].includes(t)) {
                    btn.click();
                    return true;
                }
            }
            return false;
        }"""
    )


def _confirm_batch_update_playwright(page: Page) -> None:
    """
    用 Playwright 原生 click 确认 Batch Update 弹窗。

    JS element.click() 在 Element Plus v2 / Vue 3 的 el-button 上不触发完整事件链，
    导致 Vue click handler 不执行、弹窗不关闭、改动不提交。
    必须通过 Playwright locator.click() 分发真实浏览器事件。

    查找顺序：el-button--primary（主按钮样式）→ 文本含 Confirm → 文本含 确定。
    """
    batch = page.locator('[aria-label="Batch Update"]')
    # 优先找主按钮（primary 样式通常是 Confirm/确定）
    btn = batch.locator("button.el-button--primary")
    if btn.count() == 0:
        btn = batch.locator("button").filter(has_text="Confirm")
    if btn.count() == 0:
        btn = batch.locator("button").filter(has_text="确定")
    assert btn.count() > 0, (
        "Batch Update 弹窗中未找到 Confirm/确定 按钮，"
        "请检查按钮文字或 aria-label 是否匹配"
    )
    btn.first.click()


def _get_field_by_label(page: Page, label_text: str) -> str:
    """按 label 文字找到对应 el-form-item 内的 input，读取并返回当前值。"""
    return page.evaluate(
        """(lbl_text) => {
            const labels = document.querySelectorAll('.el-form-item__label');
            for (const lbl of labels) {
                if (lbl.textContent.trim() === lbl_text) {
                    const fi = lbl.closest('.el-form-item');
                    const inp = fi && fi.querySelector('input.el-input__inner');
                    return inp ? inp.value : '';
                }
            }
            return '';
        }""",
        label_text,
    )


def _set_field_by_label(page: Page, label_text: str, value: str) -> None:
    """按 label 文字找到对应 el-form-item 内的 input，清空并填写值。"""
    page.evaluate(
        """([lbl_text, val]) => {
            const labels = document.querySelectorAll('.el-form-item__label');
            for (const lbl of labels) {
                if (lbl.textContent.trim() === lbl_text) {
                    const fi = lbl.closest('.el-form-item');
                    const inp = fi && fi.querySelector('input.el-input__inner');
                    if (inp) {
                        inp.focus();
                        inp.value = val;
                        inp.dispatchEvent(new Event('input', {bubbles: true}));
                        inp.dispatchEvent(new Event('change', {bubbles: true}));
                    }
                    return;
                }
            }
        }""",
        [label_text, value],
    )


def _get_select_options_by_label(page: Page, label_text: str) -> list[str]:
    """
    按 label 文字找到对应 el-form-item 内的 El Plus v2 Select，
    点击展开后读取所有下拉选项文字，最后关闭下拉。
    返回选项文字列表（不含 disabled 选项）。
    """
    # 点击展开
    page.evaluate(
        """(lbl_text) => {
            const labels = document.querySelectorAll('.el-form-item__label');
            for (const lbl of labels) {
                if (lbl.textContent.trim() === lbl_text) {
                    const fi = lbl.closest('.el-form-item');
                    const sel = fi && fi.querySelector('.el-select');
                    if (sel) sel.click();
                    return;
                }
            }
        }""",
        label_text,
    )
    # 等待展开（最多 10 × 300ms）
    for _ in range(10):
        expanded: bool = page.evaluate(
            """(lbl_text) => {
                const labels = document.querySelectorAll('.el-form-item__label');
                for (const lbl of labels) {
                    if (lbl.textContent.trim() === lbl_text) {
                        const fi = lbl.closest('.el-form-item');
                        const inp = fi && fi.querySelector('.el-select__input');
                        return inp ? inp.getAttribute('aria-expanded') === 'true' : false;
                    }
                }
                return false;
            }""",
            label_text,
        )
        if expanded:
            break
        page.wait_for_timeout(300)

    options: list[str] = page.evaluate(
        """(lbl_text) => {
            const labels = document.querySelectorAll('.el-form-item__label');
            for (const lbl of labels) {
                if (lbl.textContent.trim() === lbl_text) {
                    const fi = lbl.closest('.el-form-item');
                    const inp = fi && fi.querySelector('.el-select__input');
                    const listId = inp && inp.getAttribute('aria-controls');
                    if (!listId) return [];
                    const list = document.getElementById(listId);
                    if (!list) return [];
                    return Array.from(
                        list.querySelectorAll('.el-select-dropdown__item:not(.is-disabled)')
                    ).map(el => el.textContent.trim()).filter(Boolean);
                }
            }
            return [];
        }""",
        label_text,
    )
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)
    return options


def _set_select_option_by_label(page: Page, label_text: str, option_text: str) -> bool:
    """
    按 label 文字找到 El Plus v2 Select，点击展开后选择指定文字的选项。
    返回是否成功选择。
    """
    page.evaluate(
        """(lbl_text) => {
            const labels = document.querySelectorAll('.el-form-item__label');
            for (const lbl of labels) {
                if (lbl.textContent.trim() === lbl_text) {
                    const fi = lbl.closest('.el-form-item');
                    const sel = fi && fi.querySelector('.el-select');
                    if (sel) sel.click();
                    return;
                }
            }
        }""",
        label_text,
    )
    for _ in range(10):
        expanded: bool = page.evaluate(
            """(lbl_text) => {
                const labels = document.querySelectorAll('.el-form-item__label');
                for (const lbl of labels) {
                    if (lbl.textContent.trim() === lbl_text) {
                        const fi = lbl.closest('.el-form-item');
                        const inp = fi && fi.querySelector('.el-select__input');
                        return inp ? inp.getAttribute('aria-expanded') === 'true' : false;
                    }
                }
                return false;
            }""",
            label_text,
        )
        if expanded:
            break
        page.wait_for_timeout(300)

    selected: bool = page.evaluate(
        """([lbl_text, opt]) => {
            const labels = document.querySelectorAll('.el-form-item__label');
            for (const lbl of labels) {
                if (lbl.textContent.trim() === lbl_text) {
                    const fi = lbl.closest('.el-form-item');
                    const inp = fi && fi.querySelector('.el-select__input');
                    const listId = inp && inp.getAttribute('aria-controls');
                    if (!listId) return false;
                    const list = document.getElementById(listId);
                    if (!list) return false;
                    for (const item of list.querySelectorAll(
                        '.el-select-dropdown__item:not(.is-disabled)'
                    )) {
                        if (item.textContent.trim() === opt) {
                            item.click();
                            return true;
                        }
                    }
                    return false;
                }
            }
            return false;
        }""",
        [label_text, option_text],
    )
    page.wait_for_timeout(300)
    return selected


def _enable_polling_for_all_visible_rows(page: Page) -> int:
    """
    在已打开的 Parameter Config 弹窗中，对所有可见行使能 Polling Enable（cells[1]）。
    返回成功使能的行数。
    """
    return page.evaluate(
        """() => {
            const rows = document.querySelectorAll(
                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
            );
            let count = 0;
            for (const row of rows) {
                const cells = row.querySelectorAll('td');
                if (cells.length < 2) continue;
                const sw = cells[1].querySelector('.el-switch__input');
                if (sw && sw.getAttribute('aria-checked') !== 'true') {
                    sw.click();
                }
                count++;
            }
            return count;
        }"""
    )


def _batch_update_select_all_params(page: Page) -> bool:
    """
    在已打开的 Batch Update 弹窗中，尝试通过以下方式选中所有参数：
    1. 优先查找 "Select All" / "全选" 按钮或勾选框并点击
    2. 若不存在则展开下拉，逐一点击所有可用选项
    返回是否至少选中了一个参数。
    """
    # 尝试找 Select All 按钮/勾选框
    found_select_all: bool = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Batch Update"]');
            if (!dlg) return false;
            // 按钮形式
            for (const btn of dlg.querySelectorAll('button')) {
                const t = btn.textContent.trim().toLowerCase();
                if (t === 'select all' || t === '全选') {
                    btn.click();
                    return true;
                }
            }
            // 勾选框形式（label 含 Select All）
            for (const lbl of dlg.querySelectorAll('label, span')) {
                const t = lbl.textContent.trim().toLowerCase();
                if (t === 'select all' || t === '全选') {
                    lbl.click();
                    return true;
                }
            }
            // el-checkbox 全选
            const checkAll = dlg.querySelector('.el-checkbox__input[aria-label="select all"], .el-checkbox-group .check-all');
            if (checkAll) { checkAll.click(); return true; }
            return false;
        }"""
    )
    if found_select_all:
        page.wait_for_timeout(500)
        return True

    # 回退：展开下拉，逐一选择所有选项
    page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Batch Update"]');
            if (!dlg) return;
            const input = dlg.querySelector('.el-select__input');
            if (input) input.click();
        }"""
    )
    for _ in range(10):
        expanded: bool = page.evaluate(
            """() => {
                const dlg = document.querySelector('[aria-label="Batch Update"]');
                const inp = dlg && dlg.querySelector('.el-select__input');
                return inp ? inp.getAttribute('aria-expanded') === 'true' : false;
            }"""
        )
        if expanded:
            break
        page.wait_for_timeout(300)

    selected_count: int = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Batch Update"]');
            if (!dlg) return 0;
            const input = dlg.querySelector('.el-select__input');
            const listId = input && input.getAttribute('aria-controls');
            if (!listId) return 0;
            const list = document.getElementById(listId);
            if (!list) return 0;
            let cnt = 0;
            for (const item of list.querySelectorAll(
                '.el-select-dropdown__item:not(.is-disabled)'
            )) {
                item.click();
                cnt++;
            }
            return cnt;
        }"""
    )
    page.wait_for_timeout(300)
    return selected_count > 0


def _dismiss_blocking_overlays(page: Page) -> None:
    """
    关闭所有可能阻挡导航点击的 overlay 弹窗。

    已知会出现的拦截弹窗：
    - aria-label="Warning"（el-overlay-message-box）：出现在 Parameter Config 关闭后
      或导航时，内容为 "Are you sure want to log out..." 之类
    - aria-label="Batch Update"：Batch Update 弹窗未正常关闭时残留

    处理策略：
    1. Batch Update — 点击 Cancel
    2. el-overlay-message-box / Warning — 点击 Cancel 按钮（不执行操作）
    3. el-message / el-notification —关闭 toast
    """
    # 关闭 Batch Update（若残留）
    page.evaluate(
        """() => {
            const batch = document.querySelector('[aria-label="Batch Update"]');
            if (!batch) return;
            const cancel = Array.from(batch.querySelectorAll('button'))
                .find(b => ['Cancel', '取消', 'Close', '关闭'].includes(b.textContent.trim()));
            if (cancel) cancel.click();
            else { const hb = batch.querySelector('.el-dialog__headerbtn'); if (hb) hb.click(); }
        }"""
    )
    page.wait_for_timeout(300)

    # 关闭所有 message-box overlay（Warning / Confirm 等）— 点 Cancel（不提交）
    page.evaluate(
        """() => {
            const boxes = document.querySelectorAll('.el-overlay-message-box');
            for (const box of boxes) {
                // 优先 Cancel，避免触发破坏性操作（如 logout）
                const cancel = Array.from(box.querySelectorAll('button'))
                    .find(b => ['Cancel', '取消', '否', 'No'].includes(b.textContent.trim()));
                if (cancel) { cancel.click(); continue; }
                // 找关闭图标
                const close = box.querySelector('.el-message-box__headerbtn');
                if (close) close.click();
            }
        }"""
    )
    page.wait_for_timeout(500)

    # 关闭 toast
    page.evaluate(
        """() => {
            const close = document.querySelector(
                '.el-message__closeBtn, .el-notification__closeBtn'
            );
            if (close) close.click();
        }"""
    )
    page.wait_for_timeout(200)


def _navigate_to_bacnet(page: Page) -> None:
    """重新导航到 BACnet/IP 配置页面（用于 Save 后刷新验证）。

    使用 JS click 绕过可能残留的 overlay 拦截，并在导航前自动处理阻挡弹窗。
    """
    # 先清理所有阻挡 overlay
    _dismiss_blocking_overlays(page)

    # 用 JS click 绕过 overlay 拦截（Playwright click 在 overlay 存在时会超时）
    page.evaluate(
        """() => {
            // 找产品管理菜单项（PX-EMD-G）：只遍历 .nav-item 并精确匹配，
            // 避开左上角同名但不可导航的 header-title。
            for (const el of document.querySelectorAll('.nav-item')) {
                if (el.textContent.trim() === 'PX-EMD-G') { el.click(); return; }
            }
        }"""
    )
    page.wait_for_timeout(1500)

    # Protocols 左侧导航
    page.evaluate(
        """() => {
            for (const el of document.querySelectorAll('.left-nav-item')) {
                if (el.textContent.trim() === 'Protocols') { el.click(); return; }
            }
        }"""
    )
    page.wait_for_timeout(1000)
    try:
        page.wait_for_load_state("domcontentloaded", timeout=8000)
    except Exception:
        pass

    # BACnet/IP 子菜单
    page.evaluate(
        """() => {
            for (const el of document.querySelectorAll('li')) {
                if (el.textContent.trim() === 'BACnet/IP') { el.click(); return; }
            }
        }"""
    )
    page.wait_for_timeout(2000)
    try:
        page.wait_for_load_state("domcontentloaded", timeout=8000)
    except Exception:
        pass


# ═════════════════════════════════════════════════════════════════════════════
# 测试用例
# ═════════════════════════════════════════════════════════════════════════════

class TestBACnetUIConfig:
    """BACnet/IP 配置页面功能验证。"""

    # ── LV1 用例 ─────────────────────────────────────────────────────────────

    def test_001_page_load_and_default_state(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_001: 页面入口与默认状态

        断言：
        - BACnet/IP 页面关键元素存在（Port 输入框、Network Number 输入框、Save 按钮）
        - BACnet Enable 字段的 radio-group 存在
        - Port 和 Network Number 字段有值（不为空）
        """
        # 确认关键输入框存在
        port_input = hmi_page.locator('input[placeholder="Enter BACnet Port"]')
        assert port_input.count() > 0, "BACnet Port 输入框不存在，页面可能未正确加载"

        nn_input = hmi_page.locator('input[placeholder="Enter Network Number"]')
        assert nn_input.count() > 0, "Network Number 输入框不存在，页面可能未正确加载"

        # 确认 Save 按钮存在
        save_btn_count: int = hmi_page.evaluate(
            """() => Array.from(document.querySelectorAll('button'))
                    .filter(b => b.textContent.trim() === 'Save').length"""
        )
        assert save_btn_count > 0, "Save 按钮不存在"

        # 确认 BACnet Enable radio-group 存在
        has_enable_group: bool = hmi_page.evaluate(
            """() => {
                const radios = document.querySelectorAll('.el-radio__original');
                for (const r of radios) {
                    const formItem = r.closest('.el-form-item');
                    const label = formItem && formItem.querySelector('.el-form-item__label');
                    if (label && label.textContent.trim() === 'BACnet Enable') return true;
                }
                return false;
            }"""
        )
        assert has_enable_group, "BACnet Enable radio-group 不存在"

        # 记录当前 Port 和 Network Number 值（确保字段有值）
        port_val = _get_field_value(hmi_page, "Enter BACnet Port")
        nn_val = _get_field_value(hmi_page, "Enter Network Number")
        assert port_val, f"BACnet Port 字段值为空，预期有默认值（如 {_DEFAULT_PORT}）"
        assert nn_val, "Network Number 字段值为空，预期有默认值"

    def test_004_device_object_name_save(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_004: Device Object Name 配置保存展示正确

        读取当前值，修改为测试值，Save 后重新导航，验证值持久化，最后恢复原值。
        字段未找到时自动 skip（兼容页面结构差异）。
        """
        original = _get_field_by_label(hmi_page, "Device Object Name")
        if not original:
            pytest.skip("未找到 Device Object Name 字段，页面结构可能与预期不符")

        test_value = "TestHMI-AutoTest"
        try:
            _set_field_by_label(hmi_page, "Device Object Name", test_value)
            _click_save(hmi_page)

            has_error = _has_validation_error(hmi_page)
            _dismiss_toast(hmi_page)
            assert not has_error, f"Device Object Name = {test_value!r} 保存时出现验证错误"

            _navigate_to_bacnet(hmi_page)

            actual = _get_field_by_label(hmi_page, "Device Object Name")
            assert actual == test_value, (
                f"Device Object Name 保存后未持久化。预期 {test_value!r}，实际 {actual!r}"
            )
        finally:
            _set_field_by_label(hmi_page, "Device Object Name", original)
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)

    def test_005_device_instance_save(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_005: Device Instance 唯一编号配置正常

        读取当前值，修改为合法测试值，Save 后重新导航，验证值持久化，最后恢复原值。
        BACnet Device Instance 合法范围：0–4194302。
        """
        original = _get_field_by_label(hmi_page, "Device Instance")
        if not original:
            pytest.skip("未找到 Device Instance 字段，页面结构可能与预期不符")

        test_value = "12345"
        try:
            _set_field_by_label(hmi_page, "Device Instance", test_value)
            _click_save(hmi_page)

            has_error = _has_validation_error(hmi_page)
            _dismiss_toast(hmi_page)
            assert not has_error, f"Device Instance = {test_value!r} 保存时出现验证错误"

            _navigate_to_bacnet(hmi_page)

            actual = _get_field_by_label(hmi_page, "Device Instance")
            assert actual == test_value, (
                f"Device Instance 保存后未持久化。预期 {test_value!r}，实际 {actual!r}"
            )
        finally:
            _set_field_by_label(hmi_page, "Device Instance", original)
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)

    def test_006_apdu_timeout_valid_boundary(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_006: Advertised APDU Timeout 合法边界值配置正常

        Advertised APDU Timeout 是 El Plus v2 Select 下拉（只能选预定义值）。
        选择最小值（"3 seconds"）和最大值（"60 seconds"），Save 后期望无验证错误。
        """
        for test_option in ["3 seconds", "60 seconds"]:
            ok = _set_select_option_by_label(hmi_page, "Advertised APDU Timeout", test_option)
            if not ok:
                pytest.skip(
                    f"Advertised APDU Timeout 下拉中未找到选项 {test_option!r}，"
                    "可能页面结构与预期不符"
                )
            _click_save(hmi_page)

            has_error = _has_validation_error(hmi_page)
            _dismiss_toast(hmi_page)
            assert not has_error, (
                f"Advertised APDU Timeout = {test_option!r} 应为合法值，但出现验证错误"
            )

    def test_007_apdu_retries_valid_boundary(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_007: Advertised APDU Retries 合法边界值配置正常

        Advertised APDU Retries 是 El Plus v2 Select 下拉（只能选预定义值）。
        选择最小值（"0"）和最大值（"10"），Save 后期望无验证错误。
        """
        for test_option in ["0", "10"]:
            ok = _set_select_option_by_label(hmi_page, "Advertised APDU Retries", test_option)
            if not ok:
                pytest.skip(
                    f"Advertised APDU Retries 下拉中未找到选项 {test_option!r}，"
                    "可能页面结构与预期不符"
                )
            _click_save(hmi_page)

            has_error = _has_validation_error(hmi_page)
            _dismiss_toast(hmi_page)
            assert not has_error, (
                f"Advertised APDU Retries = {test_option!r} 应为合法值，但出现验证错误"
            )

    def test_008_foreign_device_toggle(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_008: Enable Foreign Device 开关联动

        断言：
        - 勾选 Foreign Device Function = Enable 后，BBMD IP / BBMD Port / Time To Live 可见
        - 恢复 Disable 后，上述字段不可见（或不存在）
        - 恢复初始状态
        """
        initial_state = _get_foreign_device_state(hmi_page)

        try:
            # 确保从 Disable 状态开始测试
            if initial_state:
                _set_foreign_device_enable(hmi_page, False)
                hmi_page.wait_for_timeout(800)

            # 步骤1: 勾选 Enable
            _set_foreign_device_enable(hmi_page, True)
            hmi_page.wait_for_timeout(1000)

            # 断言 BBMD 相关字段出现（至少一个 BBMD 或 Time To Live 相关 input 可见）
            bbmd_visible: bool = hmi_page.evaluate(
                """() => {
                    // Foreign Device 启用后应出现 BBMD IP / BBMD Port / TTL 输入框
                    const inputs = document.querySelectorAll('input.el-input__inner');
                    for (const inp of inputs) {
                        const ph = (inp.getAttribute('placeholder') || '').toLowerCase();
                        if (ph.includes('bbmd') || ph.includes('time to live') || ph.includes('ttl')) {
                            return inp.offsetParent !== null;  // 可见
                        }
                    }
                    // 也可能通过 label 找
                    const labels = document.querySelectorAll('.el-form-item__label');
                    for (const lbl of labels) {
                        const t = lbl.textContent.trim().toLowerCase();
                        if (t.includes('bbmd') || t.includes('time to live')) {
                            const fi = lbl.closest('.el-form-item');
                            return fi ? fi.offsetParent !== null : false;
                        }
                    }
                    return false;
                }"""
            )
            assert bbmd_visible, (
                "Foreign Device Function = Enable 后，BBMD IP / BBMD Port / Time To Live "
                "字段应出现但未找到可见的相关输入框或标签"
            )

            # 步骤2: 恢复 Disable
            _set_foreign_device_enable(hmi_page, False)
            hmi_page.wait_for_timeout(800)

            bbmd_visible_after_disable: bool = hmi_page.evaluate(
                """() => {
                    const inputs = document.querySelectorAll('input.el-input__inner');
                    for (const inp of inputs) {
                        const ph = (inp.getAttribute('placeholder') || '').toLowerCase();
                        if (ph.includes('bbmd') || ph.includes('time to live') || ph.includes('ttl')) {
                            return inp.offsetParent !== null;
                        }
                    }
                    const labels = document.querySelectorAll('.el-form-item__label');
                    for (const lbl of labels) {
                        const t = lbl.textContent.trim().toLowerCase();
                        if (t.includes('bbmd') || t.includes('time to live')) {
                            const fi = lbl.closest('.el-form-item');
                            return fi ? fi.offsetParent !== null : false;
                        }
                    }
                    return false;
                }"""
            )
            assert not bbmd_visible_after_disable, (
                "Foreign Device Function = Disable 后，BBMD 相关字段应消失，但仍然可见"
            )

        finally:
            # 恢复初始状态
            _set_foreign_device_enable(hmi_page, initial_state)
            hmi_page.wait_for_timeout(500)

    def test_009_time_to_live_valid_boundary(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_009: Time To Live 合法边界值保存正常

        前置条件：启用 Foreign Device（以便显示 TTL 字段）
        测试：输入 TTL=5 保存无错，输入 TTL=1440 保存无错
        清理：恢复 Foreign Device 初始状态
        """
        initial_state = _get_foreign_device_state(hmi_page)

        try:
            # 启用 Foreign Device
            if not initial_state:
                _set_foreign_device_enable(hmi_page, True)
                hmi_page.wait_for_timeout(1000)

            # 检查 TTL 字段是否出现
            ttl_exists: bool = hmi_page.evaluate(
                """() => {
                    const inputs = document.querySelectorAll('input.el-input__inner');
                    for (const inp of inputs) {
                        const ph = (inp.getAttribute('placeholder') || '').toLowerCase();
                        if (ph.includes('time to live') || ph.includes('ttl')) return true;
                    }
                    const labels = document.querySelectorAll('.el-form-item__label');
                    for (const lbl of labels) {
                        if (lbl.textContent.toLowerCase().includes('time to live')) return true;
                    }
                    return false;
                }"""
            )
            if not ttl_exists:
                pytest.skip(
                    "Foreign Device 启用后 Time To Live 字段未出现，"
                    "可能页面结构与预期不同，需真机确认"
                )

            # Foreign Device 启用后 BBMD IP/Port 为必填项，需先填入合法值
            # 否则空字段会触发必填验证，干扰 TTL 合法值的断言
            hmi_page.evaluate(
                """() => {
                    const labels = document.querySelectorAll('.el-form-item__label');
                    for (const lbl of labels) {
                        const t = lbl.textContent.trim();
                        if (t === 'BBMD IP') {
                            const fi = lbl.closest('.el-form-item');
                            const inp = fi && fi.querySelector('input.el-input__inner');
                            if (inp && !inp.value) {
                                inp.focus(); inp.value = '192.168.1.100';
                                inp.dispatchEvent(new Event('input', {bubbles: true}));
                                inp.dispatchEvent(new Event('change', {bubbles: true}));
                            }
                        }
                    }
                }"""
            )

            for ttl_value in ["5", "1440"]:
                # 设置 TTL 值
                hmi_page.evaluate(
                    """(val) => {
                        const inputs = document.querySelectorAll('input.el-input__inner');
                        for (const inp of inputs) {
                            const ph = (inp.getAttribute('placeholder') || '').toLowerCase();
                            if (ph.includes('time to live') || ph.includes('ttl')) {
                                inp.focus();
                                inp.value = val;
                                inp.dispatchEvent(new Event('input', {bubbles: true}));
                                inp.dispatchEvent(new Event('change', {bubbles: true}));
                                return;
                            }
                        }
                        // 通过 label 找对应 input
                        const labels = document.querySelectorAll('.el-form-item__label');
                        for (const lbl of labels) {
                            if (lbl.textContent.toLowerCase().includes('time to live')) {
                                const fi = lbl.closest('.el-form-item');
                                const inp = fi && fi.querySelector('input.el-input__inner');
                                if (inp) {
                                    inp.focus();
                                    inp.value = val;
                                    inp.dispatchEvent(new Event('input', {bubbles: true}));
                                    inp.dispatchEvent(new Event('change', {bubbles: true}));
                                }
                                return;
                            }
                        }
                    }""",
                    ttl_value,
                )
                _click_save(hmi_page)

                has_error = _has_validation_error(hmi_page)
                _dismiss_toast(hmi_page)
                assert not has_error, (
                    f"Time To Live = {ttl_value} 应为合法值，但出现了验证错误提示"
                )

        finally:
            # 恢复 Foreign Device 初始状态
            _set_foreign_device_enable(hmi_page, initial_state)
            hmi_page.wait_for_timeout(500)

    def test_021_epics_and_cov_linkage(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_021: EPICS 与 COV 联动规则

        说明：
        - 快照显示主配置区有 "EPICS File Download" 按钮，无独立 EPICS Enable 开关
        - COV Enable / COV Increment 在 Parameter Config 弹窗的参数表格中（每行独立）
        - 本用例验证：弹窗中 COV Enable = off 时，同行 COV Increment 输入框处于 disabled 状态；
          COV Enable = on 时，COV Increment 可编辑
        - 需要设备列表中有至少一个设备，否则 skip
        """
        rows = _get_all_device_rows(hmi_page)
        if not rows:
            pytest.skip("设备列表为空，无法验证 COV 联动")

        device_keyword = rows[0]["name"]
        was_checked = rows[0]["checked"]
        with _ensure_device_mapped(hmi_page, device_keyword, was_checked):
            if not _open_param_dialog(hmi_page, device_keyword):
                pytest.skip(f"无法打开设备 {device_keyword!r} 的 Parameter Config 弹窗")

            try:
                # 读取第一行的 COV Enable 状态和 COV Increment disabled 状态
                row_state: dict = hmi_page.evaluate(
                    """() => {
                        const rows = document.querySelectorAll(
                            '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                        );
                        if (rows.length === 0) return {found: false};
                        const row = rows[0];
                        const cells = row.querySelectorAll('td');
                        if (cells.length < 4) return {found: false};

                        const covSwitch = cells[2].querySelector('.el-switch__input');
                        const covInpWrapper = cells[3].querySelector('.el-input');
                        const covInp = cells[3].querySelector('input.el-input__inner');

                        return {
                            found: true,
                            paramName: cells[0].textContent.trim(),
                            covEnabled: covSwitch ? covSwitch.getAttribute('aria-checked') === 'true' : null,
                            covIncrDisabled: covInp ? covInp.disabled : null,
                            covIncrWrapperDisabled: covInpWrapper
                                ? covInpWrapper.classList.contains('is-disabled') : null
                        };
                    }"""
                )

                assert row_state.get("found"), (
                    "Parameter Config 弹窗中未找到参数行，无法验证 COV 联动"
                )

                param_name = row_state.get("paramName", "Row 0")

                if row_state.get("covEnabled") is True:
                    # 当前 COV Enable = on → Increment 应可编辑
                    assert not row_state.get("covIncrDisabled"), (
                        f"[{param_name}] COV Enable = ON 时，COV Increment 应可编辑，但当前为 disabled"
                    )
                    # 关闭 COV Enable，验证 Increment 变 disabled
                    hmi_page.evaluate(
                        """() => {
                            const rows = document.querySelectorAll(
                                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                            );
                            if (rows.length === 0) return;
                            const cells = rows[0].querySelectorAll('td');
                            const covSwitch = cells[2] && cells[2].querySelector('.el-switch__input');
                            if (covSwitch) covSwitch.click();
                        }"""
                    )
                    hmi_page.wait_for_timeout(600)
                    disabled_after_off: bool = hmi_page.evaluate(
                        """() => {
                            const rows = document.querySelectorAll(
                                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                            );
                            if (rows.length === 0) return false;
                            const cells = rows[0].querySelectorAll('td');
                            const inp = cells[3] && cells[3].querySelector('input.el-input__inner');
                            const wrapper = cells[3] && cells[3].querySelector('.el-input');
                            return (inp && inp.disabled) || (wrapper && wrapper.classList.contains('is-disabled'));
                        }"""
                    )
                    assert disabled_after_off, (
                        f"[{param_name}] COV Enable 关闭后，COV Increment 应变为 disabled"
                    )
                    # 恢复
                    hmi_page.evaluate(
                        """() => {
                            const rows = document.querySelectorAll(
                                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                            );
                            if (rows.length === 0) return;
                            const cells = rows[0].querySelectorAll('td');
                            const covSwitch = cells[2] && cells[2].querySelector('.el-switch__input');
                            if (covSwitch) covSwitch.click();
                        }"""
                    )
                else:
                    # 当前 COV Enable = off → Increment 应 disabled
                    assert row_state.get("covIncrDisabled") or row_state.get("covIncrWrapperDisabled"), (
                        f"[{param_name}] COV Enable = OFF 时，COV Increment 应为 disabled，"
                        f"但当前 input.disabled={row_state.get('covIncrDisabled')}, "
                        f"wrapper.is-disabled={row_state.get('covIncrWrapperDisabled')}"
                    )
                    # 开启 COV Enable，验证 Increment 可编辑
                    hmi_page.evaluate(
                        """() => {
                            const rows = document.querySelectorAll(
                                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                            );
                            if (rows.length === 0) return;
                            const cells = rows[0].querySelectorAll('td');
                            const covSwitch = cells[2] && cells[2].querySelector('.el-switch__input');
                            if (covSwitch) covSwitch.click();
                        }"""
                    )
                    hmi_page.wait_for_timeout(600)
                    enabled_after_on: bool = hmi_page.evaluate(
                        """() => {
                            const rows = document.querySelectorAll(
                                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                            );
                            if (rows.length === 0) return false;
                            const cells = rows[0].querySelectorAll('td');
                            const inp = cells[3] && cells[3].querySelector('input.el-input__inner');
                            const wrapper = cells[3] && cells[3].querySelector('.el-input');
                            const isDisabled = (inp && inp.disabled)
                                || (wrapper && wrapper.classList.contains('is-disabled'));
                            return !isDisabled;
                        }"""
                    )
                    assert enabled_after_on, (
                        f"[{param_name}] COV Enable 开启后，COV Increment 应变为可编辑"
                    )
                    # 恢复
                    hmi_page.evaluate(
                        """() => {
                            const rows = document.querySelectorAll(
                                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                            );
                            if (rows.length === 0) return;
                            const cells = rows[0].querySelectorAll('td');
                            const covSwitch = cells[2] && cells[2].querySelector('.el-switch__input');
                            if (covSwitch) covSwitch.click();
                        }"""
                    )

            finally:
                _close_param_config_dialog(hmi_page)

    def test_022_cov_increment_valid_range(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_022: COV Increment 合法范围值保存

        在 Parameter Config 弹窗中设置第一行 COV Increment = 0.000 / 0.123，
        点击 Save（弹窗内若有 Save，否则用外部 Save），断言无验证错误。
        需要有设备接入，否则 skip。
        """
        rows = _get_all_device_rows(hmi_page)
        if not rows:
            pytest.skip("设备列表为空，无法验证 COV Increment")

        device_keyword = rows[0]["name"]
        was_checked = rows[0]["checked"]
        with _ensure_device_mapped(hmi_page, device_keyword, was_checked):
            for cov_val in ["0.000", "0.123"]:
                if not _open_param_dialog(hmi_page, device_keyword):
                    pytest.skip(f"无法打开设备 {device_keyword!r} 的 Parameter Config 弹窗")

                try:
                    ok = _set_cov_increment_first_row_in_dialog(hmi_page, cov_val)
                    assert ok, (
                        f"COV Increment = {cov_val}: 未能在弹窗表格第一行找到 COV Increment 输入框"
                    )

                    # 尝试点击弹窗内的 Save（若存在），否则用外部 Save
                    saved_in_dialog: bool = hmi_page.evaluate(
                        """() => {
                            const dlg = document.querySelector('[aria-label="Parameter Config"]');
                            if (!dlg) return false;
                            for (const btn of dlg.querySelectorAll('button')) {
                                const t = btn.textContent.trim();
                                if (t === 'Save' || t === '保存') { btn.click(); return true; }
                            }
                            return false;
                        }"""
                    )
                    if not saved_in_dialog:
                        _close_param_config_dialog(hmi_page)
                        _click_save(hmi_page)
                    else:
                        hmi_page.wait_for_timeout(_SAVE_WAIT_MS)

                    has_error = _has_validation_error(hmi_page)
                    _dismiss_toast(hmi_page)
                    assert not has_error, (
                        f"COV Increment = {cov_val} 应为合法值，但出现了验证错误"
                    )

                finally:
                    # 确保弹窗关闭
                    if hmi_page.locator('[aria-label="Parameter Config"]').count() > 0:
                        _close_param_config_dialog(hmi_page)

    def test_033_cov_batch_update_4100_save(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_033: COV Batch Update 配置 AcuRev-4100 参数保存正常

        流程：
        1. 找到 AcuRev-4100 设备，打开 Parameter Config，为 realtime 组所有可见参数使能
           Polling Enable
        2. 点击 COV Batch Update，在 Parameters 输入框点击 Select All 选择所有参数，
           输入 COV Increment=2.000，点击 Apply
        3. 关闭弹窗，重新导航，重新打开 Parameter Config
        4. 验证参数 COV Increment 已按批量值更新为 2.000
        设备未接入时 pytest.skip。
        """
        # 在全部设备行中按关键词解析 4100（含未勾选行）；未勾选则先勾选映射，使其
        # Parameter Selection 弹窗可开（本用例随后会保存，勾选状态一并持久化）。
        keyword_4100, was_checked, reason = _resolve_device_by_keyword(hmi_page, "4100")
        if keyword_4100 is None:
            pytest.skip(reason)
        if not was_checked:
            _set_device_checked(hmi_page, keyword_4100, True)
            hmi_page.wait_for_timeout(500)
            # 持久化设备映射：勾选状态需主页面 Save 才能在重新导航后保留，
            # 否则后续 _navigate_to_bacnet 重载页面后该行又变回未勾选、弹窗打不开。
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)

        # 步骤1: 打开 Parameter Config，使能所有可见行 Polling Enable
        if not _open_param_dialog(hmi_page, keyword_4100):
            pytest.skip("无法打开 AcuRev-4100 的 Parameter Config 弹窗")

        enabled_rows = _enable_polling_for_all_visible_rows(hmi_page)
        if enabled_rows == 0:
            _close_param_config_dialog(hmi_page)
            pytest.skip("Parameter Config 弹窗中未找到任何参数行，需真机确认")
        hmi_page.wait_for_timeout(500)

        # first_param 在 Select All 之后从 Batch Update 已选 tag 读取，
        # 避免表格第一行与 Batch Update 实际可选参数不一致
        first_param: str = ""

        try:
            # 步骤2: 打开 COV Batch Update
            batch_opened = _open_cov_batch_update(hmi_page)
            if not batch_opened:
                pytest.skip("COV Batch Update 弹窗未出现，可能按钮未找到")

            # 在 Parameters 输入框点击 Select All 选择所有参数
            selected_any = _playwright_batch_select_all(hmi_page)
            if not selected_any:
                hmi_page.keyboard.press("Escape")
                pytest.skip(
                    "COV Batch Update 下拉选项为空（可能为 lazy load 未触发），需真机确认"
                )
            hmi_page.wait_for_timeout(300)

            # 从 Batch Update 已选 tag 读第一个参数名（El-Plus v2 用 .el-select__tags-text）
            # 过滤折叠计数 tag（如 "+ 2"），取第一个实际参数名
            first_param = hmi_page.evaluate(
                """() => {
                    const dlg = document.querySelector('[aria-label="Batch Update"]');
                    if (!dlg) return '';
                    for (const t of dlg.querySelectorAll('.el-select__tags-text')) {
                        const txt = t.textContent.trim();
                        if (txt && !/^\\+\\s*\\d+$/.test(txt)) return txt;
                    }
                    return '';
                }"""
            )

            # 设置 COV Increment = 2.000
            _playwright_batch_fill_cov(hmi_page, "2.000")

            # Playwright 原生点击 Confirm/确定（JS click 不触发 Vue 3 handler）
            _confirm_batch_update_playwright(hmi_page)
            hmi_page.wait_for_timeout(_SAVE_WAIT_MS)

        finally:
            if hmi_page.evaluate(
                "() => !!document.querySelector('[aria-label=\"Batch Update\"]')"
            ):
                _close_batch_update_dialog(hmi_page)
            # COV Batch Update 的改动只存于 Parameter Config 弹窗的 Vue 状态，
            # 必须在弹窗关闭之前触发保存，否则关闭时状态丢失。
            # 优先找弹窗内 Save 按钮；找不到则在弹窗开启状态下 force=True 点主区域 Save
            # （绕过 el-overlay 遮罩），再关弹窗。
            saved_in_dlg: bool = hmi_page.evaluate(
                """() => {
                    const dlg = document.querySelector('[aria-label="Parameter Config"]');
                    if (!dlg) return false;
                    for (const btn of dlg.querySelectorAll('button')) {
                        if (btn.textContent.trim() === 'Save') { btn.click(); return true; }
                    }
                    return false;
                }"""
            )
            if saved_in_dlg:
                hmi_page.wait_for_timeout(_SAVE_WAIT_MS)
                _dismiss_toast(hmi_page)
            else:
                # 弹窗内无 Save：在弹窗开启状态下强制点击主区域 Save
                hmi_page.locator('button:has-text("Save")').first.click(force=True)
                hmi_page.wait_for_timeout(_SAVE_WAIT_MS)
                _dismiss_toast(hmi_page)
            _close_param_config_dialog(hmi_page)

        # 步骤3: 重新导航
        _navigate_to_bacnet(hmi_page)

        # 步骤4: 重新打开 Parameter Config，验证 COV Increment 已批量更新
        if not _open_param_dialog(hmi_page, keyword_4100):
            pytest.fail("重新导航后无法打开 AcuRev-4100 Parameter Config 弹窗")

        try:
            if not first_param:
                pytest.fail(
                    "未能从 Batch Update 已选参数 tag 中读取第一个参数名，无法验证"
                )
            # 先在当前 Parameter Type 下查找
            actual_val = _get_cov_increment_value_in_dialog(hmi_page, first_param)
            # 若当前视图找不到，逐一切换 Parameter Type 重查
            if actual_val is None:
                param_types = _get_select_options_by_label(hmi_page, "Parameter Type")
                for pt in param_types:
                    _set_select_option_by_label(hmi_page, "Parameter Type", pt)
                    hmi_page.wait_for_timeout(500)
                    actual_val = _get_cov_increment_value_in_dialog(hmi_page, first_param)
                    if actual_val is not None:
                        break
            assert actual_val in ("2.000", "2"), (
                f"Batch Update 保存后，参数 {first_param!r} 的 COV Increment 预期为 '2.000'，"
                f"实际为 {actual_val!r}"
            )
        finally:
            _close_param_config_dialog(hmi_page)

    def test_036_batch_update_partial_coverage(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_036: COV Batch Update 支持覆盖已有配置且未修改配置保持不变

        流程：
        1. 打开 AcuRev-4100 Parameter Config，为参数 A（第一行）使能 Polling Enable，
           设置 COV Increment=1.000 并保存，建立已知原值
        2. 重新打开弹窗，点击 COV Batch Update，仅选参数 B（第二行），
           设置 COV Increment=9.999，Apply
        3. 重新打开弹窗，确认参数 A 的 COV Increment 仍为 1.000，未被覆盖
        设备未接入时 pytest.skip。
        """
        # 在全部设备行中按关键词解析 4100（含未勾选行）；未勾选则先勾选映射，使其
        # Parameter Selection 弹窗可开（本用例随后会保存，勾选状态一并持久化）。
        keyword_4100, was_checked, reason = _resolve_device_by_keyword(hmi_page, "4100")
        if keyword_4100 is None:
            pytest.skip(reason)
        if not was_checked:
            _set_device_checked(hmi_page, keyword_4100, True)
            hmi_page.wait_for_timeout(500)
            # 持久化设备映射：勾选状态需主页面 Save 才能在重新导航后保留，
            # 否则后续 _navigate_to_bacnet 重载页面后该行又变回未勾选、弹窗打不开。
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)

        # ── 步骤1：为参数 A 使能 Polling Enable 并设置已知 COV Increment=1.000 ──
        if not _open_param_dialog(hmi_page, keyword_4100):
            pytest.skip("无法打开 AcuRev-4100 的 Parameter Config 弹窗")

        row_info: dict = hmi_page.evaluate(
            """() => {
                const rows = document.querySelectorAll(
                    '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                );
                if (rows.length < 2) return {enough: false};
                const cells0 = rows[0].querySelectorAll('td');
                const cells1 = rows[1].querySelectorAll('td');
                return {
                    enough: true,
                    paramA: cells0[0] ? cells0[0].textContent.trim() : '',
                    paramB: cells1[0] ? cells1[0].textContent.trim() : '',
                };
            }"""
        )

        if not row_info.get("enough"):
            _close_param_config_dialog(hmi_page)
            pytest.skip("Parameter Config 弹窗中少于 2 行参数，无法测试部分覆盖场景")

        param_a: str = row_info["paramA"]
        param_b: str = row_info["paramB"]
        known_incr_a = "1.000"

        # 为参数 A（第一行）使能 Polling Enable，再设 COV Increment=1.000
        hmi_page.evaluate(
            """() => {
                const rows = document.querySelectorAll(
                    '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                );
                if (!rows.length) return;
                const cells = rows[0].querySelectorAll('td');
                // Polling Enable（cells[1]）
                const pollingSw = cells[1] && cells[1].querySelector('.el-switch__input');
                if (pollingSw && pollingSw.getAttribute('aria-checked') !== 'true') {
                    pollingSw.click();
                }
                // COV Enable（cells[2]）
                const covSw = cells[2] && cells[2].querySelector('.el-switch__input');
                if (covSw && covSw.getAttribute('aria-checked') !== 'true') {
                    covSw.click();
                }
            }"""
        )
        hmi_page.wait_for_timeout(500)

        set_ok: bool = hmi_page.evaluate(
            """(val) => {
                const rows = document.querySelectorAll(
                    '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                );
                if (!rows.length) return false;
                const cells = rows[0].querySelectorAll('td');
                const inp = cells[3] && cells[3].querySelector('input.el-input__inner');
                if (!inp || inp.disabled) return false;
                inp.removeAttribute('disabled');
                inp.focus();
                inp.value = val;
                inp.dispatchEvent(new Event('input', {bubbles: true}));
                inp.dispatchEvent(new Event('change', {bubbles: true}));
                return true;
            }""",
            known_incr_a,
        )
        if not set_ok:
            _close_param_config_dialog(hmi_page)
            pytest.skip(
                "参数 A 的 COV Increment 输入框不可编辑（disabled），需真机确认"
            )

        # 保存：优先用弹窗内 Save，否则关闭后用外部 Save
        saved_in_dlg: bool = hmi_page.evaluate(
            """() => {
                const dlg = document.querySelector('[aria-label="Parameter Config"]');
                if (!dlg) return false;
                for (const btn of dlg.querySelectorAll('button')) {
                    if (btn.textContent.trim() === 'Save') { btn.click(); return true; }
                }
                return false;
            }"""
        )
        if saved_in_dlg:
            hmi_page.wait_for_timeout(_SAVE_WAIT_MS)
            _close_param_config_dialog(hmi_page)
        else:
            _close_param_config_dialog(hmi_page)
            _click_save(hmi_page)
        _dismiss_toast(hmi_page)

        # 重新导航，确认参数 A 已有配置可正常显示
        _navigate_to_bacnet(hmi_page)
        if not _open_param_dialog(hmi_page, keyword_4100):
            pytest.fail("确认前置状态：无法打开 AcuRev-4100 Parameter Config 弹窗")
        pre_check = _get_cov_increment_value_in_dialog(hmi_page, param_a)
        _close_param_config_dialog(hmi_page)
        assert pre_check in (known_incr_a, "1"), (
            f"前置条件未满足：参数 A {param_a!r} 的 COV Increment 预期 {known_incr_a!r}，"
            f"实际 {pre_check!r}"
        )

        # ── 步骤2：仅对参数 B 执行 Batch Update（不选参数 A） ──
        if not _open_param_dialog(hmi_page, keyword_4100):
            pytest.fail("步骤2：无法打开 AcuRev-4100 Parameter Config 弹窗")

        try:
            batch_opened = _open_cov_batch_update(hmi_page)
            if not batch_opened:
                pytest.skip("COV Batch Update 弹窗未出现")

            # 展开下拉，仅选参数 B（用 Playwright 真实点击，JS click 无法触发 El-Plus 下拉）
            select_input_loc = hmi_page.locator(
                '[aria-label="Batch Update"] .el-select__input'
            )
            if select_input_loc.count() > 0:
                select_input_loc.first.click()
            for _ in range(10):
                expanded: bool = hmi_page.evaluate(
                    """() => {
                        const dlg = document.querySelector('[aria-label="Batch Update"]');
                        const inp = dlg && dlg.querySelector('.el-select__input');
                        return inp ? inp.getAttribute('aria-expanded') === 'true' : false;
                    }"""
                )
                if expanded:
                    break
                hmi_page.wait_for_timeout(300)

            selected_b: bool = hmi_page.evaluate(
                """(name) => {
                    const dlg = document.querySelector('[aria-label="Batch Update"]');
                    if (!dlg) return false;
                    const input = dlg.querySelector('.el-select__input');
                    const listId = input && input.getAttribute('aria-controls');
                    if (!listId) return false;
                    const list = document.getElementById(listId);
                    if (!list) return false;
                    for (const item of list.querySelectorAll('.el-select-dropdown__item')) {
                        if (item.textContent.trim() === name) {
                            item.click();
                            return true;
                        }
                    }
                    return false;
                }""",
                param_b,
            )
            if not selected_b:
                hmi_page.keyboard.press("Escape")
                pytest.skip(
                    f"参数 {param_b!r} 在 Batch Update 下拉中未找到，需真机确认"
                )

            hmi_page.wait_for_timeout(300)
            _playwright_batch_fill_cov(hmi_page, "9.999")
            applied = _click_batch_update_apply(hmi_page)
            if not applied:
                hmi_page.keyboard.press("Escape")
                pytest.skip("Batch Update 弹窗中未找到 Apply 按钮")

            hmi_page.wait_for_timeout(_SAVE_WAIT_MS)

        finally:
            if hmi_page.evaluate(
                "() => !!document.querySelector('[aria-label=\"Batch Update\"]')"
            ):
                _close_batch_update_dialog(hmi_page)
            saved_in_dlg: bool = hmi_page.evaluate(
                """() => {
                    const dlg = document.querySelector('[aria-label="Parameter Config"]');
                    if (!dlg) return false;
                    for (const btn of dlg.querySelectorAll('button')) {
                        if (btn.textContent.trim() === 'Save') { btn.click(); return true; }
                    }
                    return false;
                }"""
            )
            if saved_in_dlg:
                hmi_page.wait_for_timeout(_SAVE_WAIT_MS)
                _dismiss_toast(hmi_page)
                _close_param_config_dialog(hmi_page)
            else:
                _close_param_config_dialog(hmi_page)
                _click_save(hmi_page)
                _dismiss_toast(hmi_page)

        # ── 步骤3：重新导航，验证参数 A COV Increment 未被覆盖 ──
        _navigate_to_bacnet(hmi_page)

        if not _open_param_dialog(hmi_page, keyword_4100):
            pytest.fail("验证阶段无法打开 AcuRev-4100 Parameter Config 弹窗")

        try:
            actual_incr_a = _get_cov_increment_value_in_dialog(hmi_page, param_a)
            assert actual_incr_a in (known_incr_a, "1"), (
                f"Batch Update 只对 {param_b!r} 操作，参数 {param_a!r} 的 COV Increment 不应改变。"
                f"预期 {known_incr_a!r}，当前值：{actual_incr_a!r}"
            )
        finally:
            _close_param_config_dialog(hmi_page)

    # ── LV2 用例（边界值验证）────────────────────────────────────────────────

    def test_029_empty_cov_increment_blocks_param_type_switch(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_029: COV Increment 为空时切换 Parameter Type 被阻止

        在 Parameter Config 弹窗中开启 COV Enable，清空 COV Increment，
        尝试切换 Parameter Type 下拉，期望触发验证错误（阻止切换）。
        设备未接入或弹窗内无 Parameter Type 下拉时自动 skip。
        """
        rows = _get_all_device_rows(hmi_page)
        if not rows:
            pytest.skip("设备列表为空，无法测试")

        device_keyword = rows[0]["name"]
        was_checked = rows[0]["checked"]
        with _ensure_device_mapped(hmi_page, device_keyword, was_checked):
            if not _open_param_dialog(hmi_page, device_keyword):
                pytest.skip(f"无法打开设备 {device_keyword!r} 的 Parameter Config 弹窗")

            try:
                # 开启第一行 COV Enable
                hmi_page.evaluate(
                    """() => {
                        const rows = document.querySelectorAll(
                            '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                        );
                        if (!rows.length) return;
                        const cells = rows[0].querySelectorAll('td');
                        const sw = cells[2] && cells[2].querySelector('.el-switch__input');
                        if (sw && sw.getAttribute('aria-checked') !== 'true') sw.click();
                    }"""
                )
                hmi_page.wait_for_timeout(500)

                # 清空 COV Increment
                cleared: bool = hmi_page.evaluate(
                    """() => {
                        const rows = document.querySelectorAll(
                            '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                        );
                        if (!rows.length) return false;
                        const cells = rows[0].querySelectorAll('td');
                        const inp = cells[3] && cells[3].querySelector('input.el-input__inner');
                        if (!inp || inp.disabled) return false;
                        inp.focus();
                        inp.value = '';
                        inp.dispatchEvent(new Event('input', {bubbles: true}));
                        inp.dispatchEvent(new Event('change', {bubbles: true}));
                        inp.dispatchEvent(new Event('blur', {bubbles: true}));
                        return true;
                    }"""
                )
                if not cleared:
                    pytest.skip("COV Increment 输入框未能清空（disabled 或未找到），需真机确认")

                hmi_page.wait_for_timeout(300)

                # 确认弹窗内有 Parameter Type 下拉
                has_param_type: bool = hmi_page.evaluate(
                    """() => {
                        const dlg = document.querySelector('[aria-label="Parameter Config"]');
                        return dlg ? !!dlg.querySelector('.el-select') : false;
                    }"""
                )
                if not has_param_type:
                    pytest.skip("Parameter Config 弹窗内未找到 Parameter Type 下拉")

                # 点击展开 Parameter Type 下拉，尝试选择第一个非当前项
                hmi_page.evaluate(
                    """() => {
                        const dlg = document.querySelector('[aria-label="Parameter Config"]');
                        const sel = dlg && dlg.querySelector('.el-select');
                        if (sel) sel.click();
                    }"""
                )
                hmi_page.wait_for_timeout(800)

                hmi_page.evaluate(
                    """() => {
                        const dlg = document.querySelector('[aria-label="Parameter Config"]');
                        if (!dlg) return;
                        const inp = dlg.querySelector('.el-select__input');
                        const listId = inp && inp.getAttribute('aria-controls');
                        if (!listId) return;
                        const list = document.getElementById(listId);
                        if (!list) return;
                        const items = list.querySelectorAll(
                            '.el-select-dropdown__item:not(.is-disabled)'
                        );
                        if (items.length > 1) items[1].click();
                        else if (items.length === 1) items[0].click();
                    }"""
                )
                hmi_page.keyboard.press("Escape")
                hmi_page.wait_for_timeout(500)

                has_error = _has_validation_error(hmi_page)
                _dismiss_toast(hmi_page)

                assert has_error, (
                    "COV Increment 为空时，切换 Parameter Type 应触发验证错误，"
                    "但未检测到任何错误提示"
                )

            finally:
                if hmi_page.locator('[aria-label="Parameter Config"]').count() > 0:
                    _close_param_config_dialog(hmi_page)

    def test_030_empty_cov_increment_blocks_page_switch(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_030: COV Increment 为空时切换分页被阻止

        在 Parameter Config 弹窗中开启 COV Enable，清空 COV Increment，
        尝试点击分页下一页，期望出现验证错误或页码未发生变化。
        弹窗内无分页组件时自动 skip。
        """
        rows = _get_all_device_rows(hmi_page)
        if not rows:
            pytest.skip("设备列表为空，无法测试")

        device_keyword = rows[0]["name"]
        was_checked = rows[0]["checked"]
        with _ensure_device_mapped(hmi_page, device_keyword, was_checked):
            if not _open_param_dialog(hmi_page, device_keyword):
                pytest.skip(f"无法打开设备 {device_keyword!r} 的 Parameter Config 弹窗")

            try:
                # 确认弹窗内有分页组件
                has_pagination: bool = hmi_page.evaluate(
                    """() => {
                        const dlg = document.querySelector('[aria-label="Parameter Config"]');
                        return dlg ? !!dlg.querySelector('.el-pagination') : false;
                    }"""
                )
                if not has_pagination:
                    pytest.skip(
                        "Parameter Config 弹窗内无分页组件，参数数量在单页内，无法验证分页阻止"
                    )

                # 记录当前页码
                current_page: str = hmi_page.evaluate(
                    """() => {
                        const dlg = document.querySelector('[aria-label="Parameter Config"]');
                        const active = dlg && dlg.querySelector('.el-pager .is-active');
                        return active ? active.textContent.trim() : '1';
                    }"""
                )

                # 开启第一行 COV Enable 并清空 COV Increment
                hmi_page.evaluate(
                    """() => {
                        const rows = document.querySelectorAll(
                            '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                        );
                        if (!rows.length) return;
                        const cells = rows[0].querySelectorAll('td');
                        const sw = cells[2] && cells[2].querySelector('.el-switch__input');
                        if (sw && sw.getAttribute('aria-checked') !== 'true') sw.click();
                    }"""
                )
                hmi_page.wait_for_timeout(500)

                cleared: bool = hmi_page.evaluate(
                    """() => {
                        const rows = document.querySelectorAll(
                            '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                        );
                        if (!rows.length) return false;
                        const cells = rows[0].querySelectorAll('td');
                        const inp = cells[3] && cells[3].querySelector('input.el-input__inner');
                        if (!inp || inp.disabled) return false;
                        inp.focus();
                        inp.value = '';
                        inp.dispatchEvent(new Event('input', {bubbles: true}));
                        inp.dispatchEvent(new Event('change', {bubbles: true}));
                        inp.dispatchEvent(new Event('blur', {bubbles: true}));
                        return true;
                    }"""
                )
                if not cleared:
                    pytest.skip("COV Increment 输入框未能清空（disabled 或未找到）")

                # 点击分页下一页
                clicked_next: bool = hmi_page.evaluate(
                    """() => {
                        const dlg = document.querySelector('[aria-label="Parameter Config"]');
                        const nextBtn = dlg && (
                            dlg.querySelector('.btn-next') ||
                            dlg.querySelector('[aria-label="Go to next page"]')
                        );
                        if (nextBtn && !nextBtn.disabled) {
                            nextBtn.click();
                            return true;
                        }
                        return false;
                    }"""
                )
                hmi_page.wait_for_timeout(800)

                if not clicked_next:
                    pytest.skip("未找到可点击的下一页按钮（可能只有一页）")

                has_error = _has_validation_error(hmi_page)
                _dismiss_toast(hmi_page)

                new_page: str = hmi_page.evaluate(
                    """() => {
                        const dlg = document.querySelector('[aria-label="Parameter Config"]');
                        const active = dlg && dlg.querySelector('.el-pager .is-active');
                        return active ? active.textContent.trim() : '1';
                    }"""
                )

                assert has_error or new_page == current_page, (
                    f"COV Increment 为空时，分页切换应被阻止，"
                    f"但页码从 {current_page!r} 变为 {new_page!r} 且无验证错误"
                )

            finally:
                if hmi_page.locator('[aria-label="Parameter Config"]').count() > 0:
                    _close_param_config_dialog(hmi_page)

    def test_040_bacnet_port_invalid_values(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_040: BACnet Port 非法值被拦截

        Port 合法范围：47808-49000
        测试值：47807（低于下限）、49001（高于上限）
        清理：恢复默认值 47808
        """
        original_port = _get_field_value(hmi_page, "Enter BACnet Port")

        try:
            for invalid_port in ["47807", "49001"]:
                _set_field_value(hmi_page, "Enter BACnet Port", invalid_port)
                _click_save(hmi_page)

                has_error = _has_validation_error(hmi_page)
                _dismiss_toast(hmi_page)

                assert has_error, (
                    f"BACnet Port = {invalid_port} 超出范围 47808-49000，"
                    f"应出现验证错误，但未检测到任何错误提示"
                )

        finally:
            # 恢复原始端口值
            restore_val = original_port if original_port else _DEFAULT_PORT
            _set_field_value(hmi_page, "Enter BACnet Port", restore_val)
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)

    def test_041_network_number_invalid_values(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_041: Network Number 非法值被拦截

        Network Number 合法范围：1-65534
        测试值：0（低于下限）、65535（高于上限）
        """
        original_nn = _get_field_value(hmi_page, "Enter Network Number")

        try:
            for invalid_nn in ["0", "65535"]:
                _set_field_value(hmi_page, "Enter Network Number", invalid_nn)
                _click_save(hmi_page)

                has_error = _has_validation_error(hmi_page)
                _dismiss_toast(hmi_page)

                assert has_error, (
                    f"Network Number = {invalid_nn} 超出范围 1-65534，"
                    f"应出现验证错误，但未检测到任何错误提示"
                )

        finally:
            # 恢复原始值
            restore_val = original_nn if original_nn else _DEFAULT_NETWORK_NUMBER
            _set_field_value(hmi_page, "Enter Network Number", restore_val)
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)

    def test_045_cov_increment_invalid_value(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_045: COV Increment 非法值被拦截

        在 Parameter Config 弹窗中输入 COV Increment = -0.001（负值），
        期望出现验证错误。
        需要设备接入，否则 skip。
        """
        rows = _get_all_device_rows(hmi_page)
        if not rows:
            pytest.skip("设备列表为空，无法验证 COV Increment 非法值")

        device_keyword = rows[0]["name"]
        was_checked = rows[0]["checked"]
        with _ensure_device_mapped(hmi_page, device_keyword, was_checked):
            if not _open_param_dialog(hmi_page, device_keyword):
                pytest.skip(f"无法打开设备 {device_keyword!r} 的 Parameter Config 弹窗")

            try:
                # 先启用 COV Enable（否则 COV Increment 为 disabled，无法输入）
                hmi_page.evaluate(
                    """() => {
                        const rows = document.querySelectorAll(
                            '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                        );
                        if (rows.length === 0) return;
                        const cells = rows[0].querySelectorAll('td');
                        const covSwitch = cells[2] && cells[2].querySelector('.el-switch__input');
                        if (covSwitch && covSwitch.getAttribute('aria-checked') !== 'true') {
                            covSwitch.click();
                        }
                    }"""
                )
                hmi_page.wait_for_timeout(500)

                # 尝试输入 -0.001
                set_ok: bool = hmi_page.evaluate(
                    """() => {
                        const rows = document.querySelectorAll(
                            '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                        );
                        if (rows.length === 0) return false;
                        const cells = rows[0].querySelectorAll('td');
                        const inp = cells[3] && cells[3].querySelector('input.el-input__inner');
                        if (!inp || inp.disabled) return false;
                        inp.focus();
                        inp.value = '-0.001';
                        inp.dispatchEvent(new Event('input', {bubbles: true}));
                        inp.dispatchEvent(new Event('change', {bubbles: true}));
                        inp.dispatchEvent(new Event('blur', {bubbles: true}));
                        return true;
                    }"""
                )

                if not set_ok:
                    pytest.skip(
                        "COV Increment 输入框未能写入（可能仍为 disabled），需真机确认"
                    )

                # 尝试弹窗内 Save（若存在），否则用外部 Save
                saved_in_dialog: bool = hmi_page.evaluate(
                    """() => {
                        const dlg = document.querySelector('[aria-label="Parameter Config"]');
                        if (!dlg) return false;
                        for (const btn of dlg.querySelectorAll('button')) {
                            const t = btn.textContent.trim();
                            if (t === 'Save' || t === '保存') { btn.click(); return true; }
                        }
                        return false;
                    }"""
                )
                if not saved_in_dialog:
                    _close_param_config_dialog(hmi_page)
                    _click_save(hmi_page)
                else:
                    hmi_page.wait_for_timeout(_SAVE_WAIT_MS)

                has_error = _has_validation_error(hmi_page)
                _dismiss_toast(hmi_page)

                assert has_error, (
                    "COV Increment = -0.001 为非法负值，应出现验证错误，但未检测到任何错误提示"
                )

            finally:
                if hmi_page.locator('[aria-label="Parameter Config"]').count() > 0:
                    _close_param_config_dialog(hmi_page)

    def test_042_apdu_timeout_options(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_042: Advertised APDU Timeout 选项集合符合规格

        页面探查确认：Advertised APDU Timeout 是 El Plus v2 Select 下拉框（非自由输入框），
        用户只能选择预定义的离散值，无法输入任意数字。
        合法选项（来自规格）：3、6、10、20、30、45、60（秒）。

        测试策略：验证下拉选项与规格完全一致（既无遗漏合法值，也无规格外的非法值）。
        若出现规格外选项（如 2 秒或 61 秒），说明前端约束有误。

        注意：原测试设计"输入 2 / 61 期望报错"基于字段为文本输入框的假设，
        已通过页面探查证伪——该字段不接受自由输入，本用例以选项集合校验替代。
        """
        # 规格定义的合法 Timeout 选项（文字与页面显示格式一致）
        expected_options = {
            "3 seconds", "6 seconds", "10 seconds",
            "20 seconds", "30 seconds", "45 seconds", "60 seconds",
        }

        actual_options = _get_select_options_by_label(hmi_page, "Advertised APDU Timeout")
        actual_set = set(actual_options)

        missing = expected_options - actual_set
        extra = actual_set - expected_options

        assert not missing and not extra, (
            f"Advertised APDU Timeout 下拉选项与规格不符。\n"
            f"  规格有但页面无：{sorted(missing)}\n"
            f"  页面有但规格无：{sorted(extra)}\n"
            f"  当前页面选项：{sorted(actual_set)}"
        )

    def test_043_apdu_retries_options(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_043: Advertised APDU Retries 选项集合符合规格

        页面探查确认：Advertised APDU Retries 是 El Plus v2 Select 下拉框（非自由输入框），
        用户只能选择预定义的离散值，无法输入 -1 或 11 等任意数字。
        合法选项（来自规格）：0、1、2、3、5、10。

        测试策略：验证下拉选项与规格完全一致（既无遗漏合法值，也无规格外的非法值）。
        若出现规格外选项（如 -1 或 11），说明前端约束有误。

        注意：原测试设计"输入 -1 / 11 期望报错"基于字段为文本输入框的假设，
        已通过页面探查证伪——该字段不接受自由输入，本用例以选项集合校验替代。
        """
        # 规格定义的合法 Retries 选项
        expected_options = {"0", "1", "2", "3", "5", "10"}

        actual_options = _get_select_options_by_label(hmi_page, "Advertised APDU Retries")
        actual_set = set(actual_options)

        missing = expected_options - actual_set
        extra = actual_set - expected_options

        assert not missing and not extra, (
            f"Advertised APDU Retries 下拉选项与规格不符。\n"
            f"  规格有但页面无：{sorted(missing)}\n"
            f"  页面有但规格无：{sorted(extra)}\n"
            f"  当前页面选项：{sorted(actual_set)}"
        )

    def test_044_time_to_live_invalid_values(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_044: Time To Live 非法边界外值被拦截

        合法范围：5-1440（分钟）
        测试值：4（低于下限）、1441（高于上限）
        前置条件：启用 Foreign Device 使 TTL 字段可见
        清理：恢复 Foreign Device 初始状态
        """
        initial_foreign = _get_foreign_device_state(hmi_page)

        try:
            if not initial_foreign:
                _set_foreign_device_enable(hmi_page, True)
                hmi_page.wait_for_timeout(1000)

            ttl_exists: bool = hmi_page.evaluate(
                """() => {
                    const labels = document.querySelectorAll('.el-form-item__label');
                    for (const lbl of labels) {
                        if (lbl.textContent.trim() === 'Time To Live') {
                            const fi = lbl.closest('.el-form-item');
                            return fi ? fi.offsetParent !== null : false;
                        }
                    }
                    return false;
                }"""
            )
            if not ttl_exists:
                pytest.skip("Foreign Device 启用后 Time To Live 字段未出现，需真机确认")

            for invalid_val in ["4", "1441"]:
                hmi_page.evaluate(
                    """(val) => {
                        const labels = document.querySelectorAll('.el-form-item__label');
                        for (const lbl of labels) {
                            if (lbl.textContent.trim() === 'Time To Live') {
                                const fi = lbl.closest('.el-form-item');
                                const inp = fi && fi.querySelector('input.el-input__inner');
                                if (inp) {
                                    inp.focus();
                                    inp.value = val;
                                    inp.dispatchEvent(new Event('input', {bubbles: true}));
                                    inp.dispatchEvent(new Event('change', {bubbles: true}));
                                }
                                return;
                            }
                        }
                    }""",
                    invalid_val,
                )
                _click_save(hmi_page)

                has_error = _has_validation_error(hmi_page)
                _dismiss_toast(hmi_page)

                assert has_error, (
                    f"Time To Live = {invalid_val} 超出范围 5-1440，"
                    f"应出现验证错误，但未检测到任何错误提示"
                )

        finally:
            _set_foreign_device_enable(hmi_page, initial_foreign)
            hmi_page.wait_for_timeout(500)

    def test_011_invalid_bbmd_config_rejected(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_011: 非法 BBMD 配置保存被阻止

        在 Foreign Device 启用后，同时输入非法值：
          BBMD IP   = "300.1.1.1"（非法 IP）
          BBMD Port = "70000"（超出范围 47808-49000）
          Time To Live = "0"（低于下限 5）
        期望 Save 后出现验证错误。
        清理：恢复 Foreign Device 初始状态。
        """
        initial_foreign = _get_foreign_device_state(hmi_page)

        try:
            if not initial_foreign:
                _set_foreign_device_enable(hmi_page, True)
                hmi_page.wait_for_timeout(1000)

            # 设置 BBMD IP 为非法值
            hmi_page.evaluate(
                """() => {
                    const inputs = document.querySelectorAll('input.el-input__inner');
                    for (const inp of inputs) {
                        const ph = (inp.getAttribute('placeholder') || '').toLowerCase();
                        if (ph.includes('bbmd ip') || ph.includes('enter bbmd ip')) {
                            inp.focus();
                            inp.value = '300.1.1.1';
                            inp.dispatchEvent(new Event('input', {bubbles: true}));
                            inp.dispatchEvent(new Event('change', {bubbles: true}));
                            return;
                        }
                    }
                    // 通过 label 回退
                    const labels = document.querySelectorAll('.el-form-item__label');
                    for (const lbl of labels) {
                        if (lbl.textContent.trim() === 'BBMD IP') {
                            const fi = lbl.closest('.el-form-item');
                            const inp = fi && fi.querySelector('input.el-input__inner');
                            if (inp) {
                                inp.focus(); inp.value = '300.1.1.1';
                                inp.dispatchEvent(new Event('input', {bubbles: true}));
                                inp.dispatchEvent(new Event('change', {bubbles: true}));
                            }
                        }
                    }
                }"""
            )

            # 设置 BBMD Port 为超范围值
            hmi_page.evaluate(
                """() => {
                    const labels = document.querySelectorAll('.el-form-item__label');
                    for (const lbl of labels) {
                        if (lbl.textContent.trim() === 'BBMD Port') {
                            const fi = lbl.closest('.el-form-item');
                            const inp = fi && fi.querySelector('input.el-input__inner');
                            if (inp) {
                                inp.focus(); inp.value = '70000';
                                inp.dispatchEvent(new Event('input', {bubbles: true}));
                                inp.dispatchEvent(new Event('change', {bubbles: true}));
                            }
                            return;
                        }
                    }
                }"""
            )

            # 设置 Time To Live = 0（低于下限 5）
            hmi_page.evaluate(
                """() => {
                    const labels = document.querySelectorAll('.el-form-item__label');
                    for (const lbl of labels) {
                        if (lbl.textContent.trim() === 'Time To Live') {
                            const fi = lbl.closest('.el-form-item');
                            const inp = fi && fi.querySelector('input.el-input__inner');
                            if (inp) {
                                inp.focus(); inp.value = '0';
                                inp.dispatchEvent(new Event('input', {bubbles: true}));
                                inp.dispatchEvent(new Event('change', {bubbles: true}));
                            }
                            return;
                        }
                    }
                }"""
            )

            _click_save(hmi_page)

            has_error = _has_validation_error(hmi_page)
            _dismiss_toast(hmi_page)

            assert has_error, (
                "BBMD IP=300.1.1.1 / BBMD Port=70000 / TTL=0 均为非法值，"
                "Save 后应出现验证错误，但未检测到任何错误提示"
            )

        finally:
            _set_foreign_device_enable(hmi_page, initial_foreign)
            hmi_page.wait_for_timeout(500)

    def test_032_cov_batch_update_2100_save(self, hmi_page: Page) -> None:
        """
        TestCase_AcuHMI-1-7_033_001_032: COV Batch Update 配置 AcuRev2100 参数保存正常

        流程：
        1. 找到 AcuRev-2100 设备，打开 Parameter Config，为 realtime 组所有可见参数使能
           Polling Enable
        2. 点击 COV Batch Update，在 Parameters 输入框点击 Select All 选择所有参数，
           输入 COV Increment=1.500，点击 Apply
        3. 关闭弹窗，重新导航，重新打开 Parameter Config
        4. 验证参数 COV Increment 已按批量值更新为 1.500
        设备未接入时 pytest.skip。
        """
        keyword_2100, was_checked, reason = _resolve_device_by_keyword(hmi_page, "2100")
        if keyword_2100 is None:
            pytest.skip(reason)
        if not was_checked:
            _set_device_checked(hmi_page, keyword_2100, True)
            hmi_page.wait_for_timeout(500)
            # 持久化设备映射：勾选状态需主页面 Save 才能在重新导航后保留
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)

        # 步骤1: 打开 Parameter Config，使能所有可见行 Polling Enable
        if not _open_param_dialog(hmi_page, keyword_2100):
            pytest.skip("无法打开 AcuRev-2100 的 Parameter Config 弹窗")

        enabled_rows = _enable_polling_for_all_visible_rows(hmi_page)
        if enabled_rows == 0:
            _close_param_config_dialog(hmi_page)
            pytest.skip("Parameter Config 弹窗中未找到任何参数行，需真机确认")
        hmi_page.wait_for_timeout(500)

        # 在打开 COV Batch Update 之前从表格读第一行参数名（与 test_033 保持一致）
        first_param: str = hmi_page.evaluate(
            """() => {
                const rows = document.querySelectorAll(
                    '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                );
                if (!rows.length) return '';
                const cells = rows[0].querySelectorAll('td');
                return cells[0] ? cells[0].textContent.trim() : '';
            }"""
        )
        if not first_param:
            _close_param_config_dialog(hmi_page)
            pytest.skip("Parameter Config 弹窗第一行参数名为空，需真机确认")

        try:
            # 步骤2: 打开 COV Batch Update
            batch_opened = _open_cov_batch_update(hmi_page)
            if not batch_opened:
                pytest.skip("COV Batch Update 弹窗未出现，可能按钮未找到")

            # 在 Parameters 输入框点击 Select All 选择所有参数
            selected_any = _playwright_batch_select_all(hmi_page)
            if not selected_any:
                hmi_page.keyboard.press("Escape")
                pytest.skip(
                    "COV Batch Update 下拉选项为空（可能为 lazy load 未触发），需真机确认"
                )
            hmi_page.wait_for_timeout(300)

            # 设置 COV Increment = 1.500
            _playwright_batch_fill_cov(hmi_page, "1.500")

            # Playwright 原生点击 Confirm/确定（JS click 不触发 Vue 3 handler）
            _confirm_batch_update_playwright(hmi_page)
            hmi_page.wait_for_timeout(_SAVE_WAIT_MS)

        finally:
            if hmi_page.evaluate(
                "() => !!document.querySelector('[aria-label=\"Batch Update\"]')"
            ):
                _close_batch_update_dialog(hmi_page)
            # COV Batch Update 的改动只存于 Parameter Config 弹窗的 Vue 状态，
            # 必须在弹窗关闭之前触发保存，否则关闭时状态丢失。
            # 优先找弹窗内 Save 按钮；找不到则在弹窗开启状态下 force=True 点主区域 Save
            # （绕过 el-overlay 遮罩），再关弹窗。
            saved_in_dlg: bool = hmi_page.evaluate(
                """() => {
                    const dlg = document.querySelector('[aria-label="Parameter Config"]');
                    if (!dlg) return false;
                    for (const btn of dlg.querySelectorAll('button')) {
                        if (btn.textContent.trim() === 'Save') { btn.click(); return true; }
                    }
                    return false;
                }"""
            )
            if saved_in_dlg:
                hmi_page.wait_for_timeout(_SAVE_WAIT_MS)
                _dismiss_toast(hmi_page)
            else:
                # 弹窗内无 Save：在弹窗开启状态下强制点击主区域 Save
                hmi_page.locator('button:has-text("Save")').first.click(force=True)
                hmi_page.wait_for_timeout(_SAVE_WAIT_MS)
                _dismiss_toast(hmi_page)
            _close_param_config_dialog(hmi_page)

        # 步骤3: 重新导航
        _navigate_to_bacnet(hmi_page)

        # 步骤4: 重新打开 Parameter Config，验证 COV Increment 已批量更新
        if not _open_param_dialog(hmi_page, keyword_2100):
            pytest.fail("重新导航后无法打开 AcuRev-2100 Parameter Config 弹窗")

        try:
            actual_val = _get_cov_increment_value_in_dialog(hmi_page, first_param)
            assert actual_val in ("1.500", "1.5"), (
                f"Batch Update 保存后，参数 {first_param!r} 的 COV Increment 预期为 '1.500'，"
                f"实际为 {actual_val!r}"
            )
        finally:
            _close_param_config_dialog(hmi_page)
