# -*- coding: utf-8 -*-
"""
test_bacnet_ui_basic.py — BACnet/IP 参数列表一致性用例（P0）

用例覆盖：
  TestCase_AcuHMI-1-7_033_001_019: AcuRev-2100 参数列表与模板一致（设备未接入时自动跳过）
  TestCase_AcuHMI-1-7_033_001_020: AcuRev-4100 参数列表与北向模板一致
  TestCase_AcuHMI-1-7_033_001_034: COV Batch Update AcuRev-2100 参数列表一致（设备未接入时自动跳过）
  TestCase_AcuHMI-1-7_033_001_035: COV Batch Update AcuRev-4100 参数列表一致

运行：
  pytest projects/AcuHMI_1_7/tests/bacnet/test_bacnet_ui_basic.py -v
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest
from playwright.sync_api import Page

# ── 路径 ─────────────────────────────────────────────────────────────────────
_REPO_ROOT = str(Path(__file__).parents[4])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from projects.AcuHMI_1_7.helpers.template_matcher import (  # noqa: E402
    get_bacnet_descriptions_4100,
    get_bacnet_descriptions_2100,
)


# ═════════════════════════════════════════════════════════════════════════════
# 页面操作辅助函数
# ═════════════════════════════════════════════════════════════════════════════

# 原子读取「Devices Selection To Mapping」表中已勾选设备名的 JS。
# 必须用单次 evaluate 快照：设备表的 Online 状态会定时重渲染，若用 Playwright 逐行
# nth(i)+get_attribute 迭代，重渲染会令中途的行句柄失效，导致只读到第一行（曾因此
# 误判"设备未接入"）。返回 null=未找到设备表（仍在导航/加载）；[]=表在但无已勾选行。
_DEVICE_NAMES_JS = """() => {
    let table = null;
    for (const fi of document.querySelectorAll('.el-form-item')) {
        const lbl = fi.querySelector('.el-form-item__label');
        if (lbl && lbl.textContent.trim().includes('Devices Selection To Mapping')) {
            table = fi.querySelector('.el-table');
            break;
        }
    }
    if (!table) {
        const grp = document.querySelector('[role="group"] .el-table');
        if (grp) table = grp;
    }
    if (!table) return null;
    const rows = table.querySelectorAll('.el-table__body tr.el-table__row');
    const out = [];
    for (const row of rows) {
        const cb = row.querySelector('td:first-child .el-checkbox__input');
        if (!cb || !(cb.className || '').includes('is-checked')) continue;
        const cells = row.querySelectorAll('td');
        if (cells.length >= 2) {
            const name = (cells[1].textContent || '').trim();
            if (name) out.push(name);
        }
    }
    return out;
}"""


def _get_device_names(page: Page) -> list[str]:
    """
    从 BACnet/IP 页面的 Devices Selection To Mapping 表格读取已勾选设备的名称。
    仅返回 checkbox 处于 is-checked 状态的行（即已映射设备）。

    用单次原子 JS 快照读取（避免逐行迭代被表格重渲染打断只读到首行），并轮询到
    设备表渲染出非空结果再返回（兼顾冷启动的异步渲染窗口）。表格本身是整批渲染
    （要么 0 行要么全部），故首个非空快照即为完整列表。
    """
    for _ in range(20):  # 最多约 6s
        snapshot = page.evaluate(_DEVICE_NAMES_JS)
        if snapshot:
            return snapshot
        page.wait_for_timeout(300)
    return []


def _open_param_dialog(page: Page, device_keyword: str) -> bool:
    """
    在 Devices Selection To Mapping 表格中找到包含 device_keyword（模糊匹配）
    且已勾选的行，点击其最后一列的 Parameter Selection 按钮打开弹窗。
    找不到则返回 False。
    """
    rows = page.locator('[role="group"] .el-table__body tr.el-table__row')
    for i in range(rows.count()):
        row = rows.nth(i)
        checkbox = row.locator('td:first-child .el-checkbox__input')
        if checkbox.count() == 0:
            continue
        cls = checkbox.get_attribute("class") or ""
        if "is-checked" not in cls:
            continue
        cells = row.locator("td")
        if cells.count() < 2:
            continue
        name = (cells.nth(1).text_content() or "").strip()
        if device_keyword.lower() not in name.lower():
            continue
        btn = row.locator("td:last-child button")
        if btn.count() > 0:
            btn.first.click()
            page.wait_for_timeout(1000)
            return True
    return False


def _get_select_options_via_js(page: Page) -> list[str]:
    """
    用 JS 直接读取当前页面上所有可见 el-select-dropdown 的选项文本。
    比 Playwright locator 更可靠（Element Plus 把选项渲染在 body 级 portal）。
    """
    return page.evaluate(
        """() => {
            const items = document.querySelectorAll(
                '.el-select-dropdown__item:not(.is-disabled)'
            );
            return Array.from(items).map(el => el.textContent.trim()).filter(Boolean);
        }"""
    )


def _click_select_option(page: Page, text: str) -> bool:
    """点击下拉中文本匹配的选项，返回是否成功。"""
    result: bool = page.evaluate(
        """(text) => {
            const items = document.querySelectorAll('.el-select-dropdown__item');
            for (const el of items) {
                if (el.textContent.trim() === text) {
                    el.click();
                    return true;
                }
            }
            return false;
        }""",
        text,
    )
    return result


def _collect_table_page_params(page: Page) -> list[str]:
    """收集 Parameter Config 弹窗内当前表格页第一列（参数名称）的所有行文本。"""
    rows = page.locator(
        '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
    )
    params: list[str] = []
    for i in range(rows.count()):
        cells = rows.nth(i).locator("td")
        if cells.count() > 0:
            name = (cells.nth(0).text_content() or "").strip()
            if name:
                params.append(name)
    return params


def _collect_all_pages(page: Page) -> list[str]:
    """翻页收集 Parameter Config 弹窗内所有分页的参数名，直到下一页按钮禁用为止。
    用 JS click 绕过 Playwright 的 visibility 检查（弹窗内分页按钮可能被遮盖）。
    """
    all_params: list[str] = []
    page_num = 1
    while True:
        all_params.extend(_collect_table_page_params(page))

        is_disabled: bool = page.evaluate(
            """() => {
                const dlg = document.querySelector('[aria-label="Parameter Config"]');
                if (!dlg) return true;
                const btn = dlg.querySelector('.el-pagination .btn-next');
                if (!btn) return true;
                return btn.disabled || btn.classList.contains('is-disabled')
                    || btn.getAttribute('aria-disabled') === 'true';
            }"""
        )
        if is_disabled:
            break

        clicked: bool = page.evaluate(
            """() => {
                const dlg = document.querySelector('[aria-label="Parameter Config"]');
                if (!dlg) return false;
                const btn = dlg.querySelector('.el-pagination .btn-next');
                if (!btn) return false;
                btn.click();
                return true;
            }"""
        )
        if not clicked:
            break

        page.wait_for_timeout(300)
        page_num += 1
        if page_num > 200:
            break

    return all_params


def _get_parameter_types(page: Page, dlg_locator) -> list[str]:
    """
    打开 Parameter Config 弹窗内的 Parameter Type 下拉，读取所有选项。
    通过 aria-controls 精确定位弹窗专属的下拉列表，避免读到页面上其他
    Select（APDU Timeout、APDU Retries）的选项。
    """
    select_el = dlg_locator.locator(".el-select")
    if select_el.count() == 0:
        return []

    select_el.first.click()
    page.wait_for_timeout(600)

    opts: list[str] = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Parameter Config"]');
            if (!dlg) return [];
            const inp = dlg.querySelector('.el-select__input');
            if (!inp) return [];
            const listId = inp.getAttribute('aria-controls');
            if (!listId) return [];
            const list = document.getElementById(listId);
            if (!list) return [];
            return Array.from(
                list.querySelectorAll('.el-select-dropdown__item:not(.is-disabled)')
            ).map(el => el.textContent.trim()).filter(Boolean);
        }"""
    )
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)
    return opts


def _select_param_type(page: Page, dlg_locator, param_type: str) -> None:
    """在 Parameter Config 弹窗的 Parameter Type 下拉中选择指定值。
    通过 aria-controls 定位弹窗专属下拉列表后用 JS click 点击选项。
    """
    select_el = dlg_locator.locator(".el-select")
    if select_el.count() == 0:
        return
    select_el.first.click()
    page.wait_for_timeout(600)

    clicked: bool = page.evaluate(
        """(text) => {
            const dlg = document.querySelector('[aria-label="Parameter Config"]');
            if (!dlg) return false;
            const inp = dlg.querySelector('.el-select__input');
            if (!inp) return false;
            const listId = inp.getAttribute('aria-controls');
            if (!listId) return false;
            const list = document.getElementById(listId);
            if (!list) return false;
            for (const item of list.querySelectorAll('.el-select-dropdown__item')) {
                if (item.textContent.trim() === text) {
                    item.click();
                    return true;
                }
            }
            return false;
        }""",
        param_type,
    )
    if not clicked:
        page.keyboard.press("Escape")
    page.wait_for_timeout(800)


def _collect_all_params_for_device(page: Page, device_keyword: str) -> set[str]:
    """
    打开设备的 Parameter Config 弹窗，遍历所有 Parameter Type 和分页，
    返回全量参数名称集合。
    找不到设备时返回空集合。
    """
    if not _open_param_dialog(page, device_keyword):
        return set()

    dlg = page.locator(".el-dialog:visible").first
    param_types = _get_parameter_types(page, dlg)

    all_params: set[str] = set()
    for pt in param_types:
        _select_param_type(page, dlg, pt)
        dlg = page.locator(".el-dialog:visible").first
        all_params.update(_collect_all_pages(page))

    # 关闭弹窗（JS click 绕过 overlay visibility 检查）
    closed: bool = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Parameter Config"]');
            if (!dlg) return false;
            for (const btn of dlg.querySelectorAll('button')) {
                const t = btn.textContent.trim();
                if (t === 'Close' || t === '关闭') { btn.click(); return true; }
            }
            return false;
        }"""
    )
    if not closed:
        page.keyboard.press("Escape")
    page.wait_for_timeout(1000)

    return all_params


def _js_click_button_by_text(page: Page, *texts: str) -> bool:
    """
    用 JS 直接点击页面上文本匹配的第一个按钮，绕过 Playwright overlay 拦截检测。
    返回是否找到并点击了按钮。
    """
    return page.evaluate(
        """(texts) => {
            for (const btn of document.querySelectorAll('button')) {
                const t = btn.textContent.trim();
                if (texts.includes(t)) { btn.click(); return true; }
            }
            return false;
        }""",
        list(texts),
    )


def _close_batch_update_dialog(page: Page) -> None:
    """关闭 Batch Update 内层弹窗（用 JS 点击 Cancel / 取消，绕过 overlay 拦截）。"""
    clicked = page.evaluate(
        """() => {
            const overlay = document.querySelector('[aria-label="Batch Update"]');
            if (!overlay) return false;
            for (const btn of overlay.querySelectorAll('button')) {
                const t = btn.textContent.trim();
                if (['Cancel', '取消', 'Close', '关闭'].includes(t)) {
                    btn.click();
                    return true;
                }
            }
            return false;
        }"""
    )
    if not clicked:
        page.keyboard.press("Escape")
    page.wait_for_timeout(800)


def _close_param_config_dialog(page: Page) -> None:
    """关闭 Parameter Config 外层弹窗（用 JS 点击 Close / 关闭）。"""
    clicked = _js_click_button_by_text(page, "Close", "关闭")
    if not clicked:
        page.keyboard.press("Escape")
    page.wait_for_timeout(800)


def _enable_polling_all_pages(page: Page) -> None:
    """
    在已打开的 Parameter Config 弹窗中，对当前 Parameter Type 的所有分页、
    所有行批量开启 Polling Enable（第 1 列 checkbox）。
    COV Batch Update 的 Parameters 下拉只显示 Polling Enable 已开启的参数，
    因此必须先调用本函数，再打开 COV Batch Update。
    """
    page_num = 1
    while True:
        page.evaluate(
            """() => {
                const rows = document.querySelectorAll(
                    '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                );
                for (const row of rows) {
                    const cells = row.querySelectorAll('td');
                    if (cells.length < 2) continue;
                    const sw = cells[1].querySelector('.el-switch__input');
                    if (sw && sw.getAttribute('aria-checked') !== 'true') sw.click();
                }
            }"""
        )
        page.wait_for_timeout(200)

        is_last: bool = page.evaluate(
            """() => {
                const dlg = document.querySelector('[aria-label="Parameter Config"]');
                if (!dlg) return true;
                const btn = dlg.querySelector('.el-pagination .btn-next');
                if (!btn) return true;
                return btn.disabled || btn.classList.contains('is-disabled')
                    || btn.getAttribute('aria-disabled') === 'true';
            }"""
        )
        if is_last:
            break

        page.evaluate(
            """() => {
                const dlg = document.querySelector('[aria-label="Parameter Config"]');
                const btn = dlg && dlg.querySelector('.el-pagination .btn-next');
                if (btn) btn.click();
            }"""
        )
        page.wait_for_timeout(400)
        page_num += 1
        if page_num > 200:
            break

    # 翻回第一页
    page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Parameter Config"]');
            const first = dlg && dlg.querySelector('.el-pager .number');
            if (first) first.click();
        }"""
    )
    page.wait_for_timeout(300)


def _collect_cov_batch_params(page: Page, device_keyword: str) -> set[str]:
    """
    打开设备 Parameter Config 弹窗，遍历所有 Parameter Type：
    1. 切换到该 Type
    2. 翻页批量开启所有行的 Polling Enable（COV Batch Update 只显示已启用的参数）
    3. 打开 COV Batch Update，用 Playwright 真实点击展开 Parameters 下拉
    4. 通过 aria-controls 读取该 Type 的全量参数名
    返回跨所有 Type 的参数名称集合。
    """
    if not _open_param_dialog(page, device_keyword):
        return set()

    dlg_locator = page.locator('[aria-label="Parameter Config"]')
    param_types = _get_parameter_types(page, dlg_locator)

    all_params: set[str] = set()

    for pt in param_types:
        _select_param_type(page, dlg_locator, pt)

        # 批量开启当前 Type 所有行的 Polling Enable
        _enable_polling_all_pages(page)

        # 点 COV Batch Update 按钮
        clicked: bool = page.evaluate(
            """() => {
                for (const btn of document.querySelectorAll('button')) {
                    const t = btn.textContent.trim();
                    if (t.includes('COV Batch') || t === 'COV Batch Update') {
                        btn.click(); return true;
                    }
                }
                return false;
            }"""
        )
        if not clicked:
            continue

        page.wait_for_timeout(1000)
        batch_visible: bool = page.evaluate(
            "() => !!document.querySelector('[aria-label=\"Batch Update\"]')"
        )
        if not batch_visible:
            continue

        # 用 Playwright 真实点击展开 Parameters 下拉，等待展开
        batch_input = page.locator('[aria-label="Batch Update"] .el-select__input')
        if batch_input.count() > 0:
            batch_input.first.click()
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

        # 通过 aria-controls 读取该 Type 专属下拉列表的全量选项
        opts: list[str] = page.evaluate(
            """() => {
                const dlg = document.querySelector('[aria-label="Batch Update"]');
                if (!dlg) return [];
                const input = dlg.querySelector('.el-select__input');
                const listId = input && input.getAttribute('aria-controls');
                const list = listId ? document.getElementById(listId) : null;
                if (!list) return [];
                return Array.from(
                    list.querySelectorAll('.el-select-dropdown__item:not(.is-disabled)')
                ).map(el => el.textContent.trim()).filter(Boolean);
            }"""
        )
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)
        all_params.update(opts)

        _close_batch_update_dialog(page)
        page.wait_for_timeout(300)

    _close_param_config_dialog(page)

    return all_params


def _assert_param_consistency(
    page_params: set[str],
    template_params: set[str],
    device_label: str,
    context: str = "Parameter List",
) -> None:
    """比对页面参数集与模板参数集，不一致时给出明确的 diff 信息。"""
    missing = template_params - page_params
    extra = page_params - template_params
    assert not missing and not extra, (
        f"\n[{device_label} {context}] 参数不一致！\n"
        f"  模板有但页面无（共 {len(missing)} 条）：\n"
        + "".join(f"    - {p}\n" for p in sorted(missing))
        + f"  页面有但模板无（共 {len(extra)} 条）：\n"
        + "".join(f"    + {p}\n" for p in sorted(extra))
    )


# ═════════════════════════════════════════════════════════════════════════════
# 测试用例
# ═════════════════════════════════════════════════════════════════════════════

class TestBACnetParamListConsistency:
    """BACnet/IP 参数列表与模板一致性验证（P0）。"""

    def test_019_acurev2100_param_list_matches_template(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_019: AcuRev-2100 参数列表与模板展示一致。"""
        devices = _get_device_names(hmi_page)
        device_2100 = next(
            (d for d in devices if "2100" in d),
            None,
        )
        if device_2100 is None:
            pytest.skip(f"AcuRev-2100 设备未接入（当前设备列表：{devices}），跳过此用例")

        page_params = _collect_all_params_for_device(hmi_page, device_2100)
        template_params = get_bacnet_descriptions_2100()

        _assert_param_consistency(page_params, template_params, "AcuRev-2100")

    def test_020_acurev4100_param_list_matches_template(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_020: AcuRev-4100 参数列表与北向模板一致。"""
        devices = _get_device_names(hmi_page)
        keyword = next(
            (d for d in devices if "4100" in d),
            None,
        )
        if keyword is None:
            pytest.skip(f"AcuRev-4100 设备未接入（当前设备列表：{devices}），跳过此用例")

        page_params = _collect_all_params_for_device(hmi_page, keyword)
        template_params = get_bacnet_descriptions_4100()

        assert page_params, "页面未采集到任何参数，请检查弹窗选择器是否正确"
        _assert_param_consistency(page_params, template_params, "AcuRev-4100")

    def test_034_acurev2100_cov_batch_update_matches_template(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_034: COV Batch Update AcuRev-2100 参数列表与模板一致。"""
        devices = _get_device_names(hmi_page)
        device_2100 = next(
            (d for d in devices if "2100" in d),
            None,
        )
        if device_2100 is None:
            pytest.skip(f"AcuRev-2100 设备未接入（当前设备列表：{devices}），跳过此用例")

        page_params = _collect_cov_batch_params(hmi_page, device_2100)
        template_params = get_bacnet_descriptions_2100()

        _assert_param_consistency(
            page_params, template_params, "AcuRev-2100", "COV Batch Update"
        )

    def test_035_acurev4100_cov_batch_update_matches_template(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_035: COV Batch Update AcuRev-4100 参数列表与模板一致。"""
        devices = _get_device_names(hmi_page)
        keyword = next(
            (d for d in devices if "4100" in d),
            None,
        )
        if keyword is None:
            pytest.skip(f"AcuRev-4100 设备未接入（当前设备列表：{devices}），跳过此用例")

        page_params = _collect_cov_batch_params(hmi_page, keyword)
        template_params = get_bacnet_descriptions_4100()

        _assert_param_consistency(
            page_params, template_params, "AcuRev-4100", "COV Batch Update"
        )
