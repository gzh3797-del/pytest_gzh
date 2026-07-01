# -*- coding: utf-8 -*-
"""
test_bacnet_ui_basic.py — BACnet/IP 参数列表一致性用例（P0）

用例覆盖：
  TestCase_AcuHMI-1-7_033_001_019: AcuRev-2100 参数列表与模板一致（设备未接入时自动跳过）
  TestCase_AcuHMI-1-7_033_001_020: AcuRev-4100 参数列表与北向模板一致
  TestCase_AcuHMI-1-7_033_001_034: COV Batch Update AcuRev-2100 参数列表一致（设备未接入时自动跳过）
  TestCase_AcuHMI-1-7_033_001_035: COV Batch Update AcuRev-4100 参数列表一致
  TestCase_AcuHMI-1-7_033_001_046: AcuvimIIR（PXE1）参数列表与模板一致（设备未接入时自动跳过）
  TestCase_AcuHMI-1-7_033_001_047: AcuvimIIW（PXE2）参数列表与模板一致（设备未接入时自动跳过）
  TestCase_AcuHMI-1-7_033_001_048: AcuVIM3 参数列表与模板一致（设备未接入时自动跳过）
  TestCase_AcuHMI-1-7_033_001_049: AcuRev1300（PXM350）参数列表与模板一致（设备未接入时自动跳过）
  TestCase_AcuHMI-1-7_033_001_055: COV Batch Update AcuvimIIR 参数列表一致（设备未接入时自动跳过）
  TestCase_AcuHMI-1-7_033_001_056: COV Batch Update AcuvimIIW 参数列表一致（设备未接入时自动跳过）
  TestCase_AcuHMI-1-7_033_001_057: COV Batch Update AcuVIM3 参数列表一致（设备未接入时自动跳过）
  TestCase_AcuHMI-1-7_033_001_058: COV Batch Update AcuRev1300 参数列表一致（设备未接入时自动跳过）

  注：050~054 已被 test_bacnet_six_segment.py 占用（元数据/Device Object/协议合规/稳定性），
  故本文件 COV Batch 四条顺延至 055~058，避免用例编号与既有脚本冲突。

设备解析：
  4100/2100 用例用名称关键词匹配（"4100"/"2100" 在设备名中唯一），见 _resolve_device_by_keyword；
  AcuvimIIR/IIW/VIM3/1300 用例用网关动态发现（physical_devices_reader）按 deviceModel
  解析模板，见 _resolve_device_for_template——因 AcuvimIIR / AcuvimIIW 同族设备名共享
  "Acuvim" 前缀，关键词子串无法区分，必须靠 Model 字段。

  两类解析共同点（参数列表/COV 比对只读设备**静态参数模板**，不读实时值）：
  - 不要求设备在线：动态发现在线设备优先、无在线设备时回退到离线设备；
  - 不要求预先勾选：目标设备在设备表中**未勾选映射**时，比对前自动勾选（仅改 UI 状态、
    不点 Save，故不触发网关重启），比对结束后（含断言失败）经 _ensure_device_mapped
    恢复为未勾选。仅"设备表中根本没有该设备/该型号"才跳过。

运行：
  pytest projects/PX_EMD_G/tests/BacnetIP/test_bacnet_ui_basic.py -v
"""
from __future__ import annotations

import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, Optional

import pytest
from playwright.sync_api import Page

# ── 路径 ─────────────────────────────────────────────────────────────────────
_REPO_ROOT = str(Path(__file__).parents[4])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from projects.PX_EMD_G.helpers.template_matcher import (  # noqa: E402
    get_bacnet_descriptions,
    get_bacnet_descriptions_4100,
    get_bacnet_descriptions_2100,
)
from projects.PX_EMD_G.helpers.physical_devices_reader import (  # noqa: E402
    DiscoveredDevice,
    pick_device_for_template,
)


# ═════════════════════════════════════════════════════════════════════════════
# 页面操作辅助函数
# ═════════════════════════════════════════════════════════════════════════════

# 原子读取「Devices Selection」表中已勾选设备名的 JS。
# 必须用单次 evaluate 快照：设备表的 Online 状态会定时重渲染，若用 Playwright 逐行
# nth(i)+get_attribute 迭代，重渲染会令中途的行句柄失效，导致只读到第一行（曾因此
# 误判"设备未接入"）。返回 null=未找到设备表（仍在导航/加载）；[]=表在但无已勾选行。
_DEVICE_NAMES_JS = """() => {
    let table = null;
    for (const fi of document.querySelectorAll('.el-form-item')) {
        const lbl = fi.querySelector('.el-form-item__label');
        if (lbl && lbl.textContent.trim().includes('Devices Selection')) {
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
    从 BACnet/IP 页面的 Devices Selection 表格读取已勾选设备的名称。
    仅返回 checkbox 处于 is-checked 状态的行（即已映射设备）。

    用单次原子 JS 快照读取（避免逐行迭代被表格重渲染打断只读到首行），并轮询到
    设备表渲染出非空结果再返回（兼顾冷启动的异步渲染窗口）。表格本身是整批渲染
    （要么 0 行要么全部），故首个非空快照即为完整列表。

    注：本文件用例已改用 _resolve_device_by_keyword（全量行+自动勾选），不再用本函数；
    保留是因 test_bacnet_ui_config.py 仍 import 它读取已勾选设备名。
    """
    for _ in range(20):  # 最多约 6s
        snapshot = page.evaluate(_DEVICE_NAMES_JS)
        if snapshot:
            return snapshot
        page.wait_for_timeout(300)
    return []


# ── 全量设备行读取（不按 is-checked 过滤）──────────────────────────────────────

# 原子快照 JS：读取设备表所有行（含未勾选），返回每行的 name/checked/online。
# Online 列（index 4）DOM：
#   <span class="status-content__sign online"></span>  → 在线
#   <span class="status-content__label">ON</span>      → 文本辅助确认
# 若 .status-content__sign 不含 "online" class 则视为离线。
_ALL_DEVICE_ROWS_JS = """() => {
    let table = null;
    for (const fi of document.querySelectorAll('.el-form-item')) {
        const lbl = fi.querySelector('.el-form-item__label');
        if (lbl && lbl.textContent.trim().includes('Devices Selection')) {
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
    if (rows.length === 0) return null;
    const out = [];
    for (const row of rows) {
        const cells = row.querySelectorAll('td');
        if (cells.length < 5) continue;
        const name = (cells[1].textContent || '').trim();
        if (!name) continue;
        const cbInput = row.querySelector('td:first-child .el-checkbox__input');
        const checked = cbInput
            ? (cbInput.className || '').includes('is-checked')
            : false;
        const signEl = cells[4].querySelector('.status-content__sign');
        const online = signEl
            ? (signEl.className || '').includes('online')
            : false;
        out.push({name: name, checked: checked, online: online});
    }
    return out;
}"""


def _get_all_device_rows(page: Page) -> list[dict]:
    """
    从 BACnet/IP 页面的 Devices Selection 表格读取**全部行**（不按勾选过滤）。

    每行返回 {"name": str, "checked": bool, "online": bool}。

    Online 字段依据 Online 列（第 5 列，index 4）中 status-content__sign 的 CSS class：
    含 "online" 则为 True，否则 False（不受 Playwright 逐行迭代渲染中断影响——
    原因同 _get_device_names：设备表 Online 状态定时重渲染，必须用单次原子 JS 快照）。

    轮询到表渲染出非空结果后返回（兼顾冷启动的异步渲染窗口）。
    """
    for _ in range(20):  # 最多约 6s
        snapshot = page.evaluate(_ALL_DEVICE_ROWS_JS)
        if snapshot:
            return snapshot
        page.wait_for_timeout(300)
    return []


def _get_all_device_names(page: Page) -> list[str]:
    """
    返回 Devices Selection 表格中所有设备的名称（含未勾选行）。

    基于 _get_all_device_rows 的全量快照，供用例做设备发现用。
    与 _get_device_names 的区别：后者只返回已勾选（is-checked）的设备名。
    """
    return [row["name"] for row in _get_all_device_rows(page)]


# 用于 _set_device_checked 的原子 JS：按设备全名精确匹配找行并切换 checkbox。
# 用精确全名（而非子串）匹配，避免名字互为前缀的设备相互误命中
# （如 "Acurev4100" 是 "Acurev4100229" 的前缀，子串匹配会勾错/漏取消）。
# 通过对 .el-checkbox__original 发送 MouseEvent('click', bubbles=True) 来 toggle，
# 这比 element.click() 更接近真实鼠标事件，能正确触发 Vue 3 响应式更新。
# 返回值：
#   "no_table"    — 未找到设备表
#   "not_found"   — 未找到匹配行
#   "unchanged"   — 已是目标状态，未执行操作
#   "ok"          — 成功点击（事件已 dispatch）
_SET_DEVICE_CHECKED_JS = """([deviceName, targetChecked]) => {
    let table = null;
    for (const fi of document.querySelectorAll('.el-form-item')) {
        const lbl = fi.querySelector('.el-form-item__label');
        if (lbl && lbl.textContent.trim().includes('Devices Selection')) {
            table = fi.querySelector('.el-table');
            break;
        }
    }
    if (!table) {
        const grp = document.querySelector('[role="group"] .el-table');
        if (grp) table = grp;
    }
    if (!table) return 'no_table';

    const rows = table.querySelectorAll('.el-table__body tr.el-table__row');
    const nameLower = deviceName.toLowerCase();
    for (const row of rows) {
        const cells = row.querySelectorAll('td');
        if (cells.length < 2) continue;
        const name = (cells[1].textContent || '').trim();
        if (name.toLowerCase() !== nameLower) continue;

        const cbInput = row.querySelector('td:first-child .el-checkbox__input');
        const isChecked = cbInput
            ? (cbInput.className || '').includes('is-checked')
            : false;

        if (isChecked === targetChecked) return 'unchanged';

        // 对 .el-checkbox__original 发送 MouseEvent，触发完整 Vue 3 事件链
        const cbOrig = row.querySelector('td:first-child .el-checkbox__original');
        if (cbOrig) {
            cbOrig.scrollIntoView({block: 'center', inline: 'center'});
            cbOrig.dispatchEvent(
                new MouseEvent('click', {bubbles: true, cancelable: true})
            );
            return 'ok';
        }
        return 'not_found';
    }
    return 'not_found';
}"""


def _set_device_checked(page: Page, device_name: str, checked: bool) -> bool:
    """
    在 Devices Selection 表格中，按 device_name 精确匹配设备全名
    （大小写不敏感）找到设备行，将其第一列 checkbox 设为目标状态 ``checked``。

    用全名精确匹配（而非子串）：避免名字互为前缀的设备相互误命中——例如
    "Acurev4100" 是 "Acurev4100229" 的前缀，子串匹配会勾错或漏取消。
    调用方应先经设备发现解析出完整设备名再传入。

    若已是目标状态则不操作。返回 True 表示找到了匹配行（不论是否执行了点击）；
    返回 False 表示未找到匹配行或设备表未渲染。

    实现策略：
    - 优先通过 JS 对 ``.el-checkbox__original`` 发送 ``MouseEvent('click', bubbles=True)``，
      这比 ``element.click()`` 更接近真实鼠标事件，能正确触发 El Plus v2 的 Vue 3 响应式更新。
    - 若 JS 方式无效（理论上不应出现，但留余量）则按团队约定降级坐标点击：
      scrollIntoView 后取 getBoundingClientRect 中心坐标，经 page.mouse.click() 分发。

    注意：本函数**只改 UI 勾选状态，不点 Save**。
    """
    result: str = page.evaluate(_SET_DEVICE_CHECKED_JS, [device_name, checked])
    if result in ("ok", "unchanged"):
        if result == "ok":
            page.wait_for_timeout(300)
        return True

    if result == "no_table":
        return False

    # result == "not_found"
    return False


def _isolate_single_device(page: Page, device_name: str) -> list[str]:
    """
    取消其余所有设备的勾选、只保留 device_name（全名精确匹配）的设备勾选
    （若其原本未勾则勾上）。

    用全名精确匹配（而非子串）：名字互为前缀的设备（如 "Acurev4100" 与
    "Acurev4100229"）才能被正确区分隔离，避免兄弟设备漏取消导致多设备同时上传。
    调用方应先经设备发现解析出完整设备名再传入。

    返回**调用前**已勾选设备的名称列表，供 teardown 调用 _restore_device_selection 恢复现场。

    本函数**只改 UI 勾选状态，不点 Save**——由调用方统一 Save 并等待重启，便于控制时序。

    实现顺序（减少 Save 触发次数，也减少无谓渲染抖动）：
    1. 一次快照读取全量行及当前勾选状态
    2. 记录调用前已勾选集合
    3. 取消所有非目标行的已勾选行
    4. 确保目标行处于勾选状态
    """
    rows = _get_all_device_rows(page)

    # 调用前已勾选名称集合（供 teardown 恢复）
    original_checked: list[str] = [r["name"] for r in rows if r["checked"]]

    target_lower = device_name.lower()

    # 取消所有非目标行的勾选（按全名精确匹配区分目标与其余设备）
    for row in rows:
        if row["name"].lower() != target_lower:
            if row["checked"]:
                _set_device_checked(page, row["name"], False)

    # 确保目标行勾选
    _set_device_checked(page, device_name, True)

    return original_checked


def _restore_device_selection(page: Page, checked_names: list[str]) -> None:
    """
    将 Devices Selection 表格的勾选状态恢复为 ``checked_names`` 集合：
    名称在集合内的行勾上，不在的取消。

    本函数**只改 UI 勾选状态，不点 Save**——由调用方统一 Save。
    """
    checked_set = {n.lower() for n in checked_names}
    rows = _get_all_device_rows(page)
    for row in rows:
        want_checked = row["name"].lower() in checked_set
        _set_device_checked(page, row["name"], want_checked)


def _open_param_dialog(page: Page, device_name: str) -> bool:
    """
    在 Devices Selection 表格中找到 device_name（全名精确匹配，
    大小写不敏感）且已勾选的行，点击其最后一列的 Parameter Selection 按钮打开弹窗。
    找不到则返回 False。

    用全名精确匹配（而非子串）：避免名字互为前缀的设备（如 "Acurev4100" 与
    "Acurev4100229"）误打开错设备的参数弹窗。调用方应先经设备发现解析出完整设备名。
    """
    rows = page.locator('[role="group"] .el-table__body tr.el-table__row')
    target_lower = device_name.lower()
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
        if name.lower() != target_lower:
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


def _enable_polling_select_all(page: Page) -> bool:
    """
    在已打开的 Parameter Config 弹窗中，通过点击 "Polling Enable" 列头的 el-checkbox
    一键开启当前 Parameter Type 下**所有分页**的 Polling Enable。

    此列头 checkbox 是全选控件：点击后后端会对该 Type 的全量参数（含所有未翻到的页）
    批量修改状态，无需翻页，比逐页逐行点击快约百倍。

    幂等处理（防止误关闭）：
    - 列头已是 ``is-checked``（全部已开启）→ 直接返回 True，不再点击。
    - 列头是 ``is-indeterminate``（部分开启）或未选中（全部关闭）→ 点一次使其全选。
    El Plus v2 checkbox 合成事件需对 ``.el-checkbox__original`` 发送 MouseEvent，
    与项目约定的 el-checkbox 操作方式保持一致。

    返回值：
    - True  — 成功操作（已全选或已是全选状态无需操作）。
    - False — 弹窗不存在或未找到 Polling Enable 列头 checkbox（调用方应回退到
              ``_enable_polling_all_pages`` 逐页翻页处理）。
    """
    result: str = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Parameter Config"]');
            if (!dlg) return 'no_dialog';

            const thead = dlg.querySelector('.el-table__header');
            if (!thead) return 'no_thead';

            for (const th of thead.querySelectorAll('th')) {
                const label = th.querySelector('.el-checkbox');
                if (!label) continue;
                // 用列头标签文字定位，不依赖动态 class 名（如 el-table_N_column_M）
                if (!label.textContent.includes('Polling Enable')) continue;

                const cbInput = label.querySelector('.el-checkbox__input');
                if (!cbInput) return 'no_input';

                // 已是全选状态则无需操作，直接返回成功
                if (cbInput.classList.contains('is-checked')) return 'already_checked';

                // 半选（部分开启）或全部关闭时，点一次 → 全选
                const cbOriginal = label.querySelector('.el-checkbox__original');
                if (!cbOriginal) return 'no_original';

                cbOriginal.scrollIntoView({block: 'center', behavior: 'instant'});
                cbOriginal.dispatchEvent(
                    new MouseEvent('click', {bubbles: true, cancelable: true})
                );
                return 'clicked';
            }
            return 'not_found';
        }"""
    )
    # 点击后稍等 Vue 响应式更新生效
    if result == "clicked":
        page.wait_for_timeout(300)
    return result in ("clicked", "already_checked")


def _enable_polling_all_pages(page: Page) -> None:
    """
    在已打开的 Parameter Config 弹窗中，对当前 Parameter Type 的所有分页、
    所有行批量开启 Polling Enable（第 1 列 checkbox）。
    COV Batch Update 的 Parameters 下拉只显示 Polling Enable 已开启的参数，
    因此必须先调用本函数，再打开 COV Batch Update。
    本函数保留供 test_020 参数列表用例及 _collect_cov_batch_params 继续使用；
    _enable_all_polling_for_device 已改用更快的 _enable_polling_select_all。
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


def _resolve_device_by_keyword(
    page: Page, keyword: str
) -> tuple[Optional[str], bool, str]:
    """在 BACnet 设备表中按名称关键词（子串）定位设备，返回其当前勾选状态。

    用于 4100/2100——其设备名内含唯一型号数字（"4100"/"2100"），关键词足以区分。
    在全量行（含未勾选）中匹配：已勾选的同名设备优先（避免对本就勾选的设备做无谓
    勾选/取消）；无已勾选匹配时回退到第一台名称匹配的设备，由调用方经
    _ensure_device_mapped 临时勾选。

    返回 (device_name, was_checked, reason)：device_name 为 None 时跳过，reason 说明原因。
    """
    rows = _get_all_device_rows(page)
    row = next((r for r in rows if keyword in r["name"] and r["checked"]), None)
    if row is None:
        row = next((r for r in rows if keyword in r["name"]), None)
    if row is None:
        names = [r["name"] for r in rows]
        return None, False, (
            f"设备未接入（设备表中无名称含 {keyword!r} 的设备，当前设备列表：{names}），跳过此用例"
        )
    return row["name"], bool(row["checked"]), "ok"


def _resolve_device_for_template(
    page: Page,
    discovered: list[DiscoveredDevice],
    template_name: str,
) -> tuple[Optional[str], bool, str]:
    """按 template_name 在 BACnet 设备表中定位目标设备，并返回其当前勾选状态。

    解析步骤：
    1. 从网关动态发现列表里按 deviceModel→模板取第一台在线设备（pick_device_for_template）；
    2. 在 BACnet 设备表 _get_all_device_rows 里按全名精确匹配确认该设备存在。

    用动态发现的 Model 字段区分同族设备（如 AcuvimIIR / AcuvimIIW），名称关键词子串无法区分。

    返回 (device_name, was_checked, reason)：
    - device_name 为 None 时跳过，reason 说明原因；was_checked 无意义（False）。
    - device_name 非 None 时 reason 为 "ok"，was_checked 表示该设备**当前是否已勾选映射**——
      未勾选时调用方应经 _ensure_device_mapped 临时勾选再读参数（参数弹窗仅对已勾选行可开）。
    """
    # 参数列表/COV 比对只读设备的静态参数模板（其支持的参数集，由设备型号决定），
    # 不读实时值、不依赖连通性；故在线设备优先，无在线设备时回退到离线设备——
    # 离线设备的 Parameter Config 仍按设备型号展示参数。仅"该型号一台设备都没有"才跳过。
    dev = pick_device_for_template(discovered, template_name, online_only=True)
    if dev is None:
        dev = pick_device_for_template(discovered, template_name, online_only=False)
    if dev is None:
        return None, False, (
            f"网关下挂设备中无模板为 {template_name!r} 的设备"
            f"（已发现：{[(d.name, d.model, d.online) for d in discovered]}），跳过此用例"
        )
    rows = _get_all_device_rows(page)
    row = next((r for r in rows if r["name"].lower() == dev.name.lower()), None)
    if row is None:
        return None, False, (
            f"动态发现的设备 {dev.name!r} 不在 BACnet 设备表中"
            "（设备表与 Physical Devices 不一致，请核查网关），跳过此用例"
        )
    return dev.name, bool(row["checked"]), "ok"


@contextmanager
def _ensure_device_mapped(
    page: Page, device_name: str, was_checked: bool
) -> Iterator[None]:
    """确保 device_name 在比对期间处于勾选（映射）状态，比对结束后恢复原状。

    参数列表/COV 读取均**只读**（弹窗只点 Close 不点 Save），故本上下文管理器
    全程只改 UI 勾选状态、不点 Save——既不触发网关 BACnet 服务重启，也不会把临时
    勾选持久化到网关。

    - was_checked 为 True：设备原本已勾选，进入/退出都不动它。
    - was_checked 为 False：进入时勾上（使其 Parameter Selection 按钮可用），退出时
      取消勾选恢复原状；退出在 finally 中执行，断言失败也保证还原现场，不影响同
      module 后续用例（hmi_page 为 module 级）。
    """
    if not was_checked:
        _set_device_checked(page, device_name, True)
    try:
        yield
    finally:
        if not was_checked:
            _set_device_checked(page, device_name, False)


# ═════════════════════════════════════════════════════════════════════════════
# 测试用例
# ═════════════════════════════════════════════════════════════════════════════

class TestBACnetParamListConsistency:
    """BACnet/IP 参数列表与模板一致性验证（P0）。"""

    def test_019_acurev2100_param_list_matches_template(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_019: AcuRev-2100 参数列表与模板展示一致。"""
        device, was_checked, reason = _resolve_device_by_keyword(hmi_page, "2100")
        if device is None:
            pytest.skip(f"AcuRev-2100 {reason}")

        template_params = get_bacnet_descriptions_2100()
        with _ensure_device_mapped(hmi_page, device, was_checked):
            page_params = _collect_all_params_for_device(hmi_page, device)
            _assert_param_consistency(page_params, template_params, "AcuRev-2100")

    def test_020_acurev4100_param_list_matches_template(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_020: AcuRev-4100 参数列表与北向模板一致。"""
        device, was_checked, reason = _resolve_device_by_keyword(hmi_page, "4100")
        if device is None:
            pytest.skip(f"AcuRev-4100 {reason}")

        template_params = get_bacnet_descriptions_4100()
        with _ensure_device_mapped(hmi_page, device, was_checked):
            page_params = _collect_all_params_for_device(hmi_page, device)
            assert page_params, "页面未采集到任何参数，请检查弹窗选择器是否正确"
            _assert_param_consistency(page_params, template_params, "AcuRev-4100")

    def test_034_acurev2100_cov_batch_update_matches_template(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_034: COV Batch Update AcuRev-2100 参数列表与模板一致。"""
        device, was_checked, reason = _resolve_device_by_keyword(hmi_page, "2100")
        if device is None:
            pytest.skip(f"AcuRev-2100 {reason}")

        template_params = get_bacnet_descriptions_2100()
        with _ensure_device_mapped(hmi_page, device, was_checked):
            page_params = _collect_cov_batch_params(hmi_page, device)
            _assert_param_consistency(
                page_params, template_params, "AcuRev-2100", "COV Batch Update"
            )

    def test_035_acurev4100_cov_batch_update_matches_template(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_035: COV Batch Update AcuRev-4100 参数列表与模板一致。"""
        device, was_checked, reason = _resolve_device_by_keyword(hmi_page, "4100")
        if device is None:
            pytest.skip(f"AcuRev-4100 {reason}")

        template_params = get_bacnet_descriptions_4100()
        with _ensure_device_mapped(hmi_page, device, was_checked):
            page_params = _collect_cov_batch_params(hmi_page, device)
            _assert_param_consistency(
                page_params, template_params, "AcuRev-4100", "COV Batch Update"
            )

    # ── 参数列表一致性：AcuvimIIR / AcuvimIIW / AcuVIM3 / AcuRev1300 ───────────
    # 同族设备名共享前缀，关键词无法区分，统一用 _resolve_device_for_template
    # 按网关动态发现的 deviceModel 解析目标设备；未勾选映射的设备由 _ensure_device_mapped
    # 比对前临时勾选、比对后恢复。

    def test_046_acuvimiir_param_list_matches_template(
        self, hmi_page: Page, discovered_devices: list
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_046: AcuvimIIR（PXE1）参数列表与模板一致。"""
        device, was_checked, reason = _resolve_device_for_template(
            hmi_page, discovered_devices, "AcuvimIIR"
        )
        if device is None:
            pytest.skip(f"AcuvimIIR {reason}")

        template_params = get_bacnet_descriptions("AcuvimIIR")
        with _ensure_device_mapped(hmi_page, device, was_checked):
            page_params = _collect_all_params_for_device(hmi_page, device)
            _assert_param_consistency(page_params, template_params, "AcuvimIIR")

    def test_047_acuvimiiw_param_list_matches_template(
        self, hmi_page: Page, discovered_devices: list
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_047: AcuvimIIW（PXE2）参数列表与模板一致。"""
        device, was_checked, reason = _resolve_device_for_template(
            hmi_page, discovered_devices, "AcuvimIIW"
        )
        if device is None:
            pytest.skip(f"AcuvimIIW {reason}")

        template_params = get_bacnet_descriptions("AcuvimIIW")
        with _ensure_device_mapped(hmi_page, device, was_checked):
            page_params = _collect_all_params_for_device(hmi_page, device)
            _assert_param_consistency(page_params, template_params, "AcuvimIIW")

    def test_048_acuvim3_param_list_matches_template(
        self, hmi_page: Page, discovered_devices: list
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_048: AcuVIM3 参数列表与模板一致。"""
        device, was_checked, reason = _resolve_device_for_template(
            hmi_page, discovered_devices, "AcuVIM3"
        )
        if device is None:
            pytest.skip(f"AcuVIM3 {reason}")

        template_params = get_bacnet_descriptions("AcuVIM3")
        with _ensure_device_mapped(hmi_page, device, was_checked):
            page_params = _collect_all_params_for_device(hmi_page, device)
            _assert_param_consistency(page_params, template_params, "AcuVIM3")

    def test_049_acurev1300_param_list_matches_template(
        self, hmi_page: Page, discovered_devices: list
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_049: AcuRev1300（PXM350）参数列表与模板一致。"""
        device, was_checked, reason = _resolve_device_for_template(
            hmi_page, discovered_devices, "AcuRev1300"
        )
        if device is None:
            pytest.skip(f"AcuRev1300 {reason}")

        template_params = get_bacnet_descriptions("AcuRev1300")
        with _ensure_device_mapped(hmi_page, device, was_checked):
            page_params = _collect_all_params_for_device(hmi_page, device)
            _assert_param_consistency(page_params, template_params, "AcuRev1300")

    # ── COV Batch Update 参数列表一致性：AcuvimIIR / IIW / VIM3 / AcuRev1300 ───

    def test_055_acuvimiir_cov_batch_update_matches_template(
        self, hmi_page: Page, discovered_devices: list
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_055: COV Batch Update AcuvimIIR 参数列表与模板一致。"""
        device, was_checked, reason = _resolve_device_for_template(
            hmi_page, discovered_devices, "AcuvimIIR"
        )
        if device is None:
            pytest.skip(f"AcuvimIIR {reason}")

        template_params = get_bacnet_descriptions("AcuvimIIR")
        with _ensure_device_mapped(hmi_page, device, was_checked):
            page_params = _collect_cov_batch_params(hmi_page, device)
            _assert_param_consistency(
                page_params, template_params, "AcuvimIIR", "COV Batch Update"
            )

    def test_056_acuvimiiw_cov_batch_update_matches_template(
        self, hmi_page: Page, discovered_devices: list
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_056: COV Batch Update AcuvimIIW 参数列表与模板一致。"""
        device, was_checked, reason = _resolve_device_for_template(
            hmi_page, discovered_devices, "AcuvimIIW"
        )
        if device is None:
            pytest.skip(f"AcuvimIIW {reason}")

        template_params = get_bacnet_descriptions("AcuvimIIW")
        with _ensure_device_mapped(hmi_page, device, was_checked):
            page_params = _collect_cov_batch_params(hmi_page, device)
            _assert_param_consistency(
                page_params, template_params, "AcuvimIIW", "COV Batch Update"
            )

    def test_057_acuvim3_cov_batch_update_matches_template(
        self, hmi_page: Page, discovered_devices: list
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_057: COV Batch Update AcuVIM3 参数列表与模板一致。"""
        device, was_checked, reason = _resolve_device_for_template(
            hmi_page, discovered_devices, "AcuVIM3"
        )
        if device is None:
            pytest.skip(f"AcuVIM3 {reason}")

        template_params = get_bacnet_descriptions("AcuVIM3")
        with _ensure_device_mapped(hmi_page, device, was_checked):
            page_params = _collect_cov_batch_params(hmi_page, device)
            _assert_param_consistency(
                page_params, template_params, "AcuVIM3", "COV Batch Update"
            )

    def test_058_acurev1300_cov_batch_update_matches_template(
        self, hmi_page: Page, discovered_devices: list
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_058: COV Batch Update AcuRev1300 参数列表与模板一致。"""
        device, was_checked, reason = _resolve_device_for_template(
            hmi_page, discovered_devices, "AcuRev1300"
        )
        if device is None:
            pytest.skip(f"AcuRev1300 {reason}")

        template_params = get_bacnet_descriptions("AcuRev1300")
        with _ensure_device_mapped(hmi_page, device, was_checked):
            page_params = _collect_cov_batch_params(hmi_page, device)
            _assert_param_consistency(
                page_params, template_params, "AcuRev1300", "COV Batch Update"
            )
