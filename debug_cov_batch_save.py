# -*- coding: utf-8 -*-
"""
debug_cov_batch_save.py
阶段 1–7：逐步调试 AcuRev-2100 COV Batch Update 保存路径
运行：python debug_cov_batch_save.py
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, Page

SCREENSHOTS_DIR = Path(
    "C:/Users/ZihanGao/Desktop/testing-team/test_case/AcuHMI_1_7/bacnet_ui/screenshots"
)
SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)

HMI_URL = "https://192.168.2.8"
USERNAME = "q"
PASSWORD = "1"
LOG = []  # 收集所有输出，最后汇总打印


def log(msg: str) -> None:
    print(msg)
    LOG.append(msg)


def ss(page: Page, name: str) -> str:
    """截图并返回路径。"""
    p = str(SCREENSHOTS_DIR / name)
    page.screenshot(path=p, full_page=False)
    log(f"  [screenshot] {p}")
    return p


def login(page: Page) -> None:
    page.goto(f"{HMI_URL}/#/login", wait_until="domcontentloaded", timeout=20000)
    page.wait_for_load_state("domcontentloaded", timeout=15000)
    time.sleep(1)

    # 使用 fill 填写表单
    try:
        page.fill('input[type="text"]', USERNAME)
    except Exception:
        page.fill('input', USERNAME)
    try:
        page.fill('input[type="password"]', PASSWORD)
    except Exception:
        pass

    # 点击登录按钮
    for btn_sel in ['button:has-text("Sign in")', 'button:has-text("Login")',
                    'button[type="submit"]']:
        try:
            btn = page.locator(btn_sel)
            if btn.count() > 0:
                btn.first.click()
                break
        except Exception:
            pass

    try:
        page.wait_for_url(lambda u: "login" not in u, timeout=15000)
    except Exception:
        pass
    page.wait_for_load_state("domcontentloaded", timeout=10000)
    time.sleep(2)
    log(f"  [login] Current URL: {page.url}")


def navigate_to_bacnet(page: Page) -> None:
    """按 conftest.py 的正确选择器导航到 BACnet/IP"""
    # 点击顶部 AcuHMI-1-7
    hmi_nav = page.locator('.nav-item:has-text("AcuHMI-1-7")')
    if hmi_nav.count() > 0:
        hmi_nav.first.click()
        time.sleep(1.5)
        log("  [nav] clicked .nav-item AcuHMI-1-7")
    else:
        log("  [nav] .nav-item:has-text('AcuHMI-1-7') not found, trying fallback")
        # 回退：找任何含 AcuHMI-1-7 的元素
        page.evaluate(
            """() => {
                for (const el of document.querySelectorAll('*')) {
                    if (el.children.length === 0 && el.textContent.trim() === 'AcuHMI-1-7') {
                        el.click(); return true;
                    }
                }
            }"""
        )
        time.sleep(1.5)

    # 点击 Protocols
    protocols = page.locator('.left-nav-item:has-text("Protocols")')
    if protocols.count() > 0:
        protocols.first.click()
        time.sleep(1)
        page.wait_for_load_state("domcontentloaded", timeout=8000)
        log("  [nav] clicked .left-nav-item Protocols")
    else:
        log("  [nav] .left-nav-item:has-text('Protocols') not found, trying fallback")
        # 打印所有左侧导航项
        nav_items = page.evaluate(
            """() => {
                const items = document.querySelectorAll(
                    '.left-nav-item, .nav-item, [class*="nav"] li, aside li, .el-menu-item'
                );
                return Array.from(items).map(el => ({
                    tag: el.tagName,
                    text: el.textContent.trim().substring(0, 50),
                    class: el.className,
                })).filter(i => i.text);
            }"""
        )
        log(f"  [nav] nav items found: {json.dumps(nav_items[:20], ensure_ascii=False)}")
        page.evaluate(
            """() => {
                for (const el of document.querySelectorAll('*')) {
                    if (el.children.length <= 1 && el.textContent.trim() === 'Protocols') {
                        el.click(); return true;
                    }
                }
            }"""
        )
        time.sleep(1)

    # 点击 BACnet/IP
    bacnet = page.locator('li:has-text("BACnet/IP")')
    if bacnet.count() > 0:
        bacnet.first.click()
        time.sleep(2)
        page.wait_for_load_state("domcontentloaded", timeout=8000)
        log("  [nav] clicked li BACnet/IP")
    else:
        log("  [nav] li:has-text('BACnet/IP') not found, trying fallback")
        page.evaluate(
            """() => {
                for (const el of document.querySelectorAll('li, a, span')) {
                    if (el.textContent.trim() === 'BACnet/IP') { el.click(); return; }
                }
            }"""
        )
        time.sleep(2)

    log(f"  [nav] current URL after navigation: {page.url}")


# ─────────────────────────────────────────────────────────────────
# 阶段 1：设备列表
# ─────────────────────────────────────────────────────────────────
def phase1_device_list(page: Page) -> list:
    log("\n=== 阶段1：设备列表 ===")
    ss(page, "phase1_device_list.png")

    # 读取设备列表行
    rows_info = page.evaluate(
        """() => {
            const rows = document.querySelectorAll('.el-table__body tr.el-table__row');
            return Array.from(rows).map((row, i) => {
                const cells = row.querySelectorAll('td');
                return {
                    index: i,
                    cells: Array.from(cells).map(c => c.textContent.trim().substring(0, 30)),
                    buttonTexts: Array.from(row.querySelectorAll('button'))
                        .map(b => b.textContent.trim()),
                    linkTexts: Array.from(row.querySelectorAll('a, .el-button'))
                        .map(a => a.textContent.trim()),
                };
            });
        }"""
    )
    log(f"  设备行数: {len(rows_info)}")
    for r in rows_info:
        log(f"  Row[{r['index']}]: cells={r['cells']}")
        log(f"           buttons={r['buttonTexts']}, links={r['linkTexts']}")

    # 同时检查页面上所有可见按钮（找 Parameter 相关）
    all_btns = page.evaluate(
        """() => {
            return Array.from(document.querySelectorAll('button'))
                .filter(b => b.offsetParent !== null)
                .map(b => b.textContent.trim())
                .filter(t => t.length > 0);
        }"""
    )
    log(f"  页面可见按钮: {all_btns}")
    return rows_info


# ─────────────────────────────────────────────────────────────────
# 阶段 2：打开 2100 的 Parameter Config
# ─────────────────────────────────────────────────────────────────
def phase2_open_param_config(page: Page) -> bool:
    log("\n=== 阶段2：打开 2100 Parameter Config ===")

    # 找到 2100 行，检查行内所有可交互元素
    row_elems = page.evaluate(
        """() => {
            const rows = document.querySelectorAll('.el-table__body tr.el-table__row');
            for (const row of rows) {
                if (!row.textContent.includes('2100')) continue;
                return {
                    rowText: row.textContent.trim().substring(0, 100),
                    buttons: Array.from(row.querySelectorAll('button')).map(b => ({
                        text: b.textContent.trim(), class: b.className, disabled: b.disabled
                    })),
                    links: Array.from(row.querySelectorAll('a')).map(a => ({
                        text: a.textContent.trim(), href: a.href
                    })),
                    elButtons: Array.from(row.querySelectorAll('.el-button')).map(b => ({
                        text: b.textContent.trim(), class: b.className
                    })),
                    icons: Array.from(row.querySelectorAll('[class*="icon"], i, svg')).map(i => ({
                        class: i.className, title: i.getAttribute('title') || ''
                    })).slice(0, 10),
                    // clickable spans
                    spans: Array.from(row.querySelectorAll('span[class*="btn"], span[class*="action"]'))
                        .map(s => ({text: s.textContent.trim(), class: s.className})),
                };
            }
            return null;
        }"""
    )
    log(f"  2100 行元素: {json.dumps(row_elems, ensure_ascii=False, indent=2)}")

    # 尝试通过 _open_param_dialog 逻辑（来自 test_bacnet_ui_basic）
    # 首先检查行内有没有 "Parameter" 文字的按钮
    opened = page.evaluate(
        """() => {
            const rows = document.querySelectorAll('.el-table__body tr.el-table__row');
            for (const row of rows) {
                if (!row.textContent.includes('2100')) continue;
                // 优先找含 Parameter/Config/Param 的按钮
                for (const btn of row.querySelectorAll('button, .el-button, a')) {
                    const t = btn.textContent.trim();
                    if (t.includes('Param') || t.includes('Config') || t.includes('设置')) {
                        btn.click();
                        return 'button: ' + t;
                    }
                }
                // 找最后一个按钮（通常是 Parameter Config 按钮）
                const btns = Array.from(row.querySelectorAll('button, .el-button'));
                if (btns.length > 0) {
                    const last = btns[btns.length - 1];
                    last.click();
                    return 'last button: ' + last.textContent.trim();
                }
                // 没有按钮，找 td 中的可点击元素
                const cells = row.querySelectorAll('td');
                const lastCell = cells[cells.length - 1];
                if (lastCell) {
                    const clickable = lastCell.querySelector('[class*="btn"], [class*="icon"], span, a');
                    if (clickable) {
                        clickable.click();
                        return 'clickable in last cell: ' + clickable.className;
                    }
                }
                return 'no clickable in 2100 row';
            }
            return '2100 row not found';
        }"""
    )
    log(f"  点击结果: {opened}")
    time.sleep(2.5)
    ss(page, "phase2_after_click.png")

    # 检查是否有弹窗打开
    dialog_found = page.evaluate(
        """() => {
            // 检查多种弹窗选择器
            const selectors = [
                '[aria-label="Parameter Config"]',
                '.el-dialog',
                '[role="dialog"]',
                '.el-overlay',
            ];
            const results = {};
            for (const sel of selectors) {
                const el = document.querySelector(sel);
                results[sel] = el ? {
                    found: true,
                    ariaLabel: el.getAttribute('aria-label') || '',
                    visible: el.offsetParent !== null || el.style.display !== 'none',
                    text: el.textContent.trim().substring(0, 100),
                } : {found: false};
            }
            return results;
        }"""
    )
    log(f"  弹窗检测: {json.dumps(dialog_found, ensure_ascii=False, indent=2)}")

    # 获取弹窗详细信息
    param_dlg = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Parameter Config"]');
            if (!dlg) {
                // 尝试找任何 el-dialog
                const anyDlg = document.querySelector('.el-dialog');
                if (!anyDlg) return {found: false, anyDialog: false};
                return {
                    found: false,
                    anyDialog: true,
                    dialogTitle: (anyDlg.querySelector('.el-dialog__title') || {}).textContent || '',
                    dialogButtons: Array.from(anyDlg.querySelectorAll('button'))
                        .map(b => b.textContent.trim()),
                };
            }
            const headers = Array.from(dlg.querySelectorAll('.el-table__header th'))
                .map(th => th.textContent.trim()).filter(t => t);
            const buttons = Array.from(dlg.querySelectorAll('button'))
                .map(b => ({text: b.textContent.trim(), class: b.className, disabled: b.disabled}));
            return {found: true, headers, buttons};
        }"""
    )
    log(f"  Parameter Config 弹窗详情: {json.dumps(param_dlg, ensure_ascii=False, indent=2)}")
    return param_dlg.get("found", False)


# ─────────────────────────────────────────────────────────────────
# 从 test_bacnet_ui_basic 复用的 _open_param_dialog 逻辑
# ─────────────────────────────────────────────────────────────────
def open_param_dialog_robust(page: Page, keyword: str = "2100") -> bool:
    """完整的 Parameter Config 打开逻辑（来自 helpers）"""
    log(f"\n--- 尝试用健壮逻辑打开 {keyword} 的 Parameter Config ---")

    # 读取辅助函数（来自 test_bacnet_ui_basic）
    sys.path.insert(0, "C:/Users/ZihanGao/Desktop/testing-team")
    try:
        from test_case.AcuHMI_1_7.bacnet_ui.test_bacnet_ui_basic import _open_param_dialog
        result = _open_param_dialog(page, keyword)
        log(f"  _open_param_dialog result: {result}")
        return result
    except Exception as e:
        log(f"  import _open_param_dialog failed: {e}")
        return False


# ─────────────────────────────────────────────────────────────────
# 阶段 3：Polling Enable 状态
# ─────────────────────────────────────────────────────────────────
def phase3_check_polling(page: Page) -> None:
    log("\n=== 阶段3：检查并开启 Polling Enable ===")
    row_state = page.evaluate(
        """() => {
            const rows = document.querySelectorAll(
                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
            );
            if (rows.length === 0) return {found: false, rowCount: 0};
            const row = rows[0];
            const cells = row.querySelectorAll('td');
            const paramName = cells[0] ? cells[0].textContent.trim() : '';

            // cells[1] = 第2列，可能是 Polling Enable 或 EPICS Enable
            const sw1 = cells[1] ? cells[1].querySelector('.el-switch__input') : null;
            const sw2 = cells[2] ? cells[2].querySelector('.el-switch__input') : null;
            const sw3 = cells[3] ? cells[3].querySelector('.el-switch__input') : null;

            return {
                found: true,
                rowCount: rows.length,
                paramName,
                cellCount: cells.length,
                sw1: sw1 ? sw1.getAttribute('aria-checked') : 'no switch',
                sw2: sw2 ? sw2.getAttribute('aria-checked') : 'no switch',
                sw3: sw3 ? sw3.getAttribute('aria-checked') : 'no switch',
                // cells[4] = COV Increment?
                inp4: cells[4] ? cells[4].querySelector('input')?.value : 'no cell',
            };
        }"""
    )
    log(f"  行状态: {json.dumps(row_state, ensure_ascii=False)}")

    if row_state.get('found') and row_state.get('sw1') != 'true':
        log("  sw1 (cells[1]) = false, 开启...")
        page.evaluate(
            """() => {
                const rows = document.querySelectorAll(
                    '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                );
                if (!rows.length) return;
                const cells = rows[0].querySelectorAll('td');
                const sw = cells[1] && cells[1].querySelector('.el-switch__input');
                if (sw) sw.click();
            }"""
        )
        time.sleep(0.6)
        ss(page, "phase3_sw1_enabled.png")


# ─────────────────────────────────────────────────────────────────
# 阶段 4：打开 COV Batch Update
# ─────────────────────────────────────────────────────────────────
def phase4_open_batch_update(page: Page) -> bool:
    log("\n=== 阶段4：打开 COV Batch Update ===")
    ss(page, "phase4_before_batch_click.png")

    # 先列出弹窗内所有按钮
    all_buttons = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Parameter Config"]');
            if (!dlg) return [];
            return Array.from(dlg.querySelectorAll('button')).map(b => ({
                text: b.textContent.trim(),
                class: b.className,
                disabled: b.disabled,
                visible: b.offsetParent !== null,
            }));
        }"""
    )
    log(f"  Parameter Config 弹窗内按钮: {json.dumps(all_buttons, ensure_ascii=False, indent=2)}")

    # 找并点击 COV Batch Update
    clicked = page.evaluate(
        """() => {
            for (const btn of document.querySelectorAll('button')) {
                const t = btn.textContent.trim();
                if (t.includes('COV Batch') || t.includes('Batch Update')) {
                    btn.click();
                    return 'clicked: ' + t;
                }
            }
            return 'not found';
        }"""
    )
    log(f"  点击结果: {clicked}")
    time.sleep(2)
    ss(page, "phase4_batch_update_opened.png")

    batch_info = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Batch Update"]');
            if (!dlg) {
                // 找所有 el-dialog
                const dialogs = Array.from(document.querySelectorAll('[role="dialog"], .el-dialog'))
                    .map(d => ({
                        ariaLabel: d.getAttribute('aria-label') || '',
                        title: (d.querySelector('.el-dialog__title') || {}).textContent || '',
                        visible: d.offsetParent !== null,
                    }));
                return {found: false, allDialogs: dialogs};
            }
            const buttons = Array.from(dlg.querySelectorAll('button')).map(b => ({
                text: b.textContent.trim(), class: b.className, disabled: b.disabled
            }));
            const inputs = Array.from(dlg.querySelectorAll('input')).map(inp => ({
                type: inp.type, value: inp.value, placeholder: inp.placeholder,
                disabled: inp.disabled, class: inp.className,
            }));
            const selects = Array.from(dlg.querySelectorAll('.el-select, .el-select__wrapper')).map(sel => {
                const inp = sel.querySelector('input, .el-select__input');
                return {
                    class: sel.className,
                    ariaExpanded: inp ? inp.getAttribute('aria-expanded') : null,
                    ariaControls: inp ? inp.getAttribute('aria-controls') : null,
                    currentValue: (sel.querySelector('.el-select__selected-item, .el-input__inner') || {}).textContent || '',
                };
            });
            return {found: true, buttons, inputs, selects};
        }"""
    )
    log(f"  Batch Update: {json.dumps(batch_info, ensure_ascii=False, indent=2)}")
    return batch_info.get("found", False)


# ─────────────────────────────────────────────────────────────────
# 阶段 5：展开 Select parameters
# ─────────────────────────────────────────────────────────────────
def phase5_select_parameters(page: Page) -> None:
    log("\n=== 阶段5：展开 Select parameters ===")

    # 用 Playwright click（不用 JS）触发 select
    select_input = page.locator('[aria-label="Batch Update"] .el-select__input')
    if select_input.count() > 0:
        log(f"  找到 .el-select__input, 数量={select_input.count()}")
        select_input.first.click()
    else:
        # 尝试找 .el-select 本身
        sel_wrap = page.locator('[aria-label="Batch Update"] .el-select')
        if sel_wrap.count() > 0:
            log(f"  找到 .el-select, 点击之")
            sel_wrap.first.click()
        else:
            log("  未找到 select，尝试 JS click")
            page.evaluate(
                """() => {
                    const dlg = document.querySelector('[aria-label="Batch Update"]');
                    if (!dlg) return;
                    const inp = dlg.querySelector('input');
                    if (inp) inp.click();
                }"""
            )

    # 等待展开
    for i in range(15):
        expanded = page.evaluate(
            """() => {
                const dlg = document.querySelector('[aria-label="Batch Update"]');
                if (!dlg) return false;
                const inp = dlg.querySelector('.el-select__input, input');
                return inp ? (inp.getAttribute('aria-expanded') === 'true') : false;
            }"""
        )
        if expanded:
            log(f"  Select 在第 {i+1} 次等待（{(i+1)*300}ms）后展开")
            break
        time.sleep(0.3)
    else:
        log("  Select 展开超时（15次 × 300ms = 4500ms）")

    ss(page, "phase5_select_expanded.png")

    # 读取下拉选项（用 aria-controls 精确定位）
    dropdown = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Batch Update"]');
            if (!dlg) return {error: 'no dialog'};
            const inp = dlg.querySelector('.el-select__input, input');
            if (!inp) return {error: 'no input'};
            const ariaExpanded = inp.getAttribute('aria-expanded');
            const listId = inp.getAttribute('aria-controls');
            if (!listId) {
                // 全局找展开的 popper
                const poppers = document.querySelectorAll('.el-select-dropdown');
                return {
                    ariaExpanded,
                    listId: null,
                    allPoppers: Array.from(poppers).map(p => ({
                        id: p.id, visible: p.offsetParent !== null,
                        items: Array.from(p.querySelectorAll('.el-select-dropdown__item'))
                            .map(i => i.textContent.trim()).slice(0, 10),
                    })),
                };
            }
            const list = document.getElementById(listId);
            if (!list) return {ariaExpanded, listId, listFound: false};
            const items = Array.from(list.querySelectorAll('.el-select-dropdown__item'))
                .map(el => ({
                    text: el.textContent.trim(),
                    disabled: el.classList.contains('is-disabled'),
                    selected: el.classList.contains('is-selected'),
                }));
            return {ariaExpanded, listId, listFound: true, itemCount: items.length, items};
        }"""
    )
    log(f"  下拉内容: {json.dumps(dropdown, ensure_ascii=False, indent=2)}")

    # 如果有选项，选第一个；如果没有，检查 Select All
    items = dropdown.get('items', [])
    if items:
        log(f"  有 {len(items)} 个选项，选第一个: {items[0].get('text', '?')}")
        page.evaluate(
            """() => {
                const dlg = document.querySelector('[aria-label="Batch Update"]');
                const inp = dlg && dlg.querySelector('.el-select__input, input');
                const listId = inp && inp.getAttribute('aria-controls');
                if (listId) {
                    const list = document.getElementById(listId);
                    const first = list && list.querySelector(
                        '.el-select-dropdown__item:not(.is-disabled)');
                    if (first) { first.click(); return; }
                }
                // fallback: 全局找
                const any = document.querySelector(
                    '.el-select-dropdown__item:not(.is-disabled)');
                if (any) any.click();
            }"""
        )
        time.sleep(0.5)
        ss(page, "phase5_first_param_selected.png")
    else:
        log("  下拉为空！检查是否有 Select All 选项或其他UI元素")
        # 截图后查看完整 DOM
        ss(page, "phase5_empty_dropdown.png")

        # 打印 Batch Update 弹窗完整 DOM（前 3000 chars）
        dlg_html = page.evaluate(
            """() => {
                const dlg = document.querySelector('[aria-label="Batch Update"]');
                return dlg ? dlg.innerHTML.substring(0, 3000) : 'no dialog';
            }"""
        )
        log(f"  Batch Update DOM: {dlg_html}")


# ─────────────────────────────────────────────────────────────────
# 阶段 6：设置 COV Increment 并点击 Confirm
# ─────────────────────────────────────────────────────────────────
def phase6_set_cov_and_confirm(page: Page) -> None:
    log("\n=== 阶段6：设置 COV Increment = 1.500 ===")

    # 检查所有 input 状态
    inp_state = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Batch Update"]');
            if (!dlg) return {error: 'no dialog'};
            return Array.from(dlg.querySelectorAll('input')).map(inp => ({
                type: inp.type, value: inp.value, placeholder: inp.placeholder,
                disabled: inp.disabled, readOnly: inp.readOnly, class: inp.className,
            }));
        }"""
    )
    log(f"  Batch Update 输入框: {json.dumps(inp_state, ensure_ascii=False, indent=2)}")

    # 设置 COV Increment（多种方式）
    set_result = page.evaluate(
        """(val) => {
            const dlg = document.querySelector('[aria-label="Batch Update"]');
            if (!dlg) return 'no dialog';
            const inputs = Array.from(dlg.querySelectorAll('input'));
            // 方式1：placeholder 含 cov/increment
            for (const inp of inputs) {
                const ph = (inp.getAttribute('placeholder') || '').toLowerCase();
                if (ph.includes('cov') || ph.includes('increment')) {
                    inp.removeAttribute('disabled');
                    inp.focus();
                    inp.value = val;
                    inp.dispatchEvent(new Event('input', {bubbles: true}));
                    inp.dispatchEvent(new Event('change', {bubbles: true}));
                    inp.dispatchEvent(new Event('blur', {bubbles: true}));
                    return 'set via placeholder: ' + ph;
                }
            }
            // 方式2：第一个非 hidden、非 readonly 的 input
            for (const inp of inputs) {
                if (inp.type !== 'hidden' && !inp.readOnly) {
                    inp.removeAttribute('disabled');
                    inp.focus();
                    inp.value = val;
                    inp.dispatchEvent(new Event('input', {bubbles: true}));
                    inp.dispatchEvent(new Event('change', {bubbles: true}));
                    inp.dispatchEvent(new Event('blur', {bubbles: true}));
                    return 'set via fallback type=' + inp.type + ' ph=' + inp.placeholder;
                }
            }
            return 'no writable input';
        }""",
        "1.500",
    )
    log(f"  设置结果: {set_result}")
    time.sleep(0.3)
    ss(page, "phase6_cov_set.png")

    # 读取所有按钮
    btns = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Batch Update"]');
            if (!dlg) return [];
            return Array.from(dlg.querySelectorAll('button')).map(b => ({
                text: b.textContent.trim(), class: b.className, disabled: b.disabled,
            }));
        }"""
    )
    log(f"  Batch Update 按钮: {json.dumps(btns, ensure_ascii=False, indent=2)}")

    # Playwright 原生点击 primary 按钮
    batch = page.locator('[aria-label="Batch Update"]')
    primary = batch.locator("button.el-button--primary")
    log(f"  primary 按钮数量: {primary.count()}")

    if primary.count() > 0:
        log("  点击 el-button--primary ...")
        primary.first.click()
    else:
        for txt in ["Confirm", "确定", "Apply", "OK", "Save"]:
            btn = batch.locator(f"button").filter(has_text=txt)
            if btn.count() > 0:
                log(f"  点击 {txt} ...")
                btn.first.click()
                break
        else:
            log("  未找到任何操作按钮！")

    time.sleep(2.5)
    ss(page, "phase6_after_confirm.png")

    batch_closed = not page.evaluate(
        "() => !!document.querySelector('[aria-label=\"Batch Update\"]')"
    )
    param_open = page.evaluate(
        "() => !!document.querySelector('[aria-label=\"Parameter Config\"]')"
    )
    log(f"  Batch Update 已关闭: {batch_closed}")
    log(f"  Parameter Config 仍打开: {param_open}")


# ─────────────────────────────────────────────────────────────────
# 阶段 7：检查表格后的 COV Increment 值
# ─────────────────────────────────────────────────────────────────
def phase7_check_table(page: Page) -> None:
    log("\n=== 阶段7：Confirm 后查看 Parameter Config 表格 ===")
    ss(page, "phase7_table_after_confirm.png")
    table_state = page.evaluate(
        """() => {
            const rows = document.querySelectorAll(
                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
            );
            return Array.from(rows).slice(0, 5).map((row, i) => {
                const cells = row.querySelectorAll('td');
                return {
                    index: i,
                    paramName: cells[0] ? cells[0].textContent.trim() : '',
                    cells: Array.from(cells).map((c, ci) => {
                        const sw = c.querySelector('.el-switch__input');
                        const inp = c.querySelector('input');
                        return {
                            ci,
                            text: c.textContent.trim().substring(0, 30),
                            switch: sw ? sw.getAttribute('aria-checked') : null,
                            inputValue: inp ? inp.value : null,
                            inputDisabled: inp ? inp.disabled : null,
                        };
                    }),
                };
            });
        }"""
    )
    log(f"  表格前5行:")
    for row in table_state:
        log(f"    [{row['index']}] {row['paramName']}")
        for cell in row['cells']:
            if cell.get('switch') is not None or cell.get('inputValue') is not None:
                log(f"      col[{cell['ci']}]: switch={cell['switch']}, "
                    f"input={cell['inputValue']}, disabled={cell['inputDisabled']}")


# ─────────────────────────────────────────────────────────────────
# 阶段 8：寻找 Save 按钮
# ─────────────────────────────────────────────────────────────────
def phase8_find_save(page: Page) -> None:
    log("\n=== 阶段8：列出所有 Save 按钮 ===")
    save_info = page.evaluate(
        """() => {
            const result = {
                inParamDlg: [],
                mainPage: [],
                allVisible: [],
            };
            for (const btn of document.querySelectorAll('button')) {
                const t = btn.textContent.trim();
                if (!t) continue;
                const visible = btn.offsetParent !== null;
                const inParamDlg = !!btn.closest('[aria-label="Parameter Config"]');
                const inBatchDlg = !!btn.closest('[aria-label="Batch Update"]');
                const info = {text: t, class: btn.className, disabled: btn.disabled,
                              visible, inParamDlg, inBatchDlg};
                if (t === 'Save') {
                    if (inParamDlg) result.inParamDlg.push(info);
                    else result.mainPage.push(info);
                }
                if (visible) result.allVisible.push(info);
            }
            return result;
        }"""
    )
    log(f"  Parameter Config 弹窗内 Save 按钮: {json.dumps(save_info['inParamDlg'], ensure_ascii=False)}")
    log(f"  主页面 Save 按钮: {json.dumps(save_info['mainPage'], ensure_ascii=False)}")
    log(f"  所有可见按钮 (前15): {json.dumps(save_info['allVisible'][:15], ensure_ascii=False)}")
    ss(page, "phase8_save_buttons.png")


# ─────────────────────────────────────────────────────────────────
# 阶段 9：点击 Save
# ─────────────────────────────────────────────────────────────────
def phase9_click_save(page: Page) -> None:
    log("\n=== 阶段9：点击 Save ===")

    # 先尝试弹窗内 Save
    in_dlg = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Parameter Config"]');
            if (!dlg) return false;
            for (const btn of dlg.querySelectorAll('button')) {
                if (btn.textContent.trim() === 'Save' && !btn.disabled) {
                    btn.click();
                    return true;
                }
            }
            return false;
        }"""
    )
    log(f"  弹窗内 Save 点击: {in_dlg}")

    if not in_dlg:
        log("  弹窗内无 Save，用 Playwright force click 主区域 Save...")
        save_btn = page.locator('button:has-text("Save")')
        cnt = save_btn.count()
        log(f"  主区域 Save 按钮数量: {cnt}")
        if cnt > 0:
            save_btn.first.click(force=True)
        else:
            page.evaluate(
                """() => {
                    for (const btn of document.querySelectorAll('button')) {
                        if (btn.textContent.trim() === 'Save') { btn.click(); return; }
                    }
                }"""
            )

    time.sleep(2.5)
    ss(page, "phase9_after_save.png")

    toast = page.evaluate(
        """() => {
            const msg = document.querySelector('.el-message');
            const notif = document.querySelector('.el-notification');
            return {
                message: msg ? {text: msg.textContent.trim(), class: msg.className} : null,
                notification: notif ? {text: notif.textContent.trim().substring(0, 100)} : null,
            };
        }"""
    )
    log(f"  Toast: {json.dumps(toast, ensure_ascii=False)}")


# ─────────────────────────────────────────────────────────────────
# 阶段 10：关闭弹窗、重新导航、验证
# ─────────────────────────────────────────────────────────────────
def phase10_verify(page: Page) -> None:
    log("\n=== 阶段10：关闭弹窗 + 重新导航 + 验证 ===")

    # 关闭 Parameter Config 弹窗
    closed = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Parameter Config"]');
            if (!dlg) return 'no dialog open';
            // 找 Close/Cancel/× 按钮
            for (const btn of dlg.querySelectorAll('button')) {
                const t = btn.textContent.trim();
                if (t === 'Close' || t === '关闭' || t === 'Cancel') {
                    btn.click();
                    return 'closed via: ' + t;
                }
            }
            const headerbtn = dlg.querySelector('.el-dialog__headerbtn');
            if (headerbtn) { headerbtn.click(); return 'closed via headerbtn'; }
            return 'no close button';
        }"""
    )
    log(f"  关闭弹窗: {closed}")
    time.sleep(1.5)

    # 重新导航
    navigate_to_bacnet(page)
    ss(page, "phase10_after_navigate.png")

    # 重新打开 2100 Parameter Config
    opened = open_param_dialog_robust(page, "2100")
    time.sleep(2)
    ss(page, "phase10_param_config_reopened.png")

    # 验证 COV Increment
    result = page.evaluate(
        """() => {
            const rows = document.querySelectorAll(
                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
            );
            return Array.from(rows).slice(0, 5).map((row, i) => {
                const cells = row.querySelectorAll('td');
                const covInp = cells[3] ? cells[3].querySelector('input') : null;
                return {
                    i,
                    paramName: cells[0] ? cells[0].textContent.trim() : '',
                    covIncrement: covInp ? covInp.value : 'no input',
                };
            });
        }"""
    )
    log(f"\n  === 最终验证（重新打开后 COV Increment 值）===")
    for r in result:
        is_target = "1.500" in str(r.get('covIncrement', ''))
        marker = " <-- TARGET" if is_target else ""
        log(f"    [{r['i']}] {r['paramName']}: COV Increment = {r['covIncrement']}{marker}")


# ─────────────────────────────────────────────────────────────────
# 主入口
# ─────────────────────────────────────────────────────────────────
def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,  # 无头模式，输出结果到控制台
            args=["--ignore-certificate-errors"],
        )
        ctx = browser.new_context(ignore_https_errors=True)
        page = ctx.new_page()
        page.set_default_timeout(15000)

        try:
            login(page)
            navigate_to_bacnet(page)
            ss(page, "phase0_bacnet_page.png")

            # 阶段1
            phase1_device_list(page)

            # 阶段2 - 打开 Parameter Config
            param_opened = open_param_dialog_robust(page, "2100")
            if not param_opened:
                # 手动尝试
                param_opened = phase2_open_param_config(page)

            if not param_opened:
                log("ERROR: Parameter Config 弹窗未能打开")
                # 打印完整页面 DOM 片段供分析
                dom_snap = page.evaluate(
                    """() => document.body.innerHTML.substring(0, 5000)"""
                )
                log(f"  页面 DOM 片段: {dom_snap}")
                return

            ss(page, "phase2_dialog_opened.png")

            # 阶段3
            phase3_check_polling(page)

            # 阶段4
            batch_ok = phase4_open_batch_update(page)
            if not batch_ok:
                log("ERROR: Batch Update 弹窗未找到，检查截图")
                return

            # 阶段5
            phase5_select_parameters(page)

            # 阶段6
            phase6_set_cov_and_confirm(page)

            # 阶段7
            phase7_check_table(page)

            # 阶段8
            phase8_find_save(page)

            # 阶段9
            phase9_click_save(page)

            # 阶段10
            phase10_verify(page)

        except Exception as e:
            import traceback
            log(f"\n[EXCEPTION] {e}")
            log(traceback.format_exc())
            try:
                ss(page, "exception_state.png")
            except Exception:
                pass
        finally:
            browser.close()

    log(f"\n\n=== 全部截图 ===")
    log(f"目录: {SCREENSHOTS_DIR}")
    for f in sorted(SCREENSHOTS_DIR.glob("phase*.png")):
        log(f"  {f}")


if __name__ == "__main__":
    main()
