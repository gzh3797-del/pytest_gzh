# -*- coding: utf-8 -*-
"""
debug_cov_fix_verify.py
端到端验证修复后的完整保存路径：
  1. 打开 2100 Parameter Config
  2. 开启 Polling Enable
  3. 打开 Batch Update，选参数，设置 COV Increment = 1.500，Confirm
  4. 弹窗内 Save
  5. 重新导航，重新打开验证 COV Increment = 1.500
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, Page

sys.path.insert(0, "C:/Users/ZihanGao/Desktop/testing-team")

SCREENSHOTS_DIR = Path(
    "C:/Users/ZihanGao/Desktop/testing-team/test_case/AcuHMI_1_7/bacnet_ui/screenshots"
)
HMI_URL = "https://192.168.2.8"
USERNAME = "q"
PASSWORD = "1"


def ss(page: Page, name: str) -> None:
    p = str(SCREENSHOTS_DIR / name)
    page.screenshot(path=p)
    print(f"  [ss] {p}")


def login_and_nav_to_bacnet(page: Page) -> None:
    page.goto(f"{HMI_URL}/#/login", wait_until="domcontentloaded", timeout=20000)
    page.wait_for_load_state("domcontentloaded", timeout=15000)
    time.sleep(1)
    page.fill('input[type="text"]', USERNAME)
    page.fill('input[type="password"]', PASSWORD)
    page.click('button:has-text("Sign in")')
    try:
        page.wait_for_url(lambda u: "login" not in u, timeout=15000)
    except Exception:
        pass
    time.sleep(2)

    hmi_nav = page.locator('.nav-item:has-text("AcuHMI-1-7")')
    if hmi_nav.count() > 0:
        hmi_nav.first.click()
        time.sleep(1.5)
    protocols = page.locator('.left-nav-item:has-text("Protocols")')
    if protocols.count() > 0:
        protocols.first.click()
        time.sleep(1)
    bacnet = page.locator('li:has-text("BACnet/IP")')
    if bacnet.count() > 0:
        bacnet.first.click()
        time.sleep(2)
    print(f"  [nav] URL: {page.url}")


def main() -> None:
    from test_case.AcuHMI_1_7.bacnet_ui.test_bacnet_ui_basic import (
        _open_param_dialog,
        _close_param_config_dialog,
    )
    from test_case.AcuHMI_1_7.bacnet_ui.test_bacnet_ui_config import (
        _open_cov_batch_update,
        _batch_update_select_all_params,
        _set_batch_update_cov_increment,
        _confirm_batch_update_playwright,
        _navigate_to_bacnet,
        _get_cov_increment_value_in_dialog,
        _dismiss_blocking_overlays,
    )

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--ignore-certificate-errors"])
        ctx = browser.new_context(ignore_https_errors=True)
        page = ctx.new_page()
        page.set_default_timeout(15000)

        try:
            login_and_nav_to_bacnet(page)
            ss(page, "verify_00_bacnet.png")

            # Step 1: 打开 2100 Parameter Config
            opened = _open_param_dialog(page, "2100")
            print(f"  [1] Parameter Config 打开: {opened}")
            assert opened, "Parameter Config 未打开"
            time.sleep(1)
            ss(page, "verify_01_param_config.png")

            # Step 2: 确保第一行 Polling Enable = true
            page.evaluate(
                """() => {
                    const rows = document.querySelectorAll(
                        '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                    );
                    if (!rows.length) return;
                    const cells = rows[0].querySelectorAll('td');
                    const sw = cells[1] && cells[1].querySelector('.el-switch__input');
                    if (sw && sw.getAttribute('aria-checked') !== 'true') sw.click();
                }"""
            )
            time.sleep(0.5)

            # Step 3: 打开 Batch Update
            batch_ok = _open_cov_batch_update(page)
            print(f"  [2] Batch Update 打开: {batch_ok}")
            assert batch_ok, "Batch Update 未打开"

            # Step 4: 选择一个参数（选第一个可用参数）
            # 用 Playwright locator 精确点击展开 select（避免 JS click 展开不稳定）
            sel_input = page.locator('[aria-label="Batch Update"] .el-select__input')
            assert sel_input.count() > 0, "未找到 .el-select__input"
            sel_input.first.click()
            time.sleep(0.5)

            # 等待展开
            for _ in range(10):
                expanded = page.evaluate(
                    """() => {
                        const dlg = document.querySelector('[aria-label="Batch Update"]');
                        const inp = dlg && dlg.querySelector('.el-select__input');
                        return inp ? inp.getAttribute('aria-expanded') === 'true' : false;
                    }"""
                )
                if expanded:
                    break
                time.sleep(0.3)

            # 选第一个参数
            first_param_name = page.evaluate(
                """() => {
                    const dlg = document.querySelector('[aria-label="Batch Update"]');
                    const inp = dlg && dlg.querySelector('.el-select__input');
                    const listId = inp && inp.getAttribute('aria-controls');
                    const list = listId && document.getElementById(listId);
                    if (!list) return null;
                    const first = list.querySelector('.el-select-dropdown__item:not(.is-disabled)');
                    if (first) { first.click(); return first.textContent.trim(); }
                    return null;
                }"""
            )
            print(f"  [3] 选中参数: {first_param_name!r}")
            assert first_param_name, "未能选中任何参数"
            time.sleep(0.3)

            # Step 5: 设置 COV Increment = 1.500（使用修复后函数）
            _set_batch_update_cov_increment(page, "1.500")
            time.sleep(0.3)
            ss(page, "verify_02_cov_set.png")

            # 验证输入框值
            cov_val = page.evaluate(
                """() => {
                    const dlg = document.querySelector('[aria-label="Batch Update"]');
                    const inp = dlg && dlg.querySelector('input.el-input__inner');
                    return inp ? inp.value : null;
                }"""
            )
            print(f"  [4] COV Increment 输入值: {cov_val!r}")
            assert cov_val == "1.500", f"COV Increment 未设置为 1.500，当前 {cov_val!r}"

            # Step 6: 点击 Confirm（Playwright 原生 click）
            _confirm_batch_update_playwright(page)
            time.sleep(2.5)
            ss(page, "verify_03_after_confirm.png")

            batch_closed = not page.evaluate(
                "() => !!document.querySelector('[aria-label=\"Batch Update\"]')"
            )
            print(f"  [5] Batch Update 关闭: {batch_closed}")
            assert batch_closed, "Batch Update 未关闭，Confirm 失败（可能有验证错误）"

            # 检查表格中的 COV Increment 值
            cov_in_table = page.evaluate(
                """(name) => {
                    const rows = document.querySelectorAll(
                        '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
                    );
                    for (const row of rows) {
                        const cells = row.querySelectorAll('td');
                        if (cells[0] && cells[0].textContent.trim() === name) {
                            const inp = cells[3] && cells[3].querySelector('input');
                            return inp ? inp.value : 'no input';
                        }
                    }
                    return 'not found in current page';
                }""",
                first_param_name,
            )
            print(f"  [5] 表格中 {first_param_name!r} COV Increment: {cov_in_table!r}")

            # Step 7: 点击 Parameter Config 弹窗内的 Save
            saved = page.evaluate(
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
            print(f"  [6] 弹窗内 Save 点击: {saved}")
            time.sleep(2.5)
            ss(page, "verify_04_after_save.png")

            toast = page.evaluate(
                """() => {
                    const msg = document.querySelector('.el-message');
                    return msg ? msg.textContent.trim() : null;
                }"""
            )
            print(f"  [6] Toast: {toast!r}")
            assert toast and "success" in (toast or "").lower(), (
                f"Save 后未出现 success toast，当前 toast: {toast!r}"
            )

            # Step 8: 关闭弹窗
            _close_param_config_dialog(page)
            time.sleep(1)
            ss(page, "verify_05_after_close.png")

            # 检查是否有 Warning 弹窗（logout 警告等）并 Cancel 掉
            _dismiss_blocking_overlays(page)
            time.sleep(0.5)

            # Step 9: 重新导航
            _navigate_to_bacnet(page)
            ss(page, "verify_06_after_navigate.png")
            print(f"  [7] 重新导航后 URL: {page.url}")

            # Step 10: 重新打开 2100 Parameter Config 验证
            opened2 = _open_param_dialog(page, "2100")
            print(f"  [8] 重新打开 Parameter Config: {opened2}")
            assert opened2, "重新打开 Parameter Config 失败"
            time.sleep(1.5)
            ss(page, "verify_07_reopened.png")

            # 验证 COV Increment 值
            # 找到当前页面中的参数（可能需要翻页）
            actual = _get_cov_increment_value_in_dialog(page, first_param_name)

            if actual is None:
                # 在所有页面中找
                total_pages = page.evaluate(
                    """() => {
                        const dlg = document.querySelector('[aria-label="Parameter Config"]');
                        const last = dlg && dlg.querySelector('.el-pager li:last-child');
                        return last ? parseInt(last.textContent.trim()) || 1 : 1;
                    }"""
                )
                print(f"  [9] 参数在当前页未找到，总页数: {total_pages}，翻页查找...")
                for page_num in range(2, min(total_pages + 1, 6)):
                    page.evaluate(
                        f"""() => {{
                            const dlg = document.querySelector('[aria-label="Parameter Config"]');
                            const pages = dlg && dlg.querySelectorAll('.el-pager li');
                            for (const p of (pages || [])) {{
                                if (p.textContent.trim() === '{page_num}') {{ p.click(); return; }}
                            }}
                        }}"""
                    )
                    time.sleep(0.8)
                    actual = _get_cov_increment_value_in_dialog(page, first_param_name)
                    if actual is not None:
                        print(f"  [9] 在第 {page_num} 页找到参数")
                        break

            print(f"\n  === 最终验证结果 ===")
            print(f"  参数: {first_param_name!r}")
            print(f"  COV Increment: {actual!r}")
            if actual in ("1.500", "1.5"):
                print("  PASS: COV Increment 保存为 1.500，保存路径验证通过！")
            else:
                print(f"  FAIL: 预期 '1.500'，实际 {actual!r}")

            ss(page, "verify_08_final.png")
            _close_param_config_dialog(page)

        except AssertionError as e:
            print(f"\n[ASSERTION ERROR] {e}")
            ss(page, "verify_assertion_error.png")
        except Exception as e:
            import traceback
            print(f"\n[EXCEPTION] {e}")
            print(traceback.format_exc())
            ss(page, "verify_exception.png")
        finally:
            browser.close()


if __name__ == "__main__":
    main()
