"""
HMI1-7 接线检查页面结构调试脚本

用途：打开 HMI 页面，执行一次 Wiring Check，逐步 dump：
  1. 登录后页面元素（有无 Device 下拉）
  2. 弹窗完整 HTML（确认是否有 Save 按钮、Start 按钮文字）
  3. 电压结果表每行的列数和内容
  4. 电流结果表每行的列数和内容
  5. 分页器数量

运行（需已有 4100 正常上电，HMI 可访问）：
    python projects/RPP/tests/Wiring_check/debug_page.py

结果写入同目录 debug_output.txt，并在控制台同步打印。
"""
import sys
import os
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../../..'))

from playwright.sync_api import sync_playwright

HMI_IP   = '192.168.2.8'
HMI_USER = 'q'
HMI_PASS = '1'
LOGIN_URL = f'https://{HMI_IP}'

OUT_FILE = os.path.join(os.path.dirname(__file__), 'debug_output.txt')

_lines: list[str] = []


def log(msg: str = ''):
    print(msg)
    _lines.append(msg)


def dump_all_tables(page):
    all_tbls = page.locator('table.el-table__body').count()
    log(f'  el-table__body 总数：{all_tbls}')
    for idx in range(all_tbls):
        log(f'\n  ── 表格[{idx}] ──')
        tbl = page.locator('table.el-table__body').nth(idx)
        rows = tbl.locator('tbody tr').all()
        log(f'  行数：{len(rows)}')
        for i, tr in enumerate(rows):
            cells = [td.inner_text().strip().replace('\n', '↵')
                     for td in tr.locator('td').all()]
            log(f'    行[{i:02d}] 列数={len(cells)}  {cells}')


def list_buttons(page, label=''):
    btns = page.locator('button').all()
    log(f'  [{label}] button 数量：{len(btns)}')
    for i, b in enumerate(btns):
        try:
            txt = b.inner_text().strip()
            visible = b.is_visible()
            if visible:
                cls = b.get_attribute('class') or ''
                log(f'    [{i:02d}] text={txt!r}  class={cls[:60]!r}')
        except Exception:
            pass


def main():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, slow_mo=200)
        page = browser.new_page(ignore_https_errors=True)

        # ── 1. 登录 ──────────────────────────────────────────────────────────
        log('='*60)
        log('STEP 1: 登录')
        log('='*60)
        page.goto(LOGIN_URL, timeout=20_000)
        page.wait_for_load_state('networkidle', timeout=20_000)

        if page.locator("input[placeholder='Enter User Name']").count() > 0:
            page.locator("input[placeholder='Enter User Name']").fill(HMI_USER)
            page.locator("input[placeholder='Enter Password']").fill(HMI_PASS)
            page.locator('button:has-text("Sign In")').click()
            page.wait_for_selector("input[placeholder='Enter User Name']",
                                   state='hidden', timeout=15_000)
            log('  登录成功')
        else:
            log('  已登录，跳过')

        time.sleep(2)
        log(f'  当前 URL: {page.url}')

        # ── 2. 进入 Settings 侧边栏 ───────────────────────────────────────────
        log('\n' + '='*60)
        log('STEP 2: 进入 Settings 侧边栏')
        log('='*60)

        hmi_btn = page.locator('button:has-text("AcuHMI-1-7")')
        if hmi_btn.count() > 0:
            hmi_btn.first.click()
            time.sleep(2)
            page.wait_for_load_state('networkidle', timeout=10_000)
            log(f'  点击后 URL: {page.url}')
        else:
            log('  未找到 AcuHMI-1-7 按钮，直接导航到 Settings')
            page.goto(f'https://{HMI_IP}/#/systemSettings/dateTime', timeout=15_000)
            page.wait_for_load_state('networkidle', timeout=10_000)
            time.sleep(1.5)
            log(f'  导航后 URL: {page.url}')

        # ── 3. Settings 侧边栏：列出所有菜单项 ───────────────────────────────
        log('\n' + '='*60)
        log('STEP 3: Settings 侧边栏菜单项')
        log('='*60)

        for sel in ['.el-menu-item', 'li[class*="menu-item"]', '[class*="menu-item"]',
                    'li.el-menu-item', '.el-menu > li', 'aside li']:
            items = page.locator(sel).all()
            if items:
                log(f'  [{sel}] count={len(items)}')
                for i, m in enumerate(items[:20]):
                    try:
                        txt = m.inner_text().strip()
                        vis = m.is_visible()
                        log(f'    [{i:02d}] {txt!r}  visible={vis}')
                    except Exception:
                        pass
                break

        # ── 4. 点击 Diagnostics 菜单项 ────────────────────────────────────────
        log('\n' + '='*60)
        log('STEP 4: 点击 Diagnostics')
        log('='*60)

        diag = None
        for sel in [
            'li:has-text("Diagnostics")',
            '[class*="menu"]:has-text("Diagnostics")',
            'a:has-text("Diagnostics")',
            ':text-is("Diagnostics")',
            'span:has-text("Diagnostics")',
        ]:
            cands = page.locator(sel).all()
            log(f'  [{sel}] count={len(cands)}')
            for c in cands:
                try:
                    txt = c.inner_text().strip()
                    log(f'    text={txt!r}  visible={c.is_visible()}')
                    if txt == 'Diagnostics' and c.is_visible():
                        diag = c
                        break
                except Exception:
                    pass
            if diag:
                break

        if diag:
            diag.click()
            time.sleep(2)
            page.wait_for_load_state('networkidle', timeout=10_000)
            log(f'  点击后 URL: {page.url}')
        else:
            log('  [WARN] 未找到 Diagnostics，尝试直接导航')
            for url_candidate in [
                f'https://{HMI_IP}/#/diagnostics/wiringCheck',
                f'https://{HMI_IP}/#/diagnostics',
            ]:
                page.goto(url_candidate, timeout=10_000)
                time.sleep(1.5)
                log(f'  尝试 {url_candidate} → 实际 URL: {page.url}')
                if 'diagnostic' in page.url.lower():
                    break

        # ── 5. Diagnostics 页内容分析 ─────────────────────────────────────────
        log('\n' + '='*60)
        log('STEP 5: Diagnostics 页内容分析')
        log('='*60)
        time.sleep(1.5)
        log(f'  URL: {page.url}')
        list_buttons(page, 'diag_page')

        tabs = page.locator('.el-tabs__item').all()
        log(f'  .el-tabs__item 数量：{len(tabs)}')
        for i, t in enumerate(tabs):
            try:
                log(f'    tab[{i}] {t.inner_text().strip()!r}')
            except Exception:
                pass

        for kw in ['Wiring Check', 'Wiring', 'Check']:
            found = page.locator(f'*:has-text("{kw}")').all()
            log(f'  含"{kw}"的元素: {len(found)}')
            for el in found[:5]:
                try:
                    tag = el.evaluate('el => el.tagName')
                    txt = el.inner_text().strip()[:80]
                    cls = (el.get_attribute('class') or '')[:40]
                    log(f'    <{tag}> cls={cls!r} text={txt!r}')
                except Exception:
                    pass
            if found:
                break

        # ── 6. 定位并点击 Wiring Check ────────────────────────────────────────
        log('\n' + '='*60)
        log('STEP 6: 定位并点击 Wiring Check 触发按钮')
        log('='*60)
        wc_trigger = None
        for sel_fn, desc in [
            (lambda: page.get_by_role('button', name='Wiring Check'), 'role=button name=WC'),
            (lambda: page.locator('button').filter(has_text='Wiring Check'), 'button has-text WC'),
            (lambda: page.locator('.el-tabs__item:has-text("Wiring Check")'), 'tab WC'),
            (lambda: page.locator('button:has-text("Wiring Check")'), 'button:has-text WC'),
            (lambda: page.locator(':text-is("Wiring Check")'), 'text-is WC'),
        ]:
            try:
                el = sel_fn()
                cnt = el.count()
                log(f'  [{desc}] count={cnt}')
                if cnt > 0:
                    wc_trigger = el.first
                    log(f'    命中: {wc_trigger.inner_text().strip()!r}')
                    break
            except Exception as ex:
                log(f'  [{desc}] 异常: {ex}')

        if wc_trigger is None:
            log('  [ERROR] 未找到 Wiring Check 按钮/标签')
            body_html = page.locator('body').evaluate('el => el.innerHTML')
            log(f'\n  页面 body innerHTML（前3000字符）：\n{body_html[:3000]}')
        else:
            wc_trigger.click()
            time.sleep(2)
            log(f'  点击后 URL: {page.url}')

            # ── 7. Wiring Check 页分析 ─────────────────────────────────────
            log('\n' + '='*60)
            log('STEP 7: Wiring Check 页分析')
            log('='*60)
            page.wait_for_load_state('networkidle', timeout=10_000)
            time.sleep(2)
            list_buttons(page, 'wc_page')

            el_selects = page.locator('.el-select').count()
            log(f'  .el-select 数量：{el_selects}')
            for i in range(el_selects):
                try:
                    txt = page.locator('.el-select').nth(i).inner_text().strip()
                    log(f'    select[{i}] text={txt!r}')
                except Exception:
                    pass

            # ── 8. 点击 Wiring Check 触发按钮 ─────────────────────────────
            log('\n' + '='*60)
            log('STEP 8: 点击触发 Wiring Check')
            log('='*60)
            trigger = None
            for kw in ['Wiring Check', 'Start', 'Check']:
                cand = page.locator('button').filter(has_text=kw)
                if cand.count() > 0:
                    trigger = cand.first
                    log(f'  使用按钮含"{kw}"')
                    break
            if trigger is None:
                log('  [ERROR] 无触发按钮')
            else:
                trigger.click()
                time.sleep(2)

                # ── 9. 弹窗分析 ──────────────────────────────────────────────
                log('\n' + '='*60)
                log('STEP 9: 弹窗分析')
                log('='*60)
                dialog = page.locator('.el-dialog')
                log(f'  .el-dialog count={dialog.count()}')
                visible_dlg = None
                for i in range(dialog.count()):
                    d = dialog.nth(i)
                    if d.is_visible():
                        visible_dlg = d
                        break

                if visible_dlg:
                    try:
                        log(f'  标题：{visible_dlg.locator(".el-dialog__title").inner_text()!r}')
                    except Exception:
                        pass
                    btns = visible_dlg.locator('button').all()
                    log(f'  弹窗按钮数：{len(btns)}')
                    for i, b in enumerate(btns):
                        try:
                            log(f'    [{i}] text={b.inner_text().strip()!r}  disabled={b.is_disabled()}')
                        except Exception as ex:
                            log(f'    [{i}] 失败: {ex}')
                    dlg_tbls = visible_dlg.locator('table.el-table__body').count()
                    log(f'  弹窗 el-table__body 数量：{dlg_tbls}')
                    if dlg_tbls > 0:
                        rows0 = visible_dlg.locator('table.el-table__body').nth(0).locator('tbody tr').all()
                        log(f'  弹窗表格行数：{len(rows0)}')
                        for i, tr in enumerate(rows0[:10]):
                            cells = [td.inner_text().strip() for td in tr.locator('td').all()]
                            log(f'    行[{i}] 列数={len(cells)}  {cells}')
                    try:
                        dlg_html = visible_dlg.evaluate('el => el.outerHTML')
                        log(f'\n  弹窗 HTML（前3000字符）：\n{dlg_html[:3000]}')
                    except Exception as ex:
                        log(f'  HTML 失败: {ex}')

                    # ── 10. Start ─────────────────────────────────────────────
                    log('\n' + '='*60)
                    log('STEP 10: Start Wiring Check')
                    log('='*60)
                    start_btn = None
                    for kw in ['Start Wiring Check', 'Start', 'OK', 'Confirm']:
                        cand = visible_dlg.locator(f'button:has-text("{kw}")')
                        if cand.count() > 0:
                            start_btn = cand.first
                            log(f'  使用"{kw}"')
                            break
                    if start_btn is None:
                        all_b = visible_dlg.locator('button').all()
                        if all_b:
                            start_btn = all_b[-1]
                            log(f'  fallback 最后一个按钮')
                    if start_btn:
                        start_btn.click()
                        log('  已点击')

                    # ── 11. 等待完成 ──────────────────────────────────────────
                    log('\n' + '='*60)
                    log('STEP 11: 等待检查完成')
                    log('='*60)
                    deadline = time.time() + 60
                    time.sleep(1.5)
                    while time.time() < deadline:
                        try:
                            disabled = trigger.evaluate(
                                'el => el.disabled || el.classList.contains("is-disabled")')
                            if not disabled:
                                log('  完成')
                                break
                        except Exception:
                            pass
                        time.sleep(0.5)
                    else:
                        log('  [WARN] 超时')
                    time.sleep(2)

                    # ── 12. 结果表 ────────────────────────────────────────────
                    log('\n' + '='*60)
                    log('STEP 12: 结果表结构')
                    log('='*60)
                    paginators = page.locator('.el-pagination').all()
                    log(f'  .el-pagination 数量：{len(paginators)}')
                    dump_all_tables(page)

                    all_tbls = page.locator('table.el-table__body').count()
                    log('\n[各表前5行 outerHTML]')
                    for idx in range(all_tbls):
                        tbl = page.locator('table.el-table__body').nth(idx)
                        rows = tbl.locator('tbody tr').all()
                        if not rows:
                            continue
                        log(f'\n表格[{idx}] 前5行:')
                        for tr in rows[:5]:
                            try:
                                log(f'  {tr.evaluate("el => el.outerHTML")[:600]}')
                            except Exception:
                                pass

        log('\n' + '='*60)
        log('调试完成')
        log('='*60)
        input('按 Enter 关闭浏览器...')
        browser.close()

    _save()


def _save():
    with open(OUT_FILE, 'w', encoding='utf-8') as f:
        f.write('\n'.join(_lines))
    print(f'输出已写入：{OUT_FILE}')


if __name__ == '__main__':
    main()
