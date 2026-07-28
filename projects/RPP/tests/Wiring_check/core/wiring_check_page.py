"""
RPP 接线检查页面对象（Playwright）

页面：http://192.168.2.94:3030/#/maintenance/diagnostics/wiringCheck
导航路径：RPP Maintenance 顶级组 → Diagnostics（左侧可展开）→ Wiring Check 子项
（HMI 在 Settings 侧栏下的 /diagnostics/wiringCheck；RPP 归到 Maintenance→Diagnostics）

实际流程：
  1. navigate()：
     a. goto /#/maintenance/systemStatus（激活 Maintenance 侧边栏）
     b. goto 目标 URL /#/maintenance/diagnostics/wiringCheck（组已激活，路由生效）
     c. 兜底：若被重定向，点击侧边栏 Diagnostics 展开 → Wiring Check 子项
  2. select_device()：从 Device 下拉（el-select，默认"All"）选择设备
  3. trigger_check()：点击 Wiring Check 按钮 → Confirm Nominal Voltage 弹窗（无 Save）
     → Start Wiring Check
  4. wait_for_completion()：轮询 Wiring Check 按钮 enabled 状态
  5. read_voltage_results() / read_current_results()：解析结果表

与 WEB2 的主要区别：
  - navigate() 需通过菜单导航（直接 goto URL 会被重定向）
  - 弹窗无 Save 按钮，额定电压只用于检查计算，不下发表
  - Phase Order 行 Status 列显示 '-' 代表 Pass（_status_match 已兼容）

表格结构（调试确认，与 WEB2 完全相同）：
  el-table__body[0]：弹窗设备表（dialog 始终在 DOM，不可见时跳过）
  el-table__body[1]：电压结果表
    首行 5 列：Device|WiringConfig|Phase|Measurement|WiringStatus
    合并行 3 列：Phase|Measurement|WiringStatus（Phase∈A/B/C）
    相序行 3 列：'Phase Order'|value|status（'-'→Pass，错误时显示具体文字）
  el-table__body[2]：电流结果表
    首行 7 列：MeterPoint|Device|WiringConfig|Phase|InputChannel|Measurement|WiringStatus
    合并行 4 列：Phase|InputChannel|Measurement|WiringStatus（Phase∈A/B/C）
"""
import logging
import time
from playwright.sync_api import Page, expect
from projects.RPP.tests.Wiring_check.core import config as _cfg
from projects.RPP.settings import BASE_URL

# ── 页面常量 ─────────────────────────────────────────────────────────────────
# RPP：Wiring Check 归属 Maintenance → Diagnostics（HMI 在 Settings 侧栏下）。
# 地址用统一 BASE_URL（RPP 为 http://192.168.2.94:3030），不再硬编码 https/IP。
URL           = f'{BASE_URL}/#/maintenance/diagnostics/wiringCheck'
LOGIN_URL     = BASE_URL
# 激活 Maintenance 侧边栏用的落地页（先进该组，再点 Diagnostics → Wiring Check）
SETTINGS_URL  = f'{BASE_URL}/#/maintenance/systemStatus'
DEFAULT_USER  = _cfg.HMI_USER
DEFAULT_PASS  = _cfg.HMI_PASS

TIMEOUT_PAGE  = 15_000
TIMEOUT_CHECK = 60_000

# Device 下拉容器（页面顶部"选择检查哪台表"的 el-select 所在 div）
DEVICE_SELECT_XPATH = (
    'xpath=//*[@id="app"]/div/section/section/main'
    '/div[2]/div[1]/div[1]/div/div/div[1]/div[2]'
)


class WiringCheckPage:
    def __init__(self, page: Page, device_name: str = None):
        self.page = page
        self.device_name = device_name

    # ── 登录 ─────────────────────────────────────────────────────────────────

    def login_if_needed(self):
        self.page.goto(LOGIN_URL, timeout=TIMEOUT_PAGE)
        try:
            self.page.wait_for_selector('input', timeout=5_000)
        except Exception:
            logging.info('No input found, already logged in')
            return

        if self.page.locator("input[placeholder='Enter User Name']").count() == 0:
            logging.info('Already logged in')
            return

        self.page.locator("input[placeholder='Enter User Name']").fill(DEFAULT_USER)
        self.page.locator("input[placeholder='Enter Password']").fill(DEFAULT_PASS)
        self.page.locator('button:has-text("Sign In")').click()
        self.page.wait_for_selector("input[placeholder='Enter User Name']",
                                    state='hidden', timeout=TIMEOUT_PAGE)
        logging.info('Login successful')

    # ── 导航 ─────────────────────────────────────────────────────────────────

    def navigate(self):
        """
        导航到 Wiring Check 页面（RPP：Maintenance → Diagnostics → Wiring Check）。
        直接 goto 目标 URL 会被 SPA 按当前激活组重定向，故先进 Maintenance 组落地页，
        再直接 goto 目标 URL（此时 Maintenance 组已激活，路由生效）：
          1. goto Maintenance 落地页（/maintenance/systemStatus）激活 Maintenance 侧边栏
          2. goto Wiring Check 目标 URL（/maintenance/diagnostics/wiringCheck）
          3. 兜底：若未落到目标，点击侧边栏 Diagnostics（展开）→ Wiring Check 子项
        """
        # Step 1：激活 Maintenance 侧边栏
        self.page.goto(SETTINGS_URL, timeout=TIMEOUT_PAGE)
        self.page.wait_for_load_state('networkidle', timeout=TIMEOUT_PAGE)
        time.sleep(1)
        logging.info('Maintenance sidebar activated')

        # Step 2：直接跳转 Wiring Check（组已激活，路由生效）
        self.page.goto(URL, timeout=TIMEOUT_PAGE)
        self.page.wait_for_load_state('networkidle', timeout=TIMEOUT_PAGE)
        time.sleep(1)
        logging.info('Wiring Check URL goto, URL=%s', self.page.url)

        # Step 3：兜底菜单导航——若被重定向未落到 wiringCheck，则点 Diagnostics 展开 → Wiring Check
        if 'wiringCheck' not in self.page.url:
            for item in self.page.locator('li:has-text("Diagnostics"), .left-nav-item:has-text("Diagnostics")').all():
                try:
                    if item.inner_text().strip() == 'Diagnostics' and item.is_visible():
                        item.click()
                        break
                except Exception:
                    pass
            time.sleep(0.6)
            wc = self.page.locator(':text-is("Wiring Check")')
            try:
                expect(wc.first).to_be_visible(timeout=TIMEOUT_PAGE)
                wc.first.click()
                self.page.wait_for_load_state('networkidle', timeout=TIMEOUT_PAGE)
                time.sleep(0.8)
            except Exception:
                logging.warning('Wiring Check 菜单兜底导航未命中，当前 URL=%s', self.page.url)
        logging.info('Wiring Check navigate done, URL=%s', self.page.url)

    # ── 设备选择 ─────────────────────────────────────────────────────────────

    def select_device(self):
        """从 Device 下拉选择指定设备（el-select 组件，默认 All）。

        device_name 即 HMI 页面 Device 下拉中显示的设备名，由
        config.yaml 的 wiring_device_name 配置（未配置则回退 hmi_device_name）。
        """
        if not self.device_name:
            return
        # 定位顶部 Device 下拉容器内的 el-select，避免误点页面其它 el-select
        container = self.page.locator(DEVICE_SELECT_XPATH)
        select = container.locator('.el-select')
        if select.count() == 0:
            select = self.page.locator('.el-select').first
        select.first.click()
        time.sleep(0.3)
        self.page.locator('.el-select-dropdown__item').filter(
            has_text=self.device_name
        ).first.click()
        time.sleep(0.3)
        logging.info('Selected device: %s', self.device_name)

    # ── 触发检查 ─────────────────────────────────────────────────────────────

    def trigger_check(self):
        """点击 Wiring Check 按钮 → Confirm Nominal Voltage 弹窗（无 Save）→ Start Wiring Check"""
        btn = self.page.get_by_role('button', name='Wiring Check', exact=True)
        expect(btn).to_be_enabled(timeout=TIMEOUT_PAGE)
        btn.click()
        logging.info('Wiring Check button clicked')

        # 等待弹窗
        dialog = self.page.locator('.el-dialog')
        expect(dialog).to_be_visible(timeout=TIMEOUT_PAGE)
        try:
            title = dialog.locator('.el-dialog__title').inner_text()
            logging.info('Dialog: "%s"', title)
        except Exception:
            pass

        # 直接点击 Start Wiring Check（HMI 无 Save 按钮）
        start_btn = dialog.locator('button:has-text("Start Wiring Check")')
        expect(start_btn).to_be_enabled(timeout=TIMEOUT_PAGE)
        start_btn.click()
        logging.info('Start Wiring Check clicked')

    # ── 等待完成 ─────────────────────────────────────────────────────────────

    def wait_for_completion(self):
        """轮询直到 Wiring Check 按钮恢复 enabled（检查完成）"""
        deadline = time.time() + TIMEOUT_CHECK / 1000
        time.sleep(0.5)
        btn = self.page.get_by_role('button', name='Wiring Check', exact=True)
        while time.time() < deadline:
            if not btn.is_disabled():
                logging.info('Wiring Check completed')
                return
            time.sleep(0.5)
        logging.warning('wait_for_completion timed out')

    # ── 工具 ─────────────────────────────────────────────────────────────────

    def _dismiss_message_box(self):
        """关闭页面上可能出现的 el-message-box 遮罩弹窗（alert/confirm），防止拦截后续点击。

        遮罩可能堆叠多个（上一条用例遗留 + 本次新弹），故循环清理直到无可见遮罩，
        避免只关掉最上层、底下仍有遮罩继续拦截 pointer events。
        """
        overlay = self.page.locator('.el-overlay.is-message-box')
        for _ in range(5):  # 上限 5 次，防止异常情况下死循环
            if overlay.count() == 0 or not overlay.first.is_visible():
                return
            dismissed = False
            # 依次尝试确认/关闭按钮
            for sel in [
                '.el-message-box__btns .el-button--primary',
                '.el-message-box__btns .el-button',
                '.el-message-box__headerbtn',
            ]:
                btn = overlay.first.locator(sel)
                if btn.count() > 0:
                    btn.first.click()
                    time.sleep(0.3)
                    logging.info('Dismissed message-box overlay')
                    dismissed = True
                    break
            if not dismissed:
                return

    # ── 结果解析 ─────────────────────────────────────────────────────────────

    def read_voltage_results(self) -> dict:
        """
        解析电压结果表（table[1]），返回指定设备的各相状态：
        {'A': str, 'B': str, 'C': str, 'order': str}

        列顺序：[0]Device [1]WiringConfig [2]Phase [3]Measurement [4]WiringStatus
        合并行：[0]Phase [1]Measurement [2]WiringStatus
        Phase Order 行：[0]'Phase Order' [1]value [2]status（'-'→Pass）
        """
        result = {
            'A': 'Not Checked', 'B': 'Not Checked',
            'C': 'Not Checked', 'order': 'Not Checked',
        }

        def _read_page():
            data_tbl = self.page.locator('table.el-table__body').nth(1)
            current_device = None
            for tr in data_tbl.locator('tbody tr').all():
                cells = [td.inner_text().strip() for td in tr.locator('td').all()]
                n = len(cells)

                if n == 5:
                    current_device = cells[0]
                    phase, status = cells[2], cells[4]
                elif n == 3 and cells[0] in ('A', 'B', 'C'):
                    phase, status = cells[0], cells[2]
                elif n == 3 and cells[0] == 'Phase Order':
                    phase  = 'order'
                    # cells[2] = '-' 时表示 Pass；有错误时显示错误文字
                    status = cells[2] if cells[2] else cells[1]
                else:
                    continue

                if self.device_name and current_device != self.device_name:
                    continue

                if phase in ('A', 'B', 'C', 'order'):
                    result[phase] = status or 'Pass'

        # 电压表分页（paginators[0]）
        paginators = self.page.locator('.el-pagination').all()
        v_pager = paginators[0] if paginators else None

        if v_pager:
            page_btns = v_pager.locator('.el-pager .number').all()
            num_pages = len(page_btns)
            logging.info('Voltage table: %d page(s)', num_pages)
            for pg in range(num_pages):
                _read_page()
                if pg < num_pages - 1:
                    self._dismiss_message_box()
                    self.page.locator('.el-pagination').nth(0)\
                        .locator('.el-pager .number').nth(pg + 1).click()
                    time.sleep(0.4)
        else:
            _read_page()

        return result

    def read_current_results(self) -> list[dict]:
        """
        解析电流结果表（table[2]），逐页翻页读取所有 Meter Point 状态。
        返回：[{'A': str, 'B': str, 'C': str}, ...]，按 Meter Point 顺序

        列顺序：[0]MeterPoint [1]Device [2]WiringConfig [3]Phase [4]InputChannel [5]Measurement [6]WiringStatus
        合并行：[0]Phase [1]InputChannel [2]Measurement [3]WiringStatus
        """
        groups = {}
        mp_order = []

        def _read_page():
            data_tbl = self.page.locator('table.el-table__body').nth(2)
            current_mp     = None
            current_device = None
            for tr in data_tbl.locator('tbody tr').all():
                cells = [td.inner_text().strip() for td in tr.locator('td').all()]
                n = len(cells)
                if n == 7:
                    current_mp, current_device = cells[0], cells[1]
                    phase, status = cells[3], cells[6]
                elif n == 4 and cells[0] in ('A', 'B', 'C'):
                    phase, status = cells[0], cells[3]
                else:
                    continue

                if self.device_name and current_device != self.device_name:
                    continue
                if current_mp not in groups:
                    groups[current_mp] = {}
                    mp_order.append(current_mp)
                if phase == 'A':
                    groups[current_mp]['A'] = status
                elif phase == 'B':
                    groups[current_mp]['B'] = status
                elif phase == 'C':
                    groups[current_mp]['C'] = status

        # 电流表分页（paginators[1]）
        paginators = self.page.locator('.el-pagination').all()
        current_pager = paginators[1] if len(paginators) >= 2 else None

        if current_pager:
            page_btns = current_pager.locator('.el-pager .number').all()
            num_pages = len(page_btns)
            logging.info('Current table: %d page(s)', num_pages)
            for pg in range(num_pages):
                _read_page()
                if pg < num_pages - 1:
                    self._dismiss_message_box()
                    self.page.locator('.el-pagination').nth(1)\
                        .locator('.el-pager .number').nth(pg + 1).click()
                    time.sleep(0.4)
        else:
            _read_page()

        return [groups[mp] for mp in mp_order]

    # ── 完整流程 ──────────────────────────────────────────────────────────────

    def run_check(self) -> tuple[dict, list[dict]]:
        # 先清理上一条用例可能遗留的 Warning/alert 遮罩（共享页面跨用例时会残留，
        # 否则 el-overlay-message-box 会拦截 select_device 的下拉点击导致超时）
        self._dismiss_message_box()
        self.select_device()
        self.trigger_check()
        self.wait_for_completion()
        time.sleep(1.0)
        return self.read_voltage_results(), self.read_current_results()
