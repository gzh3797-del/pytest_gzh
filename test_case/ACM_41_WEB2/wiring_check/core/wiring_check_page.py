"""
WEB2 接线检查页面对象（Playwright）

页面：https://192.168.3.9/#/diagnostics/wiringCheck
实际流程（调试确认）：
  选择设备 → 点击绿色 Wiring Check → Confirm Nominal Voltage 弹窗 → Start Wiring Check → 等按钮恢复 → 解析结果表

表格结构（调试确认）：
  弹窗设备表：el-table__body[0]（4列，dialog 始终在 DOM 中）
  电压结果：  el-table__body[1]
    首行5列：Device|WiringConfig|Phase|Measurement|Status
    合并行3列：Phase|Measurement|Status（Phase∈A/B/C）
    相序行3列：'Phase Order'|value|status（Status列空时读Measurement列，空→Pass）
  电流结果：  el-table__body[2]
    首行7列：MeterPoint|Device|WiringConfig|Phase|InputChannel|Measurement|Status
    合并行4列：Phase|InputChannel|Measurement|Status（Phase∈A/B/C）
"""
import logging
import time
from playwright.sync_api import Page, expect
from test_case.ACM_41_WEB2.wiring_check.core import config as _cfg

# ── 页面常量 ─────────────────────────────────────────────────────────────────
URL          = f'https://{_cfg.WEB2_IP}/#/diagnostics/wiringCheck'
LOGIN_URL    = f'https://{_cfg.WEB2_IP}'
DEFAULT_USER = _cfg.WEB2_USER
DEFAULT_PASS = _cfg.WEB2_PASS

TIMEOUT_PAGE  = 15_000
TIMEOUT_CHECK = 60_000


class WiringCheckPage:
    def __init__(self, page: Page, device_name: str = None):
        self.page = page
        self.device_name = device_name          # e.g. cfg.WEB2_DEVICE_NAME
        self._nominal_saved = False             # 额定电压已在弹窗中 Save 过

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
        self.page.goto(URL, timeout=TIMEOUT_PAGE)
        self.page.wait_for_load_state('networkidle', timeout=TIMEOUT_PAGE)

    # ── 设备选择 ─────────────────────────────────────────────────────────────

    def select_device(self):
        """从 Device 下拉选择指定设备（el-select 组件）"""
        if not self.device_name:
            return
        self.page.locator('.el-select').first.click()
        time.sleep(0.3)
        self.page.locator('.el-select-dropdown__item').filter(
            has_text=self.device_name
        ).click()
        time.sleep(0.3)
        logging.info('Selected device: %s', self.device_name)

    # ── 触发检查 ─────────────────────────────────────────────────────────────

    def trigger_check(self):
        """点击 Wiring Check → 处理 Confirm Nominal Voltage 弹窗 → Start Wiring Check"""
        btn = self.page.get_by_role('button', name='Wiring Check', exact=True)
        expect(btn).to_be_enabled(timeout=TIMEOUT_PAGE)
        btn.click()
        logging.info('Wiring Check button clicked')

        # ── 等待 Confirm Nominal Voltage 弹窗 ─────────────────────────────────
        dialog = self.page.locator('.el-dialog')
        expect(dialog).to_be_visible(timeout=TIMEOUT_PAGE)
        title = dialog.locator('.el-dialog__title').inner_text()
        logging.info('Dialog appeared: "%s"', title)

        # 首次触发时 Save 下发额定电压，后续用例直接跳过
        save_btn = dialog.locator('button:has-text("Save")')
        if save_btn.count() > 0 and not self._nominal_saved:
            save_btn.click()
            self._nominal_saved = True
            logging.info('Save nominal voltage clicked')
            time.sleep(0.5)
        else:
            logging.info('Nominal voltage already saved, skip Save')

        # 点击 Start Wiring Check 正式开始
        start_btn = dialog.locator('button:has-text("Start Wiring Check")')
        expect(start_btn).to_be_enabled(timeout=TIMEOUT_PAGE)
        start_btn.click()
        logging.info('Start Wiring Check clicked')

    # ── 等待完成 ─────────────────────────────────────────────────────────────

    def wait_for_completion(self):
        """轮询直到 Wiring Check 按钮恢复 enabled（检查完成）"""
        deadline = time.time() + TIMEOUT_CHECK / 1000
        time.sleep(0.5)   # 给按钮时间变灰
        btn = self.page.get_by_role('button', name='Wiring Check', exact=True)
        while time.time() < deadline:
            if not btn.is_disabled():
                logging.info('Wiring Check completed')
                return
            time.sleep(0.5)
        logging.warning('wait_for_completion timed out')

    # ── 结果解析 ─────────────────────────────────────────────────────────────

    def read_voltage_results(self) -> dict:
        """
        解析电压结果表（table[1]），返回指定设备的各相状态：
        {'A': str, 'B': str, 'C': str, 'order': str}
        列顺序：[0]Device [1]WiringConfig [2]Phase [3]Measurement [4]WiringStatus

        Phase Order 行结构（3列）：['Phase Order', <value>, <status>]
          - cells[2] 为 Status 列；若为空则读 cells[1]（部分固件版本 Status 列留空，值在 Measurement 列）
          - 空字符串视为 'Pass'
        """
        result = {
            'A': 'Not Checked', 'B': 'Not Checked',
            'C': 'Not Checked', 'order': 'Not Checked',
        }

        data_tbl = self.page.locator('table.el-table__body').nth(1)
        rows = data_tbl.locator('tbody tr').all()

        current_device = None
        for tr in rows:
            cells = [td.inner_text().strip() for td in tr.locator('td').all()]
            n = len(cells)

            if n == 5:
                # 首行：[Device, WiringConfig, Phase, Measurement, Status]
                current_device = cells[0]
                phase, status = cells[2], cells[4]
            elif n == 3 and cells[0] in ('A', 'B', 'C'):
                # A/B/C 合并行：[Phase, Measurement, Status]
                phase, status = cells[0], cells[2]
            elif n == 3 and cells[0] == 'Phase Order':
                # Phase Order 行：Status 在 cells[2]，若空则读 cells[1]
                phase  = 'order'
                status = cells[2] if cells[2] else cells[1]
            else:
                continue

            if self.device_name and current_device != self.device_name:
                continue

            if phase in ('A', 'B', 'C', 'order'):
                result[phase] = status or 'Pass'

        return result

    def read_current_results(self) -> list[dict]:
        """
        解析电流结果表（table[3]），逐页翻页读取所有 Meter Point 状态。
        返回：[{'A': str, 'B': str, 'C': str}, ...]，按 Meter Point 顺序
        列顺序：[0]MeterPoint [1]Device [2]WiringConfig [3]Phase [4]InputChannel [5]Measurement [6]WiringStatus
        """
        groups = {}          # {mp_name: {phase: status}}
        mp_order = []        # 保持 Meter Point 顺序

        def _read_page():
            data_tbl = self.page.locator('table.el-table__body').nth(2)
            current_mp     = None
            current_device = None
            for tr in data_tbl.locator('tbody tr').all():
                cells = [td.inner_text().strip() for td in tr.locator('td').all()]
                n = len(cells)
                if n == 7:
                    # 首行：[MeterPoint, Device, WiringConfig, Phase, InputCh, Meas, Status]
                    current_mp, current_device = cells[0], cells[1]
                    phase, status = cells[3], cells[6]
                elif n == 4 and cells[0] in ('A', 'B', 'C'):
                    # 合并行：[Phase, InputCh, Measurement, Status]
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

        # 找电流表分页器（第2个 .el-pagination）
        paginators = self.page.locator('.el-pagination').all()
        current_pager = paginators[1] if len(paginators) >= 2 else None

        if current_pager:
            page_btns = current_pager.locator('.el-pager .number').all()
            num_pages = len(page_btns)
            logging.info('Current table: %d page(s)', num_pages)
            for pg in range(num_pages):
                _read_page()
                if pg < num_pages - 1:
                    # 每次重新定位，防止 DOM 刷新导致 stale
                    self.page.locator('.el-pagination').nth(1)\
                        .locator('.el-pager .number').nth(pg + 1).click()
                    time.sleep(0.4)
        else:
            _read_page()

        return [groups[mp] for mp in mp_order]

    # ── 完整流程 ──────────────────────────────────────────────────────────────

    def run_check(self) -> tuple[dict, list[dict]]:
        self.select_device()
        self.trigger_check()
        self.wait_for_completion()
        time.sleep(1.0)   # 等待结果表在按钮恢复后完成渲染
        return self.read_voltage_results(), self.read_current_results()
