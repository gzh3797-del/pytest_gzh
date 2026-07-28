import logging
import re
import time

from pages.base_page import BasePage
from config.settings import BASE_URL


class AWSIoTPage(BasePage):
    """AWS IoT 配置页面（Playwright 实现），支持 AcuHMI-1-7 / PX-EMD-G 等网关"""

    # ── XPath / CSS 定位器常量 ─────────────────────────────────────────────────

    # 顶部导航栏设备按钮（默认值，构造函数可按 device_name 覆盖）
    DEVICE_SETTINGS_BTN = "xpath=//div[contains(@class,'nav-item-menu')][.//span[normalize-space(.)='AcuHMI-1-7']]"

    def __init__(self, page, device_name: str = "AcuHMI-1-7",
                 gateway_url: str = "", username: str = "admin", password: str = ""):
        super().__init__(page)
        self._device_name = device_name
        self._gateway_url = gateway_url
        self._username = username
        self._password = password
        self.DEVICE_SETTINGS_BTN = (
            f"xpath=//div[contains(@class,'nav-item-menu')]"
            f"[.//span[normalize-space(.)='{device_name}']]"
        )

    # 左侧导航（通用文本匹配，兼容 AcuHMI-1-7 / PX-EMD-G 等不同 class 名）
    PROTOCOLS_MENU = (
        "xpath=("
        "//li[contains(@class,'left-nav-item') and contains(normalize-space(.),'Protocols')]"
        " | //li[contains(normalize-space(.),'Protocols') and not(.//*[normalize-space(text())='Protocols'])]"
        " | //*[@role='menuitem' and contains(normalize-space(.),'Protocols')]"
        ")[1]"
    )
    AWS_IOT_TAB = (
        "xpath=("
        "//li[contains(@class,'el-menu-item') and contains(normalize-space(.),'AWS IoT')]"
        " | //*[@role='tab' and contains(normalize-space(.),'AWS IoT')]"
        " | //div[contains(@class,'tab') and contains(normalize-space(.),'AWS IoT')]"
        " | //span[normalize-space(.)='AWS IoT']/parent::*"
        ")[1]"
    )

    # Enable / Disable 单选
    ENABLE_RADIO  = "xpath=//label[.//span[normalize-space(.)='Enable']]"
    DISABLE_RADIO = "xpath=//label[.//span[normalize-space(.)='Disable']]"

    # 文本输入
    CLIENT_ID_INPUT = "css=input[placeholder='Enter Client Id']"
    URL_INPUT       = "css=input[placeholder='Enter URL']"
    TOPIC_INPUT     = "css=input[placeholder='Enter Topic']"

    # Interval 下拉
    INTERVAL_DROPDOWN = "xpath=//label[contains(normalize-space(.),'Interval')]/following::div[contains(@class,'el-select')][1]"

    # 文件上传
    CERT_FILE_INPUT = "xpath=(//input[@type='file'])[1]"
    KEY_FILE_INPUT  = "xpath=(//input[@type='file'])[2]"

    # 设备表格行
    DEVICE_TABLE_ROWS = "xpath=//div[contains(@class,'el-table__body-wrapper')]//tbody/tr[contains(@class,'el-table__row')]"

    # 操作按钮
    SAVE_BTN            = "xpath=//button[.//span[normalize-space(.)='Save']]"
    TEST_CONNECTION_BTN = "xpath=//button[.//span[normalize-space(.)='Test Connection']]"

    # 结果提示
    RESULT_MSG = "xpath=//div[contains(@class,'el-message') or contains(@class,'el-notification__content') or contains(@class,'el-message-box')]"

    # Virtual Devices 相关
    # RPP：Virtual Devices 归属「Monitoring」顶级组（HMI 为 'Devices' 设备菜单）
    DEVICES_NAV_BTN      = "xpath=//div[contains(@class,'nav-item-menu') and (.//span[normalize-space(.)='Monitoring'] or .//span[normalize-space(.)='Devices'])]"
    VIRTUAL_DEVICES_MENU = "xpath=//li[contains(normalize-space(.),'Virtual Devices')] | //a[contains(normalize-space(.),'Virtual Devices')]"
    READING_TAB          = "xpath=//*[contains(@class,'el-tabs__item') and contains(normalize-space(.),'Reading')]"
    READING_TABLE_ROWS   = "xpath=//div[contains(@class,'el-table__body-wrapper')]//tbody/tr[contains(@class,'el-table__row')]"

    # ── 内部辅助 ──────────────────────────────────────────────────────────────

    def _dismiss_overlay_dialog(self):
        """关闭页面可能残留的确认弹框及模态层，避免遮挡后续操作。"""
        try:
            # 覆盖两类：el-message-box 确认框 和 el-overlay-dialog 全屏模态（如 Parameter Config）
            dialogs = self.page.locator(
                "xpath=//div["
                "contains(@class,'el-overlay-message-box')"
                " or contains(@class,'el-message-box')"
                " or (@role='dialog' and @aria-modal='true')"
                " or contains(@class,'el-overlay-dialog')"
                "]"
            ).all()
            for dlg in dialogs:
                if not dlg.is_visible():
                    continue
                # 优先 Escape 键关闭模态
                try:
                    self.page.keyboard.press("Escape")
                    self.page.wait_for_timeout(400)
                    logging.info("[Web] 用 Escape 关闭残留模态层")
                    continue
                except Exception:
                    pass
                for xpath in [
                    ".//button[.//span[normalize-space(.)='Cancel']]",
                    ".//button[.//span[normalize-space(.)='取消']]",
                    ".//button[contains(@class,'el-message-box__headerbtn')]",
                    ".//button[.//i[contains(@class,'el-icon-close')]]",
                    ".//button[.//span[normalize-space(.)='Yes']]",
                    ".//button[.//span[normalize-space(.)='Continue']]",
                    ".//button[.//span[normalize-space(.)='确认']]",
                    ".//button[contains(@class,'el-button')]",
                ]:
                    btns = dlg.locator(f"xpath={xpath}").all()
                    if btns:
                        btns[0].evaluate("el => el.click()")
                        logging.info("[Web] 已关闭页面残留确认框")
                        self.page.wait_for_timeout(500)
                        break
        except Exception as e:
            logging.debug(f"[Web] 关闭残留弹框时忽略异常：{e}")

    def _clear_and_fill(self, selector: str, value: str):
        """清空并填写输入框。"""
        loc = self.page.locator(selector)
        loc.fill(value)

    # ── 公开方法 ──────────────────────────────────────────────────────────────

    @staticmethod
    def _click_first_visible(page, selectors: list, timeout: int = 10000):
        """按顺序尝试多个选择器，点击第一个可见元素，全部失败则抛出。"""
        last_err = None
        for sel in selectors:
            try:
                loc = page.locator(sel).first
                loc.wait_for(state="visible", timeout=timeout)
                # click 也使用相同超时，避免默认 30s 造成长时间阻塞
                loc.click(timeout=timeout)
                return
            except Exception as e:
                last_err = e
        raise RuntimeError(f"None of the selectors matched a visible element: {selectors}") from last_err

    def _ensure_logged_in(self):
        """若当前在登录页，自动重新登录后等待页面就绪。"""
        try:
            if "login" not in (self.page.url or ""):
                return
            logging.info("[Nav] 检测到登录页，执行重新登录")
            from urllib.parse import urlparse
            parsed = urlparse(self.page.url)
            base = f"{parsed.scheme}://{parsed.netloc}"
            target = self._gateway_url or base
            self.page.goto(target, wait_until="domcontentloaded")
            user_inp = self.page.locator("input[placeholder='Enter User Name']").first
            user_inp.wait_for(state="visible", timeout=8000)
            user_inp.fill(self._username)
            self.page.locator("input[placeholder='Enter Password']").first.fill(self._password)
            self.page.locator("xpath=//button[span[text()='Sign In']]").first.click()
            try:
                self.page.wait_for_selector(
                    "div.nav-item-menu, .left-nav-item, .el-aside",
                    state="visible", timeout=15000,
                )
            except Exception:
                self.page.wait_for_timeout(500)
            cancel = self.page.locator("button:has-text('Cancel')")
            if cancel.count() > 0 and cancel.first.is_visible():
                cancel.first.click()
                self.page.wait_for_timeout(200)
            logging.info("[Nav] 重新登录完成")
        except Exception as e:
            logging.warning(f"[Nav] 重新登录失败（将继续尝试导航）：{e}")

    def navigate_to_aws_iot(self):
        """UI 点击导航：设备按钮 → Protocols → AWS IoT Tab，约 1s 完成。"""
        self._dismiss_overlay_dialog()
        self._ensure_logged_in()

        if "awsIot" in self.page.url:
            logging.info("[Nav] 已在 AWS IoT 页，跳过导航")
            return

        logging.info(f"[Nav] 当前 URL: {self.page.url}，执行 UI 导航")

        for sel, label in [
            (self.DEVICE_SETTINGS_BTN, "设备导航按钮"),
            (self.PROTOCOLS_MENU,      "Protocols 菜单"),
            (self.AWS_IOT_TAB,         "AWS IoT Tab"),
        ]:
            try:
                el = self.page.locator(sel).first
                el.wait_for(state="visible", timeout=5000)
                el.click()
                self.page.wait_for_timeout(300)
            except Exception as e:
                logging.warning(f"[Nav] 点击 {label} 失败：{e}")

        logging.info(f"[Nav] 导航完成，当前 URL：{self.page.url}")

    def is_enabled(self) -> bool:
        """检查当前是否已处于 Enable 状态。"""
        try:
            label = self.page.locator(self.ENABLE_RADIO)
            cls = label.get_attribute('class') or ''
            return 'is-checked' in cls
        except Exception:
            return False

    def ensure_enabled(self):
        """若当前为 Disable，切换到 Enable 后等待表单出现。"""
        if not self.is_enabled():
            logging.info("当前为 Disable，切换为 Enable")
            self.page.locator(self.ENABLE_RADIO).click()
            # 等待输入框出现（表单在 Enable 后才展开）
            self.page.wait_for_selector(self.CLIENT_ID_INPUT, state="visible", timeout=10000)

    def set_client_id(self, client_id: str):
        self._clear_and_fill(self.CLIENT_ID_INPUT, client_id)

    def set_url(self, url: str):
        self._clear_and_fill(self.URL_INPUT, url)

    def set_topic(self, topic: str):
        self._clear_and_fill(self.TOPIC_INPUT, topic)

    def set_interval(self, interval_text: str):
        """interval_text 为下拉可见文字，如 '30 seconds'。"""
        logging.info(f"设置 Interval：{interval_text}")
        # 点击 el-select 容器触发下拉（用 Playwright 原生 click，勿用 JS evaluate）
        self.page.locator(self.INTERVAL_DROPDOWN).click()
        self.page.wait_for_timeout(600)
        option_xpath = (
            f"xpath=//li[contains(@class,'el-select-dropdown__item') "
            f"and normalize-space(.)='{interval_text}']"
        )
        self.page.wait_for_selector(option_xpath, state="visible", timeout=5000)
        self.page.locator(option_xpath).first.click()
        self.page.wait_for_timeout(300)

    def upload_cert_file(self, file_path: str):
        self.page.locator(self.CERT_FILE_INPUT).set_input_files(file_path)
        self.page.wait_for_timeout(1000)

    def upload_key_file(self, file_path: str):
        self.page.locator(self.KEY_FILE_INPUT).set_input_files(file_path)
        self.page.wait_for_timeout(1000)

    def select_unchecked_devices(self):
        """逐行检查设备复选框，仅勾选尚未勾选的行。"""
        rows = self.page.locator(self.DEVICE_TABLE_ROWS).all()
        logging.info(f"设备列表共 {len(rows)} 行，检查勾选状态")
        for idx, row in enumerate(rows):
            cb_labels = row.locator("xpath=.//label[contains(@class,'el-checkbox')]").all()
            if not cb_labels:
                logging.warning(f"  第 {idx + 1} 行：未找到复选框，跳过")
                continue
            cb = cb_labels[0]
            cls = cb.get_attribute('class') or ''
            if 'is-checked' in cls:
                logging.info(f"  第 {idx + 1} 行：已勾选，跳过")
            else:
                cb.evaluate("el => el.scrollIntoView({block:'center'})")
                self.page.wait_for_timeout(200)
                cb.evaluate("el => el.click()")
                logging.info(f"  第 {idx + 1} 行：已勾选设备")
                self.page.wait_for_timeout(300)

    def _transfer_left_count(self, dialog_locator) -> int:
        """返回 transfer 左侧（未选）区域的条目数，多个 XPath fallback 取最大值。"""
        for xpath in (
            # Element Plus el-transfer 左面板条目
            "xpath=(.//*[contains(@class,'el-transfer-panel')])[1]"
            "//*[contains(@class,'el-transfer-panel__item')]",
            # 自定义 transfer-left > option-item
            "xpath=.//div[contains(@class,'transfer-left')]"
            "//div[contains(@class,'option-item')]",
            # 通用：左面板内任意 checkbox label
            "xpath=(.//*[contains(@class,'el-transfer-panel')])[1]"
            "//label[contains(@class,'el-checkbox')]",
        ):
            try:
                cnt = len(dialog_locator.locator(xpath).all())
                if cnt > 0:
                    return cnt
            except Exception:
                continue
        return 0

    def _all_params_selected(self, dialog_locator) -> bool:
        """检查 transfer 左侧 Not Selected 区域是否为空。"""
        cnt = self._transfer_left_count(dialog_locator)
        logging.info(f"    [check] transfer-left 剩余未选项：{cnt}")
        return cnt == 0

    def _get_visible_dialog(self):
        """返回当前可见的 el-dialog locator。"""
        self.page.wait_for_selector(
            "xpath=//div[contains(@class,'el-overlay-dialog')]//div[contains(@class,'el-dialog__body')]",
            state="visible",
            timeout=10000,
        )
        wrappers = self.page.locator(
            "xpath=//div[contains(@class,'el-overlay-dialog')]"
        ).all()
        for w in wrappers:
            if w.is_visible():
                dialog = w.locator("xpath=.//div[contains(@class,'el-dialog')]")
                if dialog.count() > 0:
                    return dialog.first
        raise RuntimeError("未找到可见的弹窗")

    def _dismiss_message_box(self):
        """关闭可能出现的 el-message-box 确认框，不关闭参数配置弹窗（el-overlay-dialog）。"""
        try:
            msg_boxes = self.page.locator(
                "xpath=//div[contains(@class,'el-overlay-message-box')]"
            ).all()
            for mb in msg_boxes:
                if not mb.is_visible():
                    continue
                ok_btns = mb.locator(
                    "xpath=.//button[contains(@class,'el-button--primary')]"
                ).all()
                if ok_btns:
                    ok_btns[-1].evaluate("el => el.click()")
                    logging.info("已关闭 el-message-box 确认框")
                    self.page.wait_for_timeout(300)
                    return
                # 没有 primary 按钮则按 ESC
                self.page.keyboard.press("Escape")
                self.page.wait_for_timeout(300)
        except Exception:
            pass

    def _close_pt_dropdown_safely(self, dialog, pt_wrapper=None):
        """收起 Parameter Type 下拉：再次点击 wrapper 触发 el-select toggle 关闭。
        绝不使用 Escape——Escape 会同时关闭 el-dialog（close-on-press-escape）。
        """
        if pt_wrapper is not None:
            try:
                expanded = pt_wrapper.get_attribute("aria-expanded", timeout=300) or "false"
                if expanded.lower() == "true":
                    pt_wrapper.click()  # toggle close，焦点保留在 wrapper 上
                    self.page.wait_for_timeout(300)
                    return
            except Exception:
                pass
        # fallback：点击 dialog body 左上角空白处，仅在无法读取 aria-expanded 时使用
        try:
            body = dialog.locator("xpath=.//div[contains(@class,'el-dialog__body')]").first
            body.click(position={"x": 5, "y": 5})
            self.page.wait_for_timeout(300)
        except Exception:
            pass

    def _configure_parameter_dialog(self):
        """
        在 Parameter Selection 弹窗中配置参数：
        - 有 Parameter Type 下拉（普通设备）：逐个类型全选
        - 无 Parameter Type 下拉（虚拟设备）：直接点 All 一次性全选
        最后点击 Confirm 关闭弹窗。
        """
        dialog = None
        try:
            dialog = self._get_visible_dialog()

            # 判断是否有 Parameter Type 下拉（普通设备）
            pt_wrappers = dialog.locator(
                "xpath=.//div[contains(@class,'el-form-item')]"
                "[.//label[contains(normalize-space(.),'Parameter Type')]]"
                "//div[contains(@class,'el-select__wrapper')]"
            ).all()

            if not pt_wrappers:
                # 虚拟设备：无分组，直接点击 All
                logging.info("  未检测到 Parameter Type 下拉（虚拟设备），直接点击 All")
                all_btns = dialog.locator(
                    "xpath=.//button[normalize-space(.)='All' or .//span[normalize-space(.)='All']]"
                ).all()
                if all_btns:
                    all_btns[0].evaluate("el => el.click()")
                    logging.info("  [All] → 已点击 All，全选参数")
                    self.page.wait_for_timeout(300)
                else:
                    logging.warning("  [All] → 未找到 All 按钮")
                return

            # 普通设备：逐 Parameter Type 全选
            pt_wrapper = pt_wrappers[0]

            # 打开下拉，采集所有选项文本
            # 用 Playwright 原生 click（正确触发 Vue 事件），不用 keyboard.press("Escape")
            pt_wrapper.click()
            self.page.wait_for_timeout(500)

            # 通过 aria-controls 把 popper 查找范围锁定到 PT 专属下拉，隔离 Interval 等其他 select
            popper_id = None
            try:
                popper_id = pt_wrapper.get_attribute("aria-controls", timeout=1000)
            except Exception:
                pass

            # 读取当前可见的所有 Parameter Type 选项（按可见性过滤，不依赖 popper id）
            # 若首次为空，额外等待最多 2 次，防止下拉加载延迟
            option_texts = []
            for _extra_wait in (0, 500, 1000):
                if _extra_wait:
                    self.page.wait_for_timeout(_extra_wait)
                all_opt_els = self.page.locator(
                    "xpath=//li[contains(@class,'el-select-dropdown__item')]"
                ).all()
                option_texts = [
                    o.inner_text().strip()
                    for o in all_opt_els
                    if o.is_visible() and o.inner_text().strip()
                ]
                if option_texts:
                    break
            logging.info(f"Parameter Type 选项：{option_texts}")

            # 不提前关闭下拉——i=0 直接使用已打开的 popper；i>0 由选项点击后自动关闭再重新打开
            for i, opt_text in enumerate(option_texts):
                self._dismiss_message_box()

                if i > 0:
                    pt_wrapper.click()
                    self.page.wait_for_timeout(800)

                clicked = False
                opt_scan = self.page.locator(
                    "xpath=//li[contains(@class,'el-select-dropdown__item')]"
                ).all()
                for o in opt_scan:
                    try:
                        if o.is_visible() and o.inner_text().strip() == opt_text:
                            o.click()
                            clicked = True
                            break
                    except Exception:
                        continue
                if not clicked:
                    logging.warning(f"  [{opt_text}] → 未找到可见选项，关闭下拉跳过")
                    self._close_pt_dropdown_safely(dialog, pt_wrapper)
                self.page.wait_for_timeout(500)

                # 若参数已全部选中，跳过
                if self._all_params_selected(dialog):
                    logging.info(f"  [{opt_text}] → 参数已全选，跳过")
                    continue

                # 点击 All 按钮（参数列表全选）
                all_btns = dialog.locator(
                    "xpath=.//button[normalize-space(.)='All' or .//span[normalize-space(.)='All']]"
                ).all()
                if all_btns:
                    all_btns[0].evaluate("el => el.click()")
                    logging.info(f"  [{opt_text}] → 已点击 All")
                    self.page.wait_for_timeout(500)
                    # All 点击后可能触发 el-message-box 确认框（如 Realtime 档）
                    self._dismiss_message_box()
                    # 不在此处 break：必须继续遍历所有 Parameter Type 组，全部完成后才 Confirm
                else:
                    logging.warning(f"  [{opt_text}] → All 按钮未找到")

        except Exception as e:
            logging.error(f"配置 Parameter 弹窗时出错：{e}")
        finally:
            # ── 关闭弹窗：优先从 aria-label 重新定位（避免使用可能已过期的引用）────
            def _find_visible_param_dialog():
                try:
                    dlgs = self.page.locator(
                        "xpath=//div[@aria-label='AWSIoT Parameter Config']"
                    ).all()
                    for d in reversed(dlgs):
                        if d.is_visible():
                            return d
                except Exception:
                    pass
                return None

            def _dialog_is_open():
                d = _find_visible_param_dialog()
                return d is not None

            # 找 Confirm 按钮并点击（用原生 .click() 触发 Vue 事件）
            _dlg = _find_visible_param_dialog()
            confirmed = False
            if _dlg is not None:
                _CONF_XPATHS = [
                    "xpath=.//div[contains(@class,'dialog-footer')]"
                    "//button[.//span[normalize-space(.)='Confirm'] or normalize-space(.)='Confirm']",
                    "xpath=.//div[contains(@class,'el-dialog__footer')]"
                    "//button[contains(@class,'el-button--primary')][last()]",
                    "xpath=.//button[.//span[normalize-space(.)='Confirm']]",
                    "xpath=.//button[normalize-space(.)='Confirm']",
                    "xpath=.//div[contains(@class,'dialog-footer')]"
                    "//button[contains(@class,'el-button--primary')][last()]",
                ]
                for conf_xpath in _CONF_XPATHS:
                    try:
                        btns = _dlg.locator(conf_xpath).all()
                        visible_btns = [b for b in btns if b.is_visible()]
                        if visible_btns:
                            visible_btns[-1].click(timeout=5000)
                            logging.info("已点击 Confirm，关闭弹窗")
                            confirmed = True
                            self.page.wait_for_timeout(800)
                            break
                    except Exception:
                        continue

            # 若点击 Confirm 后仍然开着，再按 ESC
            if not confirmed or _dialog_is_open():
                try:
                    self.page.keyboard.press("Escape")
                    logging.info("已按 ESC 关闭弹窗")
                    self.page.wait_for_timeout(500)
                except Exception:
                    pass

            # 等待弹窗彻底消失（优先用 Playwright 的 wait_for_selector hidden/detached）
            _closed = False
            for _state in ("hidden", "detached"):
                try:
                    self.page.wait_for_selector(
                        "xpath=//div[@aria-label='AWSIoT Parameter Config']",
                        state=_state,
                        timeout=5000,
                    )
                    logging.info(f"弹窗已关闭（{_state}）")
                    _closed = True
                    break
                except Exception:
                    pass
            if not _closed:
                logging.warning("弹窗关闭等待超时，将自行继续执行")
            self.page.wait_for_timeout(300)

    def configure_all_devices_parameters(self, checked_only: bool = False):
        """
        遍历设备列表，点击 Parameter Selection 按钮配置参数。
        checked_only=True 时只处理已勾选的设备行。
        """
        rows = self.page.locator(self.DEVICE_TABLE_ROWS).all()
        logging.info(f"设备列表共 {len(rows)} 行，逐行配置参数选择（checked_only={checked_only}）")
        for idx, row in enumerate(rows):
            if checked_only:
                cb_labels = row.locator("xpath=.//label[contains(@class,'el-checkbox')]").all()
                if not cb_labels or 'is-checked' not in (cb_labels[0].get_attribute('class') or ''):
                    logging.info(f"  第 {idx + 1} 行：未勾选，跳过参数配置")
                    continue
            btns = row.locator("xpath=.//button").all()
            if not btns:
                logging.warning(f"  第 {idx + 1} 行未找到按钮，跳过")
                continue
            # 优先找包含 "Select" / "Parameter" 文字的按钮
            target_btn = None
            for b in btns:
                txt = (b.inner_text() or '').strip()
                if any(kw in txt for kw in ('Select', 'Parameter', '选择', '参数')):
                    target_btn = b
                    break
            if target_btn is None:
                target_btn = btns[-1]
            label = (target_btn.inner_text() or '').strip()
            logging.info(f"  第 {idx + 1} 行：点击按钮「{label or '(icon)'}」")
            target_btn.evaluate("el => el.scrollIntoView({block:'center'})")
            self.page.wait_for_timeout(300)
            target_btn.evaluate("el => el.click()")
            self.page.wait_for_timeout(1500)
            self._configure_parameter_dialog()

    def _clear_parameter_dialog(self):
        """
        在 Parameter Selection 弹窗中点击 Clear（清除所有已选参数），再点 Confirm 关闭。
        调用前须已打开弹窗。
        """
        try:
            dialog = self._get_visible_dialog()
            clear_btns = dialog.locator(
                "xpath=.//button[normalize-space(.)='Clear' or .//span[normalize-space(.)='Clear']]"
            ).all()
            if clear_btns:
                clear_btns[0].evaluate("el => el.click()")
                logging.info("  已点击 Clear，清除所有参数")
                self.page.wait_for_timeout(500)
                self._dismiss_message_box()
            else:
                logging.warning("  未找到 Clear 按钮，跳过")
        except Exception as e:
            logging.error(f"  _clear_parameter_dialog 出错：{e}")
        finally:
            # 点击 Confirm 关闭弹窗
            _CONF_XPATHS = [
                "xpath=//div[@aria-label='AWSIoT Parameter Config']"
                "//div[contains(@class,'dialog-footer')]"
                "//button[.//span[normalize-space(.)='Confirm'] or normalize-space(.)='Confirm']",
                "xpath=//div[@aria-label='AWSIoT Parameter Config']"
                "//button[contains(@class,'el-button--primary')][last()]",
            ]
            confirmed = False
            for xpath in _CONF_XPATHS:
                try:
                    btns = [b for b in self.page.locator(xpath).all() if b.is_visible()]
                    if btns:
                        btns[-1].click(timeout=5000)
                        logging.info("  已点击 Confirm，关闭弹窗")
                        confirmed = True
                        self.page.wait_for_timeout(800)
                        break
                except Exception:
                    continue
            if not confirmed:
                try:
                    self.page.keyboard.press("Escape")
                    self.page.wait_for_timeout(500)
                except Exception:
                    pass
            for _state in ("hidden", "detached"):
                try:
                    self.page.wait_for_selector(
                        "xpath=//div[@aria-label='AWSIoT Parameter Config']",
                        state=_state, timeout=5000,
                    )
                    break
                except Exception:
                    pass
            self.page.wait_for_timeout(300)

    def clear_all_devices_parameters(self, checked_only: bool = True):
        """
        遍历设备列表，为每台（已勾选）设备打开参数弹窗并点击 Clear + Confirm。
        checked_only=True（默认）只处理已勾选的行。
        """
        rows = self.page.locator(self.DEVICE_TABLE_ROWS).all()
        logging.info(f"设备列表共 {len(rows)} 行，逐行清除参数（checked_only={checked_only}）")
        for idx, row in enumerate(rows):
            if checked_only:
                cb_labels = row.locator("xpath=.//label[contains(@class,'el-checkbox')]").all()
                if not cb_labels or 'is-checked' not in (cb_labels[0].get_attribute('class') or ''):
                    logging.info(f"  第 {idx + 1} 行：未勾选，跳过")
                    continue
            btns = row.locator("xpath=.//button").all()
            if not btns:
                logging.warning(f"  第 {idx + 1} 行未找到按钮，跳过")
                continue
            target_btn = None
            for b in btns:
                txt = (b.inner_text() or '').strip()
                if any(kw in txt for kw in ('Select', 'Parameter', '选择', '参数')):
                    target_btn = b
                    break
            if target_btn is None:
                target_btn = btns[-1]
            label = (target_btn.inner_text() or '').strip()
            logging.info(f"  第 {idx + 1} 行：点击「{label or '(icon)'}」打开参数弹窗")
            target_btn.evaluate("el => el.scrollIntoView({block:'center'})")
            self.page.wait_for_timeout(300)
            target_btn.evaluate("el => el.click()")
            self.page.wait_for_timeout(1500)
            self._clear_parameter_dialog()

    def save(self):
        btn = self.page.locator(self.SAVE_BTN)
        btn.evaluate("el => el.scrollIntoView({block:'center'})")
        self.page.wait_for_timeout(300)
        btn.evaluate("el => el.click()")
        # 等待保存成功提示出现后立即继续，最长等 5s；无提示则兜底等 1s
        try:
            self.page.wait_for_selector(
                "xpath=//div[contains(@class,'el-message--success')]"
                " | //div[contains(@class,'el-notification--success')]",
                state="visible",
                timeout=5000,
            )
        except Exception:
            self.page.wait_for_timeout(1000)
        logging.info(f"已点击 Save，当前 URL：{self.page.url}")

    def test_connection(self) -> str:
        """每隔 1s 点击一次 Test Connection，读取页面上方提示，直到出现最终结果。"""
        self._dismiss_overlay_dialog()
        self.page.wait_for_selector(self.TEST_CONNECTION_BTN, state="visible", timeout=30000)
        btn = self.page.locator(self.TEST_CONNECTION_BTN)
        btn.evaluate("el => el.scrollIntoView({block:'center'})")
        self.page.wait_for_timeout(300)

        _SKIP_STATES = {
            "connecting", "connecting...", "testing...", "test connecting",
            "test connection", "in progress", "please wait",
        }

        def _is_in_progress(txt: str) -> bool:
            t = txt.lower().strip().rstrip(".")
            if t in _SKIP_STATES or t == "":
                return True
            if ":" in t:
                suffix = t.rsplit(":", 1)[-1].strip().rstrip(".")
                if suffix in _SKIP_STATES:
                    return True
            return False

        def _read_result() -> str:
            """读取页面上方最新的提示文字，返回非空非等待态的文本，否则返回空串。"""
            for xpath in (
                "xpath=//*[contains(normalize-space(.),'Test Result')]",
                "xpath=//*[contains(normalize-space(.),'test result')]",
            ):
                for el in self.page.locator(xpath).all():
                    try:
                        txt = el.inner_text(timeout=500).strip()
                        if not txt or len(txt) > 500:
                            continue
                        if _is_in_progress(txt):
                            continue
                        if txt.lower().strip() == "test result":
                            continue
                        return txt
                    except Exception:
                        continue
            for el in self.page.locator(self.RESULT_MSG).all():
                try:
                    txt = el.inner_text(timeout=500).strip()
                    if not txt or len(txt) > 300:
                        continue
                    if _is_in_progress(txt):
                        continue
                    if "save" in txt.lower():
                        continue
                    return txt
                except Exception:
                    continue
            return ""

        # 记录点击前页面上已有的结果文本（可能是上次 test_connection 残留的旧值）
        _stale_result = _read_result()
        if _stale_result:
            logging.info(f"检测到旧结果（可能残留自上次连接）：{_stale_result!r}，等待新结果刷新")

        # 必须先观察到「进行中」（_read_result 返回空），再接受新的最终结果
        # 这样即使页面残留旧结果也不会被误读
        _saw_progress = not _stale_result   # 若点击前就是空，视为已进入进行中状态

        _timeout_secs = 600
        _last_log_time = start_time = time.time()
        logging.info("开始循环点击 Test Connection（每 1s 一次），等待 Success…")

        deadline = start_time + _timeout_secs
        while time.time() < deadline:
            try:
                btn.click(timeout=5000)
            except Exception as e:
                logging.warning(f"点击 Test Connection 异常（将继续重试）：{e}")

            self.page.wait_for_timeout(1000)
            elapsed = int(time.time() - start_time)

            result = _read_result()
            if not result:
                # 空结果 = 进行中状态，标记已见过进行中
                _saw_progress = True
            elif result != _stale_result:
                # 结果与点击前不同，一定是本次连接的新鲜结果
                logging.info(f"[{elapsed}s] Test Connection 结果（结果已变更）：{result!r}")
                return result
            elif _saw_progress:
                # 已经历过进行中状态，此结果为本次连接的新鲜结果
                logging.info(f"[{elapsed}s] Test Connection 结果：{result!r}")
                return result
            else:
                # 尚未见到进行中状态，当前结果与旧值相同，可能仍是旧残留，继续等
                logging.debug(f"[{elapsed}s] 疑似旧结果仍显示：{result!r}，继续等待刷新")

            now = time.time()
            if now - _last_log_time >= 60:
                _last_log_time = now
                logging.info(f"[{elapsed}s] 持续点击等待结果（最长 {_timeout_secs}s）…")

        logging.warning(f"未在 {_timeout_secs} 秒内检测到 Test Connection 结果")
        return "未检测到结果提示"

    def get_checked_device_names(self) -> list:
        """返回设备表格中所有已勾选行的设备名称列表。"""
        rows = self.page.locator(self.DEVICE_TABLE_ROWS).all()
        names = []
        for row in rows:
            cb_labels = row.locator("xpath=.//label[contains(@class,'el-checkbox')]").all()
            if not cb_labels:
                continue
            if 'is-checked' not in (cb_labels[0].get_attribute('class') or ''):
                continue
            # 找到不含 button / input 的文本单元格作为设备名
            for td in row.locator("xpath=.//td").all():
                if td.locator("xpath=.//*[self::button or self::input]").count() > 0:
                    continue
                text = td.inner_text().strip().split('\n')[0].strip()
                if text:
                    names.append(text)
                    break
        return names

    def get_all_devices_from_table(self) -> list:
        """
        读取设备选择表格中所有设备。
        返回 [{"name": str, "is_virtual": bool}, ...] 列表。
        调用前请确保页面处于 Enable 状态（表格可见）。
        """
        rows = self.page.locator(self.DEVICE_TABLE_ROWS).all()
        devices = []
        for row in rows:
            cells = row.locator("xpath=.//td").all()
            is_virtual = any(
                "virtual" in (c.inner_text() or "").lower()
                for c in cells
            )
            name = ""
            for td in cells:
                if td.locator("xpath=.//*[self::button or self::input]").count() > 0:
                    continue
                text = td.inner_text().strip().split("\n")[0].strip()
                if text:
                    name = text
                    break
            if name:
                devices.append({"name": name, "is_virtual": is_virtual})
        return devices

    def select_only_device(self, device_name: str = "", skip_names: list = None,
                           exact: bool = False):
        """取消所有已勾选设备，再单独勾选名称匹配的设备。
        exact=True 时精确匹配（大小写不敏感），exact=False（默认）时子串匹配。
        device_name 为空时，从物理（非虚拟）设备中随机选取，排除 skip_names。"""
        import random
        if skip_names is None:
            skip_names = []
        rows = self.page.locator(self.DEVICE_TABLE_ROWS).all()

        if not device_name:
            physical_names = []
            for row in rows:
                cells = row.locator("xpath=.//td").all()
                if any("Virtual" in (c.inner_text() or "") for c in cells):
                    continue
                for td in cells:
                    if td.locator("xpath=.//*[self::button or self::input]").count() > 0:
                        continue
                    text = td.inner_text().strip().split("\n")[0].strip()
                    if text and text not in skip_names:
                        physical_names.append(text)
                        break
            if not physical_names:
                logging.warning("select_only_device: 未找到可用的物理设备，跳过选择")
                return
            device_name = random.choice(physical_names)
            exact = True  # 随机选出的名称必须精确匹配，防止子串命中同型其他设备
            logging.info(f"select_only_device: 自动随机选取物理设备 {device_name!r}")
            print(f"\n[Setup] 自动随机选取物理设备: {device_name!r}")

        logging.info(f"select_only_device: 目标={device_name!r}，共 {len(rows)} 行")
        for idx, row in enumerate(rows):
            cb_labels = row.locator("xpath=.//label[contains(@class,'el-checkbox')]").all()
            if not cb_labels:
                continue
            cb = cb_labels[0]
            is_checked = "is-checked" in (cb.get_attribute("class") or "")
            row_name = ""
            for td in row.locator("xpath=.//td").all():
                if td.locator("xpath=.//*[self::button or self::input]").count() > 0:
                    continue
                text = td.inner_text().strip().split("\n")[0].strip()
                if text:
                    row_name = text
                    break
            if exact:
                should_check = device_name.lower() == row_name.lower()
            else:
                should_check = device_name.lower() in row_name.lower()
            if should_check != is_checked:
                cb.evaluate("el => el.scrollIntoView({block:'center'})")
                self.page.wait_for_timeout(200)
                cb.evaluate("el => el.click()")
                action = "勾选" if should_check else "取消"
                logging.info(f"  第 {idx+1} 行 [{row_name}]：{action}")
            else:
                logging.info(f"  第 {idx+1} 行 [{row_name}]：无需操作")

    def select_devices(self, device_names: list):
        """按名称列表批量勾选设备，不在列表中的行取消勾选。
        匹配规则：device_name.lower() in row_name.lower()（与 select_only_device 一致）。"""
        if not device_names:
            logging.warning("select_devices: device_names 为空，不做任何操作")
            return
        rows = self.page.locator(self.DEVICE_TABLE_ROWS).all()
        logging.info(f"select_devices: 目标={device_names}，共 {len(rows)} 行")
        for idx, row in enumerate(rows):
            cb_labels = row.locator("xpath=.//label[contains(@class,'el-checkbox')]").all()
            if not cb_labels:
                continue
            cb = cb_labels[0]
            is_checked = "is-checked" in (cb.get_attribute("class") or "")
            row_name = ""
            for td in row.locator("xpath=.//td").all():
                if td.locator("xpath=.//*[self::button or self::input]").count() > 0:
                    continue
                text = td.inner_text().strip().split("\n")[0].strip()
                if text:
                    row_name = text
                    break
            should_check = any(dn.lower() in row_name.lower() for dn in device_names)
            if should_check != is_checked:
                cb.evaluate("el => el.scrollIntoView({block:'center'})")
                self.page.wait_for_timeout(200)
                cb.evaluate("el => el.click()")
                action = "勾选" if should_check else "取消"
                logging.info(f"  第 {idx+1} 行 [{row_name}]：{action}")
            else:
                logging.info(f"  第 {idx+1} 行 [{row_name}]：无需操作")

    def disable(self):
        """点击 Disable 单选并保存，关闭 AWS IoT 推送。已是 Disable 状态则跳过。"""
        # 先关闭可能残留的弹窗
        try:
            confirm_btn = self.page.locator(
                "xpath=//div[contains(@class,'el-overlay-message-box')]//button[contains(@class,'el-button--primary')]"
            )
            if confirm_btn.count() > 0 and confirm_btn.first.is_visible():
                confirm_btn.first.evaluate("el => el.click()")
                self.page.wait_for_timeout(500)
        except Exception:
            pass
        if self.is_enabled():
            logging.info("切换 AWS IoT 为 Disable")
            self.page.locator(self.DISABLE_RADIO).evaluate("el => el.click()")
            self.page.wait_for_timeout(500)
            self.save()
            logging.info("AWS IoT 已 Disable 并保存")
        else:
            logging.info("AWS IoT 已是 Disable 状态，跳过")

    def configure(self, client_id: str, url: str, topic: str,
                  cert_file: str, key_file: str,
                  interval: str = "30 seconds") -> str:
        """
        完整配置流程：
        1. 确保 Enable 已勾选
        2. 填写连接参数
        3. 上传证书文件
        4. 全选设备
        5. 保存
        6. 测试连接并返回结果
        """
        self.ensure_enabled()
        self.set_client_id(client_id)
        self.set_url(url)
        self.set_topic(topic)
        self.set_interval(interval)
        self.upload_cert_file(cert_file)
        self.upload_key_file(key_file)
        self.select_unchecked_devices()
        self.configure_all_devices_parameters()
        self.save()
        result = self.test_connection()
        return result

    def _click_tab(self, tab_name: str, timeout: float = 15.0):
        """点击与 Reading/Configuration 同级的 tab 标签。"""
        _xpaths = [
            f"xpath=//*[contains(@class,'el-tabs__item') and contains(normalize-space(.), '{tab_name}')]",
            f"xpath=//*[@role='tab' and contains(normalize-space(.), '{tab_name}')]",
            f"xpath=//span[normalize-space(.)='{tab_name}']/ancestor::*[@role='tab' or contains(@class,'tab')][1]",
        ]
        tab = None
        deadline = time.time() + timeout
        while time.time() < deadline:
            for xp in _xpaths:
                els = self.page.locator(xp).all()
                if els:
                    tab = els[0]
                    break
            if tab is not None:
                break
            self.page.wait_for_timeout(500)
        if tab is None:
            raise RuntimeError(f"未找到 '{tab_name}' 标签")
        tab.evaluate("el => el.click()")
        self.page.wait_for_timeout(1500)


    def get_virtual_device_readings(self, device_name: str) -> dict:
        """
        导航到 Devices → Virtual Devices → 指定设备，分别读取：
          - Configuration 标签：Post Label（MQTT 上报的参数名）
          - Reading 标签：Value 列（值 + 单位，空格分隔）
        按行索引合并后返回 {post_label: {"value": str, "unit": str}}
        """
        # 1. 进入 Devices 视图
        self.page.wait_for_selector(self.DEVICES_NAV_BTN, state="visible", timeout=10000)
        self.page.locator(self.DEVICES_NAV_BTN).evaluate("el => el.click()")
        self.page.wait_for_timeout(1500)

        # 2. 点击 Virtual Devices
        self.page.wait_for_selector(self.VIRTUAL_DEVICES_MENU, state="visible", timeout=10000)
        self.page.locator(self.VIRTUAL_DEVICES_MENU).first.evaluate("el => el.click()")
        self.page.wait_for_timeout(1500)

        # 3. 点击目标设备名
        _dev_xpaths = [
            f"xpath=//div[contains(@class,'link-url') and normalize-space(.)='{device_name}']",
            f"xpath=//td[contains(normalize-space(.),'{device_name}')]//div[contains(@class,'link-url')]",
            f"xpath=//td[contains(normalize-space(.),'{device_name}')]//a",
            f"xpath=//a[normalize-space(.)='{device_name}']",
        ]
        dev_link = None
        deadline = time.time() + 10
        while time.time() < deadline:
            for xp in _dev_xpaths:
                els = self.page.locator(xp).all()
                if els:
                    dev_link = els[0]
                    break
            if dev_link is not None:
                break
            self.page.wait_for_timeout(500)
        if dev_link is None:
            raise RuntimeError(f"未找到设备 {device_name!r} 的链接")

        dev_link.evaluate("el => el.scrollIntoView({block:'center'})")
        self.page.wait_for_timeout(300)
        dev_link.evaluate("el => el.click()")
        self.page.wait_for_timeout(2500)

        # 4. Configuration 标签 → 读 Post Label（MQTT 参数名）
        # Configuration 页用 el-form-item 表单布局，Post Label 在 label 旁的 input 里
        self._click_tab("Configuration")
        _post_label_inputs = self.page.locator(
            "xpath=//div[contains(@class,'el-form-item')][.//label[normalize-space(.)='Post Label']]//input"
        ).all()
        post_labels: list[str] = []
        for inp in _post_label_inputs:
            try:
                val = inp.input_value() or ""
            except Exception:
                val = inp.get_attribute("value") or ""
            if val.strip():
                post_labels.append(val.strip())
        logging.info(f"  [{device_name}] Configuration Post Labels: {post_labels}")

        # 5. Reading 标签 → 读 Value 列（值 单位，忽略 Parameter 列）
        self._click_tab("Reading")
        value_units: list[tuple[str, str]] = []
        for row in self.page.locator(self.READING_TABLE_ROWS).all():
            cells = row.locator("xpath=.//td").all()
            # Reading 表格: [Parameter(ignore), Value+Unit]
            raw_val = cells[-1].inner_text().strip() if cells else ""
            parts = raw_val.split(None, 1)
            value_units.append((parts[0] if parts else "", parts[1] if len(parts) > 1 else ""))
        logging.info(f"  [{device_name}] Reading value_units: {value_units}")

        # 6. 按行合并（以 Post Label 行数为准）
        params: dict = {}
        for i, label in enumerate(post_labels):
            val, unit = value_units[i] if i < len(value_units) else ("", "")
            params[label] = {"value": val, "unit": unit}

        logging.info(f"get_virtual_device_readings({device_name!r}): {len(params)} 个参数，"
                     f"keys={list(params.keys())}")
        return params
