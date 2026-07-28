# -*- coding: utf-8 -*-
"""
datalog_page.py — DataLog 页面自动化（Playwright 版）

导航路径：
  左侧导航 Data Log
    → 顶部 tab "Post Channels ▼" 下拉 → Post Channel 1 / 2 / 3
    → 顶部 tab "Data Loggers ▼"  下拉 → Data Logger  1 / 2 / 3
"""
import logging
from dataclasses import dataclass, field
from typing import Optional

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError

log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# 配置数据结构
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class PostChannelConfig:
    protocol: str          # "FTP" | "SFTP" | "HTTP" | "HTTPS"
    host: str
    port: int
    username: str = ""
    password: str = ""
    anonymous_mode: bool = False
    post_name_fixed: bool = False
    post_file_name: str = ""
    auth_required: bool = False
    meter_id: str = ""
    include_header: bool = True

    @property
    def ui_method(self) -> str:
        return "HTTP/HTTPS" if self.protocol in ("HTTP", "HTTPS") else self.protocol

    @property
    def http_scheme(self) -> str:
        return "HTTPS://" if self.protocol == "HTTPS" else "HTTP://"


@dataclass
class DataLoggerConfig:
    channel_index: int
    enabled: bool = True
    timestamp_format: str = "UTC Seconds"
    log_file_name_format: str = "Time Interval Format"
    log_file_format: str = "csv"
    log_file_name_prefix: str = ""
    log_file_length: str = "5 minute"
    log_interval: str = "1 minute"
    device_names: list = field(default_factory=list)


@dataclass
class DataLogParamConfig:
    device_names: list = field(default_factory=list)
    param_types: list = field(default_factory=list)


# ─────────────────────────────────────────────────────────────────────────────
# 基础操作
# ─────────────────────────────────────────────────────────────────────────────

class _PageBase:

    # Devices Selection 表（Data Logger / Rapid Logger 页面共用同一结构）
    _DEVICE_HEADER_CB  = "xpath=//thead//input[@type='checkbox']"
    _DEVICE_TBODY_ROWS = "xpath=//tbody/tr"

    def __init__(self, page: Page):
        self.page = page

    def _select_devices(self, target_names: list):
        """勾选 Devices Selection：target_names 为空 → 仅勾选全部物理 Modbus 设备并
        取消虚拟设备；否则按行文本子串匹配。

        禁止表头全选：虚拟设备（Virtual Device）的推送文件无法被
        datalog_server_verifier 的文件名设备识别匹配，会兜底归入 AcuRev4100 桶
        与物理表 Modbus 实时值比对，制造假失败。
        """
        try:
            self.page.locator(self._DEVICE_TBODY_ROWS).first.wait_for(state="visible", timeout=10000)
        except PlaywrightTimeoutError:
            log.warning("Devices Selection 表格未找到")
            return

        rows = self.page.locator(self._DEVICE_TBODY_ROWS).all()
        checked_any = False
        for row in rows:
            cells = row.locator("xpath=.//td").all()
            # 列结构：[checkbox] | Device Name(td1) | Device Type | ...。td[0] 是
            # checkbox 列（无文字），设备名在 td[1]。用整行文本做子串匹配，对列布局鲁棒。
            row_text = (row.text_content() or "").strip()
            device_name = cells[1].text_content().strip() if len(cells) > 1 else row_text.split("\n")[0]
            is_virtual = "virtual device" in row_text.lower()
            if target_names:
                should_check = any(t.lower() in row_text.lower() for t in target_names)
            else:
                should_check = not is_virtual
            cb = row.locator("xpath=.//input[@type='checkbox']").first
            if cb.count() == 0:
                continue
            is_checked = cb.is_checked()
            if should_check:
                checked_any = True
            if should_check and not is_checked:
                cb.evaluate("el => el.click()")
                log.info("  ✓ %s", device_name)
                self.page.wait_for_timeout(150)
            elif not should_check and is_checked:
                cb.evaluate("el => el.click()")
                log.info("  ✗ 取消 %s%s", device_name, "（虚拟设备）" if is_virtual else "")
                self.page.wait_for_timeout(150)
        if not checked_any:
            log.warning("Devices Selection：未勾选任何设备（目标 %s），可能无数据推送",
                        target_names or "全部物理设备")

    def _safe_click(self, selector: str, name: str = ""):
        try:
            loc = self.page.locator(selector).first
            loc.wait_for(state="visible", timeout=10000)
            loc.evaluate("el => el.click()")
            log.debug("点击：%s", name)
        except PlaywrightTimeoutError:
            log.warning("元素未找到：%s  %s", name, selector)

    def _fill(self, selector: str, value: str, name: str = ""):
        try:
            loc = self.page.locator(selector).first
            loc.wait_for(state="visible", timeout=8000)
            loc.evaluate("el => el.click()")
            loc.press("Control+a")
            loc.fill(str(value))
            log.debug("填写 %s = %s", name, value)
        except PlaywrightTimeoutError:
            log.warning("输入框未找到：%s  %s", name, selector)

    def _select_el_by_text(self, container_selector: str, text: str, name: str = ""):
        """操作 Element Plus el-select，点击展开后选目标 option。"""
        try:
            container = self.page.locator(container_selector).first
            container.wait_for(state="visible", timeout=8000)
            wrapper = container.locator(".el-select__wrapper")
            if wrapper.count() > 0:
                wrapper.first.evaluate("el => el.click()")
            else:
                container.evaluate("el => el.click()")
            self.page.wait_for_timeout(300)

            opt_patterns = [
                f"xpath=//li[contains(@class,'el-select-dropdown__item') and normalize-space(.)='{text}']",
                f"xpath=//li[contains(@class,'el-select-dropdown__item') and .//span[normalize-space(.)='{text}']]",
                f"xpath=//ul[contains(@class,'el-select-dropdown__list')]//li[contains(normalize-space(.),'{text}')]",
                f"xpath=//div[contains(@class,'el-popper')]//li[contains(normalize-space(.),'{text}')]",
            ]
            for pat in opt_patterns:
                try:
                    opt = self.page.locator(pat).first
                    opt.wait_for(state="visible", timeout=1500)
                    opt.evaluate("el => el.click()")
                    log.debug("el-select %s 选择：%s", name, text)
                    return
                except PlaywrightTimeoutError:
                    continue
            log.warning("el-select %s 选 '%s' 失败", name, text)
        except PlaywrightTimeoutError:
            log.warning("el-select %s 容器未找到（'%s'）", name, text)
        finally:
            try:
                self.page.keyboard.press("Escape")
                self.page.wait_for_timeout(100)
            except Exception:
                pass

    def _click_radio_in_group(self, group_label: str, option_text: str) -> bool:
        """在包含 group_label 的表单项后，点击含 option_text 的 el-radio label。"""
        xpaths = [
            (f"xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.), '{group_label}')]"
             f"/following::label[contains(@class,'el-radio') and normalize-space(.)='{option_text}'][1]"),
            (f"xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.), '{group_label}')]"
             f"/following::label[contains(@class,'el-radio') and contains(normalize-space(.), '{option_text}')][1]"),
            (f"xpath=//*[contains(@class,'el-form-item')]"
             f"[.//label[contains(normalize-space(.), '{group_label}')]]"
             f"//label[contains(@class,'el-radio') and contains(normalize-space(.), '{option_text}')]"),
        ]
        for xpath in xpaths:
            loc = self.page.locator(xpath)
            if loc.count() > 0:
                try:
                    loc.first.evaluate("el => el.click()")
                    log.debug("  Radio [%s] → %s", group_label, option_text)
                    self.page.wait_for_timeout(300)
                    return True
                except Exception:
                    continue
        log.warning("未找到 radio 组 '%s' 中的选项 '%s'", group_label, option_text)
        return False

    def _dismiss_overlay_dialog(self):
        """关闭登录后可能残留的 el-overlay-message-box 确认弹窗。"""
        try:
            for xpath in [
                "xpath=//div[contains(@class,'el-overlay-message-box')]",
                "xpath=//div[contains(@class,'el-message-box')]",
            ]:
                dlgs = self.page.locator(xpath).all()
                for dlg in dlgs:
                    if not dlg.is_visible():
                        continue
                    for btn_xpath in [
                        "xpath=.//button[.//span[normalize-space(.)='Yes']]",
                        "xpath=.//button[.//span[normalize-space(.)='Continue']]",
                        "xpath=.//button[contains(@class,'el-button--primary')]",
                    ]:
                        btns = dlg.locator(btn_xpath).all()
                        if btns:
                            btns[0].evaluate("el => el.click()")
                            log.info("已关闭确认弹窗")
                            self.page.wait_for_timeout(500)
                            break
        except Exception as e:
            log.debug("dismiss_overlay_dialog: %s", e)

    def _navigate_data_log(self):
        self._dismiss_overlay_dialog()
        self.page.wait_for_timeout(800)

    def _navigate_url(self, url_hash: str):
        """导航到目标子页面。若当前已在目标父路径下，直接跳转；否则先经由 dataLog 根路由。"""
        current = self.page.url
        base = current.split('#')[0]
        current_hash = current.split('#')[1] if '#' in current else ''
        # 提取目标父路径：/dataLog/postChannels/postChannel1 → /dataLog/postChannels
        target_parent = url_hash.rstrip('/').rsplit('/', 1)[0]
        if target_parent and current_hash.startswith(target_parent):
            # 已在目标父路径下，直接跳转，无需过 dataLog 根路由
            self.page.goto(f"{base}#{url_hash}")
            self.page.wait_for_timeout(2000)
        else:
            self.page.goto(f"{base}#/dataLog")
            self.page.wait_for_timeout(1500)
            self.page.goto(f"{base}#{url_hash}")
            self.page.wait_for_timeout(2000)
        log.info("URL 导航完成：%s", self.page.url)

    def _get_dropdown_options(self, container_selector: str) -> list:
        """展开 el-select，收集所有未禁用选项后关闭，返回文字列表。"""
        options = []
        try:
            # 先关闭页面上任何已开的下拉，避免选项混入
            self.page.keyboard.press("Escape")
            self.page.wait_for_timeout(200)

            container = self.page.locator(container_selector).first
            container.wait_for(state="visible", timeout=8000)
            wrapper = container.locator(".el-select__wrapper")
            if wrapper.count() > 0:
                wrapper.first.evaluate("el => el.click()")
            else:
                container.evaluate("el => el.click()")
            self.page.wait_for_timeout(800)
            # 只取可见项，避免收集到其他已关闭下拉残留在 DOM 中的 li
            items = self.page.locator(
                "xpath=//li[contains(@class,'el-select-dropdown__item') and not(contains(@class,'disabled'))]"
            ).all()
            options = [
                item.text_content().strip()
                for item in items
                if item.text_content().strip() and item.is_visible()
            ]
        except Exception as e:
            log.warning("获取下拉选项失败：%s", e)
        finally:
            try:
                self.page.keyboard.press("Escape")
                self.page.wait_for_timeout(200)
            except Exception:
                pass
        return options


# ─────────────────────────────────────────────────────────────────────────────
# Post Channel 页面
# ─────────────────────────────────────────────────────────────────────────────

class PostChannelPage(_PageBase):

    _POST_METHOD_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and normalize-space(.)='Post Method']"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _ENABLE_RADIO  = "xpath=//label[contains(@class,'el-radio') and .//input[@value='true']]"
    _DISABLE_RADIO = "xpath=//label[contains(@class,'el-radio') and .//input[@value='false']]"

    _FTP_URL_INPUT  = "xpath=//input[@placeholder='Enter FTP URL']"
    _FTP_PORT_INPUT = "xpath=//input[@placeholder='Enter FTP Port']"
    _FTP_ANON_CB    = "xpath=//input[@type='checkbox']"
    _FTP_USER_INPUT = "xpath=//input[@placeholder='Enter FTP User Name']"
    _FTP_PASS_INPUT = "xpath=//input[@placeholder='Enter FTP Password']"

    _SFTP_URL_INPUT  = "xpath=//input[@placeholder='Enter SFTP URL']"
    _SFTP_PORT_INPUT = "xpath=//input[@placeholder='Enter SFTP Port']"
    _SFTP_USER_INPUT = "xpath=//input[@placeholder='Enter SFTP User Name']"
    _SFTP_PASS_INPUT = "xpath=//input[@placeholder='Enter SFTP Password']"

    _HTTP_SCHEME_SELECT = (
        "xpath=//input[@placeholder='Enter HTTP/HTTPS URL']/preceding::div[contains(@class,'el-select')][1]"
    )
    _HTTP_URL_INPUT      = "xpath=//input[@placeholder='Enter HTTP/HTTPS URL'] | //input[@placeholder='Enter URL']"
    _HTTP_PORT_INPUT     = "xpath=//input[@placeholder='Enter HTTP/HTTPS Port'] | //input[@placeholder='Enter Port']"
    _HTTP_METER_ID_INPUT = "xpath=//input[@placeholder='Enter Meter ID'] | //input[@placeholder='Enter HTTP/HTTPS Meter ID']"
    _HTTP_FILE_NAME_INPUT = "xpath=//input[@placeholder='Enter Post File Name'] | //input[@placeholder='Enter File Name']"

    _SAVE_BTN = "xpath=//button[.//span[normalize-space(.)='Save'] or normalize-space(.)='Save']"
    _TEST_BTN = "xpath=//button[.//span[normalize-space(.)='Test Post Channel'] or normalize-space(.)='Test Post Channel']"

    def navigate_to_channel(self, n: int):
        self._navigate_data_log()
        # RPP：Post Channels 归属 Data Forwarding 子组，真实路径含 dataForwarding 段
        self._navigate_url(f"/dataLog/dataForwarding/postChannels/postChannel{n}")
        log.info("已进入 Post Channel %d 页面", n)

    def _set_enable(self, enabled: bool):
        selector = self._ENABLE_RADIO if enabled else self._DISABLE_RADIO
        try:
            loc = self.page.locator(selector).first
            loc.wait_for(state="visible", timeout=8000)
            loc.evaluate("el => el.click()")
        except PlaywrightTimeoutError:
            log.warning("Enable/Disable radio 未找到")

    def _set_post_method(self, ui_method: str):
        self._select_el_by_text(self._POST_METHOD_SELECT, ui_method, "Post Method")
        self.page.wait_for_timeout(800)

    def _fill_ftp(self, cfg: PostChannelConfig):
        self._fill(self._FTP_URL_INPUT, cfg.host, "FTP URL")
        self._fill(self._FTP_PORT_INPUT, str(cfg.port), "FTP Port")
        if cfg.anonymous_mode:
            try:
                cb = self.page.locator(self._FTP_ANON_CB).first
                cb.wait_for(state="visible", timeout=5000)
                if not cb.is_checked():
                    cb.evaluate("el => el.click()")
            except PlaywrightTimeoutError:
                log.warning("Enable anonymous mode 复选框未找到")
        else:
            self._fill(self._FTP_USER_INPUT, cfg.username, "FTP User Name")
            self._fill(self._FTP_PASS_INPUT, cfg.password, "FTP Password")

    def _fill_sftp(self, cfg: PostChannelConfig):
        self._fill(self._SFTP_URL_INPUT, cfg.host, "SFTP URL")
        self._fill(self._SFTP_PORT_INPUT, str(cfg.port), "SFTP Port")
        self._fill(self._SFTP_USER_INPUT, cfg.username, "SFTP User Name")
        self._fill(self._SFTP_PASS_INPUT, cfg.password, "SFTP Password")

    def _fill_http(self, cfg: PostChannelConfig):
        self._click_radio_in_group("Post Name Fixed", "Yes" if cfg.post_name_fixed else "No")
        self.page.wait_for_timeout(300)
        if cfg.post_name_fixed and cfg.post_file_name:
            self._fill(self._HTTP_FILE_NAME_INPUT, cfg.post_file_name, "Post File Name")
        self._click_radio_in_group("Authentication Required", "Yes" if cfg.auth_required else "No")
        self._select_el_by_text(self._HTTP_SCHEME_SELECT, cfg.http_scheme, "HTTP/HTTPS scheme")
        self._fill(self._HTTP_URL_INPUT, cfg.host, "HTTP/HTTPS URL")
        self._fill(self._HTTP_PORT_INPUT, str(cfg.port), "HTTP/HTTPS Port")
        self._fill(self._HTTP_METER_ID_INPUT, cfg.meter_id or "", "HTTP/HTTPS Meter ID")
        self._click_radio_in_group("Include Header", "Yes" if cfg.include_header else "No")

    def configure_channel(self, n: int, cfg: PostChannelConfig,
                          enabled: bool = True, test: bool = True) -> str:
        log.info("配置 Post Channel %d：%s  %s:%d", n, cfg.protocol, cfg.host, cfg.port)
        self.navigate_to_channel(n)
        self._set_enable(True)
        self._set_post_method(cfg.ui_method)

        if cfg.protocol == "FTP":
            self._fill_ftp(cfg)
        elif cfg.protocol == "SFTP":
            self._fill_sftp(cfg)
        else:
            self._fill_http(cfg)

        self._safe_click(self._SAVE_BTN, "Save")
        self.page.wait_for_timeout(2000)
        log.info("Post Channel %d 已保存（Enabled）", n)

        result = ""
        if test and enabled:
            result = self._test_channel()

        if not enabled:
            self._set_enable(False)
            self._safe_click(self._SAVE_BTN, "Save")
            self.page.wait_for_timeout(1500)

        return result

    def enable_channel(self, n: int):
        self.navigate_to_channel(n)
        self._set_enable(True)
        self._safe_click(self._SAVE_BTN, "Save")
        self.page.wait_for_timeout(1500)

    def disable_channel(self, n: int):
        self.navigate_to_channel(n)
        self._set_enable(False)
        self._safe_click(self._SAVE_BTN, "Save")
        self.page.wait_for_timeout(1500)

    def _test_channel(self) -> str:
        self._safe_click(self._TEST_BTN, "Test Post Channel")
        log.info("已点击 Test Post Channel，等待结果…")
        self.page.wait_for_timeout(2000)

        msgbox_selector = (
            "xpath=//div[contains(@class,'el-message-box__wrapper') and not(contains(@style,'display: none'))]"
            " | //div[@role='dialog' and contains(@class,'el-overlay')]"
        )
        try:
            box = self.page.locator(msgbox_selector).first
            box.wait_for(state="visible", timeout=28000)
            result = box.text_content().strip()
            log.info("Test Post Channel 弹窗内容：%s", result)

            result_lower = result.lower()
            is_fail = any(kw in result_lower for kw in ("fail", "error", "failed")) and \
                      not any(kw in result_lower for kw in ("success", "connected", "pass"))

            ok_xpath = (
                "xpath=.//button[contains(@class,'el-button--primary')]"
                " | .//button[normalize-space(.)='OK']"
                " | .//button[normalize-space(.)='Confirm']"
            )
            try:
                ok_btn = box.locator(ok_xpath).first
                ok_btn.evaluate("el => el.click()")
                self.page.wait_for_timeout(800)
            except Exception:
                pass

            if is_fail:
                raise RuntimeError(f"Test Post Channel 失败，请检查服务器配置。详情：{result!r}")
            return result
        except PlaywrightTimeoutError:
            log.warning("Test Post Channel 30 秒内未见弹窗")
            return ""

    def configure_all(self, channel_configs: dict, enabled: bool = True, test: bool = True) -> dict:
        results = {}
        for n, cfg in sorted(channel_configs.items()):
            results[n] = self.configure_channel(n, cfg, enabled=enabled, test=test)
        return results


# ─────────────────────────────────────────────────────────────────────────────
# Data Logger 页面
# ─────────────────────────────────────────────────────────────────────────────

class DataLoggerPage(_PageBase):

    _POST_CHANNEL_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Post Channel')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _LOG_FILE_FORMAT_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log File Format')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _LOG_FILE_NAME_PREFIX_INPUT = (
        "xpath=//input[@placeholder='Enter Log File Name Prefix']"
        " | //label[contains(normalize-space(.),'Log File Name Prefix')]/following::input[1]"
    )
    _LOG_FILE_LENGTH_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log File Length')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _LOG_INTERVAL_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log Interval')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _TIMESTAMP_FORMAT_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Timestamp Format')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _LOG_FILE_NAME_FORMAT_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log File Name Format')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _DEVICE_HEADER_CB  = "xpath=//thead//input[@type='checkbox']"
    _DEVICE_TBODY_ROWS = "xpath=//tbody/tr"
    _SAVE_BTN = "xpath=//button[.//span[normalize-space(.)='Save'] or normalize-space(.)='Save']"

    def navigate_to_logger(self, n: int):
        self._navigate_data_log()
        self._navigate_url(f"/dataLog/dataLogger/dataLogger{n}")
        log.info("已进入 Data Loggers %d 页面", n)

    def _set_enable(self, enabled: bool):
        selector = (
            "xpath=//label[contains(@class,'el-radio') and .//input[@value='true']]" if enabled
            else "xpath=//label[contains(@class,'el-radio') and .//input[@value='false']]"
        )
        try:
            loc = self.page.locator(selector).first
            loc.wait_for(state="visible", timeout=8000)
            loc.evaluate("el => el.click()")
        except PlaywrightTimeoutError:
            log.warning("Enable/Disable radio 未找到")

    def _set_timestamp_format(self, fmt: str):
        if not self._click_radio_in_group("Timestamp Format", fmt):
            self._select_el_by_text(self._TIMESTAMP_FORMAT_SELECT, fmt, "Timestamp Format")

    def _set_log_file_name_format(self, fmt: str):
        if not self._click_radio_in_group("Log File Name Format", fmt):
            self._select_el_by_text(self._LOG_FILE_NAME_FORMAT_SELECT, fmt, "Log File Name Format")

    # _select_devices 已上移至 _PageBase（与 Rapid Logger 页面共用）

    def configure_logger(self, n: int, cfg: DataLoggerConfig):
        log.info("配置 Data Loggers %d：channel=%d  format=%s  length=%s  interval=%s",
                 n, cfg.channel_index, cfg.log_file_format, cfg.log_file_length, cfg.log_interval)
        self.navigate_to_logger(n)
        self._set_enable(True)

        try:
            self.page.locator(self._POST_CHANNEL_SELECT).first.wait_for(state="visible", timeout=8000)
        except PlaywrightTimeoutError:
            log.warning("Post Channel 下拉未及时出现")

        self._select_el_by_text(self._POST_CHANNEL_SELECT, f"Post Channel {cfg.channel_index}", "Post Channel")
        self._select_el_by_text(self._LOG_FILE_FORMAT_SELECT, cfg.log_file_format, "Log File Format")

        if cfg.log_file_format.lower() != "json":
            self._set_timestamp_format(cfg.timestamp_format)

        self._set_log_file_name_format(cfg.log_file_name_format)

        if cfg.log_file_name_prefix:
            self._fill(self._LOG_FILE_NAME_PREFIX_INPUT, cfg.log_file_name_prefix, "Log File Name Prefix")

        self._select_el_by_text(self._LOG_FILE_LENGTH_SELECT, cfg.log_file_length, "Log File Length")
        self._select_el_by_text(self._LOG_INTERVAL_SELECT, cfg.log_interval, "Log Interval")
        self._select_devices(cfg.device_names)

        self._safe_click(self._SAVE_BTN, "Save")
        self.page.wait_for_timeout(2000)

        if not cfg.enabled:
            self._set_enable(False)
            self._safe_click(self._SAVE_BTN, "Save")
            self.page.wait_for_timeout(1500)
            log.info("Data Loggers %d 已配置并禁用", n)
        else:
            log.info("Data Loggers %d 已保存", n)

    def configure_all(self, logger_configs: dict):
        for n, cfg in sorted(logger_configs.items()):
            self.configure_logger(n, cfg)

    def disable_logger(self, n: int):
        log.info("禁用 Data Loggers %d", n)
        self.navigate_to_logger(n)
        self._set_enable(False)
        self._safe_click(self._SAVE_BTN, "Save")
        self.page.wait_for_timeout(1500)

    def disable_all(self, logger_configs: dict):
        for n in sorted(logger_configs.keys()):
            self.disable_logger(n)

    def get_log_interval_options(self, n: int) -> list[str]:
        """获取指定 Logger 在当前 Log File Length 下可选的 Log Interval 列表。"""
        self.navigate_to_logger(n)
        self._set_enable(True)
        return self._get_dropdown_options(self._LOG_INTERVAL_SELECT)


# ─────────────────────────────────────────────────────────────────────────────
# Data Log Parameter Config 页面
# ─────────────────────────────────────────────────────────────────────────────

class DataLogParamConfigPage(_PageBase):

    # 备选 URL hash，按顺序尝试（不同固件版本路径可能不同）
    _URL_HASH_CANDIDATES = [
        "/dataLog/dataLoggers/dataLogParameterConfig",
        "/dataLog/dataLogger/dataLogParameterConfig",
        "/dataLog/dataLogParameterConfig",
        "/dataLog/parameterConfig",
    ]
    _URL_HASH = _URL_HASH_CANDIDATES[0]
    _DEVICE_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and normalize-space(.)='Device']"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _PARAM_TYPE_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and normalize-space(.)='Parameter Type']"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _ALL_BTN  = "xpath=//button[.//span[normalize-space(.)='All'] or normalize-space(.)='All']"
    _SAVE_BTN = "xpath=//button[.//span[normalize-space(.)='Save'] or normalize-space(.)='Save']"

    def navigate(self):
        self._navigate_data_log()
        current = self.page.url
        base = current.split('#')[0]

        # 优先尝试直接跳转（不经过 #/dataLog，避免 SPA 将其重定向到 #/dashboard）
        for url_hash in self._URL_HASH_CANDIDATES:
            self.page.goto(f"{base}#{url_hash}")
            self.page.wait_for_timeout(2500)
            log.info("URL 直接导航尝试：%s", self.page.url)
            try:
                self.page.locator(self._DEVICE_SELECT).first.wait_for(state="visible", timeout=5000)
                log.info("已进入 Data Log Parameter Config 页面：%s", url_hash)
                self._URL_HASH = url_hash  # 缓存有效路径
                return
            except PlaywrightTimeoutError:
                log.warning("URL %s 未找到 Device 下拉，继续尝试", url_hash)

        # 直接跳转均失败，改用菜单导航
        log.warning("直接 URL 均失败，改用菜单导航")
        self._navigate_via_menu()

    def _navigate_via_menu(self):
        current = self.page.url
        base = current.split('#')[0]
        self.page.goto(f"{base}#/dataLog")
        self.page.wait_for_timeout(2000)
        try:
            # 先尝试直接点击菜单项（不依赖展开子菜单）
            item = self.page.locator(
                "xpath=//li[contains(@class,'el-menu-item') and contains(normalize-space(.),'Data Log Parameter Config')]"
                " | //li[contains(@class,'el-menu-item') and contains(normalize-space(.),'Parameter Config')]"
            ).first
            item.wait_for(state="visible", timeout=4000)
            item.evaluate("el => el.click()")
            self.page.wait_for_timeout(2000)
            log.info("菜单导航 Data Log Parameter Config 成功（直接点击）")
            return
        except PlaywrightTimeoutError:
            pass

        try:
            # 备选：展开 Data Loggers 子菜单再点击
            sub = self.page.locator(
                "xpath=//div[contains(@class,'el-sub-menu__title') and contains(normalize-space(.),'Data Loggers')]"
                " | //div[contains(@class,'el-sub-menu__title') and contains(normalize-space(.),'Data Logger')]"
            ).first
            sub.wait_for(state="visible", timeout=5000)
            sub.evaluate("el => el.click()")
            self.page.wait_for_timeout(1000)
            item = self.page.locator(
                "xpath=//li[contains(@class,'el-menu-item') and contains(normalize-space(.),'Parameter Config')]"
            ).first
            item.wait_for(state="visible", timeout=5000)
            item.evaluate("el => el.click()")
            self.page.wait_for_timeout(2000)
            log.info("菜单导航 Data Log Parameter Config 成功（展开子菜单）")
        except PlaywrightTimeoutError:
            log.warning("菜单导航 Data Log Parameter Config 失败")

    def _not_select_has_items(self) -> bool:
        """检查 Not Select 区域（transfer 左侧第一面板）是否有待选参数。

        与 aws_iot_page._transfer_left_count 保持一致：不依赖面板标题文字，
        直接取第一个 el-transfer-panel（左侧 = Not Select）内的条目数。
        """
        xpaths = [
            # el-transfer 左侧面板条目
            "xpath=(//div[contains(@class,'el-transfer-panel')])[1]"
            "//*[contains(@class,'el-transfer-panel__item')]",
            # 备选：自定义 transfer-left
            "xpath=//div[contains(@class,'transfer-left')]"
            "//div[contains(@class,'option-item')]",
            # 再备选：左面板 checkbox label
            "xpath=(//div[contains(@class,'el-transfer-panel')])[1]"
            "//label[contains(@class,'el-checkbox')]",
            # 再再备选：左面板任意 li
            "xpath=(//div[contains(@class,'el-transfer-panel')])[1]//li",
        ]
        for xpath in xpaths:
            try:
                count = self.page.locator(xpath).count()
                log.debug("_not_select_has_items: %d 项（%s）", count, xpath[:60])
                if count > 0:
                    return True
            except Exception:
                continue
        return False

    def configure_device(self, device_name: str, param_types: list):
        """为单台设备配置 DataLog Parameter Config。

        仿照 aws_iot_page._configure_parameter_dialog 的方式：
        - 打开 Parameter Type 下拉一次读取全部选项
        - i=0 直接使用已打开的 popper 点击选项（不用 Escape）
        - i>0 点击 wrapper 重新打开后再点击选项
        - 整个遍历过程不触发 Escape，避免取消已选中的值
        """
        log.info("Parameter Config：配置设备 [%s]", device_name)
        self._select_el_by_text(self._DEVICE_SELECT, device_name, "Device")
        self.page.wait_for_timeout(1000)

        # ── 若外部传入了类型列表则使用，否则从页面读取 ──────────────────────────
        if param_types:
            types_to_use = list(param_types)
            # 外部传参时需先打开下拉定位到第一个选项；此处不做特殊处理，直接进入循环
            has_pt = True
        else:
            types_to_use, has_pt = self._open_pt_and_read_options()

        if not has_pt:
            # 无 Parameter Type 下拉：直接全选
            if self._not_select_has_items():
                self._safe_click(self._ALL_BTN, "All")
                self.page.wait_for_timeout(500)
                log.info("  已全选（无 Parameter Type 下拉）")
            else:
                log.info("  Not Select 无参数，跳过（无 Parameter Type 下拉）")
            self._safe_click(self._SAVE_BTN, "Save")
            self.page.wait_for_timeout(1500)
            log.info("  设备 [%s] 参数配置已保存", device_name)
            return

        # ── 定位 Parameter Type 下拉 wrapper（用于后续重新打开）────────────────
        pt_container = self.page.locator(self._PARAM_TYPE_SELECT).first
        pt_wrapper = pt_container.locator(".el-select__wrapper")
        pt_trigger = pt_wrapper.first if pt_wrapper.count() > 0 else pt_container

        for i, pt in enumerate(types_to_use):
            if i > 0:
                # 上一次点击选项后 popper 已关闭，重新打开
                pt_trigger.click()
                self.page.wait_for_timeout(500)

            # 在已开的 popper 中找到并点击目标选项（不使用 Escape）
            clicked = False
            for opt in self.page.locator(
                "xpath=//li[contains(@class,'el-select-dropdown__item')]"
            ).all():
                try:
                    if opt.is_visible() and opt.inner_text().strip() == pt:
                        opt.evaluate("el => el.click()")
                        clicked = True
                        break
                except Exception:
                    continue

            if not clicked:
                log.warning("  Parameter Type '%s' 选项未找到，跳过", pt)
                # 关闭 popper（点击 trigger toggle，不用 Escape）
                try:
                    pt_trigger.click()
                    self.page.wait_for_timeout(300)
                except Exception:
                    pass
                continue

            # 等待 transfer 面板刷新
            self.page.wait_for_timeout(1000)

            if self._not_select_has_items():
                self._safe_click(self._ALL_BTN, f"All [{pt}]")
                self.page.wait_for_timeout(300)
                log.info("  已全选：%s", pt)
            else:
                log.info("  Not Select 无参数，跳过：%s", pt)

        self._safe_click(self._SAVE_BTN, "Save")
        self.page.wait_for_timeout(1500)
        log.info("  设备 [%s] 参数配置已保存", device_name)

    def _open_pt_and_read_options(self) -> "tuple[list, bool]":
        """打开 Parameter Type 下拉，读取所有可见选项，保持 popper 打开供后续直接点击。

        返回 (options, has_pt)：
        - options: 选项文字列表（下拉保持打开状态，首个选项可直接点击）
        - has_pt: False 表示该设备无 Parameter Type 下拉
        """
        try:
            container = self.page.locator(self._PARAM_TYPE_SELECT).first
            container.wait_for(state="visible", timeout=2000)
        except PlaywrightTimeoutError:
            return [], False

        wrapper = container.locator(".el-select__wrapper")
        trigger = wrapper.first if wrapper.count() > 0 else container
        trigger.click()
        self.page.wait_for_timeout(600)

        opts = self.page.locator(
            "xpath=//li[contains(@class,'el-select-dropdown__item')"
            " and not(contains(@class,'disabled'))]"
        ).all()
        option_texts = [
            o.inner_text().strip() for o in opts
            if o.is_visible() and o.inner_text().strip()
        ]
        log.info("Parameter Type 选项：%s", option_texts)

        if not option_texts:
            # 没读到选项 → 可能不是 el-select；关闭 popper
            try:
                trigger.click()
            except Exception:
                pass
            return [], False

        return option_texts, True

    def _get_param_types_safe(self) -> list:
        """获取 Parameter Type 下拉选项；若不存在或返回设备名则返回空列表。"""
        try:
            container = self.page.locator(self._PARAM_TYPE_SELECT).first
            container.wait_for(state="visible", timeout=2000)
        except PlaywrightTimeoutError:
            return []

        options = self._get_dropdown_options(self._PARAM_TYPE_SELECT)
        # 若选项与 Device 下拉相同（包含连字符格式的设备名），视为无效
        device_opts = self._get_dropdown_options(self._DEVICE_SELECT)
        if set(options) & set(device_opts):
            log.debug("Parameter Type 返回设备名，视为无 Parameter Type 下拉")
            return []
        return options

    def configure_all(self, cfg: DataLogParamConfig):
        self.navigate()
        for device_name in cfg.device_names:
            self.configure_device(device_name, cfg.param_types)

    def configure_all_devices(self, param_types: list = None):
        self.navigate()
        device_names = self._get_dropdown_options(self._DEVICE_SELECT)
        if not device_names:
            log.warning("Data Log Parameter Config：Device 下拉无可用设备")
            return
        log.info("发现 %d 台设备：%s", len(device_names), device_names)
        for device_name in device_names:
            self.configure_device(device_name, param_types or [])


# ─────────────────────────────────────────────────────────────────────────────
# Rapid Logger 页面
# ─────────────────────────────────────────────────────────────────────────────

class RapidLoggerPage(_PageBase):

    _URL_HASH_CANDIDATES = [
        "/dataLog/dataLogger/rapidLogger",
        "/dataLog/rapidLogger",
        "/dataLog/dataLoggers/rapidLogger",
    ]

    _POST_CHANNEL_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Post Channel')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _LOG_FILE_FORMAT_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log File Format')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _LOG_FILE_NAME_PREFIX_INPUT = (
        "xpath=//input[@placeholder='Enter Log File Name Prefix']"
        " | //label[contains(normalize-space(.),'Log File Name Prefix')]/following::input[1]"
    )
    _LOG_FILE_LENGTH_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log File Length')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _LOG_INTERVAL_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log Interval')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _TIMESTAMP_FORMAT_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Timestamp Format')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _LOG_FILE_NAME_FORMAT_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log File Name Format')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _SAVE_BTN = "xpath=//button[.//span[normalize-space(.)='Save'] or normalize-space(.)='Save']"

    def navigate_to_rapid_logger(self):
        self._navigate_data_log()
        current = self.page.url
        base = current.split('#')[0]
        for url_hash in self._URL_HASH_CANDIDATES:
            self.page.goto(f"{base}#{url_hash}")
            self.page.wait_for_timeout(2000)
            try:
                self.page.locator(self._SAVE_BTN).first.wait_for(state="visible", timeout=4000)
                log.info("已进入 Rapid Logger 页面：%s", url_hash)
                return
            except PlaywrightTimeoutError:
                log.warning("URL %s 未找到 Save 按钮，继续尝试", url_hash)
        log.warning("Rapid Logger 页面所有 URL 候选均失败")

    def _set_enable(self, enabled: bool):
        selector = (
            "xpath=//label[contains(@class,'el-radio') and .//input[@value='true']]" if enabled
            else "xpath=//label[contains(@class,'el-radio') and .//input[@value='false']]"
        )
        try:
            loc = self.page.locator(selector).first
            loc.wait_for(state="visible", timeout=8000)
            loc.evaluate("el => el.click()")
        except PlaywrightTimeoutError:
            log.warning("Enable/Disable radio 未找到")

    def _set_timestamp_format(self, fmt: str):
        if not self._click_radio_in_group("Timestamp Format", fmt):
            self._select_el_by_text(self._TIMESTAMP_FORMAT_SELECT, fmt, "Timestamp Format")

    def _set_log_file_name_format(self, fmt: str):
        if not self._click_radio_in_group("Log File Name Format", fmt):
            self._select_el_by_text(self._LOG_FILE_NAME_FORMAT_SELECT, fmt, "Log File Name Format")

    def configure_rapid_logger(
        self,
        channel_index: Optional[int],
        file_format: str,
        file_length: str,
        timestamp_fmt: str,
        name_fmt: str,
        prefix: str,
        log_interval: str,
        enabled: bool = True,
        device_names: Optional[list] = None,
    ):
        log.info("配置 Rapid Logger：format=%s  length=%s  interval=%s",
                 file_format, file_length, log_interval)
        self.navigate_to_rapid_logger()
        self._set_enable(True)

        if channel_index is not None:
            self._select_el_by_text(
                self._POST_CHANNEL_SELECT,
                f"Post Channel {channel_index}",
                "Post Channel",
            )
        else:
            # 清空 Post Channel（设为 None）
            try:
                container = self.page.locator(self._POST_CHANNEL_SELECT).first
                container.wait_for(state="visible", timeout=8000)
                wrapper = container.locator(".el-select__wrapper")
                trigger = wrapper.first if wrapper.count() > 0 else container
                trigger.hover()
                self.page.wait_for_timeout(1000)
                cleared = False
                for sel in [".el-select__clear", "[class*='el-select__clear']",
                             ".el-select__suffix .el-icon:last-child"]:
                    btn = container.locator(sel)
                    if btn.count() > 0:
                        try:
                            btn.first.wait_for(state="visible", timeout=500)
                            btn.first.evaluate("el => el.click()")
                            self.page.wait_for_timeout(300)
                            cleared = True
                            break
                        except Exception:
                            continue
                if not cleared:
                    log.warning("Post Channel clear 按钮未找到，可能已为空")
            except Exception as e:
                log.warning("清空 Post Channel 失败：%s", e)

        self._select_el_by_text(self._LOG_FILE_FORMAT_SELECT, file_format, "Log File Format")
        if file_format.lower() != "json":
            self._set_timestamp_format(timestamp_fmt)
        self._set_log_file_name_format(name_fmt)
        if prefix:
            self._fill(self._LOG_FILE_NAME_PREFIX_INPUT, prefix, "Log File Name Prefix")
        self._select_el_by_text(self._LOG_FILE_LENGTH_SELECT, file_length, "Log File Length")
        self._select_el_by_text(self._LOG_INTERVAL_SELECT, log_interval, "Log Interval")
        # Devices Selection 必选：一台不勾网关不会记录/推送任何文件（曾因缺此步导致
        # Rapid 系列全部"超时未收到文件"）。device_names=None → 全选
        self._select_devices(device_names or [])
        self._safe_click(self._SAVE_BTN, "Save")
        self.page.wait_for_timeout(2000)

        if not enabled:
            self._set_enable(False)
            self._safe_click(self._SAVE_BTN, "Save")
            self.page.wait_for_timeout(1500)
            log.info("Rapid Logger 已配置并禁用")
        else:
            log.info("Rapid Logger 已保存（Enabled）")

    def disable_rapid_logger(self):
        log.info("禁用 Rapid Logger")
        self.navigate_to_rapid_logger()
        self._set_enable(False)
        self._safe_click(self._SAVE_BTN, "Save")
        self.page.wait_for_timeout(1500)

    def get_log_interval_options(self, length_text: str) -> list:
        """读取指定 Log File Length 下 Rapid Logger 可用的 Log Interval 选项。"""
        self.navigate_to_rapid_logger()
        self._set_enable(True)
        self._select_el_by_text(self._LOG_FILE_LENGTH_SELECT, length_text, "Log File Length")
        self.page.wait_for_timeout(800)
        options = []
        try:
            container = self.page.locator(self._LOG_INTERVAL_SELECT).first
            inner = container.locator(".el-select__wrapper")
            if inner.count() > 0:
                inner.first.evaluate("el => el.click()")
            else:
                container.evaluate("el => el.click()")
            self.page.wait_for_timeout(800)
            all_items = self.page.locator(
                "xpath=//li[contains(@class,'el-select-dropdown__item')"
                " and not(contains(@class,'is-disabled'))"
                " and not(contains(@class,'disabled'))]"
            ).all()
            options = [
                el.text_content().strip()
                for el in all_items
                if el.text_content().strip() and el.is_visible()
            ]
        except Exception:
            pass
        finally:
            try:
                self.page.keyboard.press("Escape")
                self.page.wait_for_timeout(200)
            except Exception:
                pass
        return options


# ─────────────────────────────────────────────────────────────────────────────
# Physical Device Poll Interval 辅助类
# ─────────────────────────────────────────────────────────────────────────────

class PhysicalDevicePollHelper(_PageBase):
    """设置/恢复所有 Physical Devices 的 Poll Interval（秒）。"""

    _URL_HASH_CANDIDATES = [
        "/dataLog/physicalDevices",
        "/dataLog/physicalDevice",
        "/dataLog/physical-devices",
    ]
    _DEVICE_ROWS   = "xpath=//tbody/tr"
    _SAVE_BTN      = "xpath=//button[.//span[normalize-space(.)='Save'] or normalize-space(.)='Save']"
    _POLL_INPUT    = "xpath=.//input[@type='number' or contains(@class,'el-input__inner')]"
    _POLL_SELECT   = "xpath=.//div[contains(@class,'el-select')]"

    def _navigate(self) -> bool:
        self._navigate_data_log()
        current = self.page.url
        base = current.split('#')[0]
        for url_hash in self._URL_HASH_CANDIDATES:
            self.page.goto(f"{base}#{url_hash}")
            self.page.wait_for_timeout(2000)
            try:
                self.page.locator(self._DEVICE_ROWS).first.wait_for(state="visible", timeout=4000)
                log.info("已进入 Physical Devices 页面：%s", url_hash)
                return True
            except PlaywrightTimeoutError:
                continue
        log.warning("Physical Devices 页面导航失败")
        return False

    def set_all(self, seconds: int) -> dict:
        """将所有 Physical Device 的 Poll Interval 设为 seconds，返回 {row_index: 原始值}。"""
        originals: dict = {}
        if not self._navigate():
            return originals
        rows = self.page.locator(self._DEVICE_ROWS).all()
        changed = False
        for i, row in enumerate(rows):
            try:
                inp = row.locator(self._POLL_INPUT)
                if inp.count() > 0 and inp.first.is_visible():
                    original = inp.first.input_value() or ""
                    originals[i] = original
                    inp.first.evaluate("el => el.click()")
                    inp.first.press("Control+a")
                    inp.first.fill(str(seconds))
                    log.info("Physical Device row %d：poll interval → %ds（原 %s）", i, seconds, original)
                    changed = True
            except Exception as e:
                log.warning("row %d 设置 poll interval 失败：%s", i, e)
        if changed:
            try:
                self._safe_click(self._SAVE_BTN, "Save (PhysicalDevices)")
                self.page.wait_for_timeout(1500)
            except Exception:
                pass
        return originals

    def restore_all(self, originals: dict):
        """恢复 set_all 保存的原始 Poll Interval 值。"""
        if not originals:
            return
        if not self._navigate():
            return
        rows = self.page.locator(self._DEVICE_ROWS).all()
        changed = False
        for i, original in originals.items():
            if i >= len(rows):
                continue
            try:
                row = rows[i]
                inp = row.locator(self._POLL_INPUT)
                if inp.count() > 0 and inp.first.is_visible():
                    inp.first.evaluate("el => el.click()")
                    inp.first.press("Control+a")
                    inp.first.fill(original)
                    log.info("Physical Device row %d：poll interval 恢复为 %s", i, original)
                    changed = True
            except Exception as e:
                log.warning("row %d 恢复 poll interval 失败：%s", i, e)
        if changed:
            try:
                self._safe_click(self._SAVE_BTN, "Save (PhysicalDevices restore)")
                self.page.wait_for_timeout(1500)
            except Exception:
                pass


# ─────────────────────────────────────────────────────────────────────────────
# Post Historical Data 页面
# ─────────────────────────────────────────────────────────────────────────────

class PostHistoricalDataPage(_PageBase):

    _URL_HASH_CANDIDATES = [
        "/dataLog/dataForwarding/postHistoricalData",   # RPP：Post Historical Data 归属 Data Forwarding
        "/dataLog/postHistoricalData",
        "/dataLog/postHistorical",
        "/dataLog/historicalData",
        "/dataLog/dataLoggers/postHistoricalData",
    ]
    _POST_CHANNEL_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Post Channel')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _DEVICE_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and normalize-space(.)='Device']"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _LOG_FILE_FORMAT_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log File Format')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _TIMESTAMP_FORMAT_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Timestamp Format')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _LOG_FILE_NAME_FORMAT_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log File Name Format')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _LOG_FILE_NAME_PREFIX_INPUT = (
        "xpath=//input[@placeholder='Enter Log File Name Prefix']"
        " | //label[contains(normalize-space(.),'Log File Name Prefix')]/following::input[1]"
    )
    _LOG_FILE_LENGTH_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log File Length')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _LOG_INTERVAL_SELECT = (
        "xpath=//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log Interval')]"
        "/following::div[contains(@class,'el-select')][1]"
    )
    _POST_BTN = (
        "xpath=//button[.//span[normalize-space(.)='Post'] or normalize-space(.)='Post']"
    )
    _PROGRESS_XPATH = (
        "xpath=//*[contains(@class,'el-progress') or contains(@class,'progress-bar')]"
        " | //div[contains(@class,'el-progress__text')]"
    )
    _ERROR_XPATH = (
        "xpath=//div[contains(@class,'el-message--error')]"
        " | //div[contains(@class,'el-form-item__error')]"
        " | //span[contains(@class,'el-form-item__error')]"
        " | //*[contains(@class,'is-error')]//div[contains(@class,'el-form-item__error')]"
    )

    def navigate(self):
        self._navigate_data_log()
        current = self.page.url
        base = current.split('#')[0]
        for url_hash in self._URL_HASH_CANDIDATES:
            self.page.goto(f"{base}#{url_hash}")
            self.page.wait_for_timeout(2500)
            try:
                self.page.locator(self._POST_CHANNEL_SELECT).first.wait_for(state="visible", timeout=5000)
                log.info("已进入 Post Historical Data 页面：%s", url_hash)
                return
            except PlaywrightTimeoutError:
                log.warning("URL %s 未找到 Post Channel 下拉，继续尝试", url_hash)
        log.warning("Post Historical Data 页面导航均失败")

    def _set_timestamp_format(self, fmt: str):
        if not self._click_radio_in_group("Timestamp Format", fmt):
            self._select_el_by_text(self._TIMESTAMP_FORMAT_SELECT, fmt, "Timestamp Format")

    def _set_log_file_name_format(self, fmt: str):
        if not self._click_radio_in_group("Log File Name Format", fmt):
            self._select_el_by_text(self._LOG_FILE_NAME_FORMAT_SELECT, fmt, "Log File Name Format")

    def _select_first_device(self):
        options = self._get_dropdown_options(self._DEVICE_SELECT)
        if options:
            self._select_el_by_text(self._DEVICE_SELECT, options[0], "Device")
            log.info("Post Historical Data：已选择设备 %s", options[0])
        else:
            log.warning("Post Historical Data：Device 下拉无可用设备")

    def post(
        self,
        channel_index: int,
        file_format: str,
        file_length: str,
        timestamp_fmt: str,
        name_fmt: str,
        prefix: str,
        interval: str,
    ) -> str:
        """填写表单并点击 Post。返回错误文本（空字符串表示无错误）。"""
        self.navigate()
        self._select_el_by_text(
            self._POST_CHANNEL_SELECT,
            f"Post Channel {channel_index}",
            "Post Channel",
        )
        self.page.wait_for_timeout(500)
        self._select_first_device()
        self.page.wait_for_timeout(500)
        self._select_el_by_text(self._LOG_FILE_FORMAT_SELECT, file_format, "Log File Format")
        self.page.wait_for_timeout(300)
        if file_format.lower() != "json":
            self._set_timestamp_format(timestamp_fmt)
        self._set_log_file_name_format(name_fmt)
        if prefix:
            self._fill(self._LOG_FILE_NAME_PREFIX_INPUT, prefix, "Log File Name Prefix")
        self._select_el_by_text(self._LOG_FILE_LENGTH_SELECT, file_length, "Log File Length")
        self._select_el_by_text(self._LOG_INTERVAL_SELECT, interval, "Log Interval")
        self._safe_click(self._POST_BTN, "Post")
        self.page.wait_for_timeout(1500)

        # 检查是否有立即出现的校验错误
        err_loc = self.page.locator(self._ERROR_XPATH)
        if err_loc.count() > 0:
            for el in err_loc.all():
                txt = (el.text_content() or "").strip()
                if txt:
                    log.warning("Post Historical Data 校验错误：%s", txt)
                    return txt
        return ""

    def wait_for_completion(self, timeout_ms: int = 300000) -> bool:
        """等待 Post 进度完成（进度达 100% 或进度条消失）。"""
        elapsed = 0
        poll = 2000
        while elapsed < timeout_ms:
            self.page.wait_for_timeout(poll)
            elapsed += poll
            try:
                pb = self.page.locator(self._PROGRESS_XPATH).first
                if not pb.is_visible():
                    log.info("Post Historical Data 进度条消失，视为完成")
                    return True
                text = (pb.text_content() or "").strip()
                log.debug("Post Historical Data 进度：%s", text)
                if "100" in text:
                    log.info("Post Historical Data 进度达 100%%")
                    return True
            except Exception:
                return True
        log.warning("Post Historical Data 等待超时（%dms）", timeout_ms)
        return False
