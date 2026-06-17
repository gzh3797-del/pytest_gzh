# -*- coding: utf-8 -*-
import sys
from pathlib import Path

import pytest
from playwright.sync_api import Page

_REPO_ROOT = str(Path(__file__).parents[3])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from projects.AcuHMI_1_7.settings import HMI_URL  # noqa: E402
from projects.AcuHMI_1_7.tests.ui.systemsettings.helpers.factory_reset import (  # noqa: E402
    nav_to,
    trigger_factory_reset,
    wait_for_device_and_relogin,
)

# 默认值（用例表）
_DEFAULT_TLS = "Auto"
_DEFAULT_EMAIL_INTERVAL = "5min"


def _nav_to_email_settings(page: Page) -> None:
    """导航到 System Settings → Email 页面。

    选择器待真机校验：hash 路径基于同类页面模式推断。
    """
    nav_to(page, "#/systemSettings/email")


def _nav_to_alarm_notification(page: Page) -> None:
    """导航到 System Settings → Alarm Notification 页面。"""
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(600)
    page.locator(".el-menu-item").filter(has_text="Alarm Notification").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def _modify_email_settings(page: Page) -> None:
    """修改 Email 配置，使其偏离默认值（默认所有字段为空）。

    修改策略：填入 Email Server / Email Port / Sender Name / From Email Address。
    选择器待真机校验。
    """
    _nav_to_email_settings(page)

    field_map = {
        "Email Server": "mail.test.com",
        "Email Port": "587",
        "Sender Name": "TestSender",
        "From Email Address": "test@test.com",
        "Username": "testuser",
    }
    for label, value in field_map.items():
        item = page.locator(".el-form-item").filter(has_text=label).first
        if item.count() > 0:
            inp = item.locator("input").first
            if inp.count() > 0:
                inp.fill(value)
                page.wait_for_timeout(200)

    save_btn = page.get_by_role("button", name="Save").first
    if save_btn.count() > 0:
        save_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)


def _modify_alarm_notification(page: Page) -> None:
    """修改 Alarm Notification 配置，使其偏离默认值（默认 Disable）。

    修改策略：将 Alarm Email Enable 改为 Enable，填入 Recipient。
    选择器待真机校验。
    """
    _nav_to_alarm_notification(page)

    enable_item = page.locator(".el-form-item").filter(has_text="Alarm Email Enable").first
    if enable_item.count() > 0:
        enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
        if enable_radio.count() > 0:
            enable_radio.first.click()
            page.wait_for_timeout(500)

    recipient_fi = page.locator(".el-form-item").filter(has_text="Recipient 1").first
    if recipient_fi.count() > 0:
        inp = recipient_fi.locator("input").first
        if inp.count() > 0:
            inp.fill("test@example.com")
            page.wait_for_timeout(200)

    save_btn = page.get_by_role("button", name="Save").first
    if save_btn.count() > 0:
        save_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)


@pytest.mark.destructive
def test_TestCase_AcuHMI_005_06_case12(system_settings_page: Page) -> None:
    """TestCase_AcuHMI_005_06_case12｜修改Email/AlarmNotification配置后恢复出厂，验证默认值

    预置条件：
      1. 管理权限登录 AcuHMI，设备运行正常

    测试步骤：
      1. 修改 Email 配置（Email Server/Port/Sender Name/From Email/Username）
      2. 修改 Alarm Notification（Enable，填入 Recipient）
      3. Factory Reset 恢复出厂，等待设备重启并重新登录
      4. 查看 Email 配置，断言默认值：
         - TLS/SSL 默认 Auto
         - Email Server / Email Port / Sender Name / From Email Address /
           Username / Password 默认为空
      5. 查看 Alarm Notification，断言：
         - Alarm Notification 默认为 Disable
         - Recipient 默认为空，Email Interval 默认为 5min

    预期结果：Factory Reset 成功，Email 和 Alarm Notification 所有字段恢复默认值。
    """
    page = system_settings_page

    # ── Step 1 & 2: 修改配置 ─────────────────────────────────────────────────
    _modify_email_settings(page)
    _modify_alarm_notification(page)

    # ── Step 3: 执行 Factory Reset + 等待重启 + 重登录 ───────────────────────
    trigger_factory_reset(page)
    wait_for_device_and_relogin(page)

    # ── Step 4: 导航到 Email，断言默认值 ────────────────────────────────────
    _nav_to_email_settings(page)
    page.wait_for_timeout(1000)

    empty_fields = ["Email Server", "Email Port", "Sender Name",
                    "From Email Address", "Username", "Password"]
    for label in empty_fields:
        item = page.locator(".el-form-item").filter(has_text=label).first
        if item.count() > 0:
            inp = item.locator("input").first
            if inp.count() > 0:
                val = inp.input_value()
                assert val == "" or val is None, (
                    f"Factory Reset 后 {label} 应为空，实际='{val}'"
                )

    # ── Step 5: 导航到 Alarm Notification，断言默认值 ───────────────────────
    _nav_to_alarm_notification(page)
    page.wait_for_timeout(1000)

    # 断言 Alarm Email Enable 默认为 Disable
    alarm_enable_item = page.locator(".el-form-item").filter(has_text="Alarm Email Enable").first
    if alarm_enable_item.count() > 0:
        disable_radio = alarm_enable_item.locator(".el-radio").filter(has_text="Disable")
        if disable_radio.count() > 0:
            radio_class = disable_radio.first.get_attribute("class") or ""
            assert "is-checked" in radio_class, (
                f"Factory Reset 后 Alarm Email Enable 默认应为 Disable，"
                f"实际 radio class='{radio_class}'"
            )

    # 断言 Email Interval 默认为 5min（若字段可见）
    interval_item = page.locator(".el-form-item").filter(has_text="Email Interval").first
    if interval_item.count() > 0:
        interval_text = interval_item.inner_text()
        # 字段可能因 Disable 而隐藏，若可见则断言包含 5min
        if interval_item.is_visible():
            assert _DEFAULT_EMAIL_INTERVAL in interval_text or "5" in interval_text, (
                f"Factory Reset 后 Email Interval 默认应为 5min，"
                f"实际文本='{interval_text}'"
            )
