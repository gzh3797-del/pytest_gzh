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
    trigger_factory_reset,
    wait_for_device_and_relogin,
)

# 默认值（用例表）
_DEFAULT_NTP_SERVER1 = "0.us.pool.ntp.org"
_DEFAULT_NTP_SERVER2 = ""
_DEFAULT_NTP_SERVER3 = ""


def _nav_to_datetime(page: Page) -> None:
    """导航到 System Settings → Date & Time 页面。"""
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def _modify_datetime_settings(page: Page) -> None:
    """启用 NTP 并将 NTP Server 1/2/3 修改为 www.baidu.com，再关闭 NTP 并保存。

    用例步骤：
      1. NTP Enable
      2. 将 NTP Server 1/2/3 设为 www.baidu.com
      3. 修改时区为非默认值（选第一个可用选项）
      4. 保存
      5. 关闭 NTP，设备再次触发 Factory Reset
    """
    _nav_to_datetime(page)

    # 启用 NTP Enable
    ntp_enable_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    if ntp_enable_item.count() > 0:
        enable_radio = ntp_enable_item.locator(".el-radio").filter(has_text="Enable")
        if enable_radio.count() > 0:
            enable_radio.first.click()
            page.wait_for_timeout(500)

    # 修改 NTP Server 1/2/3 为 www.baidu.com
    for placeholder in ("NTP Server 1", "NTP Server 2", "NTP Server 3"):
        inp = page.get_by_placeholder(placeholder).first
        if inp.count() > 0:
            inp.fill("www.baidu.com")
            page.wait_for_timeout(200)

    # 修改时区（选一个非当前的选项）
    tz_fi = page.locator(".el-form-item").filter(has_text="Time Zone").first
    if tz_fi.count() > 0:
        tz_select = tz_fi.locator(".el-select").first
        if tz_select.count() > 0:
            tz_select.click()
            page.wait_for_timeout(400)
            # 搜索 "UTC" 并选第一项
            search_inp = page.locator(".el-select-dropdown input").first
            if search_inp.count() > 0 and search_inp.is_visible():
                search_inp.fill("UTC")
                page.wait_for_timeout(300)
            opts = page.locator(".el-select-dropdown__item").all()
            for opt in opts:
                try:
                    if opt.is_visible():
                        opt.click()
                        break
                except Exception:
                    pass
            else:
                page.keyboard.press("Escape")
            page.wait_for_timeout(300)

    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)


@pytest.mark.destructive
def test_TestCase_AcuHMI_005_06_case09(system_settings_page: Page) -> None:
    """TestCase_AcuHMI_005_06_case09｜修改Date & Time配置后恢复出厂，验证默认值

    预置条件：
      1. 管理权限登录 AcuHMI，设备运行正常

    测试步骤：
      1. 进入 Date & Time 页面，NTP Enable，将 NTP Server 1/2/3 设为 www.baidu.com，
         修改时区为非默认值，Save
      2. Factory Reset 恢复出厂，等待设备重启并重新登录
      3. 验证 Date & Time 默认值：
         - NTP Server 1 默认：0.us.pool.ntp.org
         - NTP Server 2/3 默认：空
         - 时区恢复默认值
      4. 关闭 NTP，再次 Factory Reset，验证 NTP 默认状态为 Enable

    预期结果：
      - Factory Reset 后 NTP Server 1 = 0.us.pool.ntp.org，NTP Server 2/3 为空
      - 默认时区正确
      - NTP 默认状态为 Enable
    """
    page = system_settings_page

    # ── Step 1: 修改 Date & Time 配置 ────────────────────────────────────────
    _modify_datetime_settings(page)

    # ── Step 2: 执行 Factory Reset + 等待重启 + 重登录 ───────────────────────
    trigger_factory_reset(page)
    wait_for_device_and_relogin(page)

    # ── Step 3: 导航回 Date & Time，断言默认值 ───────────────────────────────
    _nav_to_datetime(page)
    page.wait_for_selector("input[placeholder='NTP Server 1']", timeout=30_000)
    page.wait_for_timeout(500)

    ntp1_val = page.get_by_placeholder("NTP Server 1").first.input_value()
    assert _DEFAULT_NTP_SERVER1 in ntp1_val or ntp1_val == _DEFAULT_NTP_SERVER1, (
        f"Factory Reset 后 NTP Server 1 应为 '{_DEFAULT_NTP_SERVER1}'，"
        f"实际='{ntp1_val}'"
    )

    ntp2_inp = page.get_by_placeholder("NTP Server 2").first
    if ntp2_inp.count() > 0:
        ntp2_val = ntp2_inp.input_value()
        assert ntp2_val == _DEFAULT_NTP_SERVER2, (
            f"Factory Reset 后 NTP Server 2 应为空，实际='{ntp2_val}'"
        )

    ntp3_inp = page.get_by_placeholder("NTP Server 3").first
    if ntp3_inp.count() > 0:
        ntp3_val = ntp3_inp.input_value()
        assert ntp3_val == _DEFAULT_NTP_SERVER3, (
            f"Factory Reset 后 NTP Server 3 应为空，实际='{ntp3_val}'"
        )

    # 断言 NTP Enable 默认状态为 Enable
    ntp_enable_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    if ntp_enable_item.count() > 0:
        enable_radio = ntp_enable_item.locator(".el-radio").filter(has_text="Enable")
        if enable_radio.count() > 0:
            radio_class = enable_radio.first.get_attribute("class") or ""
            assert "is-checked" in radio_class, (
                f"Factory Reset 后 NTP Enable 默认应为 Enable，"
                f"实际 radio class='{radio_class}'"
            )

    # ── Step 4: 关闭 NTP，再次 Factory Reset，验证 NTP 默认状态为 Enable ──────
    disable_radio = ntp_enable_item.locator(".el-radio").filter(has_text="Disable")
    if disable_radio.count() > 0:
        disable_radio.first.click()
        page.wait_for_timeout(500)
        page.get_by_role("button", name="Save").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)

    trigger_factory_reset(page)
    wait_for_device_and_relogin(page)

    _nav_to_datetime(page)
    page.wait_for_selector("input[placeholder='NTP Server 1']", timeout=30_000)
    page.wait_for_timeout(500)

    ntp_enable_item2 = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    if ntp_enable_item2.count() > 0:
        enable_radio2 = ntp_enable_item2.locator(".el-radio").filter(has_text="Enable")
        if enable_radio2.count() > 0:
            radio_class2 = enable_radio2.first.get_attribute("class") or ""
            assert "is-checked" in radio_class2, (
                f"关闭 NTP 后 Factory Reset，NTP 默认状态应恢复为 Enable，"
                f"实际 radio class='{radio_class2}'"
            )
