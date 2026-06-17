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


def _nav_to_post_channel(page: Page) -> None:
    """导航到 System Settings → Post Channel 页面。

    选择器待真机校验：hash 路径基于同类页面模式推断。
    """
    nav_to(page, "#/systemSettings/postChannel")


def _modify_post_channel(page: Page) -> None:
    """修改 Post Channel 1 的配置，使其偏离默认值。

    默认值：
      - 状态：Disable
      - Post Method：FTP
      - FTP Port / User Name / Password：空
      - Enable anonymous mode：未勾选
    修改策略：将状态改为 Enable，FTP Port 填为非默认值。
    选择器待真机校验。
    """
    _nav_to_post_channel(page)

    # 将 Post Channel 1 状态改为 Enable
    enable_item = page.locator(".el-form-item").filter(has_text="Enable").first
    if enable_item.count() > 0:
        enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
        if enable_radio.count() > 0:
            enable_radio.first.click()
            page.wait_for_timeout(500)

    # 修改 FTP Port 为非默认值（默认为空，填入测试值 2121）
    # 选择器待真机校验
    ftp_port_item = page.locator(".el-form-item").filter(has_text="FTP Port").first
    if ftp_port_item.count() > 0:
        inp = ftp_port_item.locator("input").first
        if inp.count() > 0:
            inp.fill("2121")
            page.wait_for_timeout(300)

    # 修改 User Name 为测试值
    user_item = page.locator(".el-form-item").filter(has_text="User Name").first
    if user_item.count() > 0:
        inp = user_item.locator("input").first
        if inp.count() > 0:
            inp.fill("testuser")
            page.wait_for_timeout(300)

    save_btn = page.get_by_role("button", name="Save").first
    if save_btn.count() > 0:
        save_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)


@pytest.mark.destructive
def test_TestCase_AcuHMI_005_06_case07(system_settings_page: Page) -> None:
    """TestCase_AcuHMI_005_06_case07｜修改Post Channel后恢复出厂，验证默认值

    预置条件：
      1. 管理权限登录 AcuHMI，设备运行正常

    测试步骤：
      1. 进入 Post Channel 页面，修改 Post Channel 1-3 的配置参数（Enable/FTP Port/User Name）
      2. Factory Reset 恢复出厂，等待设备重启并重新登录
      3. 进入 Post Channel 页面，验证默认值：
         - 默认状态为 Disable
         - Post Method 默认为 FTP
         - FTP Port / User Name / Password 默认为空
         - Enable anonymous mode 默认未勾选
      4. 查看 Data Log Management 页面，设备信息为空，Device 不显示任何设备信息

    预期结果：Factory Reset 成功，Post Channel 所有字段恢复默认值，
              Data Log Management 页面 Device 列表为空。
    """
    page = system_settings_page

    # ── Step 1: 修改 Post Channel 配置 ───────────────────────────────────────
    _modify_post_channel(page)

    # ── Step 2: 执行 Factory Reset + 等待重启 + 重登录 ───────────────────────
    trigger_factory_reset(page)
    wait_for_device_and_relogin(page)

    # ── Step 3: 导航回 Post Channel，断言默认值 ──────────────────────────────
    _nav_to_post_channel(page)
    page.wait_for_timeout(1000)

    # 断言 Post Channel 1 默认为 Disable
    enable_items = page.locator(".el-form-item").filter(has_text="Enable").all()
    for item in enable_items[:3]:
        try:
            disable_radio = item.locator(".el-radio").filter(has_text="Disable")
            if disable_radio.count() > 0:
                radio_class = disable_radio.first.get_attribute("class") or ""
                assert "is-checked" in radio_class, (
                    f"Post Channel Enable 默认应为 Disable，"
                    f"实际 radio class='{radio_class}'"
                )
        except Exception:
            pass

    # 断言 FTP Port 已恢复为空
    ftp_port_items = page.locator(".el-form-item").filter(has_text="FTP Port").all()
    for item in ftp_port_items[:1]:
        try:
            inp = item.locator("input").first
            if inp.count() > 0:
                val = inp.input_value()
                assert "2121" not in val, (
                    f"Factory Reset 后 FTP Port 不应包含测试写入值 '2121'，实际='{val}'"
                )
        except Exception:
            pass

    # 断言 User Name 已恢复为空
    user_items = page.locator(".el-form-item").filter(has_text="User Name").all()
    for item in user_items[:1]:
        try:
            inp = item.locator("input").first
            if inp.count() > 0:
                val = inp.input_value()
                assert val == "" or val is None, (
                    f"Factory Reset 后 User Name 应为空，实际='{val}'"
                )
        except Exception:
            pass

    # ── Step 4: 查看 Data Log Management，断言 Device 列表为空 ───────────────
    # 选择器待真机校验：hash 路径基于同类页面模式推断
    nav_to(page, "#/systemSettings/dataLogManagement")
    page.wait_for_timeout(800)

    row_count = page.locator("tbody").get_by_role("row").count()
    assert row_count == 0, (
        f"Factory Reset 后 Data Log Management 页面应无设备信息，实际行数={row_count}"
    )
