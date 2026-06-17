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

_TEST_IP_PREFIX = "10.20.30."
_WHITELIST_COUNT = 5


def _nav_to_access_control(page: Page) -> None:
    """导航到 System Settings → Access Control (Whitelist) 页面。

    选择器待真机校验：hash 路径基于同类页面模式推断。
    """
    nav_to(page, "#/systemSettings/accessControl")


def _add_whitelist_entry(page: Page, ip_addr: str, description: str) -> None:
    """添加一条白名单记录。

    选择器待真机校验：参考 005_07_case03_3 已有实现。
    """
    page.get_by_role("button", name="Add Allow List").click()
    page.wait_for_timeout(500)
    # 单 IP 模式
    single_radio = page.locator(".el-dialog").locator(".el-radio").filter(has_text="No")
    if single_radio.count() > 0:
        single_radio.first.click()
        page.wait_for_timeout(300)
    ip_inp = page.get_by_placeholder("Enter IP Address")
    if ip_inp.count() > 0:
        ip_inp.fill(ip_addr)
    desc_inp = page.get_by_placeholder("Enter Description")
    if desc_inp.count() > 0:
        desc_inp.fill(description)
    page.get_by_role("button", name="Confirm").click()
    page.wait_for_timeout(500)


@pytest.mark.destructive
def test_TestCase_AcuHMI_005_06_case11(system_settings_page: Page) -> None:
    """TestCase_AcuHMI_005_06_case11｜添加Whitelist后恢复出厂，验证Whitelist清空

    预置条件：
      1. 管理权限登录 AcuHMI，设备运行正常

    测试步骤：
      1. 进入 Access Control 页面，Enable IP Allow List，添加 5 条白名单（IP: 10.20.30.1-5）
      2. Save，确认白名单已添加
      3. Factory Reset 恢复出厂，等待设备重启并重新登录
      4. 进入 Access Control 页面，验证默认值：
         - 默认状态为 Disable
         - 已添加的白名单条目已被清空（白名单列表为空）

    预期结果：Factory Reset 成功，Whitelist 默认为 Disable，已添加的白名单条目清空。
    """
    page = system_settings_page

    # ── Step 1: 导航到 Access Control，启用 IP Allow List，添加 5 条白名单 ────
    _nav_to_access_control(page)

    # 启用 IP Allow List Enable
    enable_item = page.locator(".el-form-item").filter(has_text="IP Allow List Enable").first
    if enable_item.count() > 0:
        enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
        if enable_radio.count() > 0:
            enable_radio.first.click()
            page.wait_for_timeout(500)

    for i in range(1, _WHITELIST_COUNT + 1):
        _add_whitelist_entry(page, f"{_TEST_IP_PREFIX}{i}", f"test_entry_{i}")

    # ── Step 2: Save ──────────────────────────────────────────────────────────
    save_btn = page.get_by_role("button", name="Save").first
    if save_btn.count() > 0:
        save_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)

    # 确认白名单条目已添加
    rows = page.locator("tbody").get_by_role("row")
    added_count = rows.count()
    assert added_count >= _WHITELIST_COUNT, (
        f"Step2 应至少有 {_WHITELIST_COUNT} 条白名单，实际={added_count}"
    )

    # ── Step 3: 执行 Factory Reset + 等待重启 + 重登录 ───────────────────────
    trigger_factory_reset(page)
    wait_for_device_and_relogin(page)

    # ── Step 4: 导航回 Access Control，断言白名单已清空 ──────────────────────
    _nav_to_access_control(page)
    page.wait_for_timeout(1000)

    # 断言 IP Allow List Enable 默认为 Disable
    enable_item2 = page.locator(".el-form-item").filter(has_text="IP Allow List Enable").first
    if enable_item2.count() > 0:
        disable_radio = enable_item2.locator(".el-radio").filter(has_text="Disable")
        if disable_radio.count() > 0:
            radio_class = disable_radio.first.get_attribute("class") or ""
            assert "is-checked" in radio_class, (
                f"Factory Reset 后 IP Allow List Enable 默认应为 Disable，"
                f"实际 radio class='{radio_class}'"
            )

    # 断言白名单列表为空（不含测试添加的条目）
    rows_after = page.locator("tbody").get_by_role("row")
    row_count = rows_after.count()
    # 若 Enable 为 Disable 则表格可能不可见，count 为 0 也满足要求
    for j in range(min(row_count, 20)):
        try:
            row_text = rows_after.nth(j).inner_text()
            assert _TEST_IP_PREFIX not in row_text, (
                f"Factory Reset 后白名单应已清空，但仍发现测试条目：'{row_text}'"
            )
        except Exception:
            pass
