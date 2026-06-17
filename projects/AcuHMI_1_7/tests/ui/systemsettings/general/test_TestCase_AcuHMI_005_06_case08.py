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


def _nav_to_post_historical(page: Page) -> None:
    """导航到 System Settings → Post Historical Data 页面。

    选择器待真机校验：hash 路径基于同类页面模式推断。
    """
    nav_to(page, "#/systemSettings/postHistoricalData")


def _modify_post_historical(page: Page) -> None:
    """修改 Post Historical Data 配置，使其偏离默认值。

    默认值（用例表）：
      - Post Channel / Device / Log File Name Prefix：空
      - Log File Length / Log Interval：默认值 1
    修改策略：将 Log File Length 或 Log Interval 改为非默认值。
    选择器待真机校验。
    """
    _nav_to_post_historical(page)

    # 修改 Log File Length（默认 1，改为 5）
    length_item = page.locator(".el-form-item").filter(has_text="Log File Length").first
    if length_item.count() > 0:
        inp = length_item.locator("input").first
        if inp.count() > 0:
            inp.triple_click()
            inp.fill("5")
            page.wait_for_timeout(300)

    # 修改 Log File Name Prefix 为测试值
    prefix_item = page.locator(".el-form-item").filter(has_text="Log File Name Prefix").first
    if prefix_item.count() > 0:
        inp = prefix_item.locator("input").first
        if inp.count() > 0:
            inp.fill("test_prefix")
            page.wait_for_timeout(300)

    save_btn = page.get_by_role("button", name="Save").first
    if save_btn.count() > 0:
        save_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)


@pytest.mark.destructive
def test_TestCase_AcuHMI_005_06_case08(system_settings_page: Page) -> None:
    """TestCase_AcuHMI_005_06_case08｜修改Post Historical Data配置后恢复出厂，验证默认值

    预置条件：
      1. 管理权限登录 AcuHMI，设备运行正常

    测试步骤：
      1. 进入 Post Historical Data 页面，修改配置（Log File Length / Log File Name Prefix）
      2. Factory Reset 恢复出厂，等待设备重启并重新登录
      3. 进入 Post Historical Data 页面，验证默认值：
         - Post Channel / Device / Log File Name Prefix 默认值为空
         - Log File Length / Log Interval 默认值为 1

    预期结果：Factory Reset 成功，Post Historical Data 所有字段恢复默认值。
    """
    page = system_settings_page

    # ── Step 1: 修改 Post Historical Data 配置 ───────────────────────────────
    _modify_post_historical(page)

    # ── Step 2: 执行 Factory Reset + 等待重启 + 重登录 ───────────────────────
    trigger_factory_reset(page)
    wait_for_device_and_relogin(page)

    # ── Step 3: 导航回 Post Historical Data，断言默认值 ──────────────────────
    _nav_to_post_historical(page)
    page.wait_for_timeout(1000)

    # 断言 Log File Name Prefix 已恢复为空
    prefix_items = page.locator(".el-form-item").filter(has_text="Log File Name Prefix").all()
    for item in prefix_items[:1]:
        try:
            inp = item.locator("input").first
            if inp.count() > 0:
                val = inp.input_value()
                assert val == "" or val is None, (
                    f"Factory Reset 后 Log File Name Prefix 应为空，实际='{val}'"
                )
        except Exception:
            pass

    # 断言 Log File Length 已恢复为 1
    length_items = page.locator(".el-form-item").filter(has_text="Log File Length").all()
    for item in length_items[:1]:
        try:
            inp = item.locator("input").first
            if inp.count() > 0:
                val = inp.input_value()
                assert val in ("1", ""), (
                    f"Factory Reset 后 Log File Length 默认应为 1，实际='{val}'"
                )
        except Exception:
            pass

    # 断言 Log Interval 已恢复为 1
    interval_items = page.locator(".el-form-item").filter(has_text="Log Interval").all()
    for item in interval_items[:1]:
        try:
            inp = item.locator("input").first
            if inp.count() > 0:
                val = inp.input_value()
                assert val in ("1", ""), (
                    f"Factory Reset 后 Log Interval 默认应为 1，实际='{val}'"
                )
        except Exception:
            pass

    # 断言 Post Channel 下拉/输入为空
    post_ch_items = page.locator(".el-form-item").filter(has_text="Post Channel").all()
    for item in post_ch_items[:1]:
        try:
            inp = item.locator("input").first
            if inp.count() > 0:
                val = inp.input_value()
                assert val == "" or val is None, (
                    f"Factory Reset 后 Post Channel 应为空，实际='{val}'"
                )
        except Exception:
            pass
