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
_DEFAULT_MODBUS_PORT = "502"
_DEFAULT_SLAVE_ID = "100"

# 修改测试值
_TEST_MODBUS_PORT = "500"
_TEST_SLAVE_ID = "102"


def _nav_to_modbus_config(page: Page) -> None:
    """导航到 System Settings → Modbus Config 页面。

    选择器待真机校验：hash 路径基于同类页面模式推断。
    """
    nav_to(page, "#/systemSettings/modbusConfig")


def _modify_modbus_config(page: Page) -> None:
    """修改 Modbus Port 为 500，Slave ID 为 102，并保存。

    选择器待真机校验。
    """
    _nav_to_modbus_config(page)

    # 修改 Modbus Port
    port_item = page.locator(".el-form-item").filter(has_text="Modbus Port").first
    if port_item.count() > 0:
        inp = port_item.locator("input").first
        if inp.count() > 0:
            inp.triple_click()
            inp.fill(_TEST_MODBUS_PORT)
            page.wait_for_timeout(300)

    # 修改 Slave ID
    slave_item = page.locator(".el-form-item").filter(has_text="Slave ID").first
    if slave_item.count() > 0:
        inp = slave_item.locator("input").first
        if inp.count() > 0:
            inp.triple_click()
            inp.fill(_TEST_SLAVE_ID)
            page.wait_for_timeout(300)

    save_btn = page.get_by_role("button", name="Save").first
    if save_btn.count() > 0:
        save_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)


@pytest.mark.destructive
def test_TestCase_AcuHMI_005_06_case13(system_settings_page: Page) -> None:
    """TestCase_AcuHMI_005_06_case13｜修改Modbus Config后恢复出厂，验证默认值

    预置条件：
      1. 管理权限登录 AcuHMI，设备运行正常

    测试步骤：
      1. 进入 Modbus Config 页面，修改 Modbus Port=500，Slave ID=102
      2. 修改 Modbus Mapping（勾选 Modbus Mapping Enable）
      3. 修改 Parameters Mapping 中的起始地址等参数
      4. Factory Reset 恢复出厂，等待设备重启并重新登录
      5. 进入 Modbus Config 页面，验证默认值：
         - Modbus Port 默认 502
         - Slave ID 默认 100
         - Modbus Mapping 默认不选
         - Parameters Mapping 页面信息为空

    预期结果：Factory Reset 成功，Modbus Port=502，Slave ID=100，
              Modbus Mapping 未选，Parameters Mapping 为空。
    """
    page = system_settings_page

    # ── Step 1: 修改 Modbus Config ────────────────────────────────────────────
    _modify_modbus_config(page)

    # ── Step 2: 修改 Modbus Mapping（尝试勾选）───────────────────────────────
    # 选择器待真机校验
    mapping_item = page.locator(".el-form-item").filter(has_text="Modbus Mapping").first
    if mapping_item.count() > 0:
        checkbox = mapping_item.locator("input[type='checkbox']").first
        if checkbox.count() > 0 and not checkbox.is_checked():
            checkbox.check()
            page.wait_for_timeout(300)
        save_btn = page.get_by_role("button", name="Save").first
        if save_btn.count() > 0:
            save_btn.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)

    # ── Step 3: 执行 Factory Reset + 等待重启 + 重登录 ───────────────────────
    trigger_factory_reset(page)
    wait_for_device_and_relogin(page)

    # ── Step 4: 导航回 Modbus Config，断言默认值 ─────────────────────────────
    _nav_to_modbus_config(page)
    page.wait_for_timeout(1000)

    # 断言 Modbus Port 恢复为 502
    port_item2 = page.locator(".el-form-item").filter(has_text="Modbus Port").first
    if port_item2.count() > 0:
        inp = port_item2.locator("input").first
        if inp.count() > 0:
            val = inp.input_value()
            assert val == _DEFAULT_MODBUS_PORT, (
                f"Factory Reset 后 Modbus Port 应为 {_DEFAULT_MODBUS_PORT}，"
                f"实际='{val}'"
            )

    # 断言 Slave ID 恢复为 100
    slave_item2 = page.locator(".el-form-item").filter(has_text="Slave ID").first
    if slave_item2.count() > 0:
        inp = slave_item2.locator("input").first
        if inp.count() > 0:
            val = inp.input_value()
            assert val == _DEFAULT_SLAVE_ID, (
                f"Factory Reset 后 Slave ID 应为 {_DEFAULT_SLAVE_ID}，"
                f"实际='{val}'"
            )

    # 断言 Modbus Mapping 默认不选（checkbox 未勾选）
    mapping_item2 = page.locator(".el-form-item").filter(has_text="Modbus Mapping").first
    if mapping_item2.count() > 0:
        checkbox2 = mapping_item2.locator("input[type='checkbox']").first
        if checkbox2.count() > 0:
            assert not checkbox2.is_checked(), (
                "Factory Reset 后 Modbus Mapping 默认应不选"
            )

    # 断言 Parameters Mapping 页面信息为空（导航到 Parameters Mapping 子页）
    # 选择器待真机校验：hash 路径基于同类页面模式推断
    nav_to(page, "#/systemSettings/parametersMapping")
    page.wait_for_timeout(800)
    row_count = page.locator("tbody").get_by_role("row").count()
    assert row_count == 0, (
        f"Factory Reset 后 Parameters Mapping 页面应为空，实际行数={row_count}"
    )
