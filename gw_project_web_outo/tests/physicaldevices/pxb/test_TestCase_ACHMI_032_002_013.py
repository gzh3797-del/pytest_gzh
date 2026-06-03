import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_TEST_NAME_CH5 = "name5"


# 用例编号：TestCase_ACHMI_032_002_013
# 用例标题：相关界面展示新名称且THD以及Harmonic页面除外
# 预置条件：
#   1. AcuHMI正常工作
#   2. AcuRev-4100设备已接入AcuHMI
# 测试步骤：
#   1. 设置User Channel 5 Description=name5并保存
#   2. 检查所有涉及User Channel的页面（THD和Harmonic页面除外）
# 预期结果：
#   1. 保存成功
#   2. 除THD/Harmonic页面外，其他相关页面均展示name5
@pytest.mark.skip(reason="需要AcuRev-4100物理设备接入AcuHMI，User and CT配置页需设备在线")
def test_TestCase_ACHMI_032_002_013(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # TODO: Navigate to the AcuRev-4100 device User and CT configuration page.
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/userAndCT")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)

    # Step 1: Set User Channel 5 Description = "name5" and save
    desc_inputs = page.locator("tbody input[type='text']")
    # Channel 5 is at index 4 (0-based)
    desc_inputs.nth(4).fill(_TEST_NAME_CH5)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    success = (
        page.locator(".el-message--success").count() > 0
        or page.get_by_text("success", exact=False).is_visible()
    )
    assert success, f"User Channel 5 Description='{_TEST_NAME_CH5}' 应保存成功，但未检测到成功提示"

    # Step 2: Check relevant pages (non-THD, non-Harmonic) show "name5"
    # TODO: Navigate to pages that show User Channel names and verify.
    # Pages to check (adjust URLs to match actual routing):
    #   - Overview / Dashboard: page.goto(BASE_URL + "/#/physicalDevice/<id>/overview")
    #   - Energy page:          page.goto(BASE_URL + "/#/physicalDevice/<id>/energy")
    #   - Demand page:          page.goto(BASE_URL + "/#/physicalDevice/<id>/demand")
    #
    # Example check on overview page:
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/overview")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)
    # assert page.get_by_text(_TEST_NAME_CH5).count() > 0, \
    #     f"Overview页面应显示User Channel 5的名称 '{_TEST_NAME_CH5}'"
    #
    # THD and Harmonic pages should NOT be required to show the custom name:
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/thd")
    # page.wait_for_load_state("networkidle")
    # (no assertion needed for THD/Harmonic pages per spec)
