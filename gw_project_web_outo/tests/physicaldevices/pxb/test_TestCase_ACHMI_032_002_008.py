import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_TEST_NAME = "MainFeed_A1-01"


# 用例编号：TestCase_ACHMI_032_002_008
# 用例标题：ASCII合法名称可保存（MainFeed_A1-01）
# 预置条件：
#   1. AcuHMI正常工作
#   2. AcuRev-4100设备已接入AcuHMI
# 测试步骤：
#   1. 设置User Channel 1 Description=MainFeed_A1-01
#   2. 保存并打开涉及User Channel 1的页面
# 预期结果：
#   1. 保存成功
#   2. 系统显示名称更新为MainFeed_A1-01
@pytest.mark.skip(reason="需要AcuRev-4100物理设备接入AcuHMI，User and CT配置页需设备在线")
def test_TestCase_ACHMI_032_002_008(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # TODO: Navigate to the AcuRev-4100 device User and CT configuration page.
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/userAndCT")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)

    # Step 1: Set User Channel 1 Description = "MainFeed_A1-01"
    desc_inputs = page.locator("tbody input[type='text']")
    channel1_input = desc_inputs.nth(0)
    channel1_input.fill(_TEST_NAME)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    success = (
        page.locator(".el-message--success").count() > 0
        or page.get_by_text("success", exact=False).is_visible()
    )
    assert success, f"ASCII合法名称 '{_TEST_NAME}' 应保存成功，但未检测到成功提示"

    # Step 2: Verify the name is updated in the UI
    # Re-navigate to the page to confirm persistence
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    # The first Description input should show the saved name
    desc_inputs_after = page.locator("tbody input[type='text']")
    saved_value = desc_inputs_after.nth(0).input_value()
    assert saved_value == _TEST_NAME, \
        f"User Channel 1 应显示保存后的名称 '{_TEST_NAME}'，实际显示: '{saved_value}'"
