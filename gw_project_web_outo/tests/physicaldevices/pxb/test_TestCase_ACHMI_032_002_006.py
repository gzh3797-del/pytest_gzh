import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


# 用例编号：TestCase_ACHMI_032_002_006
# 用例标题：Description空值回退显示规则正确
# 预置条件：
#   1. AcuHMI正常工作
#   2. AcuRev-4100设备已接入AcuHMI
# 测试步骤：
#   1. 将User Channel 3的Description清空并保存
#   2. 打开显示User Channel 3名称的页面
# 预期结果：
#   1. 系统允许空值保存
#   2. 界面显示回退为默认值（如"User Channel 3"或系统默认名称）
@pytest.mark.skip(reason="需要AcuRev-4100物理设备接入AcuHMI，User and CT配置页需设备在线")
def test_TestCase_ACHMI_032_002_006(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # TODO: Navigate to the AcuRev-4100 device User and CT configuration page.
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/userAndCT")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)

    # Step 1: Clear User Channel 3 Description and save
    desc_inputs = page.locator("tbody input[type='text']")
    # Channel 3 is at index 2 (0-based)
    channel3_input = desc_inputs.nth(2)
    channel3_input.clear()

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    # System should allow empty value to be saved
    save_error = page.locator(".el-message--error").count() > 0
    assert not save_error, "User Channel 3 Description清空后保存不应报错，系统应允许空值"

    # Step 2: Verify the display falls back to default name
    # Re-open or navigate to verify the channel name
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    # TODO: Navigate to a page that shows User Channel 3 name (e.g., measurement page)
    # and verify it shows the default name rather than empty string.
    # The exact page and selector depends on product UI.
    # Example: page.goto(BASE_URL + "/#/physicalDevice/<device_id>/measurement")
    # channel3_label = page.locator("[data-channel='3']")
    # assert channel3_label.inner_text().strip() != "", \
    #     "User Channel 3 为空时显示应回退为默认值，而非空字符串"
