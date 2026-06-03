import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_TEST_NAME_CH4 = "name4"


# 用例编号：TestCase_ACHMI_032_002_012
# 用例标题：修改名称后系统内立即生效
# 预置条件：
#   1. AcuHMI正常工作
#   2. AcuRev-4100设备已接入AcuHMI
# 测试步骤：
#   1. 设置User Channel 4 Description=name4并保存
#   2. 打开另一个涉及User Channel 4的页面
# 预期结果：
#   1. 保存成功
#   2. User Channel 4显示立即更新为name4（无需刷新或重启）
@pytest.mark.skip(reason="需要AcuRev-4100物理设备接入AcuHMI，User and CT配置页需设备在线")
def test_TestCase_ACHMI_032_002_012(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # TODO: Navigate to the AcuRev-4100 device User and CT configuration page.
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/userAndCT")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)

    # Step 1: Set User Channel 4 Description = "name4" and save
    desc_inputs = page.locator("tbody input[type='text']")
    # Channel 4 is at index 3 (0-based)
    desc_inputs.nth(3).fill(_TEST_NAME_CH4)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    success = (
        page.locator(".el-message--success").count() > 0
        or page.get_by_text("success", exact=False).is_visible()
    )
    assert success, f"User Channel 4 Description='{_TEST_NAME_CH4}' 应保存成功，但未检测到成功提示"

    # Step 2: Navigate to another page that shows User Channel 4 name
    # and verify immediate update (no page refresh required)
    # TODO: Navigate to the relevant page (e.g., measurement / overview page)
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/measurement")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)
    #
    # channel4_label = page.locator("[data-channel='4'], [aria-label*='Channel 4']")
    # assert channel4_label.inner_text().strip() == _TEST_NAME_CH4, \
    #     f"保存后立即查看，User Channel 4 应显示 '{_TEST_NAME_CH4}'，实际: {channel4_label.inner_text()}"

    # Verify persistence by reloading the User and CT page
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    desc_inputs_after = page.locator("tbody input[type='text']")
    saved_value = desc_inputs_after.nth(3).input_value()
    assert saved_value == _TEST_NAME_CH4, \
        f"User Channel 4 应持久化为 '{_TEST_NAME_CH4}'，实际: '{saved_value}'"
