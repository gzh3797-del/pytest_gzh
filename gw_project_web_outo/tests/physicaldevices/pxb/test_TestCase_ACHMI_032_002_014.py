import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_MODIFIED_NAME = "name1"


# 用例编号：TestCase_ACHMI_032_002_014
# 用例标题：未修改通道保持默认名称
# 预置条件：
#   1. AcuHMI正常工作
#   2. AcuRev-4100设备已接入AcuHMI
# 测试步骤：
#   1. 仅将User Channel 1 Description修改为name1并保存
#   2. 检查User Channel 2~12的显示名称
# 预期结果：
#   1. Channel 1保存成功，显示name1
#   2. User Channel 2~12仍显示默认名称（未被修改）
@pytest.mark.skip(reason="需要AcuRev-4100物理设备接入AcuHMI，User and CT配置页需设备在线")
def test_TestCase_ACHMI_032_002_014(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # TODO: Navigate to the AcuRev-4100 device User and CT configuration page.
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/userAndCT")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)

    desc_inputs = page.locator("tbody input[type='text']")

    # Record current default names for channels 2~12 before modification
    default_names = []
    for i in range(1, 12):  # indices 1-11 = channels 2-12
        default_names.append(desc_inputs.nth(i).input_value())

    # Step 1: Modify only Channel 1
    desc_inputs.nth(0).fill(_MODIFIED_NAME)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    success = (
        page.locator(".el-message--success").count() > 0
        or page.get_by_text("success", exact=False).is_visible()
    )
    assert success, f"User Channel 1 Description='{_MODIFIED_NAME}' 应保存成功，但未检测到成功提示"

    # Step 2: Verify Channel 1 shows modified name and Channel 2~12 show default names
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    desc_inputs_after = page.locator("tbody input[type='text']")

    # Channel 1 should show the modified name
    ch1_value = desc_inputs_after.nth(0).input_value()
    assert ch1_value == _MODIFIED_NAME, \
        f"User Channel 1 应显示修改后的名称 '{_MODIFIED_NAME}'，实际: '{ch1_value}'"

    # Channel 2~12 should retain their default names (unchanged)
    for i in range(1, 12):
        current_value = desc_inputs_after.nth(i).input_value()
        assert current_value == default_names[i - 1], \
            (f"User Channel {i + 1} 未被修改，应保持默认名称 '{default_names[i - 1]}'，"
             f"实际: '{current_value}'")
