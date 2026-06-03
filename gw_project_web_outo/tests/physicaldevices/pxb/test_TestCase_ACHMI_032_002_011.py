import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


# 用例编号：TestCase_ACHMI_032_002_011
# 用例标题：重名与特殊ASCII符号处理正确
# 预置条件：
#   1. AcuHMI正常工作
#   2. AcuRev-4100设备已接入AcuHMI
# 测试步骤：
#   1. 设置Channel 1=Channel_A, Channel 2=Channel_A, Channel 3=U-C_3#A
#   2. 保存并检查显示
# 预期结果：
#   1. 保存成功，重名（Channel_A）和特殊ASCII符号（#、-、_）均可被接受
@pytest.mark.skip(reason="需要AcuRev-4100物理设备接入AcuHMI，User and CT配置页需设备在线")
def test_TestCase_ACHMI_032_002_011(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # TODO: Navigate to the AcuRev-4100 device User and CT configuration page.
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/userAndCT")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)

    desc_inputs = page.locator("tbody input[type='text']")

    # Set Channel 1 = "Channel_A" (duplicate name, valid)
    desc_inputs.nth(0).fill("Channel_A")
    # Set Channel 2 = "Channel_A" (same name as Channel 1, should be allowed)
    desc_inputs.nth(1).fill("Channel_A")
    # Set Channel 3 = "U-C_3#A" (special ASCII chars: hyphen, underscore, hash)
    desc_inputs.nth(2).fill("U-C_3#A")

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    success = (
        page.locator(".el-message--success").count() > 0
        or page.get_by_text("success", exact=False).is_visible()
    )
    assert success, \
        "重名（Channel_A x2）和特殊ASCII符号（U-C_3#A）应保存成功，但未检测到成功提示"

    # Verify the values are persisted
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    desc_inputs_after = page.locator("tbody input[type='text']")
    assert desc_inputs_after.nth(0).input_value() == "Channel_A", \
        "Channel 1 保存后应显示 'Channel_A'"
    assert desc_inputs_after.nth(1).input_value() == "Channel_A", \
        "Channel 2 保存后应显示 'Channel_A'（允许重名）"
    assert desc_inputs_after.nth(2).input_value() == "U-C_3#A", \
        "Channel 3 保存后应显示 'U-C_3#A'（特殊ASCII字符应被接受）"
