import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_NAME_20 = "A" * 20   # exactly 20 characters (valid)
_NAME_21 = "A" * 21   # 21 characters (invalid - exceeds max)


# 用例编号：TestCase_ACHMI_032_002_010
# 用例标题：名称长度20可保存且21禁止保存
# 预置条件：
#   1. AcuHMI正常工作
#   2. AcuRev-4100设备已接入AcuHMI
# 测试步骤：
#   1. 设置User Channel 1 Description=20字符，保存
#   2. 设置User Channel 1 Description=21字符，保存
# 预期结果：
#   1. 20字符保存成功
#   2. 21字符保存失败
@pytest.mark.skip(reason="需要AcuRev-4100物理设备接入AcuHMI，User and CT配置页需设备在线")
def test_TestCase_ACHMI_032_002_010(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # TODO: Navigate to the AcuRev-4100 device User and CT configuration page.
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/userAndCT")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)

    desc_inputs = page.locator("tbody input[type='text']")

    # Step 1: 20 characters → save success
    desc_inputs.nth(0).fill(_NAME_20)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    success = (
        page.locator(".el-message--success").count() > 0
        or page.get_by_text("success", exact=False).is_visible()
    )
    assert success, f"Description=20字符 ('{_NAME_20}') 应保存成功，但未检测到成功提示"

    # Step 2: 21 characters → save failure
    desc_inputs.nth(0).fill(_NAME_21)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(800)

    has_error = (
        page.locator(".el-form-item__error").count() > 0
        or page.locator(".el-message--error").count() > 0
    )
    assert has_error, f"Description=21字符 ('{_NAME_21}') 应保存失败（超过最大长度20），但未检测到错误提示"
