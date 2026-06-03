import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


# 用例编号：TestCase_ACHMI_032_002_009
# 用例标题：非ASCII字符禁止保存（中文字符）
# 预置条件：
#   1. AcuHMI正常工作
#   2. AcuRev-4100设备已接入AcuHMI
# 测试步骤：
#   1. 设置User Channel 1 Description=主回路1，保存
#   2. 设置Description=Name中1，保存
# 预期结果：
#   1. 保存失败，提示仅允许ASCII字符
#   2. 保存失败，提示仅允许ASCII字符
@pytest.mark.skip(reason="需要AcuRev-4100物理设备接入AcuHMI，User and CT配置页需设备在线")
def test_TestCase_ACHMI_032_002_009(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # TODO: Navigate to the AcuRev-4100 device User and CT configuration page.
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/userAndCT")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)

    def _assert_save_error(label: str):
        has_error = (
            page.locator(".el-form-item__error").count() > 0
            or page.locator(".el-message--error").count() > 0
        )
        assert has_error, f"Description='{label}'（含非ASCII字符）应保存失败，但未检测到错误提示"

    # Step 1: Chinese characters only
    desc_inputs = page.locator("tbody input[type='text']")
    desc_inputs.nth(0).fill("主回路1")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(800)
    _assert_save_error("主回路1")

    # Step 2: Mixed ASCII and Chinese characters
    desc_inputs.nth(0).fill("Name中1")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(800)
    _assert_save_error("Name中1")
