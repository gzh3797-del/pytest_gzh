import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_NAMES = {
    0: "name1",   # Channel 1, index 0
    1: "name2",   # Channel 2, index 1
    11: "name12", # Channel 12, index 11
}


# 用例编号：TestCase_ACHMI_032_002_015
# 用例标题：批量修改多个通道后显示与保存一致
# 预置条件：
#   1. AcuHMI正常工作
#   2. AcuRev-4100设备已接入AcuHMI
# 测试步骤：
#   1. 将Channel 1/2/12分别设为name1/name2/name12并保存
#   2. 刷新页面
#   3. 退出重新登录并打开User and CT页面
# 预期结果：
#   1. 三个通道名称全部保存成功
#   2. 刷新后仍保持name1/name2/name12
#   3. 重登后仍保持name1/name2/name12
@pytest.mark.skip(reason="需要AcuRev-4100物理设备接入AcuHMI，User and CT配置页需设备在线")
def test_TestCase_ACHMI_032_002_015(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # TODO: Navigate to the AcuRev-4100 device User and CT configuration page.
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/userAndCT")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)

    desc_inputs = page.locator("tbody input[type='text']")

    # Step 1: Set Channel 1, 2, and 12 simultaneously
    for idx, name in _NAMES.items():
        desc_inputs.nth(idx).fill(name)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    success = (
        page.locator(".el-message--success").count() > 0
        or page.get_by_text("success", exact=False).is_visible()
    )
    assert success, "批量修改Channel 1/2/12应保存成功，但未检测到成功提示"

    def _verify_names(inputs, context_desc: str):
        for idx, expected_name in _NAMES.items():
            actual = inputs.nth(idx).input_value()
            assert actual == expected_name, \
                (f"{context_desc}: Channel {idx + 1} 应显示 '{expected_name}'，"
                 f"实际: '{actual}'")

    # Step 2: Refresh page and verify persistence
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    _verify_names(page.locator("tbody input[type='text']"), "刷新后")

    # Step 3: Log out and log back in, then verify persistence
    # Navigate to logout
    try:
        page.get_by_role("button", name="Logout").click(timeout=3000)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    except Exception:
        try:
            # Some UIs have logout in a user dropdown menu
            page.locator("header").get_by_role("button").last.click(timeout=2000)
            page.wait_for_timeout(300)
            page.get_by_role("menuitem", name="Logout").click(timeout=2000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        except Exception:
            page.goto(BASE_URL + "/#/login")
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)

    # Re-login
    page.get_by_role("textbox", name="Enter User Name").fill("admin")
    from config.settings import DEFAULT_PASSWORD
    page.get_by_role("textbox", name="Enter Password").fill(DEFAULT_PASSWORD)
    page.get_by_role("button", name="Sign In").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    try:
        page.get_by_role("button", name="Cancel").click(timeout=3000)
    except Exception:
        pass

    # TODO: Navigate back to the User and CT page
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/userAndCT")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)
    # _verify_names(page.locator("tbody input[type='text']"), "重新登录后")
