import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


# 用例编号：TestCase_ACHMI_032_002_003
# 用例标题：12个User Channel可独立配置名称
# 预置条件：
#   1. AcuHMI正常工作
#   2. AcuRev-4100设备已接入AcuHMI
# 测试步骤：
#   1. 打开User and CT页面
#   2. 检查User Channel 1~12是否全部显示，每行均有Description列
#   3. 分别设置User Channel 1=Name1、User Channel 2=Name2，点击保存
# 预期结果：
#   1. User Channel 1~12全部显示，每行均有Description列
#   3. User Channel 1~12名称可独立保存
@pytest.mark.skip(reason="需要AcuRev-4100物理设备接入AcuHMI，User and CT配置页需设备在线")
def test_TestCase_ACHMI_032_002_003(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # TODO: Navigate to the AcuRev-4100 device User and CT configuration page.
    # The URL pattern may be: /#/physicalDevice/<device_id>/userAndCT
    # page.goto(BASE_URL + "/#/physicalDevice/<device_id>/userAndCT")
    # page.wait_for_load_state("networkidle")
    # page.wait_for_timeout(1000)

    # Step 2: Verify all 12 User Channels are displayed with Description column
    rows = page.locator("tbody").get_by_role("row")
    assert rows.count() >= 12, \
        f"User Channel 1~12应全部显示，实际行数: {rows.count()}"

    for i in range(1, 13):
        row = rows.nth(i - 1)
        assert row.get_by_role("cell").count() > 0, \
            f"User Channel {i} 行应包含Description列"

    # Step 3: Set Channel 1 = "Name1", Channel 2 = "Name2"
    desc_inputs = page.locator("tbody input[type='text']")
    desc_inputs.nth(0).fill("Name1")
    desc_inputs.nth(1).fill("Name2")

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    success = (
        page.locator(".el-message--success").count() > 0
        or page.get_by_text("success", exact=False).is_visible()
    )
    assert success, "User Channel 1=Name1, Channel 2=Name2 应保存成功，但未检测到成功提示"
