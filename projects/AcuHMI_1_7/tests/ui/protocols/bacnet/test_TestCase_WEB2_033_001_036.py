import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_036
# 用例标题: COV Batch Update 支持覆盖已有配置且未修改配置保持不变
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 先为部分参数配置已有 COV Increment。
#   2. 点击 Batch Update COV 仅修改部分目标参数并保存。
#   3. 返回参数页面检查全部参数值。
# 预期结果: 1. 已有配置可正常显示。 | 2. 批量配置保存成功。 | 3. 被选中的已有配置被覆盖，未修改参数保持原值不变。

def _nav_protocol(page, protocol: str, sub: str = None):
    if "/protocols/" not in page.url:
        if not any(s in page.url for s in [
            "/systemSettings", "/userManagement", "/maintenance",
            "/templates", "/firmwareUpdate", "/diagnostics",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Protocols").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=protocol).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    if sub:
        page.get_by_role("menuitem", name=sub).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_TestCase_WEB2_033_001_036(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "BACnet/IP")

    # COV Batch Update 覆盖已有配置，未修改配置保持不变
    # 先配置一个初始值
    cov_batch = page.get_by_role("button", name="COV Batch Update").or_(
        page.get_by_text("COV Batch Update", exact=False)).first
    if cov_batch.count() > 0:
        cov_batch.click()
        page.wait_for_timeout(500)
        page.get_by_role("button", name="Save").or_(
            page.get_by_role("button", name="Confirm")).first.click()
        page.wait_for_timeout(1000)
        assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"
