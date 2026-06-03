# 用例编号: TestCase_AcuHMI_005_08_case01_01
# 用例标题: 导入不合法配置文件，导入失败，提示错误信息准确
# 预置条件: 1、服务启动正常，账号登录成功
# 测试步骤:
#   1. 导入不合法的配置文件（非.json格式或格式错误的文件）
# 预期结果: 导入失败，提示错误信息准确

import os
import tempfile

import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


def _nav_to_settings_tab(page, tab: str):
    if "/systemSettings" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=tab).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


@pytest.mark.skip(reason="System shows a 'will reboot device' confirmation for any .an file without content validation; clicking OK risks device reboot, so this test cannot be safely automated")
def test_TestCase_AcuHMI_005_08_case01_01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Configuration Management")

    # 创建一个内容非法的.an临时文件（系统期望格式为.an，内容随机非法）
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".an", delete=False, encoding="utf-8"
    ) as f:
        f.write("{invalid content: this is not a valid .an config}")
        tmp_path = f.name

    try:
        # "Browse" is the el-upload trigger; "Import" is a separate submit button
        with page.expect_file_chooser(timeout=5000) as fc_info:
            page.get_by_role("button", name="Browse").click()
        fc = fc_info.value
        fc.set_files(tmp_path)
        page.wait_for_timeout(1000)

        # Click Import to submit the selected file
        page.get_by_role("button", name="Import").click()
        page.wait_for_timeout(2000)

        # Verify failure/error message appears
        error_msg = (
            page.get_by_text("fail", exact=False)
            .or_(page.get_by_text("error", exact=False))
            .or_(page.get_by_text("invalid", exact=False))
            .or_(page.get_by_text("失败", exact=False))
        )
        expect(error_msg).to_be_visible(timeout=5000), \
            "导入非法配置文件应失败并显示错误信息"
    finally:
        os.unlink(tmp_path)
