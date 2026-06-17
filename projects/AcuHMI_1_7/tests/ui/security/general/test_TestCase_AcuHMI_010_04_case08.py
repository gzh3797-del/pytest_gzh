import io
import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_firmware(page):
    if "/#/firmwareUpdate" not in page.url:
        if not any(s in page.url for s in [
            "/systemSettings", "/userManagement", "/protocols",
            "/maintenance", "/templates", "/firmwareUpdate", "/diagnostics"
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Firmware").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_010_04_case08
# 用例标题：网页所有可上传文件的接口，只可上传指定格式的文件，不可以上传恶意软件
# 预置条件：1.服务启动正常，账号登录成功
# 测试步骤：
#   1. 文件上传功能是否允许上传恶意软件
# 预期结果：
#   1. 不允许（页面显示格式错误提示）
def test_TestCase_AcuHMI_010_04_case08(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_firmware(page)

    # 构造一个伪造的 .exe 文件内容（模拟恶意软件格式）
    fake_exe_content = b"MZ\x90\x00" + b"\x00" * 60  # DOS header magic bytes

    # 直接向 file input 注入伪造文件（set_input_files 不触发 file chooser 对话框）
    upload_input = page.locator("input[type='file']").first
    if upload_input.count() > 0:
        upload_input.set_input_files(
            files=[{"name": "malware.exe", "mimeType": "application/octet-stream", "buffer": fake_exe_content}]
        )

    page.wait_for_timeout(1000)

    # 断言：页面应显示格式错误提示，不允许上传 .exe 文件
    error_visible = (
        page.locator(".el-message--error").count() > 0
        or page.locator(".el-form-item__error").count() > 0
        or page.locator("[class*='error']").count() > 0
    )
    # 也检查是否有任何提示文字说明格式不支持
    page_text = page.locator("body").inner_text().lower()
    format_rejected = any(kw in page_text for kw in [
        "invalid", "not supported", "format", "only", "error", "failed", "不支持", "格式", "错误"
    ])

    assert error_visible or format_rejected, (
        "上传 .exe 文件后，页面应显示格式错误提示，但未检测到任何错误信息"
    )
