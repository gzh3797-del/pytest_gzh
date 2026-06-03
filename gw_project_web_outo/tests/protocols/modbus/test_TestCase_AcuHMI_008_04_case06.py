import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


def _nav_device_mirror(page):
    """Navigate to Protocols > Modbus > Device Mirror."""
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
    page.get_by_role("menuitem", name="Modbus").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    page.get_by_role("menuitem", name="Device Mirror").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_008_04_case06
# 用例标题：点击Download All按钮，成功下载设备寄存器映射表文件
# 预置条件：
#   1. 管理员账号登录
#   2. Device Mirror功能已开启
#   3. 至少一台设备已配置Slave ID
# 测试步骤：
#   1. Protocols→Modbus→Device Mirror
#   2. 点击Download All按钮
#   3. 查看下载文件内容
# 预期结果：
#   2. 浏览器触发文件下载，文件正常保存至本地
#   3. 文件内容包含已配置设备的Slave ID与寄存器地址映射信息
def test_TestCase_AcuHMI_008_04_case06(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_device_mirror(page)

    # Verify the Download All button is present
    download_btn = page.get_by_role("button", name="Download All", exact=False)
    if download_btn.count() == 0:
        # Try alternative locator if button text differs slightly
        download_btn = page.locator("button").filter(has_text="Download")
    expect(download_btn.first).to_be_visible(timeout=5000)

    # Use expect_download context manager to capture the download event
    with page.expect_download(timeout=15000) as download_info:
        download_btn.first.click()

    download = download_info.value

    # Assert a download was triggered (download object should exist)
    assert download is not None, "点击 Download All 后未触发文件下载"

    # Assert the downloaded file has a name (not empty)
    assert download.suggested_filename, \
        "下载文件名为空，期望包含设备寄存器映射信息的文件"

    # The file should have a reasonable extension (csv, xlsx, txt, etc.)
    filename = download.suggested_filename.lower()
    assert any(filename.endswith(ext) for ext in [".csv", ".xlsx", ".xls", ".txt", ".zip"]), \
        f"下载文件格式异常，文件名: {download.suggested_filename}"

    # Save to temp path and verify file is non-empty
    import tempfile, os
    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(filename)[1]) as tmp:
        tmp_path = tmp.name

    try:
        download.save_as(tmp_path)
        file_size = os.path.getsize(tmp_path)
        assert file_size > 0, \
            f"下载的文件为空（0字节），期望包含设备寄存器映射信息，文件名: {download.suggested_filename}"
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
