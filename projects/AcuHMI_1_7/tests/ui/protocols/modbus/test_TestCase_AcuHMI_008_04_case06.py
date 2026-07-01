import os
import tempfile

from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


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


def _is_device_mirror_enabled(page) -> bool:
    """Device Mirror 是否已开启（Enable 单选处于选中态，el-radio 选中态带 is-checked）。"""
    return page.locator(".el-radio.is-checked").filter(has_text="Enable").count() > 0


def _set_device_mirror(page, enable: bool):
    """选中 Enable/Disable 单选并点 Save 保存。

    Device Mirror 关闭时页面不渲染 Download All 按钮，因此下载用例必须先确保其开启。
    el-radio 点击沿用同目录 008_02 用例已验证的写法（直接点 .el-radio）。
    """
    label = "Enable" if enable else "Disable"
    page.locator(".el-radio").filter(has_text=label).first.click()
    page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)


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

    # 预置条件②：确保 Device Mirror 已开启（关闭时页面无 Download All 按钮）。
    # 记录进入时的原始状态，测试结束后还原，避免污染设备配置。
    was_enabled = _is_device_mirror_enabled(page)
    if not was_enabled:
        _set_device_mirror(page, True)

    try:
        # Download All 按钮：el-button--success，文案精确为 "Download All"
        download_btn = page.locator("button.el-button--success").filter(has_text="Download All")
        expect(download_btn).to_be_visible(timeout=8000)

        # 用 expect_download 捕获下载事件
        with page.expect_download(timeout=15000) as download_info:
            download_btn.click()

        download = download_info.value
        assert download is not None, "点击 Download All 后未触发文件下载"

        # 实测下载文件名为 MirrorConfALL.csv
        filename = download.suggested_filename
        assert filename, "下载文件名为空，期望包含设备寄存器映射信息的文件"
        assert filename.lower().endswith(".csv"), \
            f"期望 .csv 格式（实测文件名 MirrorConfALL.csv），实际文件名: {filename}"

        # 保存到系统临时目录并校验非空，随后删除（不在项目目录留文件）
        tmp_path = os.path.join(tempfile.gettempdir(), filename)
        try:
            download.save_as(tmp_path)
            file_size = os.path.getsize(tmp_path)
            assert file_size > 0, \
                f"下载文件为空（0字节），期望包含设备寄存器映射信息，文件名: {filename}"
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    finally:
        # 还原 Device Mirror 到进入时的原始状态
        if not was_enabled:
            _set_device_mirror(page, False)
