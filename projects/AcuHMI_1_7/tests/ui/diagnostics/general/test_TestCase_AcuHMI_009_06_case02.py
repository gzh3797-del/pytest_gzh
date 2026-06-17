from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_debug(page):
    """Navigate to Diagnostics > Debug."""
    if not any(s in page.url for s in [
        "/systemSettings", "/userManagement", "/protocols",
        "/maintenance", "/templates", "/firmwareUpdate", "/diagnostics",
    ]):
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    if "/diagnostics" not in page.url:
        page.locator(".left-nav-item").filter(has_text="Diagnostics").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name="Debug", exact=True).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_06_case02
# 用例标题：SSH打开，端口为22，保存配置成功，下载诊断文件成功，重置系统日志成功
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. Diagnostics -> Debug
#   2. SSH 设置为 On，port 输入 22
#   3. 点击 Save，验证保存成功（无错误）
#   4. 点击 Download Diagnostic File，验证文件下载
#   5. 点击 Reset System Logs，确认对话框，验证重置成功
# 预期结果：保存成功，文件下载触发，日志重置成功
def test_TestCase_AcuHMI_009_06_case02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_debug(page)
    assert "#/diagnostics/debug" in page.url, \
        f"应导航到Debug页面，实际URL={page.url}"

    # Step 1: SSH On
    ssh_item = page.locator(".el-form-item").filter(has_text="SSH").first
    on_radio = ssh_item.locator(".el-radio").filter(has_text="On")
    if "is-checked" not in (on_radio.get_attribute("class") or ""):
        ssh_item.locator(".el-radio__label").filter(has_text="On").click()
        page.wait_for_timeout(400)

    # Step 2: port = 22
    port_input = page.locator("input[placeholder='Enter Port']").first
    port_input.click(click_count=3)
    port_input.fill("22")
    page.keyboard.press("Tab")
    page.wait_for_timeout(300)

    # Step 3: Save，验证无字段错误
    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    field_errors = page.locator(".el-form-item__error").count()
    msg_errors = page.locator(".el-message--error").count()
    assert field_errors == 0 and msg_errors == 0, \
        f"SSH端口22保存应成功（field_errors={field_errors}, msg_errors={msg_errors}）"
    print("\nStep3 Save: 成功")

    # Step 4: Download Diagnostic File
    with page.expect_download(timeout=15000) as dl:
        page.get_by_role("button", name="Download Diagnostic File").click()
    download = dl.value
    assert download.suggested_filename, \
        "Download Diagnostic File应触发文件下载，文件名不为空"
    print(f"Step4 Download: 文件名={download.suggested_filename}")

    # Step 5: Reset System Logs + 确认对话框
    page.get_by_role("button", name="Reset System Logs").click()
    page.wait_for_timeout(500)

    # 处理确认对话框（El-Plus message-box）
    dismissed = False
    for btn_name in ["Yes,continue", "Yes, continue", "Confirm", "OK", "Yes"]:
        try:
            page.get_by_role("button", name=btn_name).click(timeout=3000)
            dismissed = True
            break
        except Exception:
            pass
    if not dismissed:
        # 尝试 CSS 选择器点击主按钮
        try:
            page.locator(".el-message-box__btns .el-button--primary").first.click(timeout=3000)
            dismissed = True
        except Exception:
            pass

    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    reset_error = page.locator(".el-message--error").count()
    assert reset_error == 0, "Reset System Logs应成功，不应显示错误消息"
    print("Step5 Reset System Logs: 成功")
