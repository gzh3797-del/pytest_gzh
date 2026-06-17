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


# 用例编号：TestCase_AcuHMI_009_06_case01_21
# 用例标题：SSH打开，端口为-1，保存失败，系统提示端口超出范围（Range: 6000-9999）
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. Diagnostics -> Debug
#   2. SSH 设置为 On
#   3. port 输入 -1（无效值，有效范围 6000-9999）
#   4. 点击 Save
# 预期结果：显示字段验证错误，保存不成功
def test_TestCase_AcuHMI_009_06_case01_21(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_debug(page)
    assert "#/diagnostics/debug" in page.url, \
        f"应导航到Debug页面，实际URL={page.url}"

    # 确保 SSH 为 On
    ssh_item = page.locator(".el-form-item").filter(has_text="SSH").first
    on_radio = ssh_item.locator(".el-radio").filter(has_text="On")
    if "is-checked" not in (on_radio.get_attribute("class") or ""):
        ssh_item.locator(".el-radio__label").filter(has_text="On").click()
        page.wait_for_timeout(400)

    # 输入非法端口 -1（有效范围 6000-9999）
    port_input = page.locator("input[placeholder='Enter Port']").first
    port_input.click(click_count=3)
    port_input.fill("-1")
    page.keyboard.press("Tab")
    page.wait_for_timeout(300)

    # 点击 Save
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    # 验证出现字段错误或错误消息
    has_error = (
        page.locator(".el-form-item__error").count() > 0
        or page.locator(".el-message--error").count() > 0
    )
    assert has_error, \
        "SSH端口-1超出有效范围(6000-9999)，应显示字段验证错误"

    err_texts = [e.inner_text() for e in page.locator(".el-form-item__error").all()
                 if e.inner_text().strip()]
    print(f"\n验证错误信息: {err_texts}")
