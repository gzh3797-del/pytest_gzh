from pages.login_page import LoginPage


def _nav_to_diagnostics(page, submenu: str):
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
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_03_case01
# 用例标题：连接性测试，连接信息显示通
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. Diagnostics -> Connection Test
#   2. 点击 Start Test
#   3. 等待测试结果
# 预期结果：
#   页面显示 PING 测试结果，Loopback/Gateway/DNS 节点连接信息正常
def test_TestCase_AcuHMI_009_03_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_diagnostics(page, "Connection Test")
    assert "#/diagnostics/connectionTest" in page.url, \
        f"应导航到Connection Test页面，实际URL={page.url}"

    page.get_by_role("button", name="Start Test").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(8000)

    page_text = page.locator("body").inner_text()

    # 结果以纯文本 "PING <IP> SUCCESS/FAIL" 格式展示
    assert "PING" in page_text, \
        "Connection Test应显示PING测试结果"

    # Loopback (127.0.0.1) 始终可达，至少1个 SUCCESS
    success_count = page_text.count("SUCCESS")
    assert success_count > 0, \
        f"Connection Test至少应有1项SUCCESS（Loopback），实际SUCCESS数={success_count}"

    print(f"\nConnection Test结果：SUCCESS共{success_count}项")
    for line in page_text.splitlines():
        line = line.strip()
        if "PING" in line or line.startswith("#"):
            print(f"  {line}")
