from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_diagnostics_ntp(page):
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
    page.get_by_role("menuitem", name="NTP Sync Test").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_04_case01
# 用例标题：刷新NTP同步测试，验证页面输出符合NTP协议
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. Diagnostics -> NTP Sync Test
#   2. 点击 Refresh 触发 ntpd 同步测试
#   3. 等待日志输出完成
# 预期结果：
#   页面显示 ntpd 进程日志，包含 NTP 协议关键字段：
#   ntpd 版本、NTP 标准端口 123、接口绑定信息
def test_TestCase_AcuHMI_009_04_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_diagnostics_ntp(page)
    assert "#/diagnostics/ntpSyncTest" in page.url, \
        f"应导航到NTP Sync Test页面，实际URL={page.url}"

    # 进入页面会自动触发一次 ntpd，Refresh 按钮一进来就 is-loading/disabled，
    # 需先等这次自动同步完成（按钮恢复 enabled，实测约 15s）再点，否则点禁用按钮会超时。
    refresh_btn = page.get_by_role("button", name="Refresh")
    expect(refresh_btn).to_be_enabled(timeout=45000)

    # 点 Refresh 触发一次 ntpd -dgq 同步，等本次跑完（按钮再次恢复 enabled）
    refresh_btn.click()
    page.wait_for_timeout(1500)  # 让本次同步进入 loading 状态
    expect(refresh_btn).to_be_enabled(timeout=45000)

    # 从日志容器（pre）读取 ntpd 输出
    page_text = page.locator("pre.common-card-info").first.inner_text()

    # NTP协议合规验证：ntpd进程输出存在
    assert "ntpd" in page_text, \
        "NTP Sync Test日志应包含ntpd进程输出"

    # NTP标准端口123
    assert ":123" in page_text, \
        "NTP日志应包含标准NTP端口 :123（RFC 5905规定）"

    # ntpd绑定到本地或网络接口（合法IP地址出现）
    assert "Listen" in page_text or "192.168" in page_text or "127.0.0.1" in page_text, \
        "NTP日志应显示接口绑定信息（Listen normally on ...）"

    # 日志内容非空，说明同步测试已执行
    assert len(page_text.strip()) > 200, \
        "NTP Sync Test应输出详细的ntpd日志内容"

    # 打印同步结果供参考
    lower = page_text.lower()
    if "synchronized" in lower:
        print("\nNTP同步状态：已成功同步到时间服务器")
    elif "no servers" in lower or "no suitable" in lower:
        print("\nNTP同步状态：未能连接到NTP服务器（网络问题）")
    else:
        print("\nNTP同步状态：日志已输出，同步结果待确认")
