import re
from projects.AcuHMI_1_7.pages.login_page import LoginPage

HOST = "www.qq.com"


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


# 用例编号：TestCase_AcuHMI_009_02_case01_01
# 用例标题：输入域名www.qq.com，勾选ping，点击Lookup，验证ping结果正确
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. Diagnostics -> Host Lookup
#   2. 输入域名 www.qq.com
#   3. 勾选 ping 复选框
#   4. 点击 Lookup
# 预期结果：
#   <pre>区域显示 ping 结果，包含 PING 头、bytes from/packets transmitted 字段，
#   至少1个合法 IPv4，无 unknown host / unreachable 等错误
def test_TestCase_AcuHMI_009_02_case01_01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_diagnostics(page, "Host Lookup")
    assert "#/diagnostics/hostLookup" in page.url, \
        f"应导航到Host Lookup页面，实际URL={page.url}"

    # 输入域名
    host_input = page.locator("input[placeholder*='domain']").first
    host_input.fill(HOST)
    page.wait_for_timeout(200)

    # 勾选 ping 复选框
    ping_cb = page.locator(".el-checkbox").filter(has_text="ping")
    if "is-checked" not in (ping_cb.get_attribute("class") or ""):
        ping_cb.locator(".el-checkbox__label").click()
        page.wait_for_timeout(300)

    # 点击 Lookup（ping 需要等待更长时间）
    page.get_by_role("button", name="Lookup").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(10000)

    # 验证没有"请选择"校验错误
    page_text = page.locator("body").inner_text()
    assert "Please select at lease one" not in page_text, \
        "ping应已勾选，不应出现'Please select at lease one'错误"

    # 结果在 <pre> 元素中
    pre = page.locator("pre").first
    assert pre.count() > 0 and pre.is_visible(), \
        "ping查询结果应显示在<pre>元素中"

    result = pre.inner_text()

    # 验证 ping 协议字段完整性
    assert "PING" in result, \
        f"ping结果应包含'PING'头部行，实际结果=\n{result}"
    assert "bytes from" in result or "packets transmitted" in result, \
        f"ping结果应包含'bytes from'或'packets transmitted'字段，实际结果=\n{result}"

    # 验证无错误关键词
    for err_kw in ["unknown host", "Network is unreachable", "Name or service not known",
                   "connect: No route"]:
        assert err_kw.lower() not in result.lower(), \
            f"ping结果不应包含错误关键词'{err_kw}'，实际结果=\n{result}"

    # 验证至少返回一个合法 IPv4 地址
    ips = re.findall(r'\b(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b', result)
    assert len(ips) > 0, \
        f"ping结果应包含至少一个IPv4地址，实际结果=\n{result}"

    print(f"\nping结果验证通过：域名={HOST}，IP={ips}")
