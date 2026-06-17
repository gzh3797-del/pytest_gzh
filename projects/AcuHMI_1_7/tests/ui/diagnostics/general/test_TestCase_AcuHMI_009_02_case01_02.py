import re
from projects.AcuHMI_1_7.pages.login_page import LoginPage

HOST = "www.accuenergy.com"


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


# 用例编号：TestCase_AcuHMI_009_02_case01_02
# 用例标题：输入域名www.accuenergy.com，勾选traceroute，点击Lookup，验证traceroute结果正确
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. Diagnostics -> Host Lookup
#   2. 输入域名 www.accuenergy.com
#   3. 勾选 traceroute 复选框
#   4. 点击 Lookup，等待结果（traceroute 较慢）
# 预期结果：
#   <pre>区域显示 traceroute 结果，包含路由跳数信息，
#   至少1个合法 IPv4，无 unknown host 等错误
def test_TestCase_AcuHMI_009_02_case01_02(login_page: LoginPage):
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

    # 勾选 traceroute 复选框
    traceroute_cb = page.locator(".el-checkbox").filter(has_text="traceroute")
    if "is-checked" not in (traceroute_cb.get_attribute("class") or ""):
        traceroute_cb.locator(".el-checkbox__label").click()
        page.wait_for_timeout(300)

    # 点击 Lookup（traceroute 需要更长等待时间）
    page.get_by_role("button", name="Lookup").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(20000)

    # 验证没有"请选择"校验错误
    page_text = page.locator("body").inner_text()
    assert "Please select at lease one" not in page_text, \
        "traceroute应已勾选，不应出现'Please select at lease one'错误"

    # 结果在 <pre> 元素中
    pre = page.locator("pre").first
    assert pre.count() > 0 and pre.is_visible(), \
        "traceroute查询结果应显示在<pre>元素中"

    result = pre.inner_text()

    # 验证 traceroute 协议字段
    assert "traceroute" in result.lower(), \
        f"traceroute结果应包含'traceroute'头部行，实际结果=\n{result}"

    # 验证无错误关键词
    for err_kw in ["unknown host", "Name or service not known", "bad address"]:
        assert err_kw.lower() not in result.lower(), \
            f"traceroute结果不应包含错误关键词'{err_kw}'，实际结果=\n{result}"

    # 验证至少返回一个合法 IPv4（路由跳点）
    ips = re.findall(r'\b(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b', result)
    assert len(ips) > 0, \
        f"traceroute结果应包含至少一个IPv4跳点地址，实际结果=\n{result}"

    print(f"\ntraceroute结果验证通过：域名={HOST}，路由跳点IP={list(set(ips))}")
