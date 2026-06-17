import re
from projects.AcuHMI_1_7.pages.login_page import LoginPage

HOST = "www.baidu.com"


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


# 用例编号：TestCase_AcuHMI_009_02_case01
# 用例标题：输入域名www.baidu.com，勾选nslookup，点击Lookup，验证查询结果正确
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. Diagnostics -> Host Lookup
#   2. 输入域名 www.baidu.com
#   3. 勾选 nslookup 复选框
#   4. 点击 Lookup
# 预期结果：
#   <pre>区域显示 nslookup 结果，包含 Server/Name/Address 字段，
#   无错误关键词，至少返回一个合法 IPv4 地址
def test_TestCase_AcuHMI_009_02_case01(login_page: LoginPage):
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

    # 勾选 nslookup 复选框
    nslookup_cb = page.locator(".el-checkbox").filter(has_text="nslookup")
    if "is-checked" not in (nslookup_cb.get_attribute("class") or ""):
        nslookup_cb.locator(".el-checkbox__label").click()
        page.wait_for_timeout(300)

    # 点击 Lookup
    page.get_by_role("button", name="Lookup").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(6000)

    # 验证没有"请选择"校验错误
    page_text = page.locator("body").inner_text()
    assert "Please select at lease one" not in page_text, \
        "nslookup应已勾选，不应出现'Please select at lease one'错误"

    # 结果在 <pre> 元素中
    pre = page.locator("pre").first
    assert pre.count() > 0 and pre.is_visible(), \
        "nslookup查询结果应显示在<pre>元素中"

    result = pre.inner_text()

    # 验证 nslookup 协议字段完整性
    assert "Server:" in result, \
        f"nslookup结果应包含'Server:'（DNS服务器信息），实际结果=\n{result}"
    assert "Name:" in result, \
        f"nslookup结果应包含'Name:'（域名解析信息），实际结果=\n{result}"
    assert "Address" in result, \
        f"nslookup结果应包含'Address'（IP地址信息），实际结果=\n{result}"
    assert "baidu.com" in result, \
        f"nslookup结果应包含查询域名'baidu.com'，实际结果=\n{result}"

    # 验证无错误关键词
    for err_kw in ["can't find", "NXDOMAIN", "SERVFAIL", "timed out", "connection refused"]:
        assert err_kw.lower() not in result.lower(), \
            f"nslookup结果不应包含错误关键词'{err_kw}'，实际结果=\n{result}"

    # 验证至少返回一个合法 IPv4 地址
    ips = re.findall(r'\b(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b', result)
    assert len(ips) > 0, \
        f"nslookup结果应包含至少一个IPv4地址，实际结果=\n{result}"

    print(f"\nnslookup结果验证通过：域名={HOST}，解析IP={ips}")
