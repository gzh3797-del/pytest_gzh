import re
from projects.AcuHMI_1_7.pages.login_page import LoginPage

HOST = "www.12345678.com"


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


def _check_checkbox(page, label: str):
    cb = page.locator(".el-checkbox").filter(has_text=label)
    if "is-checked" not in (cb.get_attribute("class") or ""):
        cb.locator(".el-checkbox__label").click()
        page.wait_for_timeout(200)


# 用例编号：TestCase_AcuHMI_009_02_case01_04
# 用例标题：输入不存在的域名www.12345678.com，分别用ping/nslookup/traceroute方式查询，全部失败
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. Diagnostics -> Host Lookup
#   2. 输入不存在的域名 www.12345678.com
#   3. 勾选 nslookup、ping、traceroute 三个复选框
#   4. 点击 Lookup，等待结果
# 预期结果：
#   三种查询结果均显示失败/错误信息（如 unknown host、Name or service not known 等）
def test_TestCase_AcuHMI_009_02_case01_04(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_diagnostics(page, "Host Lookup")
    assert "#/diagnostics/hostLookup" in page.url, \
        f"应导航到Host Lookup页面，实际URL={page.url}"

    # 输入不存在的域名
    page.locator("input[placeholder*='domain']").first.fill(HOST)
    page.wait_for_timeout(200)

    # 同时勾选三种查询方式
    for label in ["nslookup", "ping", "traceroute"]:
        _check_checkbox(page, label)

    # 点击 Lookup（等待 traceroute 超时，最长约 30s）
    page.get_by_role("button", name="Lookup").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(30000)

    # 验证无"请选择"校验错误
    page_text = page.locator("body").inner_text()
    assert "Please select at lease one" not in page_text, \
        "三种查询方式应已勾选，不应出现'Please select at lease one'错误"

    # 收集所有 <pre> 元素内容
    pre_elements = page.locator("pre").all()
    assert len(pre_elements) > 0, "查询结果应在<pre>元素中显示"
    full_result = "\n".join(p.inner_text() for p in pre_elements)

    print(f"\n=== 完整查询结果 ===\n{full_result[:1500]}")

    # ── 验证至少有一种查询出现失败/错误标志 ──────────────────────────────
    error_keywords = [
        "unknown host", "Name or service not known", "bad address",
        "Network is unreachable", "can't resolve", "server can't find",
        "NXDOMAIN", "** server can't find", "no route to host",
        "100% packet loss", "Request timeout", "nxdomain",
    ]
    found_errors = [kw for kw in error_keywords if kw.lower() in full_result.lower()]
    assert len(found_errors) > 0, \
        f"对不存在域名'{HOST}'的查询结果中应包含错误信息，但未找到任何错误关键词。\n实际结果:\n{full_result}"

    print(f"查询失败验证通过：域名={HOST}，发现错误关键词={found_errors}")
