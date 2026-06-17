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


def _check_checkbox(page, label: str):
    cb = page.locator(".el-checkbox").filter(has_text=label)
    if "is-checked" not in (cb.get_attribute("class") or ""):
        cb.locator(".el-checkbox__label").click()
        page.wait_for_timeout(200)


# 用例编号：TestCase_AcuHMI_009_02_case01_03
# 用例标题：输入域名www.accuenergy.com，同时勾选ping/nslookup/traceroute，全部查询成功
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. Diagnostics -> Host Lookup
#   2. 输入域名 www.accuenergy.com
#   3. 勾选 nslookup、ping、traceroute 三个复选框
#   4. 点击 Lookup，等待结果（traceroute最慢）
# 预期结果：
#   <pre>区域显示三种查询结果，各自包含正确协议字段，
#   均有合法 IPv4 地址，无 unknown host 等错误
def test_TestCase_AcuHMI_009_02_case01_03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_diagnostics(page, "Host Lookup")
    assert "#/diagnostics/hostLookup" in page.url, \
        f"应导航到Host Lookup页面，实际URL={page.url}"

    # 输入域名
    page.locator("input[placeholder*='domain']").first.fill(HOST)
    page.wait_for_timeout(200)

    # 同时勾选三种查询方式
    for label in ["nslookup", "ping", "traceroute"]:
        _check_checkbox(page, label)

    # 点击 Lookup（等待 traceroute 完成，最长约 25s）
    page.get_by_role("button", name="Lookup").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(25000)

    # 验证无"请选择"校验错误
    page_text = page.locator("body").inner_text()
    assert "Please select at lease one" not in page_text, \
        "三种查询方式应已勾选，不应出现'Please select at lease one'错误"

    # 收集所有 <pre> 元素内容（可能每种查询各一个）
    pre_elements = page.locator("pre").all()
    assert len(pre_elements) > 0, "查询结果应在<pre>元素中显示"
    full_result = "\n".join(p.inner_text() for p in pre_elements)

    print(f"\n=== 完整查询结果 ===\n{full_result[:1000]}")

    # ── nslookup 结果验证 ──────────────────────────────────────────────
    assert "Server:" in full_result, \
        f"nslookup结果应包含'Server:'字段"
    assert "Name:" in full_result, \
        f"nslookup结果应包含'Name:'（域名解析信息）"
    assert "accuenergy.com" in full_result, \
        f"nslookup结果应包含查询域名'accuenergy.com'"

    # ── ping 结果验证 ──────────────────────────────────────────────────
    assert "PING" in full_result, \
        f"ping结果应包含'PING'头部行"
    assert "bytes from" in full_result or "packets transmitted" in full_result, \
        f"ping结果应包含'bytes from'或'packets transmitted'字段"

    # ── traceroute 结果验证 ────────────────────────────────────────────
    assert "traceroute" in full_result.lower(), \
        f"traceroute结果应包含'traceroute'头部行"

    # ── 通用：无错误关键词 ─────────────────────────────────────────────
    for err_kw in ["unknown host", "Name or service not known", "bad address",
                   "Network is unreachable"]:
        assert err_kw.lower() not in full_result.lower(), \
            f"查询结果不应包含错误关键词'{err_kw}'"

    # ── 通用：至少有合法 IPv4 ──────────────────────────────────────────
    ips = re.findall(r'\b(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b', full_result)
    assert len(ips) > 0, "查询结果应包含至少一个合法IPv4地址"

    print(f"nslookup/ping/traceroute 全部验证通过：域名={HOST}，IP={list(set(ips))[:5]}")
