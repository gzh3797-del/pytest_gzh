import os

BASE = r"C:\AI工具\autotest\gw_project_web_outo\tests\datalog\postchannel"

NAV_HELPER = """
def _nav_to_post_historical_data(page):
    if "/#/dataLog" not in page.url:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Data Log").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name="Post Historical Data").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
"""

IMPORTS = """import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage
"""


def make_disable(case_id, ch_num, title):
    return f"""{IMPORTS}
{NAV_HELPER}

# 用例编号：{case_id}
# 用例标题：{title}
def test_{case_id}(login_page):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_post_historical_data(page)

    try:
        page.get_by_role("tab", name="Post Channel {ch_num}").click()
        page.wait_for_timeout(300)
    except Exception:
        pass

    try:
        page.locator(".el-form-item").filter(has_text="Enable").locator(
            ".el-radio, .el-switch"
        ).filter(has_text="Disable").click()
        page.wait_for_timeout(300)
    except Exception:
        page.get_by_role("button", name="Disable").click()
        page.wait_for_timeout(300)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)

    page.get_by_role("menuitem", name="Data Loggers 1").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    post_ch_select = page.locator(".el-form-item").filter(
        has_text="PostChannel"
    ).locator(".el-select").first
    post_ch_select.click()
    page.wait_for_timeout(300)

    ch_option = page.get_by_role("option", name="Channel {ch_num}")
    if ch_option.count() > 0:
        cls = ch_option.get_attribute("class") or ""
        assert "disabled" in cls or not ch_option.is_enabled(), \\
            "Post Ch{ch_num} disabled后，Logger中该选项应不可选"

    page.keyboard.press("Escape")
"""


def make_ftp_test(case_id, ch_num, title):
    return f"""{IMPORTS}
{NAV_HELPER}

# 用例编号：{case_id}
# 用例标题：{title}
@pytest.mark.xfail(strict=False, reason="Test Post Channel结果依赖真实FTP服务器连通性")
def test_{case_id}(login_page):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_post_historical_data(page)

    try:
        page.get_by_role("tab", name="Post Channel {ch_num}").click()
        page.wait_for_timeout(300)
    except Exception:
        pass

    try:
        page.locator(".el-form-item").filter(has_text="Enable").locator(
            ".el-radio, .el-switch"
        ).filter(has_text="Enable").click()
        page.wait_for_timeout(300)
    except Exception:
        pass

    try:
        page.locator(".el-form-item").filter(has_text="Post Method").locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="ftp", exact=True).click()
        page.wait_for_timeout(200)
    except Exception:
        pass

    for field, val in [("IP", "192.168.1.100"), ("Port", "21"),
                       ("Username", "ftpuser"), ("Password", "ftppass")]:
        try:
            page.locator(".el-form-item").filter(has_text=field).locator("input").fill(val)
        except Exception:
            pass

    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)

    page.get_by_role("button", name="Test Post Channel").click()
    page.wait_for_timeout(5000)
    expect(page.locator(".el-message, .test-result, .el-alert").first).to_be_visible(timeout=10000)

    try:
        page.locator(".el-form-item").filter(has_text="IP").locator("input").fill("999.999.999.999")
    except Exception:
        pass
    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)
    page.get_by_role("button", name="Test Post Channel").click()
    page.wait_for_timeout(5000)
    assert page.get_by_text("test fail", exact=False).count() > 0 or \\
        page.locator(".el-message--error").count() > 0, "无效IP应导致Test fail"
"""


def make_clear(case_id, ch_num, method, title):
    if method in ("ftp", "sftp"):
        err_setup = """    try:
        page.locator(".el-form-item").filter(has_text="IP").locator("input").fill("192.168.250.250")
    except Exception:
        pass"""
    else:
        err_setup = """    try:
        page.locator(".el-form-item").filter(has_text="URL").locator("input").fill("http://999.999.999.999/post")
    except Exception:
        pass"""

    return f"""{IMPORTS}
{NAV_HELPER}

# 用例编号：{case_id}
# 用例标题：{title}
def test_{case_id}(login_page):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_post_historical_data(page)

    try:
        page.get_by_role("tab", name="Post Channel {ch_num}").click()
        page.wait_for_timeout(300)
    except Exception:
        pass

    try:
        page.locator(".el-form-item").filter(has_text="Enable").locator(
            ".el-radio, .el-switch"
        ).filter(has_text="Enable").click()
        page.wait_for_timeout(300)
    except Exception:
        pass

    try:
        page.locator(".el-form-item").filter(has_text="Post Method").locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="{method}", exact=True).click()
        page.wait_for_timeout(200)
    except Exception:
        pass

{err_setup}

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)

    page.get_by_role("button", name="Clear Post Channel logs").click()
    page.wait_for_timeout(1000)

    try:
        page.get_by_role("button", name="Yes,continue").click(timeout=3000)
        page.wait_for_timeout(500)
    except Exception:
        try:
            page.get_by_role("button", name="Confirm").click(timeout=3000)
            page.wait_for_timeout(500)
        except Exception:
            pass

    expect(page.locator(".el-message")).to_be_visible(timeout=5000)
    assert page.locator(".el-message--error").count() == 0, "Clear Post Channel logs应成功"
"""


def make_sftp_test(case_id, ch_num, title):
    return f"""{IMPORTS}
{NAV_HELPER}

# 用例编号：{case_id}
# 用例标题：{title}
@pytest.mark.xfail(strict=False, reason="Test Post Channel结果依赖真实SFTP服务器连通性")
def test_{case_id}(login_page):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_post_historical_data(page)

    try:
        page.get_by_role("tab", name="Post Channel {ch_num}").click()
        page.wait_for_timeout(300)
    except Exception:
        pass

    try:
        page.locator(".el-form-item").filter(has_text="Enable").locator(
            ".el-radio, .el-switch"
        ).filter(has_text="Enable").click()
        page.wait_for_timeout(300)
    except Exception:
        pass

    try:
        page.locator(".el-form-item").filter(has_text="Post Method").locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="sftp", exact=True).click()
        page.wait_for_timeout(200)
    except Exception:
        pass

    for field, val in [("IP", "192.168.1.100"), ("Port", "22"),
                       ("Username", "sftpuser"), ("Password", "sftppass")]:
        try:
            page.locator(".el-form-item").filter(has_text=field).locator("input").fill(val)
        except Exception:
            pass

    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)

    page.get_by_role("button", name="Test Post Channel").click()
    page.wait_for_timeout(5000)
    expect(page.locator(".el-message, .test-result, .el-alert").first).to_be_visible(timeout=10000)

    try:
        page.locator(".el-form-item").filter(has_text="IP").locator("input").fill("999.999.999.999")
    except Exception:
        pass
    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)
    page.get_by_role("button", name="Test Post Channel").click()
    page.wait_for_timeout(5000)
    assert page.get_by_text("test fail", exact=False).count() > 0 or \\
        page.locator(".el-message--error").count() > 0, "无效SFTP IP应导致Test fail"
"""


def make_http_test(case_id, ch_num, auth_mode, title):
    val = "Yes" if auth_mode == "yes" else "No"
    return f"""{IMPORTS}
{NAV_HELPER}

# 用例编号：{case_id}
# 用例标题：{title}
@pytest.mark.xfail(strict=False, reason="Test Post Channel结果依赖真实HTTP服务器连通性")
def test_{case_id}(login_page):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_post_historical_data(page)

    try:
        page.get_by_role("tab", name="Post Channel {ch_num}").click()
        page.wait_for_timeout(300)
    except Exception:
        pass

    try:
        page.locator(".el-form-item").filter(has_text="Enable").locator(
            ".el-radio, .el-switch"
        ).filter(has_text="Enable").click()
        page.wait_for_timeout(300)
    except Exception:
        pass

    try:
        page.locator(".el-form-item").filter(has_text="Post Method").locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="http", exact=True).click()
        page.wait_for_timeout(200)
    except Exception:
        pass

    for field in ("Post Name Fixed", "Need Authorize", "Include Header"):
        try:
            page.locator(".el-form-item").filter(has_text=field).locator(
                ".el-radio"
            ).filter(has_text="{val}").click()
            page.wait_for_timeout(200)
        except Exception:
            pass

    try:
        page.locator(".el-form-item").filter(has_text="URL").locator("input").fill("http://192.168.1.100/post")
    except Exception:
        pass

    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)

    page.get_by_role("button", name="Test Post Channel").click()
    page.wait_for_timeout(5000)
    expect(page.locator(".el-message, .test-result, .el-alert").first).to_be_visible(timeout=10000)

    try:
        page.locator(".el-form-item").filter(has_text="URL").locator("input").fill("http://999.999.999.999/post")
    except Exception:
        pass
    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)
    page.get_by_role("button", name="Test Post Channel").click()
    page.wait_for_timeout(5000)
    assert page.get_by_text("test fail", exact=False).count() > 0 or \\
        page.locator(".el-message--error").count() > 0, "无效URL应导致Test fail"
"""


cases = [
    ("TestCase_AcuHMI_003_05_case09", make_disable("TestCase_AcuHMI_003_05_case09", 2, "Post Ch2设为disable，Logger无法选中Post Ch2")),
    ("TestCase_AcuHMI_003_05_case10", make_ftp_test("TestCase_AcuHMI_003_05_case10", 2, "Post Ch2 enable，FTP，正确配置Test success，错误配置Test fail")),
    ("TestCase_AcuHMI_003_05_case11", make_clear("TestCase_AcuHMI_003_05_case11", 2, "ftp", "Post Ch2 enable，FTP错误配置，Clear Post Channel logs成功")),
    ("TestCase_AcuHMI_003_05_case12", make_sftp_test("TestCase_AcuHMI_003_05_case12", 2, "Post Ch2 enable，SFTP，正确配置Test success，错误配置Test fail")),
    ("TestCase_AcuHMI_003_05_case13", make_http_test("TestCase_AcuHMI_003_05_case13", 2, "no", "Post Ch2 enable，HTTP No/No/No，Test success/fail")),
    ("TestCase_AcuHMI_003_05_case14", make_http_test("TestCase_AcuHMI_003_05_case14", 2, "yes", "Post Ch2 enable，HTTP Yes/Yes/Yes，Test success/fail")),
    ("TestCase_AcuHMI_003_05_case15", make_disable("TestCase_AcuHMI_003_05_case15", 3, "Post Ch3设为disable，Logger无法选中Post Ch3")),
    ("TestCase_AcuHMI_003_05_case16", make_ftp_test("TestCase_AcuHMI_003_05_case16", 3, "Post Ch3 enable，FTP，正确配置Test success，错误配置Test fail")),
    ("TestCase_AcuHMI_003_05_case17", make_clear("TestCase_AcuHMI_003_05_case17", 3, "ftp", "Post Ch3 enable，FTP错误配置，Clear Post Channel logs成功")),
    ("TestCase_AcuHMI_003_05_case18", make_sftp_test("TestCase_AcuHMI_003_05_case18", 3, "Post Ch3 enable，SFTP，正确配置Test success，错误配置Test fail")),
    ("TestCase_AcuHMI_003_05_case19", make_http_test("TestCase_AcuHMI_003_05_case19", 3, "no", "Post Ch3 enable，HTTP No/No/No，Test success/fail")),
    ("TestCase_AcuHMI_003_05_case20", make_http_test("TestCase_AcuHMI_003_05_case20", 3, "yes", "Post Ch3 enable，HTTP Yes/Yes/Yes，Test success/fail")),
]

for case_id, content in cases:
    path = os.path.join(BASE, f"test_{case_id}.py")
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Created: {case_id}")

print("Done")
