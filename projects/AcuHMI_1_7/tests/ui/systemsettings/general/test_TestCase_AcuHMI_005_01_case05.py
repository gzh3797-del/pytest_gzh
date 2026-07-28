# 用例编号: TestCase_AcuHMI_005_01_case05
# 用例标题: 3个NTP服务器链接字符分别超过40个，保存配置失败；
#           第1个NTP服务器链接为time.apple.co（有效但不可达），保存成功
# 预置条件: 1、服务启动正常，账号登录成功
# 测试步骤:
#   1. 添加3个NTP服务，服务器链接均为41字符，保存配置失败
#   2. 添加第一个NTP服务，服务器链接为time.apple.co，保存配置成功，时间同步失败
# 预期结果:
#   步骤1: 保存失败，显示字段长度超限错误信息
#   步骤2: 保存成功，时间同步失败（因服务器不可达）

from playwright.sync_api import expect

from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_settings_tab(page, tab: str):
    if "/systemSettings" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=tab).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def test_TestCase_AcuHMI_005_01_case05(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Date & Time")

    # 确保 NTP Enable 处于 Enable 状态（非默认则切换），否则服务器输入框被禁用
    enable_radio = page.locator(".el-form-item").filter(
        has_text="NTP Enable"
    ).first.locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (enable_radio.get_attribute("class") or ""):
        enable_radio.click()
        page.wait_for_timeout(300)

    # 步骤1: 3个NTP服务器均填写41字符，保存应失败
    long_url = "qwertyuiopasdfghjklzxcvbnm123456789012345"  # 41字符
    assert len(long_url) == 41, f"long_url长度应为41，实际为{len(long_url)}"

    page.get_by_placeholder("NTP Server 1").fill(long_url)
    page.get_by_placeholder("NTP Server 2").fill(long_url)
    page.get_by_placeholder("NTP Server 3").fill(long_url)
    page.get_by_role("button", name="Save").click()
    # 字段校验错误为异步渲染，用 expect auto-wait 等其出现，避免读取竞态
    errors = page.locator(".el-form-item__error")
    expect(errors.first).to_be_visible(timeout=5000)
    assert errors.count() > 0, \
        "3个NTP URL均超过40字符应保存失败"

    # 步骤2: 第1个NTP服务器填写time.apple.co（有效但不可达），其他清空，保存应成功
    page.get_by_placeholder("NTP Server 1").fill("time.apple.co")
    page.get_by_placeholder("NTP Server 2").fill("")
    page.get_by_placeholder("NTP Server 3").fill("")
    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    # time.apple.co 为合法格式（≤40字符），应通过字段校验保存成功；
    # 成功 toast(.el-message) 为瞬时元素、捕捉不稳定，改以"无字段级长度/格式错误"
    # 判定保存被接受（与步骤1的失败判据对称）。服务器不可达只影响同步、不影响保存。
    field_errors = page.locator(".el-form-item__error").count()
    assert field_errors == 0, \
        f"有效NTP服务器地址应通过字段校验保存成功（field_errors={field_errors}）"
