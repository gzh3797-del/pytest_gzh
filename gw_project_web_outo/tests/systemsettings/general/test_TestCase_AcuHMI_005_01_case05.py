import pytest
from pages.login_page import LoginPage

# 用例编号: TestCase_AcuHMI_005_01_case05
# 用例标题: 验证NTP URL超过40字符时行为：截断为40字符 或 保存失败报错
# 预置条件: 服务启动正常，账号登录成功
# 测试步骤:
#   1. 进入 System Settings → Date & Time
#   2. 检查 NTP Enable 已启用，NTP Server 1 输入框存在，Time Zone 下拉存在
#   3. 在 NTP Server 1 输入41字符，读回实际值，判断行为：
#      a. 若自动截断为40字符 → 验证截断后长度=40
#      b. 若允许输入41字符 → Save 后应出现验证错误
# 预期结果:
#   输入超过40字符后，要么被截断为40字符，要么 Save 时报错


def _nav_to_datetime(page):
    """Navigate directly to System Settings → Date & Time."""
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def test_TestCase_AcuHMI_005_01_case05(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 导航到 Date & Time 页面
    _nav_to_datetime(page)

    # Step 2: 检查 NTP Enable 默认已启用
    ntp_enable_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    assert ntp_enable_item.count() > 0, "未找到 NTP Enable 字段"
    enable_radio = ntp_enable_item.locator(".el-radio").filter(has_text="Enable")
    assert "is-checked" in (enable_radio.get_attribute("class") or ""), \
        "NTP Enable 默认应为 Enable 状态"

    # 检查 NTP Server 1 输入框存在
    ntp_inp = page.get_by_placeholder("NTP Server 1").first
    assert ntp_inp.count() > 0 and ntp_inp.is_visible(), "NTP Server 1 输入框应存在且可见"

    # 检查 Time Zone 下拉存在
    tz_fi = page.locator(".el-form-item").filter(has_text="Time Zone").first
    assert tz_fi.count() > 0, "Time Zone 字段应存在"
    assert tz_fi.locator(".el-select").count() > 0, "Time Zone 应为下拉选择框"

    # Step 3: 输入41字符的 NTP 地址，读回实际值
    long_url = "q" * 41
    ntp_inp.fill(long_url)
    page.wait_for_timeout(300)

    actual_val = ntp_inp.input_value()
    actual_len = len(actual_val)

    if actual_len <= 40:
        # 输入框有 maxlength=40，自动截断
        assert actual_len == 40, \
            f"NTP Server 1 输入41字符后应截断为40字符，实际长度={actual_len}"
    else:
        # 输入框允许超长，保存时应报验证错误
        page.get_by_role("button", name="Save").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)

        field_errors = page.locator(".el-form-item__error").count()
        msg_errors = page.locator(".el-message--error").count()
        assert field_errors > 0 or msg_errors > 0, \
            f"NTP URL 长度={actual_len}（>40）保存后应显示验证错误"
