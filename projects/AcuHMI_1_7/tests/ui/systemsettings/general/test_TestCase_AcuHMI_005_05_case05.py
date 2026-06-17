import pytest
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_AcuHMI_005_05_case05
# 用例标题: 三个收件人邮件均为有效的电子邮件(长度100)配置成功，长度101输入被自动截断为100
# 预置条件: 管理权限登录AcuHMI网页
# 测试步骤:
#   1. System Settings → Alarm Notification
#   2. 点击 Alarm Email Enable = Enable 开启配置页面
#   3. 配置3个Recipient分别为长度100的有效邮件，Email Interval=5，Save 验证保存成功
#   4. 配置3个Recipient分别为长度101的有效邮件，读回输入框实际值，验证被截断为100字符
# 预期结果:
#   3. 长度100的邮件配置成功
#   4. 输入101字符后输入框自动截断为100字符


def _nav_to_alarm_notification(page):
    """Navigate to System Settings → Alarm Notification."""
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(600)
    page.locator(".el-menu-item").filter(has_text="Alarm Notification").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def _fill_recipients(page, email: str):
    """Fill all three Recipient fields with the given email."""
    for i in range(1, 4):
        fi = page.locator(".el-form-item").filter(has_text=f"Recipient {i}").first
        if fi.count() > 0:
            fi.locator("input").first.fill(email)
            page.wait_for_timeout(200)


def _read_recipients(page):
    """Read back actual values from all three Recipient input fields."""
    values = []
    for i in range(1, 4):
        fi = page.locator(".el-form-item").filter(has_text=f"Recipient {i}").first
        if fi.count() > 0:
            val = fi.locator("input").first.input_value()
            values.append(val)
        else:
            values.append(None)
    return values


def test_TestCase_AcuHMI_005_05_case05(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 导航到 Alarm Notification 页面
    _nav_to_alarm_notification(page)

    # Step 2: 点击 Alarm Email Enable = Enable 开启配置页面
    enable_item = page.locator(".el-form-item").filter(has_text="Alarm Email Enable").first
    assert enable_item.count() > 0, "未找到 Alarm Email Enable 字段"
    enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (enable_radio.get_attribute("class") or ""):
        enable_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(800)

    # 构造长度100的有效邮件: "a"*93 + "@qq.com" = 100字符
    email_100 = "a" * 93 + "@qq.com"
    assert len(email_100) == 100

    # 构造长度101的有效邮件: "a"*94 + "@qq.com" = 101字符
    email_101 = "a" * 94 + "@qq.com"
    assert len(email_101) == 101

    # Step 3: 三个 Recipient 填写100字符邮件，Email Interval=5，保存
    _fill_recipients(page, email_100)

    interval_fi = page.locator(".el-form-item").filter(has_text="Email Interval").first
    if interval_fi.count() > 0:
        interval_sel = interval_fi.locator(".el-select").first
        if interval_sel.count() > 0:
            interval_sel.click()
            page.wait_for_timeout(400)
            opts = page.locator(".el-select-dropdown__item").all()
            for opt in opts:
                try:
                    if opt.is_visible() and "5" in opt.inner_text():
                        opt.click()
                        break
                except Exception:
                    pass
            else:
                page.keyboard.press("Escape")
        else:
            interval_fi.locator("input").first.fill("5")
        page.wait_for_timeout(300)

    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    field_errors_100 = page.locator(".el-form-item__error").count()
    msg_errors_100 = page.locator(".el-message--error").count()
    assert field_errors_100 == 0 and msg_errors_100 == 0, \
        f"长度100的邮件应配置成功（field_errors={field_errors_100}, msg_errors={msg_errors_100}）"

    # Step 4: 三个 Recipient 填写101字符邮件，验证输入框自动截断为100字符
    _fill_recipients(page, email_101)
    page.wait_for_timeout(300)

    actual_values = _read_recipients(page)
    truncation_errors = []
    for i, val in enumerate(actual_values, start=1):
        if val is None:
            truncation_errors.append(f"Recipient {i}: 未找到输入框")
        elif len(val) != 100:
            truncation_errors.append(
                f"Recipient {i}: 输入101字符后实际长度={len(val)}，期望被截断为100"
            )

    assert len(truncation_errors) == 0, \
        f"输入101字符后应自动截断为100字符：{truncation_errors}"
