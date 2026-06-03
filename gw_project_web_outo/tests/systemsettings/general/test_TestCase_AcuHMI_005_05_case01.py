import pytest
from pages.login_page import LoginPage

# 用例编号：TestCase_AcuHMI_005_05_case01
# 用例标题：Alarm Email Enable启用，配置报警邮件通知接收者，Save成功
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. System Settings → Alarm Notification
#   2. 点击 Alarm Email Enable = Enable 开启配置页面
#   3. 配置 Recipient 1=Recipient1@163.com，Email Interval 选第一项
#   4. Save
# 预期结果：
#   4. 显示配置保存成功，无错误提示


def _nav_to_alarm_notification(page):
    """Navigate to System Settings → Alarm Notification."""
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(600)
    page.locator(".el-menu-item").filter(has_text="Alarm Notification").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def test_TestCase_AcuHMI_005_05_case01(login_page: LoginPage):
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

    # Step 3: 配置 Recipient 1
    recipient_fi = page.locator(".el-form-item").filter(has_text="Recipient 1").first
    assert recipient_fi.count() > 0, "Enable 后未找到 Recipient 1 字段"
    recipient_input = recipient_fi.locator("input").first
    recipient_input.fill("Recipient1@163.com")
    page.wait_for_timeout(300)

    # 设置 Email Interval（下拉或输入框）
    interval_fi = page.locator(".el-form-item").filter(has_text="Email Interval").first
    if interval_fi.count() > 0:
        interval_sel = interval_fi.locator(".el-select").first
        if interval_sel.count() > 0:
            interval_sel.click()
            page.wait_for_timeout(400)
            first_opt = page.locator(".el-select-dropdown__item").first
            if first_opt.count() > 0 and first_opt.is_visible():
                first_opt.click()
                page.wait_for_timeout(300)
            else:
                page.keyboard.press("Escape")
        else:
            interval_fi.locator("input").first.fill("1")
            page.wait_for_timeout(300)

    # Step 4: 保存
    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    field_errors = page.locator(".el-form-item__error").count()
    msg_errors = page.locator(".el-message--error").count()
    assert field_errors == 0 and msg_errors == 0, \
        f"Alarm Email Enable 保存应成功（field_errors={field_errors}, msg_errors={msg_errors}）"
