from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号：TestCase_AcuHMI_005_01_case02
# 用例标题：NTP禁用后，手动修改设备时间
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. 进入 System Settings → Date & Time
#   2. NTP Enable 切换为 Disable
#   3. 查看 NTP Server 选项是否可选
#   4. 手动修改 Device Clock，点击 Save
# 预期结果：
#   3. NTP Server 选项不可选（已禁用）
#   4. 设备时间随修改时间同步，并显示更新成功

MANUAL_DATE = "2026/03/15"
MANUAL_TIME = "09:25 AM"   # 与当前真实时间不同，月/日/时均不同


def _nav_to_datetime(page):
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def _set_device_clock(page, date_str: str, time_str: str):
    clock_input = page.get_by_placeholder("--Select Device Clock--").first
    assert clock_input.count() > 0, "未找到 Device Clock 输入框"
    clock_input.click()
    page.wait_for_timeout(500)
    date_inp = page.get_by_placeholder("Select date").first
    time_inp = page.get_by_placeholder("Select time").first
    if date_inp.is_visible():
        date_inp.click(click_count=3)
        date_inp.fill(date_str)
        page.keyboard.press("Tab")
        page.wait_for_timeout(300)
        time_inp.click(click_count=3)
        time_inp.fill(time_str)
        page.keyboard.press("Tab")
        page.wait_for_timeout(300)
        for btn_name in ["OK", "Ok", "确定", "Confirm"]:
            ok_btn = page.locator(".el-picker-panel").get_by_role("button", name=btn_name)
            if ok_btn.count() == 0:
                ok_btn = page.get_by_role("button", name=btn_name)
            if ok_btn.count() > 0 and ok_btn.first.is_visible():
                ok_btn.first.click()
                break
        page.wait_for_timeout(300)
    else:
        clock_input.fill(f"{date_str} {time_str}")
        page.keyboard.press("Enter")
        page.wait_for_timeout(300)
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)


def test_TestCase_AcuHMI_005_01_case02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # ── Step 1: 导航到 Date & Time ──────────────────────────────────────────
    _nav_to_datetime(page)

    # ── Step 2: NTP Enable → Disable ────────────────────────────────────────
    ntp_enable_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    assert ntp_enable_item.count() > 0, "未找到 NTP Enable 字段"
    disable_radio = ntp_enable_item.locator(".el-radio").filter(has_text="Disable")
    if "is-checked" not in (disable_radio.get_attribute("class") or ""):
        ntp_enable_item.locator(".el-radio__label").filter(has_text="Disable").click()
        page.wait_for_timeout(600)

    # ── Step 3: 验证 NTP Server 选项不可选（隐藏或禁用）────────────────────
    # NTP Disable 后，NTP Server 配置区域可能被完全隐藏（从 DOM 移除）
    # 或保留但处于 disabled 状态，两者均满足"不可选"要求
    ntp_inp = page.get_by_placeholder("NTP Server 1").first
    ntp_count = ntp_inp.count()
    if ntp_count > 0:
        # 仍在 DOM 中，则应为 disabled
        assert ntp_inp.is_disabled(), \
            "Step3: NTP Disable 后，NTP Server 1 应处于禁用（不可选）状态"
    # else: count==0 → 输入框已隐藏，也满足"不可选"预期

    # ── Step 4: 手动修改 Device Clock，点击 Save ─────────────────────────────
    _set_device_clock(page, MANUAL_DATE, MANUAL_TIME)

    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    field_errors = page.locator(".el-form-item__error").count()
    msg_errors = page.locator(".el-message--error").count()
    assert field_errors == 0 and msg_errors == 0, \
        f"Step4 Save 应成功（field_errors={field_errors}, msg_errors={msg_errors}）"

    # 重新导航，读取 Device Clock 验证日期已更新为手动设定值
    _nav_to_datetime(page)
    clock_str = page.get_by_placeholder("--Select Device Clock--").first.input_value().strip()
    assert clock_str, "Device Clock 不应为空"
    assert MANUAL_DATE in clock_str, \
        f"Step4 Device Clock 应更新为手动设定日期 '{MANUAL_DATE}'，实际='{clock_str}'"
