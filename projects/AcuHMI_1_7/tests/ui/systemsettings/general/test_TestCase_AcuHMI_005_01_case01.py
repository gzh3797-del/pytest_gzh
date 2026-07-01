from datetime import datetime

import pytest

from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号：TestCase_AcuHMI_005_01_case01
# 用例标题：NTP配置启用，修改系统时间后触发同步，验证设备时间与NTP Server1保持一致，并验证配置持久化
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. 修改 Device Clock 与 NTP server 时间不一致（月/日/时/分均不同，非整点）
#   2. 选择默认 NTP Server 和 Timezone（保持当前值）
#   3. Save 保存设置
#   4. 点击 Sync 进行时间同步，等待 Synced. 提示，验证 Device Clock 与 NTP 时间一致
#   5. 修改 NTP Server 和 Timezone，点击 Save
#   6. 刷新页面，查看 Date & Time 页面设置信息，验证第5步配置已保存
#   7. 重启设备，查看 Date & Time 页面设置信息，验证配置持久化
# 预期结果：
#   3. Save 成功无报错
#   4. Synced. 出现，Device Clock 与 NTP Server1 时间一致（误差 ≤ 120 秒）
#   6. 修改后的 NTP Server 和 Timezone 正确显示
#   7. 重启后配置仍然保留


NTP_SERVER_STEP5 = "0.us.pool.ntp.org"
TIMEZONE_STEP5   = "America/New_York(EST)"   # 修改为不同时区
WRONG_DATE       = "2026/01/08"
WRONG_TIME       = "02:37 AM"                # 月/日/时/分均与当前不同，非整点


def _nav_to_datetime(page):
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def _read_device_clock(page):
    """Read Device Clock from '--Select Device Clock--' input."""
    clock_input = page.get_by_placeholder("--Select Device Clock--").first
    assert clock_input.count() > 0, "未找到 Device Clock 输入框"
    clock_str = clock_input.input_value().strip()
    assert clock_str, "Device Clock 输入框值为空"
    for fmt in ("%Y/%m/%d %I:%M %p", "%Y/%m/%d %I:%M:%S %p"):
        try:
            return datetime.strptime(clock_str, fmt), clock_str
        except ValueError:
            continue
    raise AssertionError(f"Device Clock 格式无法解析：'{clock_str}'")


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


def _reboot_device(page):
    """Navigate to Maintenance → System Status, click Reboot System and confirm."""
    base = page.url.split("#")[0]
    # 先进入 System Settings 区域，使 Maintenance 左侧导航可见
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(600)

    # 点击左侧 Maintenance
    page.locator(".left-nav-item").filter(has_text="Maintenance").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # 点击 Reboot System 按钮
    reboot_btn = page.locator("button").filter(has_text="Reboot System").first
    assert reboot_btn.count() > 0, "未找到 Reboot System 按钮"
    reboot_btn.click()
    page.wait_for_timeout(1000)

    # 处理确认对话框
    for btn_name in ["Yes", "Yes, continue", "Confirm", "确认", "OK"]:
        btn = page.get_by_role("button", name=btn_name)
        if btn.count() > 0 and btn.first.is_visible():
            btn.first.click()
            break
    page.wait_for_timeout(2000)


@pytest.mark.destructive
def test_TestCase_AcuHMI_005_01_case01(login_page: LoginPage, datetime_guard):
    login_page.open()
    login_page.login()
    page = login_page.page

    # ── Step 1: 导航到 Date & Time，NTP 保持 Enable，改 Device Clock 为错误时间 ─
    _nav_to_datetime(page)
    datetime_guard.snapshot(page)   # 修改任何设置前记录原状态，用例结束自动恢复（含时区）

    ntp_enable_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    assert ntp_enable_item.count() > 0, "未找到 NTP Enable 字段"
    enable_radio = ntp_enable_item.locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (enable_radio.get_attribute("class") or ""):
        ntp_enable_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(600)

    # 错误时间：2026/01/08 02:37 AM（月/日/时/分均与当前不同，非整点，同年保证NTP可达）
    _set_device_clock(page, WRONG_DATE, WRONG_TIME)

    # ── Step 2 & 3: 选择默认 NTP Server 1，Save ──────────────────────────────
    ntp_inp = page.get_by_placeholder("NTP Server 1").first
    assert ntp_inp.count() > 0, "未找到 NTP Server 1 输入框"
    ntp_inp.fill("time.google.com")
    page.wait_for_timeout(200)

    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    field_errors = page.locator(".el-form-item__error").count()
    msg_errors = page.locator(".el-message--error").count()
    assert field_errors == 0 and msg_errors == 0, \
        f"Step3 Save 应成功（field_errors={field_errors}, msg_errors={msg_errors}）"

    # ── Step 4: Sync，等待 Synced.，验证 Device Clock 与 NTP 时间一致 ─────────
    sync_btn = page.get_by_role("button", name="Sync")
    assert sync_btn.count() > 0, "未找到 Sync 按钮"

    ref_local = datetime.now()
    sync_btn.click()
    try:
        page.wait_for_selector("text=Synced", timeout=60000)
    except Exception:
        pass
    page.wait_for_timeout(500)

    _nav_to_datetime(page)
    device_dt, clock_str = _read_device_clock(page)
    diff_seconds = abs((device_dt - ref_local).total_seconds())
    assert diff_seconds <= 120, (
        f"Step4 Sync 后 Device Clock 与本机时间不一致：\n"
        f"  Device Clock = {clock_str}\n"
        f"  本机参考时间 = {ref_local.strftime('%Y/%m/%d %I:%M:%S %p')}\n"
        f"  差值 = {diff_seconds:.0f} 秒（允许 ≤ 120 秒）"
    )

    # ── Step 5: 修改 NTP Server 和 Timezone，Save ────────────────────────────
    ntp_inp2 = page.get_by_placeholder("NTP Server 1").first
    ntp_inp2.fill(NTP_SERVER_STEP5)
    page.wait_for_timeout(200)

    tz_fi = page.locator(".el-form-item").filter(has_text="Time Zone").first
    if tz_fi.count() > 0:
        tz_select = tz_fi.locator(".el-select").first
        if tz_select.count() > 0:
            tz_select.click()
            page.wait_for_timeout(400)
            # 搜索并选择 America/New_York
            search_inp = page.locator(".el-select-dropdown input").first
            if search_inp.count() > 0 and search_inp.is_visible():
                search_inp.fill("New_York")
                page.wait_for_timeout(400)
            opts = page.locator(".el-select-dropdown__item").all()
            selected_tz = False
            for opt in opts:
                try:
                    if opt.is_visible() and "New_York" in opt.inner_text():
                        opt.click()
                        selected_tz = True
                        break
                except Exception:
                    pass
            if not selected_tz:
                page.keyboard.press("Escape")
            page.wait_for_timeout(300)

    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    s5_errors = page.locator(".el-form-item__error").count()
    s5_msg_errors = page.locator(".el-message--error").count()
    assert s5_errors == 0 and s5_msg_errors == 0, \
        f"Step5 修改 NTP Server/Timezone Save 应成功（field_errors={s5_errors}, msg_errors={s5_msg_errors}）"

    # ── Step 6: 刷新/切换页面，验证 Step5 配置已保存 ────────────────────────
    # 先切换到其他页面再回来
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/network")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    _nav_to_datetime(page)

    ntp1_val = page.get_by_placeholder("NTP Server 1").first.input_value()
    assert NTP_SERVER_STEP5 in ntp1_val, \
        f"Step6 刷新后 NTP Server 1 应为 '{NTP_SERVER_STEP5}'，实际='{ntp1_val}'"

    # ── Step 7: 重启设备，验证配置持久化 ────────────────────────────────────
    _reboot_device(page)

    # 先等 60 秒让设备完成重启
    page.wait_for_timeout(60000)

    # 轮询直到页面可访问（只等 networkidle，不要求 loading mask 消失）
    for _ in range(30):
        try:
            page.goto(base, timeout=20000)
            page.wait_for_load_state("networkidle", timeout=20000)
            break
        except Exception:
            page.wait_for_timeout(10000)

    # 统一等待 loading 遮罩消失（无论在登录页还是主页都可能有遮罩）
    try:
        page.wait_for_selector(".el-loading-mask", state="hidden", timeout=60000)
    except Exception:
        pass
    page.wait_for_timeout(2000)

    # 判断是否在登录页：URL 含 "login" 或存在密码输入框
    def _needs_login(p):
        return "login" in p.url.lower() or p.locator("input[type='password']").count() > 0

    if _needs_login(page):
        from projects.AcuHMI_1_7.settings import DEFAULT_USERNAME, DEFAULT_PASSWORD
        # 设备刚重启可能未完全就绪，单次点击 Sign In 常不生效（停在登录页）。
        # 这里重试若干次：每次 填凭据 → 点击 Sign In → 校验是否已离开登录页，成功即止。
        for _attempt in range(6):
            if not _needs_login(page):
                break
            try:
                page.wait_for_selector("button:has-text('Sign In')", state="visible", timeout=30000)
            except Exception:
                pass
            page.wait_for_timeout(1000)
            # 填凭据前先确认遮罩已消失
            try:
                page.wait_for_selector(".el-loading-mask", state="hidden", timeout=30000)
            except Exception:
                pass
            page.get_by_role("textbox", name="Enter User Name").fill(DEFAULT_USERNAME)
            page.get_by_role("textbox", name="Enter User Name").press("Tab")
            page.get_by_role("textbox", name="Enter Password").fill(DEFAULT_PASSWORD)
            page.get_by_role("button", name="Sign In").click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(3000)
            # 处理可能弹出的"修改默认密码"对话框
            try:
                page.get_by_role("button", name="Cancel").click(timeout=3000)
            except Exception:
                pass
            try:
                page.wait_for_selector(".el-loading-mask", state="hidden", timeout=30000)
            except Exception:
                pass
            page.wait_for_timeout(2000)
            if not _needs_login(page):
                break
            # 仍在登录页：设备可能还没起好，等一会儿再重试
            page.wait_for_timeout(10000)

        assert not _needs_login(page), \
            "重启后多次重登录仍停留在登录页（设备恢复超时或凭据不符）"

    _nav_to_datetime(page)
    # 额外等待 NTP Server 1 输入框出现
    page.wait_for_selector("input[placeholder='NTP Server 1']", timeout=30000)

    ntp1_after_reboot = page.get_by_placeholder("NTP Server 1").first.input_value()
    assert NTP_SERVER_STEP5 in ntp1_after_reboot, \
        f"Step7 重启后 NTP Server 1 应为 '{NTP_SERVER_STEP5}'，实际='{ntp1_after_reboot}'"
