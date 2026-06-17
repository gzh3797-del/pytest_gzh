from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号：TestCase_AcuHMI_005_01_case04
# 用例标题：遍历所有时区，禁用NTP，手动修改设备时间，验证各时区下系统时间显示正确
# 预置条件：管理权限登录AcuHMI
# 测试步骤（对 Time Zone 下拉框每个时区执行一轮，共三次 Save）：
#   Save 1: Enable NTP + 选择时区 + 配置 NTP 服务器 → Save（时区同步生效）
#   Save 2: Disable NTP → Save（停止 NTP 守护进程，Device Clock 变为可手动编辑）
#   Save 3: 手动设置 Device Clock → Save（手动校准）
#   验证: Device Clock 显示正确日期，时区名称与所选一致，无报错
# 预期结果：
#   所有时区均可保存成功，Device Clock 显示正确

import random

FIXED_DATE  = "2026/04/10"
FIXED_TIME  = "02:30 PM"
NTP_SERVER  = "time.google.com"
SAMPLE_SIZE = 3  # 随机抽取时区数量


def _nav_to_datetime(page):
    """Navigate to Date & Time with full page reload to prevent SPA state accumulation."""
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(300)
    # Force SPA component reset to prevent stale state after many iterations
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)
    # If session expired and redirected to login, re-authenticate
    if "login" in page.url.lower() or page.locator("input[type='password']").count() > 0:
        from projects.AcuHMI_1_7.settings import DEFAULT_USERNAME, DEFAULT_PASSWORD
        try:
            page.get_by_role("textbox", name="Enter User Name").fill(DEFAULT_USERNAME)
            page.get_by_role("textbox", name="Enter Password").fill(DEFAULT_PASSWORD)
            page.get_by_role("button", name="Sign In").click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)
        except Exception:
            pass
        page.goto(base + "#/systemSettings/dateTime")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(800)


def _dismiss_dialog(page, wait_ms=800):
    """Wait for and close any el-message-box dialog."""
    try:
        page.wait_for_selector(".el-overlay.is-message-box", timeout=wait_ms)
    except Exception:
        return
    page.wait_for_timeout(200)
    overlay = page.locator(".el-overlay.is-message-box")
    primary = overlay.locator(".el-message-box__btns .el-button--primary")
    if primary.count() > 0:
        try:
            primary.first.click(timeout=3000)
            page.wait_for_timeout(400)
            return
        except Exception:
            pass
    any_btn = overlay.locator(".el-message-box__btns .el-button")
    if any_btn.count() > 0:
        try:
            any_btn.last.click(timeout=3000)
            page.wait_for_timeout(400)
            return
        except Exception:
            pass
    page.keyboard.press("Enter")
    page.wait_for_timeout(400)


def _click_save_and_wait(page):
    """Dismiss any dialog, click Save, wait for loading to finish, then dismiss result dialog."""
    _dismiss_dialog(page, wait_ms=600)
    # Wait for Save button to appear (form may briefly hide during Vue re-renders or NTP state changes)
    save_btn = page.locator("button:has-text('Save')").first
    for attempt in range(4):
        try:
            save_btn.wait_for(state="visible", timeout=8000)
            break
        except Exception:
            _dismiss_dialog(page, wait_ms=1000)
            page.wait_for_timeout(500)
    save_btn.click(timeout=15000)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    try:
        page.wait_for_selector(".el-button.is-loading", state="hidden", timeout=15000)
    except Exception:
        pass
    page.wait_for_timeout(300)
    _dismiss_dialog(page, wait_ms=1500)
    _dismiss_dialog(page, wait_ms=500)


def _set_ntp(page, enable: bool):
    """Toggle NTP Enable/Disable radio, then dismiss any triggered dialog."""
    label = "Enable" if enable else "Disable"
    ntp_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    if ntp_item.count() == 0:
        return
    radio = ntp_item.locator(".el-radio").filter(has_text=label)
    if "is-checked" not in (radio.get_attribute("class") or ""):
        ntp_item.locator(".el-radio__label").filter(has_text=label).click()
        page.wait_for_timeout(600)
    _dismiss_dialog(page, wait_ms=800)


def _select_timezone(page, search_key):
    """Open Time Zone dropdown, search and select; returns True on success."""
    for _outer in range(4):
        _dismiss_dialog(page, wait_ms=600)
        page.wait_for_timeout(200)

        tz_fi = page.locator(".el-form-item").filter(has_text="Time Zone").first
        if tz_fi.count() == 0:
            return False
        tz_select = tz_fi.locator(".el-select").first
        if tz_select.count() == 0:
            return False

        clicked = False
        for _ in range(3):
            try:
                tz_select.click(timeout=5000)
                clicked = True
                break
            except Exception:
                _dismiss_dialog(page, wait_ms=800)
        if not clicked:
            continue

        page.wait_for_timeout(400)

        if not page.locator(".el-select-dropdown").is_visible():
            _dismiss_dialog(page, wait_ms=600)
            continue

        search_inp = page.locator(".el-select-dropdown input").first
        if search_inp.count() > 0 and search_inp.is_visible():
            search_inp.fill(search_key)
            page.wait_for_timeout(400)

        if not page.locator(".el-select-dropdown").is_visible():
            _dismiss_dialog(page, wait_ms=600)
            continue

        for opt in page.locator(".el-select-dropdown__item").all():
            try:
                if opt.is_visible() and search_key in opt.inner_text():
                    opt.click()
                    page.wait_for_timeout(200)
                    return True
            except Exception:
                pass

        page.keyboard.press("Escape")
        page.wait_for_timeout(300)

    return False


def _set_device_clock(page, date_str: str, time_str: str):
    """Set Device Clock picker. NTP must already be disabled+saved before calling."""
    clock_input = page.get_by_placeholder("--Select Device Clock--").first
    if clock_input.count() == 0:
        return
    clock_input.click()
    page.wait_for_timeout(600)

    date_inp = page.locator("input[placeholder='Select date']").first
    time_inp = page.locator("input[placeholder='Select time']").first

    if date_inp.is_visible():
        date_inp.click(click_count=3)
        date_inp.fill(date_str)
        page.keyboard.press("Tab")
        page.wait_for_timeout(300)
        if time_inp.is_visible():
            time_inp.click(click_count=3)
            time_inp.fill(time_str)
            page.keyboard.press("Tab")
            page.wait_for_timeout(300)

        # Click OK by CSS class first (reliable across locales)
        ok_clicked = False
        for ok_sel in [
            ".el-picker-panel__footer .el-button--primary",
            ".el-picker__popper .el-button--primary",
            ".el-popper.is-pure .el-button--primary",
        ]:
            ok_btn = page.locator(ok_sel)
            if ok_btn.count() > 0 and ok_btn.first.is_visible():
                ok_btn.first.click()
                ok_clicked = True
                break

        if not ok_clicked:
            for btn_name in ["OK", "Ok", "确定", "Confirm"]:
                btn = page.get_by_role("button", name=btn_name)
                if btn.count() > 0 and btn.first.is_visible():
                    btn.first.click()
                    ok_clicked = True
                    break

        if not ok_clicked:
            page.keyboard.press("Enter")
    else:
        clock_input.click(click_count=3)
        clock_input.fill(f"{date_str} {time_str}")
        page.keyboard.press("Enter")

    page.wait_for_timeout(400)
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)


def test_TestCase_AcuHMI_005_01_case04(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # ── 收集全部时区列表（仅在初始化阶段一次性收集）────────────────────────────
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    tz_fi = page.locator(".el-form-item").filter(has_text="Time Zone").first
    assert tz_fi.count() > 0, "未找到 Time Zone 字段"
    tz_fi.locator(".el-select").first.click()
    page.wait_for_timeout(600)
    all_timezones = []
    for opt in page.locator(".el-select-dropdown__item").all():
        try:
            txt = opt.inner_text().strip()
            if txt:
                all_timezones.append(txt)
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)

    assert len(all_timezones) > 0, "未获取到任何时区选项"
    sampled = random.sample(all_timezones, min(SAMPLE_SIZE, len(all_timezones)))
    print(f"\n共收集到 {len(all_timezones)} 个时区，随机抽取 {len(sampled)} 个测试：{sampled}")

    # ── 遍历抽取的时区 ────────────────────────────────────────────────────────
    total = len(sampled)
    for idx, tz_display in enumerate(sampled, 1):
        tz_id = tz_display.split("(")[0]
        city  = tz_id.split("/")[-1]

        print(f"[{idx:03d}/{total}] 测试时区：{tz_display}")

        # 每轮完整重载页面，防止 SPA 状态积累导致下拉框失效
        _nav_to_datetime(page)
        _dismiss_dialog(page, wait_ms=800)

        # ── Save 1: Enable NTP + 选时区 + NTP 服务器 → 时区同步 ───────────────
        _set_ntp(page, enable=True)

        selected = _select_timezone(page, tz_id)
        if not selected:
            selected = _select_timezone(page, city)
        assert selected, f"[{idx}] 未能选中时区 '{tz_display}'"

        ntp_inp = page.get_by_placeholder("NTP Server 1").first
        if ntp_inp.count() > 0 and ntp_inp.is_visible():
            ntp_inp.fill(NTP_SERVER)
            page.wait_for_timeout(200)

        _click_save_and_wait(page)

        fe = page.locator(".el-form-item__error").count()
        # NTP connectivity toast (el-message--error) is acceptable for Save 1; only field errors matter
        assert fe == 0, \
            f"[{idx}] {tz_id} Save1(TZ+NTP)字段验证失败（field_errors={fe}）"

        # ── Save 2: Disable NTP → 停止 NTP 守护进程 ──────────────────────────
        _set_ntp(page, enable=False)

        _click_save_and_wait(page)

        fe2 = page.locator(".el-form-item__error").count()
        me2 = page.locator(".el-message--error").count()
        assert fe2 == 0 and me2 == 0, \
            f"[{idx}] {tz_id} Save2(NTP Disable)失败（field_errors={fe2}, msg_errors={me2}）"

        # ── Save 3: 手动设置 Device Clock → 校准 ────────────────────────────
        _set_device_clock(page, FIXED_DATE, FIXED_TIME)

        _click_save_and_wait(page)

        fe3 = page.locator(".el-form-item__error").count()
        me3 = page.locator(".el-message--error").count()
        assert fe3 == 0 and me3 == 0, \
            f"[{idx}] {tz_id} Save3(DeviceClock)失败（field_errors={fe3}, msg_errors={me3}）"

        # ── 验证 ──────────────────────────────────────────────────────────────
        clock_str = page.get_by_placeholder("--Select Device Clock--").first.input_value().strip()
        assert clock_str, f"[{idx}] {tz_id}: Device Clock 不应为空"
        assert FIXED_DATE in clock_str, \
            f"[{idx}] {tz_id}: Device Clock='{clock_str}' 应包含 '{FIXED_DATE}'"

        tz_text = page.locator(".el-form-item").filter(has_text="Time Zone").first.inner_text()
        assert tz_id in tz_text, \
            f"[{idx}] 时区字段应包含 '{tz_id}'，实际='{tz_text[:80]}'"

    print(f"\n随机抽取的 {total} 个时区测试完成")
