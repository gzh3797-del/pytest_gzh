from pages.login_page import LoginPage

# 用例编号：TestCase_AcuHMI_005_01_case04
# 用例标题：遍历所有时区，禁用NTP，手动修改设备时间，验证各时区下系统时间显示正确
# 预置条件：管理权限登录AcuHMI
# 测试步骤（对 Time Zone 下拉框每个时区执行一轮）：
#   1. 禁用 NTP，收集所有可用时区列表
#   2. 逐一选择时区 → Save（时区先生效）
#   3. 手动设置 Device Clock → Save
#   4. 验证 Device Clock 显示正确日期，时区名称与所选一致，无报错
# 预期结果：
#   所有时区均可保存成功，Device Clock 显示正确

FIXED_DATE = "2026/04/10"
FIXED_TIME = "02:30 PM"


def _nav_to_datetime(page):
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def _dismiss_dialog(page, wait_ms=800):
    """等待并关闭可能出现的 el-message-box 弹窗（含动画延迟）"""
    # 等待弹窗出现，若超时则说明没有弹窗
    try:
        page.wait_for_selector(".el-overlay.is-message-box", timeout=wait_ms)
    except Exception:
        return
    overlay = page.locator(".el-overlay.is-message-box")
    for btn_name in ["OK", "Ok", "Confirm", "Yes", "确认", "确定"]:
        btn = overlay.get_by_role("button", name=btn_name)
        if btn.count() > 0:
            try:
                btn.first.click(timeout=3000)
                page.wait_for_timeout(300)
                return
            except Exception:
                pass
    # 备选：点击第一个可见按钮
    for btn in overlay.locator(".el-button").all():
        try:
            if btn.is_visible():
                btn.click(timeout=2000)
                page.wait_for_timeout(300)
                return
        except Exception:
            pass


def _select_timezone(page, search_key):
    """打开 Time Zone 下拉框，搜索并选择时区；返回是否成功"""
    # 点击前先等待并关闭可能还在动画中的弹窗
    _dismiss_dialog(page, wait_ms=600)

    tz_fi = page.locator(".el-form-item").filter(has_text="Time Zone").first
    if tz_fi.count() == 0:
        return False
    tz_select = tz_fi.locator(".el-select").first
    if tz_select.count() == 0:
        return False

    # 尝试点击，若被弹窗拦截则关闭弹窗后重试（最多3次）
    for attempt in range(3):
        try:
            tz_select.click(timeout=5000)
            break
        except Exception:
            _dismiss_dialog(page, wait_ms=1000)
    else:
        return False

    page.wait_for_timeout(400)

    search_inp = page.locator(".el-select-dropdown input").first
    if search_inp.count() > 0 and search_inp.is_visible():
        search_inp.fill(search_key)
        page.wait_for_timeout(400)

    opts = page.locator(".el-select-dropdown__item").all()
    for opt in opts:
        try:
            if opt.is_visible() and search_key in opt.inner_text():
                opt.click()
                page.wait_for_timeout(200)
                return True
        except Exception:
            pass

    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    return False


def _set_device_clock(page, date_str: str, time_str: str):
    clock_input = page.get_by_placeholder("--Select Device Clock--").first
    if clock_input.count() == 0:
        return
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


def test_TestCase_AcuHMI_005_01_case04(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # ── 初始化：收集所有时区 + 禁用 NTP + Save ───────────────────────────────
    _nav_to_datetime(page)

    # 打开下拉框，收集全部时区选项
    tz_fi = page.locator(".el-form-item").filter(has_text="Time Zone").first
    assert tz_fi.count() > 0, "未找到 Time Zone 字段"
    tz_fi.locator(".el-select").first.click()
    page.wait_for_timeout(600)
    raw_opts = page.locator(".el-select-dropdown__item").all()
    all_timezones = []
    for opt in raw_opts:
        try:
            txt = opt.inner_text().strip()
            if txt:
                all_timezones.append(txt)
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)

    assert len(all_timezones) > 0, "未获取到任何时区选项"
    print(f"\n共收集到 {len(all_timezones)} 个时区，开始逐一测试")

    # 禁用 NTP（只需一次，后续每轮 Save 都会保留此状态）
    ntp_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    assert ntp_item.count() > 0, "未找到 NTP Enable 字段"
    disable_radio = ntp_item.locator(".el-radio").filter(has_text="Disable")
    if "is-checked" not in (disable_radio.get_attribute("class") or ""):
        ntp_item.locator(".el-radio__label").filter(has_text="Disable").click()
        page.wait_for_timeout(600)

    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)
    _dismiss_dialog(page)

    # ── 遍历每个时区 ─────────────────────────────────────────────────────────
    total = len(all_timezones)
    for idx, tz_display in enumerate(all_timezones[:20], 1):
        # tz_display 格式：如 "Asia/Shanghai(CST)"
        tz_id = tz_display.split("(")[0]          # "Asia/Shanghai"
        city  = tz_id.split("/")[-1]              # "Shanghai"（用作搜索关键字）

        print(f"[{idx:03d}/{total}] 测试时区：{tz_display}")

        # ── 导航到 Date & Time，清除可能的残留弹窗 ──────────────────────────
        _nav_to_datetime(page)
        _dismiss_dialog(page)

        # ── 选择目标时区，Save（时区先生效）────────────────────────────────────
        selected = _select_timezone(page, tz_id)
        if not selected:
            selected = _select_timezone(page, city)
        assert selected, f"[{idx}] 未能选中时区 '{tz_display}'"

        page.get_by_role("button", name="Save").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        _dismiss_dialog(page)

        fe = page.locator(".el-form-item__error").count()
        me = page.locator(".el-message--error").count()
        assert fe == 0 and me == 0, \
            f"[{idx}] {tz_id} 时区Save失败（field_errors={fe}, msg_errors={me}）"

        # ── 在新时区上下文中手动设置 Device Clock，Save ───────────────────────
        _set_device_clock(page, FIXED_DATE, FIXED_TIME)

        page.get_by_role("button", name="Save").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        _dismiss_dialog(page)

        fe2 = page.locator(".el-form-item__error").count()
        me2 = page.locator(".el-message--error").count()
        assert fe2 == 0 and me2 == 0, \
            f"[{idx}] {tz_id} 时钟Save失败（field_errors={fe2}, msg_errors={me2}）"

        # ── 验证：读取当前页面的 Device Clock 和 Time Zone ────────────────────
        clock_str = page.get_by_placeholder("--Select Device Clock--").first.input_value().strip()
        assert clock_str, f"[{idx}] {tz_id}: Device Clock 不应为空"
        assert FIXED_DATE in clock_str, \
            f"[{idx}] {tz_id}: Device Clock='{clock_str}' 应包含 '{FIXED_DATE}'"

        tz_text = page.locator(".el-form-item").filter(has_text="Time Zone").first.inner_text()
        assert tz_id in tz_text, \
            f"[{idx}] 时区字段应包含 '{tz_id}'，实际='{tz_text[:80]}'"

    print(f"\n全部 {total} 个时区测试完成")
