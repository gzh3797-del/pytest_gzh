from datetime import datetime, timezone, timedelta
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号：TestCase_AcuHMI_005_01_case03_03
# 用例标题：验证通过NTP服务器同步设备时间的功能
# 预置条件：管理权限登录AcuHMI，NTP Enable
# 测试步骤（重复3轮）：
#   1/3/5. 修改系统时间为错误时间
#   2/4/6. 配置 Server1=time.google.com，Server2=time.nist.gov，Server3=time.apple.com，Save → Sync
# 预期结果：
#   2/4/6. 时间同步成功，系统时间与真实时间一致（误差 ≤ 120 秒）

WRONG_DATE = "2026/01/08"
WRONG_TIME = "02:37 AM"

NTP_SERVERS = ["time.google.com", "time.nist.gov", "time.apple.com"]


def _nav_to_datetime(page):
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def _get_tz_offset(page):
    tz_fi = page.locator(".el-form-item").filter(has_text="Time Zone").first
    if tz_fi.count() == 0:
        return timedelta(hours=-4)
    tz_text = tz_fi.inner_text()
    if "Toronto" in tz_text or "Eastern" in tz_text or "New_York" in tz_text:
        return timedelta(hours=-4)
    tz_map = {
        "EST": timedelta(hours=-5), "EDT": timedelta(hours=-4),
        "CST": timedelta(hours=-6), "CDT": timedelta(hours=-5),
        "MST": timedelta(hours=-7), "MDT": timedelta(hours=-6),
        "PST": timedelta(hours=-8), "PDT": timedelta(hours=-7),
        "UTC": timedelta(hours=0),  "GMT": timedelta(hours=0),
        "CET": timedelta(hours=1),  "CEST": timedelta(hours=2),
    }
    for abbr, offset in tz_map.items():
        if abbr in tz_text:
            return offset
    return timedelta(hours=-4)


def _read_device_clock(page):
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


def _fill_ntp_servers(page, servers):
    for i, srv in enumerate(servers, 1):
        inp = page.get_by_placeholder(f"NTP Server {i}").first
        assert inp.count() > 0, f"未找到 NTP Server {i} 输入框"
        inp.fill(srv)
        page.wait_for_timeout(100)


def _one_round(page, label, tz_offset):
    """设错误时间 → 配置 Server1/2/3 → Save → Sync → 验证"""
    _set_device_clock(page, WRONG_DATE, WRONG_TIME)
    _fill_ntp_servers(page, NTP_SERVERS)

    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    field_errors = page.locator(".el-form-item__error").count()
    msg_errors   = page.locator(".el-message--error").count()
    assert field_errors == 0 and msg_errors == 0, \
        f"{label} Save 应成功（field_errors={field_errors}, msg_errors={msg_errors}）"

    sync_btn = page.get_by_role("button", name="Sync")
    assert sync_btn.count() > 0, "未找到 Sync 按钮"

    ref_utc = datetime.now(timezone.utc).replace(tzinfo=None)
    sync_btn.click()
    try:
        page.wait_for_selector("text=Synced", timeout=90000)
    except Exception:
        pass
    page.wait_for_timeout(500)

    _nav_to_datetime(page)
    device_dt, clock_str = _read_device_clock(page)
    device_utc = device_dt - tz_offset
    diff = abs((device_utc - ref_utc).total_seconds())
    assert diff <= 120, (
        f"{label} Sync 后 Device Clock 与 NTP 时间不一致：\n"
        f"  Device Clock（本地）= {clock_str}\n"
        f"  Device Clock（UTC）= {device_utc.strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"  参考 UTC = {ref_utc.strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"  差值 = {diff:.0f} 秒（允许 ≤ 120 秒）"
    )


def test_TestCase_AcuHMI_005_01_case03_03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # 启用 NTP
    _nav_to_datetime(page)
    ntp_enable_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    assert ntp_enable_item.count() > 0, "未找到 NTP Enable 字段"
    enable_radio = ntp_enable_item.locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (enable_radio.get_attribute("class") or ""):
        ntp_enable_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(600)

    tz_offset = _get_tz_offset(page)

    # 第1轮
    _one_round(page, "第1轮", tz_offset)

    # 第2轮
    _nav_to_datetime(page)
    tz_offset = _get_tz_offset(page)
    _one_round(page, "第2轮", tz_offset)

    # 第3轮
    _nav_to_datetime(page)
    tz_offset = _get_tz_offset(page)
    _one_round(page, "第3轮", tz_offset)
