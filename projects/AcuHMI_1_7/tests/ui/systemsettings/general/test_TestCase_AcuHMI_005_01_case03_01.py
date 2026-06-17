from datetime import datetime, timezone, timedelta
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号：TestCase_AcuHMI_005_01_case03_01
# 用例标题：添加三个NTP服务器，同步设备时间与日期
# 预置条件：NTP Enable；Server1不可用，Server2和Server3可用
# 测试步骤：
#   1. 修改设备时间与NTP server时间不一致
#   2. 点击Sync进行时间同步（Server1不可用，Server2/3可用）
#   3. 修改NTP Server1、Server2不可用，Server3可用，点击Sync
# 预期结果：
#   2. 时间同步成功（Device Clock与NTP时间一致，误差≤120s）
#   3. 时间同步成功（Device Clock与NTP时间一致，误差≤120s）

# 192.0.2.x 为 RFC 5737 保留地址，保证不可达（模拟不可用 NTP Server）
NTP_UNAVAIL_1 = "192.0.2.1"
NTP_UNAVAIL_2 = "192.0.2.2"
NTP_VALID_2   = "time.nist.gov"
NTP_VALID_3   = "time.apple.com"
NTP_VALID_3B  = "time.google.com"

WRONG_DATE = "2026/01/08"
WRONG_TIME = "02:37 AM"


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


def _fill_ntp_servers(page, srv1, srv2, srv3):
    for i, srv in enumerate([srv1, srv2, srv3], 1):
        inp = page.get_by_placeholder(f"NTP Server {i}").first
        assert inp.count() > 0, f"未找到 NTP Server {i} 输入框"
        inp.fill(srv)
        page.wait_for_timeout(100)


def _sync_and_verify(page, tz_offset, label):
    """设置错误时间 → Save → Sync → 等待 Synced. → 读取 Device Clock 验证"""
    _set_device_clock(page, WRONG_DATE, WRONG_TIME)

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


def test_TestCase_AcuHMI_005_01_case03_01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # ── 初始化：启用 NTP，配置 Server1=不可用，Server2/3=可用 ──────────────
    _nav_to_datetime(page)

    ntp_enable_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    assert ntp_enable_item.count() > 0, "未找到 NTP Enable 字段"
    enable_radio = ntp_enable_item.locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (enable_radio.get_attribute("class") or ""):
        ntp_enable_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(600)

    tz_offset = _get_tz_offset(page)

    # Server1=不可达，Server2=time.nist.gov，Server3=time.apple.com
    _fill_ntp_servers(page, NTP_UNAVAIL_1, NTP_VALID_2, NTP_VALID_3)

    # ── Step 1&2：设错误时间 → Save → Sync（Server1不可用，Server2/3可用）─
    _sync_and_verify(page, tz_offset, "Step2（Server1不可用/Server2&3可用）")

    # ── Step 3：改 Server1&2 不可用，Server3=可用 → Sync ────────────────────
    _nav_to_datetime(page)
    tz_offset = _get_tz_offset(page)
    _fill_ntp_servers(page, NTP_UNAVAIL_1, NTP_UNAVAIL_2, NTP_VALID_3B)

    _sync_and_verify(page, tz_offset, "Step3（Server1&2不可用/Server3可用）")
