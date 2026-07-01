from datetime import datetime
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号：TestCase_AcuHMI_005_01_case03_02
# 用例标题：通过配置NTP服务器，同步设备时间，测试多个NTP服务器场景
# 预置条件：管理权限登录AcuHMI，NTP Enable
# 测试步骤：
#   1. 修改系统时间为错误时间，配置 NTP Server1=time.google.com，Save → Sync
#   2. 修改系统时间为错误时间，Server1=time.google.co（无效），Server2=time.nist.gov，Save → Sync
#   3. 修改系统时间为错误时间，Server1=time.google.co（无效），Server2不可用，Server3=time.nist.gov，Save → Sync
# 预期结果：
#   1. 时间同步成功，系统时间与 Server1 一致
#   2. 时间同步成功，系统时间与 Server2 一致（Server1 无效后 fallback）
#   3. 时间同步成功，系统时间与 Server3 一致（Server1/2 无效后 fallback）

WRONG_DATE = "2026/01/08"
WRONG_TIME = "02:37 AM"

# time.google.co 为故意拼错的无效域名（缺少 m），保证不可达
NTP_INVALID  = "time.google.co"
# 192.0.2.x 为 RFC5737 保留地址，保证不可达
NTP_UNAVAIL  = "192.0.2.2"


def _nav_to_datetime(page):
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


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


def _fill_ntp_servers(page, srv1, srv2, srv3):
    for i, srv in enumerate([srv1, srv2, srv3], 1):
        inp = page.get_by_placeholder(f"NTP Server {i}").first
        assert inp.count() > 0, f"未找到 NTP Server {i} 输入框"
        inp.fill(srv)
        page.wait_for_timeout(100)


def _phase_sync(page, label):
    """设置错误时间 → Save → Sync → 等待 Synced. → 验证 Device Clock ≤ 120s（本机本地时间对比）"""
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

    ref_local = datetime.now()
    sync_btn.click()
    try:
        page.wait_for_selector("text=Synced", timeout=90000)
    except Exception:
        pass
    page.wait_for_timeout(500)

    _nav_to_datetime(page)
    device_dt, clock_str = _read_device_clock(page)
    diff = abs((device_dt - ref_local).total_seconds())
    assert diff <= 120, (
        f"{label} Sync 后 Device Clock 与本机时间不一致：\n"
        f"  Device Clock = {clock_str}\n"
        f"  本机参考时间 = {ref_local.strftime('%Y/%m/%d %I:%M:%S %p')}\n"
        f"  差值 = {diff:.0f} 秒（允许 ≤ 120 秒）"
    )


def test_TestCase_AcuHMI_005_01_case03_02(login_page: LoginPage, datetime_guard):
    login_page.open()
    login_page.login()
    page = login_page.page

    # ── 公共初始化：启用 NTP ────────────────────────────────────────────────
    _nav_to_datetime(page)
    datetime_guard.snapshot(page)   # 修改任何设置前记录原状态，用例结束自动恢复
    ntp_enable_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    assert ntp_enable_item.count() > 0, "未找到 NTP Enable 字段"
    enable_radio = ntp_enable_item.locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (enable_radio.get_attribute("class") or ""):
        ntp_enable_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(600)

    # ── Phase 1：仅 Server1=time.google.com，Server2/3 清空 ─────────────────
    _fill_ntp_servers(page, "time.google.com", "", "")
    _phase_sync(page, "Phase1（仅Server1有效）")

    # ── Phase 2：Server1 无效(time.google.co)，Server2=time.nist.gov ─────────
    _nav_to_datetime(page)
    _fill_ntp_servers(page, NTP_INVALID, "time.nist.gov", "")
    _phase_sync(page, "Phase2（Server1无效/Server2有效）")

    # ── Phase 3：Server1 无效，Server2 不可用，Server3=time.nist.gov ─────────
    _nav_to_datetime(page)
    _fill_ntp_servers(page, NTP_INVALID, NTP_UNAVAIL, "time.nist.gov")
    _phase_sync(page, "Phase3（Server1/2无效/Server3有效）")
