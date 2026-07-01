from datetime import datetime
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号：TestCase_AcuHMI_005_01_case03_04
# 用例标题：配置3个NTP服务器均相同，同步时间，Device Clock与NTP Server1保持一致
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. System Settings → Date & Time
#   2. 启用 NTP Enable
#   3. 配置 NTP Server 1/2/3 均为 time.google.com
#   4. Save 验证保存成功
#   5. 点击 Sync，等待完成后重新加载页面
#   6. 读取 Device Clock 值，与参考 UTC 时间对比，误差应 ≤ 120 秒
# 预期结果：
#   4. 保存成功无报错
#   5/6. Device Clock 与 NTP Server1 时间一致（误差 ≤ 120s）


def _nav_to_datetime(page):
    """Navigate directly to System Settings → Date & Time."""
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def _read_device_clock(page):
    """Read Device Clock from '--Select Device Clock--' input.
    Format: 'YYYY/MM/DD HH:MM AM/PM', e.g. '2026/05/26 10:47 PM'
    Returns (datetime, raw_string).
    """
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


def test_TestCase_AcuHMI_005_01_case03_04(login_page: LoginPage, datetime_guard):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 导航到 Date & Time 页面
    _nav_to_datetime(page)
    datetime_guard.snapshot(page)   # 修改任何设置前记录原状态，用例结束自动恢复

    # Step 2: 启用 NTP Enable
    ntp_enable_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    assert ntp_enable_item.count() > 0, "未找到 NTP Enable 字段"
    enable_radio = ntp_enable_item.locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (enable_radio.get_attribute("class") or ""):
        ntp_enable_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(600)

    # Step 3: 配置 NTP Server 1/2/3 均为 time.google.com
    ntp_server = "time.google.com"
    for i in range(1, 4):
        inp = page.get_by_placeholder(f"NTP Server {i}").first
        assert inp.count() > 0, f"未找到 NTP Server {i} 输入框"
        inp.fill(ntp_server)
        page.wait_for_timeout(200)

    # Step 4: Save
    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    field_errors = page.locator(".el-form-item__error").count()
    msg_errors = page.locator(".el-message--error").count()
    assert field_errors == 0 and msg_errors == 0, \
        f"3个相同NTP服务器配置保存应成功（field_errors={field_errors}, msg_errors={msg_errors}）"

    # Step 5: 点击 Sync，记录参考 UTC 时间
    sync_btn = page.get_by_role("button", name="Sync")
    assert sync_btn.count() > 0, "未找到 Sync 按钮"

    ref_local = datetime.now()
    sync_btn.click()

    # 等待 "Synced." 提示出现，最多等 60 秒
    try:
        page.wait_for_selector(
            "text=Synced",
            timeout=60000,
        )
    except Exception:
        pass  # 超时仍继续，后续用时间差断言兜底
    page.wait_for_timeout(500)

    # 重新导航刷新，确保 Device Clock 显示最新值
    _nav_to_datetime(page)

    # Step 6: 读取 Device Clock，与本机本地时间对比（设备与测试机同时区，免 UTC 换算）
    device_dt, clock_str = _read_device_clock(page)
    diff_seconds = abs((device_dt - ref_local).total_seconds())

    assert diff_seconds <= 120, (
        f"Sync 后 Device Clock 与本机时间不一致：\n"
        f"  Device Clock = {clock_str}\n"
        f"  本机参考时间（Sync 时刻）= {ref_local.strftime('%Y/%m/%d %I:%M:%S %p')}\n"
        f"  差值 = {diff_seconds:.0f} 秒（允许 ≤ 120 秒）"
    )
