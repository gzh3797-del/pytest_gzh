import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


# 用例编号：TestCase_ACUHMI17_BZ_003_001
# 用例标题：Enable 状态下 Alarm Logs 显示 Ack Status 列
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
#   2. 已接到了AcuRev4100，设备别名命名为Acu4100
# 测试步骤：
#   1. Physical Devices->选择Acu4100->Alarm->Alarm Config->Add Alarm，
#      填写Label=alarm_4100，Parameter=System Frequency，Min Value=70，Max Value=90，
#      Trigger DO Device=none，Trigger RO Device=none
#   2. 点击Save
#   3. Alarm->Alarm Logs页面，在Monitor Label列找到alarm_4100的告警，
#      查看Ack Status列对应是否为Unacknowledge
#   4. 返回到Alarm（全局），在System Frequency对应告警，选择Acknowledge
#   5. 再次进入物理设备Acu4100，Alarm->Alarm Logs页面，
#      在Monitor Label列找到alarm_4100的告警，查看Ack Status是否为Acknowledge
# 预期结果：
#   3. Ack Status对应为Unacknowledge
#   5. Ack Status对应为Acknowledge

_DEVICE_NAME = "Acu4100"
_ALARM_LABEL = "alarm_4100"
_ALARM_PARAM = "System Frequency"
_ALARM_MIN = "70"
_ALARM_MAX = "90"


def _nav_to_physical_devices(page):
    if "/#/physicalDevices" not in page.url or "addDevice" in page.url:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevices", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Physical Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _open_device(page, device_name: str):
    _nav_to_physical_devices(page)
    page.locator("tr.el-table__row").filter(has_text=device_name).first.locator("td").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(600)


def _nav_device_submenu(page, submenu: str):
    """Navigate to a submenu item (Alarm Config / Alarm Logs) on the device detail page."""
    item = page.get_by_role("menuitem", name=submenu)
    alarm_sub = page.locator("li.el-sub-menu").filter(has_text="Alarm")
    if alarm_sub.count() > 0 and not item.first.is_visible():
        alarm_sub.click()
        page.wait_for_timeout(400)
    item.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(600)


def _delete_alarm_if_exists(page, label: str):
    rows = page.locator("tr.el-table__row").filter(has_text=label)
    if rows.count() == 0:
        return
    rows.first.locator("button").last.click()
    page.wait_for_timeout(400)
    for txt in ["OK", "Confirm", "确认", "Yes"]:
        btn = page.locator("[role='dialog'] button").filter(has_text=txt)
        if btn.count() > 0:
            btn.first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
            break


def _get_ack_status_in_row(row) -> str:
    """Return the Ack Status cell text from a table row, or '' if not found."""
    for cell in row.locator("td").all():
        txt = cell.inner_text().strip()
        if "acknowledge" in txt.lower():
            return txt
    return ""


def test_TestCase_ACUHMI17_BZ_003_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # ── Step 1: Navigate to Acu4100 > Alarm Config ────────────────────────────
    _open_device(page, _DEVICE_NAME)
    _nav_device_submenu(page, "Alarm Config")

    # Cleanup: delete alarm_4100 if already exists
    _delete_alarm_if_exists(page, _ALARM_LABEL)

    # Click Add Alarm
    page.get_by_role("button", name="Add Alarm").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # ── Step 2: Fill form and save ────────────────────────────────────────────
    page.locator("[placeholder='Enter Label']").fill(_ALARM_LABEL)

    # Select Parameter dropdown
    param_item = page.locator(".el-form-item").filter(has_text="Parameter")
    param_item.locator(".el-select").click()
    page.wait_for_timeout(400)
    page.locator("li.el-select-dropdown__item").filter(has_text=_ALARM_PARAM).first.click()
    page.wait_for_timeout(300)

    # Fill Min Value and Max Value
    page.locator("[placeholder='Enter Min Value']").fill(_ALARM_MIN)
    page.locator("[placeholder='Enter Max Value']").fill(_ALARM_MAX)

    # Save
    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Verify alarm created (success toast OR alarm appears in list)
    created = (
        page.locator(".el-message--success").count() > 0
        or page.locator(".el-message").filter(has_text="created").count() > 0
        or page.locator("tr.el-table__row").filter(has_text=_ALARM_LABEL).count() > 0
    )
    assert created, f"添加Alarm '{_ALARM_LABEL}' 后应显示成功提示或列表中出现该条目"

    # ── Step 3: Global Alarm → Alarm Logs → check Ack Status = Unacknowledge ──
    # Navigate to global Alarm (left nav) → Alarm Logs
    page.locator(".left-nav-item").filter(has_text="Alarm").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    page.get_by_role("menuitem", name="Alarm Logs").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Wait for alarm_4100 to appear in global Alarm Logs (device polls every ~60s)
    for _ in range(9):
        if page.locator("tr.el-table__row").filter(has_text=_ALARM_LABEL).count() > 0:
            break
        page.wait_for_timeout(10000)
        page.reload()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)

    assert page.locator("tr.el-table__row").filter(has_text=_ALARM_LABEL).count() > 0, (
        f"全局Alarm Logs中应出现'{_ALARM_LABEL}'的记录（System Frequency触发UNDERFLOW，已等待90秒）"
    )

    # Verify Ack Status column is visible in global Alarm Logs
    expect(page.locator("th").filter(has_text="Ack Status")).to_be_visible(timeout=5000)

    # Most recent alarm_4100 row should show Unacknowledge
    first_alarm_row = page.locator("tr.el-table__row").filter(has_text=_ALARM_LABEL).first
    ack_status_before = _get_ack_status_in_row(first_alarm_row)
    assert "unacknowl" in ack_status_before.lower(), (
        f"alarm_4100最新记录的Ack Status应为'Unacknowledge'，实际为'{ack_status_before}'"
    )

    # ── Step 4: Global Alarm → Unacknowledged Alarms → Acknowledge ───────────
    # Switch from Alarm Logs tab to Unacknowledged Alarms (same global Alarm area)
    page.get_by_role("menuitem", name="Unacknowledged Alarms").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # Wait for alarm_4100 to appear (already triggered, should be immediate or within one cycle)
    for _ in range(9):
        if page.locator("tr").filter(has_text=_ALARM_LABEL).count() > 0:
            break
        page.wait_for_timeout(10000)
        page.reload()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)

    assert page.locator("tr").filter(has_text=_ALARM_LABEL).count() > 0, (
        f"全局 Unacknowledged Alarms 页应显示'{_ALARM_LABEL}'的未确认告警（已等待最多90秒）"
    )

    # Find alarm_4100 row and click Acknowledge
    alarm_row_global = page.locator("tr").filter(has_text=_ALARM_LABEL).first
    alarm_row_global.get_by_role("button", name="Acknowledge").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # ── Step 5: Back to Acu4100 Alarm Logs → check Ack Status = Acknowledged ─
    _open_device(page, _DEVICE_NAME)
    _nav_device_submenu(page, "Alarm Logs")
    page.wait_for_timeout(1000)

    rows_after = page.locator("tr.el-table__row").filter(has_text=_ALARM_LABEL)
    assert rows_after.count() > 0, f"Alarm Logs中仍应有'{_ALARM_LABEL}'的记录"

    # Find at least one row with "Acknowledged" (not "Unacknowledged")
    acknowledged_found = False
    for i in range(min(rows_after.count(), 10)):
        row = rows_after.nth(i)
        ack_txt = _get_ack_status_in_row(row)
        if ack_txt and "acknowledge" in ack_txt.lower() and not ack_txt.lower().startswith("unacknowl"):
            acknowledged_found = True
            break

    assert acknowledged_found, (
        f"执行Acknowledge后，'{_ALARM_LABEL}'的Ack Status应至少有一条'Acknowledged'记录"
    )
