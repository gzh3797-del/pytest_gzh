import time
import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

_TCP_DEVICE_NAME = "AutoTest_TCP_Delete_case10"


def _nav_to_physical_devices(page):
    if "/#/physicalDevice" not in page.url or "addDevice" in page.url:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Physical Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _click_visible_option(page, option_text: str = ""):
    all_items = page.locator(".el-select-dropdown__item").all()
    for item in all_items:
        try:
            if item.is_visible() and option_text in item.inner_text():
                item.click()
                page.wait_for_timeout(300)
                return True
        except Exception:
            pass
    for item in all_items:
        try:
            if item.is_visible():
                item.click()
                page.wait_for_timeout(300)
                return True
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    return False


def _add_tcp_device(page, device_name: str):
    """Add a TCP device via the Add Device form."""
    _nav_to_physical_devices(page)
    page.get_by_role("button", name="Add Device").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # 按页面从上到下顺序依次填写

    # 1. Device Name
    page.locator(".el-form-item").filter(has_text="Device Name").first.locator("input").first.fill(device_name)
    page.wait_for_timeout(100)

    # 2. Serial Number（必填，使用时间戳保证唯一）
    ts = str(int(time.time()))[-8:]
    page.locator(".el-form-item").filter(has_text="Serial Number").first.locator("input").first.fill(f"SN{ts}")
    page.wait_for_timeout(100)

    # 3. Template（选第一个可用模板）
    page.locator(".el-form-item").filter(has_text="Template").first.locator(".el-select").first.click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")
    page.wait_for_timeout(500)

    # 4. Protocol: 点击 TCP
    page.locator(".el-radio__label").filter(has_text="TCP").click()
    page.wait_for_timeout(500)

    # 5. IP Address（TCP 选中后出现，必填）
    ip_fi = page.locator(".el-form-item").filter(has_text="IP Address").first
    ip_fi.wait_for(state="visible", timeout=5000)
    ip_fi.locator("input").first.fill("192.168.99.99")
    page.wait_for_timeout(100)

    # 6. Modbus ID（必填，时间戳尾数避免 Slave ID 冲突，范围 10-209）
    slave_id = str(int(time.time()) % 200 + 10)
    page.locator(".el-form-item").filter(has_text="Modbus ID").first.locator("input").first.fill(slave_id)
    page.wait_for_timeout(100)

    # 7. Add to Logger（必填，选 No）
    page.locator(".el-form-item").filter(has_text="Add to Logger").first.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _click_visible_option(page, "No")

    # Save
    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)


def _delete_device_by_name(page, device_name: str) -> bool:
    """Delete device by name from Physical Devices list."""
    _nav_to_physical_devices(page)
    row = page.locator("tbody tr").filter(has_text=device_name)
    if row.count() == 0:
        return False

    # 优先点击红色删除按钮，否则点最后一个按钮
    danger_btn = row.first.locator(".el-button--danger")
    if danger_btn.count() > 0:
        danger_btn.first.click()
    else:
        row.first.locator("button").last.click()
    page.wait_for_timeout(500)

    # 确认对话框（依次尝试几个常见按钮名）
    for btn_name in ["Yes, continue", "Yes", "Confirm", "确认"]:
        btn = page.get_by_role("button", name=btn_name)
        if btn.count() > 0:
            try:
                if btn.first.is_visible():
                    btn.first.click()
                    page.wait_for_timeout(800)
                    break
            except Exception:
                pass

    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    return True


# 用例编号：TestCase_AcuHMI_001_01_case10
# 用例标题：设备手动删除TCP方式添加接入设备，删除成功
# 预置条件：
#   1. 接入设备支持 Modbus TCP
# 测试步骤：
#   1. 通过 TCP 方式添加设备（Device Name 唯一，无需物理设备在线）
#   2. 验证设备出现在 Physical Devices 列表
#   3. 删除该设备
#   4. 验证设备已从列表中消失
# 预期结果：
#   设备添加并删除成功，删除后列表中不再显示该设备
def test_TestCase_AcuHMI_001_01_case10(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # 清理上次残留（如有）
    _delete_device_by_name(page, _TCP_DEVICE_NAME)

    # Step 1: 添加 TCP 设备
    _add_tcp_device(page, _TCP_DEVICE_NAME)

    # Step 2: 验证设备出现在列表
    _nav_to_physical_devices(page)
    added_row = page.locator("tbody tr").filter(has_text=_TCP_DEVICE_NAME)
    assert added_row.count() > 0, \
        f"TCP 设备 '{_TCP_DEVICE_NAME}' 应出现在 Physical Devices 列表中，但未找到"

    # Step 3: 删除设备
    _delete_device_by_name(page, _TCP_DEVICE_NAME)

    # Step 4: 验证设备已删除
    _nav_to_physical_devices(page)
    remaining = page.locator("tbody tr").filter(has_text=_TCP_DEVICE_NAME)
    assert remaining.count() == 0, \
        f"TCP 设备 '{_TCP_DEVICE_NAME}' 删除后应从列表中消失，但仍然存在"
