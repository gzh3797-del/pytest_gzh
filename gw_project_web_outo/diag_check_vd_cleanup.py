"""验证 VD_CapTest_ 测试设备是否已被清理"""
from pages.login_page import LoginPage


def _nav_to_vd_list(page):
    on_list = "/#/virtualMeter" in page.url and "addVirtualMeter" not in page.url
    if not on_list:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Virtual Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_diag_check_vd_cleanup(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_vd_list(page)
    rows = page.locator("tbody tr")
    total = rows.count()
    print(f"\n当前虚拟设备总数: {total}")

    leftover = []
    for i in range(total):
        try:
            name = rows.nth(i).locator("td").first.inner_text().strip()
            print(f"  [{i+1}] {name}")
            if "VD_CapTest_" in name:
                leftover.append(name)
        except Exception:
            pass

    print(f"\n遗留 VD_CapTest_ 设备: {len(leftover)} 台")
    if leftover:
        print("  !! 清理未完成，以下设备仍存在:")
        for n in leftover:
            print(f"     - {n}")
    else:
        print("  ✓ 清理完毕，无遗留测试设备")

    assert len(leftover) == 0, f"发现 {len(leftover)} 台遗留测试设备: {leftover}"
