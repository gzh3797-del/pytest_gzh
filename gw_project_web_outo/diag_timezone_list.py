"""Diagnostic: enumerate all Time Zone options in Date & Time page."""
from pages.login_page import LoginPage


def test_diag_timezone_list(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    tz_fi = page.locator(".el-form-item").filter(has_text="Time Zone").first
    assert tz_fi.count() > 0, "未找到 Time Zone 字段"

    tz_select = tz_fi.locator(".el-select").first
    tz_select.click()
    page.wait_for_timeout(600)

    # 打印所有可见选项
    opts = page.locator(".el-select-dropdown__item").all()
    print(f"\n=== Time Zone 选项总数（初始可见）: {len(opts)} ===")
    for i, opt in enumerate(opts):
        try:
            txt = opt.inner_text().strip()
            if txt:
                print(f"  [{i:03d}] {txt}")
        except Exception:
            pass

    # 搜索 Shanghai
    search_inp = page.locator(".el-select-dropdown input").first
    if search_inp.count() > 0 and search_inp.is_visible():
        search_inp.fill("Shanghai")
        page.wait_for_timeout(400)
        opts2 = page.locator(".el-select-dropdown__item").all()
        print(f"\n=== 搜索 'Shanghai' 结果: {len(opts2)} 条 ===")
        for opt in opts2:
            try:
                txt = opt.inner_text().strip()
                if txt:
                    print(f"  {txt}")
            except Exception:
                pass

        # 再搜索 Asia
        search_inp.fill("Asia/S")
        page.wait_for_timeout(400)
        opts3 = page.locator(".el-select-dropdown__item").all()
        print(f"\n=== 搜索 'Asia/S' 结果: {len(opts3)} 条 ===")
        for opt in opts3:
            try:
                txt = opt.inner_text().strip()
                if txt:
                    print(f"  {txt}")
            except Exception:
                pass

    page.keyboard.press("Escape")
    assert True
