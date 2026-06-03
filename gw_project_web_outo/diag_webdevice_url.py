"""Diagnostic: confirm exact boundary at 291-294 chars."""
import pytest
from pages.login_page import LoginPage

# 精确构造指定长度域名：用"ab."重复 + 末尾标签
_URLS = {
    291: "ab." * 96 + "abc",       # 288+3=291
    292: "ab." * 96 + "abcd",      # 288+4=292
    293: "ab." * 96 + "abcde",     # 288+5=293
    294: "ab." * 96 + "abcdef",    # 288+6=294
}


def _nav_to_web_devices(page):
    if "/#/webDevice" not in page.url:
        try:
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        except Exception:
            pass
        page.locator(".left-nav-item").filter(has_text="Web Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_diag_exact_boundary(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_to_web_devices(page)

    results = []
    for idx, (length, url_val) in enumerate(_URLS.items()):
        assert len(url_val) == length, f"构造错误: expected {length}, got {len(url_val)}"

        page.get_by_role("button", name="Add Device").click()
        page.wait_for_timeout(500)
        dialog = page.locator(".el-dialog")
        dialog.wait_for(timeout=5000)

        def fill_by_label(text, value):
            inp = dialog.locator(".el-form-item").filter(has_text=text).locator("input").first
            inp.fill(value)

        fill_by_label("Device Name", f"BndDev{idx}")
        fill_by_label("Serial Number", f"SNBND{idx:04d}")
        fill_by_label("Model", "M1")
        dialog.locator("input[placeholder='---Enter URL---']").fill(url_val)
        page.wait_for_timeout(300)

        dialog.get_by_role("button", name="Confirm").click()
        page.wait_for_timeout(1500)

        still_open = dialog.is_visible() if dialog.count() > 0 else False
        accepted = not still_open
        results.append((length, accepted))

        if still_open:
            try:
                dialog.get_by_role("button", name="Cancel").click(timeout=2000)
                page.wait_for_timeout(500)
            except Exception:
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(500)
                except Exception:
                    pass
        else:
            page.wait_for_timeout(300)
            rows = page.locator("tbody tr")
            if rows.count() > 0:
                try:
                    rows.first.get_by_role("button").last.click()
                    page.wait_for_timeout(500)
                    page.get_by_role("button", name="Yes").click(timeout=2000)
                    page.wait_for_timeout(800)
                except Exception:
                    pass

    print("\n\n=== Exact Boundary Results ===")
    for length, ok in results:
        print(f"  {length} chars: {'ACCEPTED' if ok else 'REJECTED'}")

    accepted_lengths = [l for l, ok in results if ok]
    print(f"\nMax accepted: {max(accepted_lengths) if accepted_lengths else 'none'}")

    assert True
