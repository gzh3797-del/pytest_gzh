"""
只读页面结构探索脚本 — 不修改任何服务器配置，仅截图记录
"""
from pathlib import Path
from playwright.sync_api import sync_playwright

BASE_URL  = "https://192.168.2.199"
USERNAME  = "admin"
PASSWORD  = "Admin@110001"
SHOT_DIR  = Path(__file__).parent.parent / "product_structure_testcase_regulation" / "explore"
SHOT_DIR.mkdir(parents=True, exist_ok=True)


def shot(page, name: str):
    p = SHOT_DIR / f"{name}.png"
    page.screenshot(path=str(p), full_page=True)
    print(f"  [截图] {p.name}")


def dismiss_dialog(page):
    """关闭登录后弹出的修改默认密码提示（只点 Cancel）"""
    try:
        page.get_by_role("button", name="Cancel").click(timeout=2000)
        page.wait_for_timeout(300)
    except Exception:
        pass


def nav_to_user_management(page):
    """进入 Device -> AcuHMI -> User Management 侧边栏"""
    page.locator("header span").filter(has_text="AcuHMI").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    um = page.get_by_text("User Management").first
    um.click()
    page.wait_for_timeout(500)


def explore_submenu(page, menu_name: str, slug: str):
    """点击侧边栏子菜单并截图全部 Tab"""
    print(f"\n  >> {menu_name}")
    try:
        item = page.get_by_role("menuitem", name=menu_name)
        if item.count() == 0:
            item = page.get_by_text(menu_name, exact=True)
        item.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(600)
        shot(page, f"{slug}_main")

        # 如果有 Tab，逐一截图（只读，不操作表单）
        tabs = page.locator(".el-tabs__item").all()
        for i, tab in enumerate(tabs):
            label = (tab.inner_text() or f"tab{i}").strip()
            try:
                tab.click()
                page.wait_for_timeout(400)
                shot(page, f"{slug}_tab{i}_{label}")
            except Exception:
                pass
    except Exception as e:
        print(f"    跳过: {e}")


def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=200)
        ctx  = browser.new_context(ignore_https_errors=True,
                                   viewport={"width": 1440, "height": 900})
        page = ctx.new_page()

        # ── 1. 登录 ──────────────────────────────────────────────
        print("=== 登录 ===")
        page.goto(BASE_URL + "/#/login")
        page.wait_for_load_state("networkidle")
        page.get_by_role("textbox", name="Enter User Name").fill(USERNAME)
        page.get_by_role("textbox", name="Enter User Name").press("Tab")
        page.get_by_role("textbox", name="Enter Password").fill(PASSWORD)
        page.get_by_role("button", name="Sign In").click()
        page.wait_for_load_state("networkidle")
        dismiss_dialog(page)
        shot(page, "00_home")

        # ── 2. 进入 User Management ───────────────────────────────
        print("=== 进入 Device -> AcuHMI -> User Management ===")
        nav_to_user_management(page)
        shot(page, "01_user_management_landing")

        # ── 3. 遍历 User Management 子菜单 ───────────────────────
        submenus = [
            ("User Configuration",  "uc"),
            ("Role Configuration",  "rc"),
            ("Password Policy",     "pp"),
            ("Password Management", "pm"),
        ]
        for name, slug in submenus:
            # 确保侧边栏 User Management 保持展开
            try:
                page.get_by_text("User Management").first.click()
                page.wait_for_timeout(300)
            except Exception:
                pass
            explore_submenu(page, name, slug)

        ctx.close()
        browser.close()

    print(f"\n[完成] 截图已保存至: {SHOT_DIR}")


if __name__ == "__main__":
    run()
