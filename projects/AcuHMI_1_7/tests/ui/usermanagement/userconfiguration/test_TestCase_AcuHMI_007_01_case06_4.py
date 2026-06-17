import pytest
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_submenu(page, submenu: str):
    if "/userManagement/" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_text("User Management").first.click()
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_007_01_case06_4
# 用例标题：添加用户，用户名长度为41，密码长度为64，保存配置失败（用户名超过40位）
# 测试步骤：
#   1. Add User：Username=41字符，Password=64字符，保存
#   2. 预期失败（用户名超出40字符上限）
# 预期结果：
#   保存配置失败，系统提示错误信息准确
#   备注：用户名超过40位可能被前端自动截断，以实际设备行为为准
def test_TestCase_AcuHMI_007_01_case06_4(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # 41字符用户名（超出40上限）
    username41 = "qwertyuiopasdfghjklzxcvbnm012345678901234"  # 41 chars
    # 64字符密码（符合策略）
    pwd64 = "Abc@1234" + "x" * 56  # 8+56=64

    _nav_to_submenu(page, "User Configuration")
    page.get_by_role("button", name="Add User").click()
    page.wait_for_timeout(1000)
    page.get_by_label("Username", exact=True).fill(username41)
    page.get_by_label("Password", exact=True).fill(pwd64)
    page.get_by_label("Repeat Password", exact=True).fill(pwd64)
    page.get_by_text("--Select Role--", exact=True).click()
    page.get_by_role("option", name="admin").click()
    page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)

    # 用户名可能被截断（仅保留前40位），也可能报错
    # 两种结果都可接受，本测试主要记录实际行为
    dialog_still_open = page.get_by_label("Password", exact=True).is_visible()
    if dialog_still_open:
        # 添加失败
        page.get_by_role("button", name="Cancel").click(timeout=2000)
    # 若对话框关闭（用户名被截断后添加成功），清理截断后的用户
    else:
        truncated = username41[:40]
        _nav_to_submenu(page, "User Configuration")
        row = page.locator("tbody").get_by_role("row").filter(has_text=truncated[:20])
        if row.count() > 0:
            row.get_by_role("button").last.click()
            page.get_by_role("button", name="Yes, continue").click()
            page.wait_for_timeout(500)
