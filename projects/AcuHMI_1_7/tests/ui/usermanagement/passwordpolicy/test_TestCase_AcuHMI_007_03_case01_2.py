from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_ROLE_NAME = "pp01_2role"
_USER_NAME = "pp01_2user"
# 密码含数字+字母（满足 Numbers+Letters 策略）
_GOOD_PWD  = "12345abc"

_PERM_LABELS = [
    "User", "Device", "Data Log", "System Settings",
    "Protocol", "Alarm Log", "Maintenance", "Diagnostics", "Firmware Update",
]

_IDX_UPPER_LOWER    = 0
_IDX_NUMBERS_LETTERS = 1
_IDX_SPECIAL_CHARS   = 2


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


def _policy_checkbox(page, idx: int):
    # 用 locator 取第 idx 个策略复选框（保留 Playwright auto-wait，避免在
    # 组件尚未挂载时读取导致 querySelectorAll[idx] 为 undefined 而报错）。
    return page.locator(".el-checkbox").nth(idx)


def _is_cb_checked(page, idx: int) -> bool:
    cls = _policy_checkbox(page, idx).get_attribute("class") or ""
    return "is-checked" in cls


def _set_cb(page, idx: int, desired: bool) -> bool:
    """设置复选框到目标状态；返回是否实际发生了改动。"""
    cb = _policy_checkbox(page, idx)
    expect(cb).to_be_visible()
    if ("is-checked" in (cb.get_attribute("class") or "")) != desired:
        cb.locator(".el-checkbox__inner").click()
        page.wait_for_timeout(200)
        return True
    return False


def _set_policy(page, upper_lower: bool, numbers_letters: bool, special_chars: bool):
    """设置三项密码策略并保存。

    返回 (changed, toast_msg)：
      changed   —— 是否实际改动了任一复选框（用于区分真实保存与 "No change to save"）；
      toast_msg —— 点击 Save 后捕获的提示文本（提示已消失则为空串）。
    三个复选框都必须逐一调用 _set_cb（不可短路），故先各自求值再聚合 changed。
    """
    _nav_to_submenu(page, "Password Policy")
    # 等待策略复选框渲染完成后再读写，规避 SPA 路由切换后的渲染竞态
    expect(_policy_checkbox(page, _IDX_SPECIAL_CHARS)).to_be_visible(timeout=10000)
    c1 = _set_cb(page, _IDX_UPPER_LOWER,      upper_lower)
    c2 = _set_cb(page, _IDX_NUMBERS_LETTERS,  numbers_letters)
    c3 = _set_cb(page, _IDX_SPECIAL_CHARS,    special_chars)
    changed = c1 or c2 or c3
    page.get_by_role("button", name="Save").click()
    msg = ""
    try:
        content = page.locator(".el-message__content")
        content.last.wait_for(state="visible", timeout=3000)
        msg = (content.last.text_content() or "").strip()
    except Exception:
        pass
    page.wait_for_timeout(1000)
    return changed, msg


def _restore_default_policy(page):
    _set_policy(page, True, True, True)


def _create_role(page):
    _nav_to_submenu(page, "Role Configuration")
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(_ROLE_NAME)
    for lbl in _PERM_LABELS:
        page.locator(".el-form-item").filter(has_text=lbl).locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="view", exact=True).click()
        page.wait_for_timeout(200)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_role(page):
    _nav_to_submenu(page, "Role Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=_ROLE_NAME)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.wait_for_timeout(500)
    try:
        page.get_by_role("button", name="Yes, continue").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        pass


def _delete_user_if_exists(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_007_03_case01_2
# 用例标题：配置密码策略，仅勾选数字和字母，创建用户密码包含数字，可登录成功
# 测试步骤：
#   1. Password Policy 仅勾选 Numbers and Letters，保存
#   2. 创建用户，密码为 "12345abc"（含数字和字母）
#   3. 新用户登录系统
# 预期结果：
#   1. 提示 "Password policy configuration saved"
#   2. 添加成功
#   3. 登录成功
def test_TestCase_AcuHMI_007_03_case01_2(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    # 预清理：删除上次可能残留的同名用户/角色（崩溃中断会留下孤儿，导致重名创建失败）。
    # 顺序：先删用户、后删角色——角色被用户引用时无法直接删除。
    _delete_user_if_exists(page, _USER_NAME)
    _delete_role(page)

    _create_role(page)
    try:
        # Step 1: 仅启用 Numbers+Letters 策略，校验保存提示（区分真实保存 vs 无变更）
        changed, msg = _set_policy(
            page, upper_lower=False, numbers_letters=True, special_chars=False
        )
        low = msg.lower()
        if changed:
            assert "saved" in low, (
                "启用策略并产生变更后应提示保存成功"
                f"（Password policy configuration saved），实际: '{msg}'"
            )
        else:
            assert "no change" in low, (
                "策略已是目标状态、无变更，应提示 'No change to save'，"
                f"实际: '{msg}'"
            )

        # Step 2: 创建用户，密码含数字+字母 → 应成功
        _nav_to_submenu(page, "User Configuration")
        page.get_by_role("button", name="Add User").click()
        page.wait_for_timeout(1000)
        page.get_by_label("Username", exact=True).fill(_USER_NAME)
        page.get_by_label("Password", exact=True).fill(_GOOD_PWD)
        page.get_by_label("Repeat Password", exact=True).fill(_GOOD_PWD)
        page.get_by_text("--Select Role--", exact=True).click()
        page.get_by_role("option", name=_ROLE_NAME).click()
        page.wait_for_timeout(300)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1500)

        # 直接断言：用户列表出现该用户（强于"对话框关闭"——后者可能因其它原因消失）
        user_row = page.locator("tbody").get_by_role("row").filter(has_text=_USER_NAME)
        expect(user_row.first).to_be_visible(timeout=5000)

        # Step 3: 新用户登录
        ctx = browser.new_context(ignore_https_errors=True)
        p = ctx.new_page()
        try:
            p.goto(BASE_URL + "/#/login")
            p.wait_for_load_state("networkidle")
            p.get_by_role("textbox", name="Enter User Name").fill(_USER_NAME)
            p.get_by_role("textbox", name="Enter Password").fill(_GOOD_PWD)
            p.get_by_role("button", name="Sign In").click()
            p.wait_for_load_state("networkidle")
            p.wait_for_timeout(1000)
            try:
                p.get_by_role("button", name="Accept").click(timeout=3000)
                p.wait_for_load_state("networkidle")
                p.wait_for_timeout(500)
            except Exception:
                pass
            try:
                p.get_by_role("button", name="Cancel").click(timeout=2000)
            except Exception:
                pass
            assert "/#/login" not in p.url, \
                f"用户应能登录，当前 URL: {p.url}"
        finally:
            ctx.close()
    finally:
        _delete_user_if_exists(page, _USER_NAME)
        _delete_role(page)
        _restore_default_policy(page)
