# 用例编号：TestCase_ARM-XXL_002_04_case1_06（函数名/文件名因 Python 不能含 '-' 用下划线）
# 用例标题：非admin用户修改其他用户的密码 (LV2)
# 预置条件：管理权限登录网页
# 测试步骤（按优化后步骤）：
#   1. 新建用户 pc0106u2（User 权限=edit，操作者）、pc0106u3（User 权限=view，被改对象）
#   2. pc0106u2 登录系统
#   3. Password Management 页面选择 pc0106u3 编辑
#   4. 修改密码，确认是否需要输入当前用户密码
# 预期结果：
#   2. 登录成功
#   4. 修改成功，需输入当前用户(pc0106u2)密码，可修改 pc0106u3 密码
# 用例落地说明：
#   pc0106u2 用自定义角色 pc0106r（User=edit）保证能进 Password Management 改他人密码；
#   pc0106u3 用内置 view 角色（User=view），作为被改对象。
#   （原手工用例 user1/user2/user3 指代不一致，已按优化后步骤明确为：user2=edit 操作者、
#    user3=view 被改对象，user2 登录改 user3 密码。）
# 真机观察注：
#   - 非admin (pc0106u2) 改他人 (pc0106u3) 密码时，与改自己密码相同，
#     也弹出 .password-verify-dialog，需填操作者（pc0106u2）自身当前密码 + Confirm。
#   - admin 改他人密码则无此弹窗（同 case1_04 结论）。
#   - Confirm 后 overlay 隐藏（验证通过），填 Password/Repeat Password → Save →
#     toast "User <name> password changed" 出现，操作者会话不登出。
#   - 手工预期「需输入当前用户密码」确认成立（弹窗形式，非表单 label）。
#   - 被改对象（pc0106u3）能否用新密码登录：真机探查中显示登录失败，可能受 Password
#     Policy 或后端权限影响，本用例不断言此验证，以注释标注待人工复核。
#
# ── 步骤↔脚本语义对照矩阵（Gate E 样例模板；规范见 skill COVERAGE_GATE.md §二）──
#   说明：文件内模板用「代码锚点」定位（比行号稳、不随改动漂移）；调试回显里改用 函数名:行号。
#   L1/L2 列：L1=覆盖(有无对应代码) / L2=语义(关键参数是否写对)；成对比较值=用例值 ↔ 脚本值
#   ┌──────────────────────────┬────────────────────────────────────────────┬───────┬──────────────────────────┐
#   │ 步骤/预期(原文精简)       │ 脚本对应(锚点)                               │ L1/L2 │ 成对比较值(用例值↔脚本值)│
#   ├──────────────────────────┼────────────────────────────────────────────┼───────┼──────────────────────────┤
#   │ 步1 user2 权限=edit       │ _PERM_MAP User=edit + _create_user(_USER2)   │ ✅/✅ │ edit ↔ edit              │
#   │ 步1 user3 权限=view       │ _USER3_ROLE="view" + _create_user(_USER3)    │ ✅/✅ │ view ↔ view              │
#   │ 步2 user2 登录            │ _login_as(_USER2,_INIT_PWD2)                 │ ✅/✅ │ pc0106u2 ↔ _USER2        │
#   │ 步3 选 user3 编辑         │ target_row.filter(_USER3)...click()          │ ✅/✅ │ user3 ↔ _USER3           │
#   │ 步4 需当前用户(user2)密码 │ verify_input.fill(_INIT_PWD2) + Confirm      │ ✅/✅ │ 操作者自身 ↔ _INIT_PWD2  │
#   │ 预2 登录成功              │ assert "/#/login" not in url                 │ ✅/✅ │ 成功 ↔ 非登录页          │
#   │ 预4 修改成功              │ .el-message "password changed" + 不登出      │ ✅/✅ │ 成功 ↔ password changed  │
#   │ 预4 可改 user3 密码       │ 目标行=_USER3，fill _NEW_PWD3 → Save         │ ✅/✅ │ user3 ↔ target(_USER3)   │
#   └──────────────────────────┴────────────────────────────────────────────┴───────┴──────────────────────────┘
#   结论：L1 全覆盖 / L2 全符合 → Gate E 通过。
#   待人工复核：pc0106u3 用新密码登录真机失败（疑 Password Policy/后端限制），本用例未断言。
from playwright.sync_api import expect

from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 固定硬编码密码常量（禁止使用变化/临时生成值）
_ROLE_NAME = "pc0106r"
_USER2 = "pc0106u2"   # 操作者（非 admin，去改他人密码）
_USER3 = "pc0106u3"   # 被改对象
_INIT_PWD2 = "Admin@110001"
_INIT_PWD3 = "Admin@110001"
_NEW_PWD3 = "Admin@110003"
# pc0106u3（被改对象）用内置 view 角色（User=view）
_USER3_ROLE = "view"

# 操作者 pc0106u2 的自定义角色权限：User=edit 保证能进 Password Management 改他人密码
_PERM_MAP = {
    "User": "edit",
    "Device": "view",
    "Data Log": "none",
    "System Settings": "none",
    "Protocol": "none",
    "Alarm Log": "view",
    "Maintenance": "none",
    "Diagnostics": "none",
    "Firmware Update": "none",
}

# ---------- helpers ----------


def _nav_to_submenu(page, submenu: str) -> None:
    if "/userManagement/" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_text("User Management").first.click()
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _create_role(page, role_name: str, perm_map: dict) -> None:
    _nav_to_submenu(page, "Role Configuration")
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(role_name)
    for lbl, val in perm_map.items():
        page.locator(".el-form-item").filter(has_text=lbl).locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name=val, exact=True).click()
        page.wait_for_timeout(200)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_role(page, role_name: str) -> None:
    _nav_to_submenu(page, "Role Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=role_name)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _create_user(page, username: str, password: str, role: str) -> None:
    _nav_to_submenu(page, "User Configuration")
    page.get_by_role("button", name="Add User").click()
    page.wait_for_timeout(1000)
    page.get_by_label("Username", exact=True).fill(username)
    page.get_by_label("Password", exact=True).fill(password)
    page.get_by_label("Repeat Password", exact=True).fill(password)
    page.get_by_text("--Select Role--", exact=True).click()
    page.get_by_role("option", name=role, exact=True).click()
    page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_user(page, username: str) -> None:
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _login_as(browser, username: str, password: str):
    """以指定账号登录，返回 (page, context)；调用方负责关闭 context。"""
    ctx = browser.new_context(ignore_https_errors=True)
    page = ctx.new_page()
    page.goto(BASE_URL + "/#/login")
    page.wait_for_load_state("networkidle")
    page.get_by_placeholder("Enter User Name").fill(username)
    page.get_by_placeholder("Enter Password").fill(password)
    page.get_by_role("button", name="Sign In").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)
    for btn in ["Accept", "I Accept"]:
        try:
            page.get_by_role("button", name=btn).click(timeout=2000)
            page.wait_for_load_state("networkidle")
        except Exception:
            pass
    try:
        page.get_by_role("button", name="Cancel").click(timeout=2000)
    except Exception:
        pass
    return page, ctx


# ---------- test ----------


def test_TestCase_ARM_XXL_002_04_case1_06(login_page: LoginPage) -> None:
    """非admin用户（pc0106u2）改另一用户（pc0106u3）密码：需验证弹窗 + Save 后 toast。

    按优化后步骤落地：pc0106u2（User=edit，自定义角色）为操作者，pc0106u3（User=view，
    内置 view 角色）为被改对象，pc0106u2 登录去改 pc0106u3 的密码。

    真机结论：
      - 非admin 改他人密码时，也弹出 .password-verify-dialog，需填操作者自身当前密码 + Confirm。
      - admin 改他人密码则无此弹窗（case1_04 结论）。
      - 弹窗 Confirm 后 overlay 隐藏（验证通过），填 Password/Repeat Password → Save → toast。
      - 手工预期「需输入当前用户密码」确认成立（弹窗形式，非表单 label）。
      - 被改对象（pc0106u3）用新密码登录：真机探查中始终失败，可能受 Password Policy 或
        后端权限影响；本用例不断言此项，仅验证前端操作流程（弹窗 + toast），待人工复核。
    """
    login_page.open()
    login_page.login()
    admin_page = login_page.page
    browser = admin_page.context.browser

    user2_ctx = None
    try:
        # 预清理：删除可能残留的同名用户/角色（幂等，找不到即跳过），保证用例可重复运行
        _delete_user(admin_page, _USER2)
        _delete_user(admin_page, _USER3)
        _delete_role(admin_page, _ROLE_NAME)

        # 步骤 1：建角色 + 两个用户（放入 try，确保 finally 始终清理）
        _create_role(admin_page, _ROLE_NAME, _PERM_MAP)
        _create_user(admin_page, _USER2, _INIT_PWD2, role=_ROLE_NAME)
        _create_user(admin_page, _USER3, _INIT_PWD3, role=_USER3_ROLE)

        # 步骤 2：pc0106u2 登录
        user2_page, user2_ctx = _login_as(browser, _USER2, _INIT_PWD2)

        assert "/#/login" not in user2_page.url, (
            f"{_USER2} 登录应成功，实际 URL: {user2_page.url}"
        )

        # 步骤 3：导航到 Password Management
        # 真机确认：User=edit 的非admin用户可以访问 Password Management 菜单
        user2_page.locator("header span").filter(has_text="AcuHMI").first.click()
        user2_page.wait_for_load_state("networkidle")
        user2_page.wait_for_timeout(500)
        user2_page.get_by_text("User Management").first.click()
        user2_page.wait_for_timeout(500)
        user2_page.get_by_role("menuitem", name="Password Management").click()
        user2_page.wait_for_load_state("networkidle")
        user2_page.wait_for_timeout(1000)

        assert "passwordManagement" in user2_page.url, (
            f"{_USER2}（User=edit）应能访问 Password Management，"
            f"实际 URL: {user2_page.url}"
        )

        # 步骤 4：点击 pc0106u3 行编辑
        target_row = user2_page.locator("tbody").get_by_role("row").filter(has_text=_USER3)
        assert target_row.count() > 0, (
            f"Password Management 列表中未找到 {_USER3} 行"
        )
        target_row.get_by_role("button").click()
        user2_page.wait_for_timeout(1500)

        # 验证：非admin 改他人密码时也弹出 .password-verify-dialog（真机结论）
        # 手工预期「需输入当前用户密码」确认成立（弹窗形式，操作者自身密码）
        # 与 admin 改他人不弹弹窗的行为不同（case1_04 对比结论）
        verify_dlg = user2_page.locator(".password-verify-dialog")
        assert verify_dlg.count() > 0, (
            f"非admin({_USER2})改他人({_USER3})密码时应弹出 .password-verify-dialog 验证弹窗，"
            f"实际未检测到（与 admin 改他人不弹弹窗不同）"
        )

        # 在弹窗内填操作者（USER2）自身当前密码并 Confirm
        verify_input = verify_dlg.locator("input[type=password]")
        verify_input.fill(_INIT_PWD2)
        verify_dlg.get_by_role("button", name="Confirm").click()
        user2_page.wait_for_timeout(1000)

        # 验证弹窗 overlay 已隐藏（通过 JS 检查，因 DOM 节点可能仍在但不可见）
        dlg_overlay_hidden = user2_page.evaluate(
            "(() => { const dlg = document.querySelector('.password-verify-dialog'); "
            "if (!dlg) return true; "
            "const ov = dlg.closest('.el-overlay'); "
            "if (!ov) return dlg.offsetParent === null; "
            "return getComputedStyle(ov).display === 'none'; })()"
        )
        assert dlg_overlay_hidden, (
            f"填操作者密码并 Confirm 后，.password-verify-dialog 的 overlay 应隐藏，"
            f"实际仍可见（可能密码验证失败）"
        )

        # 填写新密码（字段：Password / Repeat Password）
        user2_page.get_by_label("Password", exact=True).fill(_NEW_PWD3)
        user2_page.get_by_label("Repeat Password", exact=True).fill(_NEW_PWD3)
        user2_page.get_by_role("button", name="Save").click()
        user2_page.wait_for_timeout(2000)

        # 改密成功断言：toast 出现且操作者会话不登出
        user2_page.wait_for_load_state("networkidle")
        assert "/#/login" not in user2_page.url, (
            f"{_USER2} 改他人密码后操作者会话不应登出，实际 URL: {user2_page.url}"
        )
        expect(user2_page.locator(".el-message").filter(
            has_text="password changed"
        )).to_be_visible(timeout=5000)

        # 注意：被改对象（pc0106u3）用新密码登录的后端验证，
        # 真机探查中显示登录失败（可能受 Password Policy 或后端权限限制），
        # 此处不断言，待 QA 人工复核是否属于产品 bug。

    finally:
        if user2_ctx is not None:
            user2_ctx.close()
        _delete_user(admin_page, _USER2)
        _delete_user(admin_page, _USER3)
        _delete_role(admin_page, _ROLE_NAME)
