# 用例编号: TestCase_AcuHMI_005_07_case03_5
# 用例标题: 白名单描述验证，描述≤40字符保存成功，=41字符被拒绝
# 预置条件: 1、管理权限登录AcuHMI网页
# 测试步骤:
#   1. 添加白名单，描述为40字符混合字符串 → 保存成功，读回描述与输入完全一致
#   2. 添加白名单，描述为41字符混合字符串 → 保存被拒绝（表单显示校验错误）
# 预期结果:
#   40字符描述：保存成功且读回内容与输入完全相等（无静默截断）
#   41字符描述：保存失败，el-form-item__error 出现（字段长度上限为40）

from projects.AcuHMI_1_7.pages.login_page import LoginPage

_TEST_IP = "192.168.1.200"


def _nav_to_settings_tab(page, tab: str) -> None:
    if "/systemSettings" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=tab).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _read_desc_via_edit_dialog(page) -> str:
    """Open the Edit Allow List dialog for _TEST_IP and return the description input value.

    Uses the first action button in the matching row (Edit), reads the 'Enter Description'
    input value, then closes the dialog via Cancel.  This reads the persisted server value
    rather than the potentially CSS-truncated table cell text.
    """
    row = page.locator("tbody").get_by_role("row").filter(has_text=_TEST_IP)
    row.locator(".el-button").first.click(force=True)
    page.wait_for_timeout(800)
    value: str = page.get_by_placeholder("Enter Description").input_value()
    page.get_by_role("button", name="Cancel").click()
    page.wait_for_timeout(300)
    return value


def _delete_entry(page) -> None:
    """Delete the _TEST_IP row via the last button (delete) and the popconfirm."""
    try:
        row = page.locator("tbody").get_by_role("row").filter(has_text=_TEST_IP)
        if row.count() > 0:
            row.locator(".el-button").last.click(force=True)
            page.wait_for_timeout(500)
            page.locator(".el-popconfirm__action .el-button--primary").click(timeout=2000)
            page.wait_for_timeout(500)
    except Exception:
        pass


def test_TestCase_AcuHMI_005_07_case03_5(login_page: LoginPage) -> None:
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Access Control")

    # Enable IP Allow List so the Add Allow List button appears
    page.locator(".el-form-item").filter(has_text="IP Allow List Enable").locator(
        ".el-radio"
    ).filter(has_text="Enable").click()
    page.wait_for_timeout(500)

    # 40-char mixed description (boundary value at the maximum allowed length)
    desc_40 = "qweRTYUIOP0123456789_ 23466789!@#RTYUIOP"
    assert len(desc_40) == 40, f"desc_40 should be 40 chars, got {len(desc_40)}"

    # 41-char description (one beyond the maximum, should be rejected)
    desc_41 = "qweRTYUIOP0123456789_ 23466789!@#RTYUIOP1"
    assert len(desc_41) == 41, f"desc_41 should be 41 chars, got {len(desc_41)}"

    try:
        # ── Case 1: 40-char description – expect save success and no truncation ──
        page.get_by_role("button", name="Add Allow List").click()
        page.wait_for_timeout(500)
        # Switch to No (single IP mode) — placeholder becomes "Enter IP Address"
        page.locator(".el-dialog").locator(".el-radio").filter(has_text="No").click()
        page.wait_for_timeout(300)
        page.get_by_placeholder("Enter IP Address").fill(_TEST_IP)
        page.get_by_placeholder("Enter Description").fill(desc_40)
        page.get_by_role("button", name="Confirm").click()
        page.wait_for_timeout(500)

        # Assert no form validation error (save accepted)
        assert page.locator(".el-form-item__error").count() == 0, (
            "40字符描述的白名单应保存成功（无表单校验错误）"
        )

        # Read back the persisted description via Edit dialog and assert exact match
        saved_desc = _read_desc_via_edit_dialog(page)
        assert saved_desc == desc_40, (
            f"40字符描述保存后读回不一致（静默截断或其他损坏）\n"
            f"  输入长度={len(desc_40)}, 读回长度={len(saved_desc)}\n"
            f"  输入值:  {repr(desc_40)}\n"
            f"  读回值:  {repr(saved_desc)}"
        )

        # Clean up the entry created for Case 1 before Case 2
        _delete_entry(page)

        # ── Case 2: 41-char description – expect form validation error ──
        page.get_by_role("button", name="Add Allow List").click()
        page.wait_for_timeout(500)
        page.locator(".el-dialog").locator(".el-radio").filter(has_text="No").click()
        page.wait_for_timeout(300)
        page.get_by_placeholder("Enter IP Address").fill(_TEST_IP)
        page.get_by_placeholder("Enter Description").fill(desc_41)
        page.get_by_role("button", name="Confirm").click()
        page.wait_for_timeout(500)

        # Assert a form validation error IS shown (41 chars exceeds the 40-char limit)
        # EP 表单校验错误异步渲染，先 auto-wait 再断言，避免即时 count 取到 0 误判
        page.locator(".el-form-item__error").first.wait_for(state="visible", timeout=5000)
        assert page.locator(".el-form-item__error").count() > 0, (
            "41字符描述应触发表单校验错误（字段长度上限为40字符），但未见任何错误提示"
        )

        # Cancel the rejected dialog so cleanup can proceed cleanly
        try:
            page.get_by_role("button", name="Cancel").click(timeout=2000)
            page.wait_for_timeout(300)
        except Exception:
            pass

    finally:
        # Clean up any leftover _TEST_IP entry (guard for unexpected flow)
        _delete_entry(page)
        # Disable IP Allow List to restore device state
        try:
            page.locator(".el-form-item").filter(has_text="IP Allow List Enable").locator(
                ".el-radio"
            ).filter(has_text="Disable").click()
            page.wait_for_timeout(300)
            page.get_by_role("button", name="Save").click()
            page.wait_for_timeout(500)
        except Exception:
            pass
