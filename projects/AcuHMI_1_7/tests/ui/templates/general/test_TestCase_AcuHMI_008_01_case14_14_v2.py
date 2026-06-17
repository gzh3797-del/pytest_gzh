import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case14_14_v2
# 用例标题：编辑模板参数，Multiplier输入框验证（0.001成功，0失败，1001失败）
# 预置条件：
#   1. 服务启动正常，账号登录成功
#   2. 至少存在一个含 Block 的自定义模板
# 测试步骤：
#   1. 进入 Edit Parameter 对话框
#   2. Multiplier 输入 0.001 → Save → 应成功
#   3. 重新打开对话框，Multiplier 输入 0 → Save → 应失败
#   4. 重新打开对话框，Multiplier 输入 1001 → Save → 应失败
# 预期结果：
#   0.001 保存成功；0 和 1001 出现验证错误


def _nav_to_templates(page):
    if "/#/templates" not in page.url:
        if not any(s in page.url for s in [
            "/#/systemSettings", "/#/templates", "/#/protocols",
            "/#/maintenance", "/#/diagnostics", "/#/userManagement", "/#/firmwareUpdate",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Templates").click()
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


def _enter_template_edit_page(page):
    _nav_to_templates(page)
    page.locator(".el-menu-item").filter(has_text="Template List").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)
    last_tbody = page.locator("tbody").last
    rows = last_tbody.locator("tr").all()
    assert len(rows) > 0, "Template List 中无自定义模板"
    last_tbody.locator("tr").first.locator(".el-button--warning").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def _open_param_edit_dialog(page) -> bool:
    tbodies = page.locator("tbody").all()
    target_tbody = None
    for i, tb in enumerate(tbodies):
        if i > 0 and tb.locator("tr").count() > 0:
            target_tbody = tb
            break
    if target_tbody is None:
        return False
    target_tbody.locator("tr").first.locator(".el-button--primary").first.click()
    page.wait_for_timeout(500)
    return page.locator(".el-dialog").filter(has_text="Edit Parameter").count() > 0


def _setup_dialog_base(page, dialog):
    """Select Block and Address Format so Multiplier field is accessible."""
    fi_block = dialog.locator(".el-form-item").filter(has_text="Block").first
    fi_block.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _click_visible_option(page, "")

    fi_addr = dialog.locator(".el-form-item").filter(has_text="Address Format").first
    fi_addr.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _click_visible_option(page, "Hex")

    addr_fi = dialog.locator(".el-form-item").filter(has_text="Address").filter(has_not_text="Format").first
    addr_fi.locator("input").first.click()
    addr_fi.locator("input").first.fill("1")
    page.wait_for_timeout(100)

    fi_fmt = dialog.locator(".el-form-item").filter(has_text="Data Format").first
    fi_fmt.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _click_visible_option(page, "UINT16")

    fi_bo = dialog.locator(".el-form-item").filter(has_text="Byte Order").first
    fi_bo.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _click_visible_option(page, "BA")


def _fill_multiplier(dialog, value: str):
    fi = dialog.locator(".el-form-item").filter(has_text="Multiplier").first
    inp = fi.locator("input").first
    inp.click()
    inp.fill(value)
    inp.press("Tab")


def _has_error(page) -> bool:
    if page.locator(".el-message--error").count() > 0:
        return True
    for el in page.locator(".el-form-item__error").all():
        try:
            if el.is_visible() and el.inner_text().strip():
                return True
        except Exception:
            pass
    return False


def test_TestCase_AcuHMI_008_01_case14_14_v2(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _enter_template_edit_page(page)

    # ── Test 1: Multiplier = 0.001 → 应成功 ─────────────────────────────────
    opened = _open_param_edit_dialog(page)
    assert opened, "未能打开 Edit Parameter 对话框"

    dialog = page.locator(".el-dialog").filter(has_text="Edit Parameter").first
    _setup_dialog_base(page, dialog)
    _fill_multiplier(dialog, "0.001")
    dialog.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    assert not _has_error(page), "Multiplier=0.001（最小有效值）应保存成功，但出现错误提示"
    # 对话框应已关闭
    assert not page.locator(".el-dialog").filter(has_text="Edit Parameter").is_visible(), \
        "Multiplier=0.001 保存后对话框应关闭"

    # ── Test 2: Multiplier = 0 → 应失败 ──────────────────────────────────────
    opened = _open_param_edit_dialog(page)
    assert opened, "重新打开 Edit Parameter 对话框失败"

    dialog = page.locator(".el-dialog").filter(has_text="Edit Parameter").first
    _setup_dialog_base(page, dialog)
    _fill_multiplier(dialog, "0")
    dialog.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    assert _has_error(page), "Multiplier=0（低于最小值）应保存失败，但未检测到错误提示"

    # 关闭对话框（Cancel 或 Escape）
    cancel_btn = dialog.locator("button").filter(has_text="Cancel")
    if cancel_btn.count() > 0:
        cancel_btn.first.click()
    else:
        page.keyboard.press("Escape")
    page.wait_for_timeout(300)

    # ── Test 3: Multiplier = 1001 → 应失败 ───────────────────────────────────
    opened = _open_param_edit_dialog(page)
    assert opened, "重新打开 Edit Parameter 对话框失败（第3次）"

    dialog = page.locator(".el-dialog").filter(has_text="Edit Parameter").first
    _setup_dialog_base(page, dialog)
    _fill_multiplier(dialog, "1001")
    dialog.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    assert _has_error(page), "Multiplier=1001（超出上限1000）应保存失败，但未检测到错误提示"
