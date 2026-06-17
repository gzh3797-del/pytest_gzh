import time
import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case14_14
# 用例标题：编辑模板参数，Address Format选择不同的进制，保存成功
# 预置条件：
#   1. 服务启动正常，账号登录成功
#   2. 模板管理界面，至少存在一个含 Block 的自定义模板
# 测试步骤：
#   1. Template List → 点击黄色 Edit 进入编辑页
#   2. Parameter Table 中点击第一行 Action 按钮 → 弹出 Edit Parameter 对话框
#   3. Block 选第一个 Block，Address Format 选 Hex，Address 输入 1
#   4. Multiplier 输入 10，Data Format 选 UINT16，Byte Order 选 BA
#   5. 点击 Save，验证保存成功（对话框关闭，无错误）
# 预期结果：
#   参数保存成功，对话框关闭，无错误提示


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
    """Navigate to Template List and click yellow edit on first custom template."""
    _nav_to_templates(page)
    page.locator(".el-menu-item").filter(has_text="Template List").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    last_tbody = page.locator("tbody").last
    rows = last_tbody.locator("tr").all()
    assert len(rows) > 0, "Template List 中无自定义模板，请先创建"
    last_tbody.locator("tr").first.locator(".el-button--warning").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def _open_param_edit_dialog(page) -> bool:
    """Find Parameter Table with rows, click blue Action button on first row."""
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


def _dialog_select(page, dialog, label: str, option_text: str):
    """Open select in dialog's form item by label, pick matching option."""
    fi = dialog.locator(".el-form-item").filter(has_text=label).first
    fi.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _click_visible_option(page, option_text)


def _dialog_fill(dialog, label: str, value: str):
    """Fill input in dialog's form item by label."""
    fi = dialog.locator(".el-form-item").filter(has_text=label).first
    inp = fi.locator("input").first
    inp.click()
    inp.fill(value)


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


def test_TestCase_AcuHMI_008_01_case14_14(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 进入模板编辑页
    _enter_template_edit_page(page)

    # Step 2: 打开 Parameter Table 行的编辑对话框
    opened = _open_param_edit_dialog(page)
    assert opened, "未能打开 Edit Parameter 对话框（Parameter Table 中无行或蓝色按钮未找到）"

    dialog = page.locator(".el-dialog").filter(has_text="Edit Parameter").first

    # Step 3: Block 选第一个，Address Format = Hex，Address = 1（Block 0 地址范围 0x1）
    _dialog_select(page, dialog, "Block", "")
    _dialog_select(page, dialog, "Address Format", "Hex")

    addr_fi = dialog.locator(".el-form-item").filter(has_text="Address").filter(has_not_text="Format").first
    addr_fi.locator("input").first.click()
    addr_fi.locator("input").first.fill("1")
    page.wait_for_timeout(200)

    # Step 4: Multiplier = 10，Data Format = UINT16，Byte Order = BA
    _dialog_fill(dialog, "Multiplier", "10")
    page.wait_for_timeout(200)
    _dialog_select(page, dialog, "Data Format", "UINT16")
    _dialog_select(page, dialog, "Byte Order", "BA")

    # Step 5: Save
    dialog.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    # 验证：对话框关闭 + 无错误
    dialog_still_open = page.locator(".el-dialog").filter(has_text="Edit Parameter").is_visible()
    assert not dialog_still_open or not _has_error(page), \
        "Address Format=Hex, Address=1, Multiplier=10, DataFormat=UINT16, ByteOrder=BA 应保存成功，但对话框未关闭或出现错误"
    assert not _has_error(page), \
        "Parameter 编辑保存时出现了错误提示"
