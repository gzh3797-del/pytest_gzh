import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case15_15
# 用例标题：编辑模板参数，Data Format选择不同参数，保存成功
# 预置条件：
#   1. 服务启动正常，账号登录成功
#   2. 至少存在一个含 Block 的自定义模板
# 测试步骤：
#   对每种 Data Format（UINT16 / UINT32 / FLOAT32）：
#   1. 打开 Edit Parameter 对话框
#   2. 选 Block、Address Format=Hex、Address=1、Multiplier=1
#   3. Data Format 选择对应格式，Byte Order=BA
#   4. Save → 应成功，对话框关闭
# 预期结果：
#   三种 Data Format 均保存成功，无错误提示


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


def test_TestCase_AcuHMI_008_01_case15_15(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _enter_template_edit_page(page)

    # 实际可用的 Data Format 选项（UI 实际值）
    data_formats = ["UINT16", "UINT32", "FLOAT32"]

    for fmt in data_formats:
        opened = _open_param_edit_dialog(page)
        assert opened, f"未能打开 Edit Parameter 对话框（Data Format={fmt}）"

        dialog = page.locator(".el-dialog").filter(has_text="Edit Parameter").first

        # Block
        dialog.locator(".el-form-item").filter(has_text="Block").first.locator(".el-select").first.click()
        page.wait_for_timeout(300)
        _click_visible_option(page, "")

        # Address Format = Hex
        dialog.locator(".el-form-item").filter(has_text="Address Format").first.locator(".el-select").first.click()
        page.wait_for_timeout(300)
        _click_visible_option(page, "Hex")

        # Address = 1
        addr_fi = dialog.locator(".el-form-item").filter(has_text="Address").filter(has_not_text="Format").first
        addr_fi.locator("input").first.click()
        addr_fi.locator("input").first.fill("1")
        page.wait_for_timeout(100)

        # Multiplier = 1
        mul_fi = dialog.locator(".el-form-item").filter(has_text="Multiplier").first
        mul_fi.locator("input").first.click()
        mul_fi.locator("input").first.fill("1")
        page.wait_for_timeout(100)

        # Data Format = fmt
        dialog.locator(".el-form-item").filter(has_text="Data Format").first.locator(".el-select").first.click()
        page.wait_for_timeout(300)
        _click_visible_option(page, fmt)

        # Byte Order = BA
        dialog.locator(".el-form-item").filter(has_text="Byte Order").first.locator(".el-select").first.click()
        page.wait_for_timeout(300)
        _click_visible_option(page, "BA")

        # Save
        dialog.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1000)

        assert not _has_error(page), \
            f"Data Format={fmt} 应保存成功，但出现错误提示"
        assert not page.locator(".el-dialog").filter(has_text="Edit Parameter").is_visible(), \
            f"Data Format={fmt} 保存后对话框应关闭"
