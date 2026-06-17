import time
import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case13_13
# 用例标题：编辑模板进入，该编辑页面有Edit按钮，Delete按钮，用户确认UI元素存在
# 预置条件：
#   1. 服务启动正常，账号登录成功
#   2. Template List 中存在至少一个自定义模板
# 测试步骤：
#   1. 进入 Template List
#   2. 在 Customized 表格中找到第一行，确认存在黄色（Edit）和红色（Delete）按钮
#   3. 点击黄色 Edit 按钮进入编辑页
#   4. 验证编辑页存在保存类按钮（Update Template）和 Block 区域
# 预期结果：
#   Template List 行操作按钮（编辑/删除）存在；进入编辑页后功能正常显示


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


def test_TestCase_AcuHMI_008_01_case13_13(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 进入 Template List
    _nav_to_templates(page)
    page.locator(".el-menu-item").filter(has_text="Template List").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Step 2: 在 Customized 表格（最后一个 tbody）中找第一行，确认 Edit/Delete 按钮存在
    last_tbody = page.locator("tbody").last
    rows = last_tbody.locator("tr").all()
    assert len(rows) > 0, "Template List Customized 表格中无模板行，请先创建自定义模板"

    first_row = rows[0]
    edit_btn = first_row.locator(".el-button--warning")
    delete_btn = first_row.locator(".el-button--danger")

    assert edit_btn.count() > 0, "模板行应有黄色 Edit 按钮（.el-button--warning）"
    assert delete_btn.count() > 0, "模板行应有红色 Delete 按钮（.el-button--danger）"

    # Step 3: 点击黄色 Edit 按钮进入编辑页
    edit_btn.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Step 4: 验证编辑页存在保存类按钮（Update Template）和 Block 区域的 Save Block
    has_update = page.get_by_role("button", name="Update Template").count() > 0
    has_save_block = page.get_by_role("button", name="Save Block").count() > 0

    assert has_update or has_save_block, \
        "编辑页应有 'Update Template' 或 'Save Block' 按钮"

    # 验证 Template Name 字段存在（只读状态）
    name_fi = page.locator(".el-form-item").filter(has_text="Template Name")
    assert name_fi.count() > 0, "编辑页应有 Template Name 字段"
