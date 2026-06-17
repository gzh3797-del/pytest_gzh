import re
import time
import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case12_12
# 用例标题：Block Table打开模板，切换显示数量查看是否显示数量预期
# 预置条件：
#   1. 服务启动正常，账号登录成功
#   2. 创建一个包含 85 个 Block 的自定义模板
# 测试步骤：
#   1. 进入 New Typical Energy Meter Template，填写 Device 区，添加 85 个 Block
#   2. Create Template 保存
#   3. Template List 中点击黄色编辑按钮进入编辑页
#   4. 依次切换 Block Table 每页数量：10/20/40/80
#   5. 验证每次切换后实际显示行数与所选数量一致
# 预期结果：
#   Block Table 分页切换正确，每页行数不超过所选数量

TARGET_BLOCKS = 85


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
                page.wait_for_timeout(400)
                return True
        except Exception:
            pass
    for item in all_items:
        try:
            if item.is_visible():
                item.click()
                page.wait_for_timeout(400)
                return True
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    return False


def _add_85_blocks(page):
    """Add 85 blocks on the create/edit template page. Function & Format selected once."""
    # Select Function (READ_HOLDING_REGISTERS) — first block only
    page.keyboard.press("Escape")
    page.wait_for_timeout(150)
    page.locator(".el-form-item").filter(has_text="Function").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "READ_HOLDING_REGISTERS")

    # Select Address Format = Decimal — first block only
    page.keyboard.press("Escape")
    page.wait_for_timeout(150)
    page.locator(".el-form-item").filter(has_text="Address Format").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "Decimal")
    page.keyboard.press("Escape")
    page.wait_for_timeout(150)

    for i in range(1, TARGET_BLOCKS + 1):
        # If Function/Format reset after Save Block, re-select (detected by checking select text)
        fn_fi = page.locator(".el-form-item").filter(has_text="Function").first
        fn_val = fn_fi.locator(".el-select__wrapper, .el-input__inner").first.inner_text().strip()
        if not fn_val or "Select" in fn_val or fn_val == "":
            fn_fi.locator(".el-select").click()
            page.wait_for_timeout(300)
            _click_visible_option(page, "READ_HOLDING_REGISTERS")
            page.keyboard.press("Escape")
            page.wait_for_timeout(100)

            page.locator(".el-form-item").filter(has_text="Address Format").first.locator(".el-select").click()
            page.wait_for_timeout(300)
            _click_visible_option(page, "Decimal")
            page.keyboard.press("Escape")
            page.wait_for_timeout(100)

        # Fill Start (unique per block: 1, 2, …, 85)
        start_input = page.locator(".el-form-item").filter(has_text="Start").first.locator("input")
        start_input.fill(str(i))
        page.wait_for_timeout(100)

        # Fill Count
        count_input = page.locator(".el-form-item").filter(has_text="Count").first.locator("input")
        count_input.fill("1")
        page.wait_for_timeout(100)

        page.get_by_role("button", name="Save Block").click()
        page.wait_for_timeout(800)


def _switch_block_table_page_size(page, size: int) -> bool:
    """Switch Block Table pagination page size."""
    # Block Table pagination is the last el-pagination on the edit page
    pgs = page.locator(".el-pagination").all()
    if not pgs:
        return False
    pg = pgs[-1]
    sel = pg.locator(".el-select").first
    if sel.count() == 0:
        return False
    sel.click()
    page.wait_for_timeout(300)
    all_items = page.locator(".el-select-dropdown__item").all()
    for item in all_items:
        try:
            if item.is_visible() and str(size) in item.inner_text():
                item.click()
                page.wait_for_timeout(600)
                return True
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    return False


def test_TestCase_AcuHMI_008_01_case12_12(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 进入创建页，填写 Device 区
    _nav_to_templates(page)
    page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    ts = str(int(time.time()))[-6:]
    template_name = f"BlkPg_{ts}"

    page.locator(".el-form-item").filter(has_text="Template Name").first.locator("input").fill(template_name)
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Version").first.locator("input").fill("v1.00")
    page.wait_for_timeout(200)

    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Typical Model").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")

    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Wiring Configuration").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)

    # Step 2: 批量添加 85 个 Block
    _add_85_blocks(page)

    # Step 3: Create Template
    page.get_by_role("button", name="Create Template").click()
    page.wait_for_timeout(2000)
    assert page.locator(".el-message--error").count() == 0, \
        f"含 {TARGET_BLOCKS} 个 Block 的模板创建应成功，但出现了错误"

    # Step 4: 进入 Template List，找到模板行，点击黄色编辑按钮
    _nav_to_templates(page)
    tl = page.locator(".el-menu-item").filter(has_text="Template List")
    tl.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Customized 表格使用最后一个 .el-pagination；切到 80条/页 再逐页搜索
    pgs = page.locator(".el-pagination").all()
    last_pg = pgs[-1] if pgs else None
    if last_pg:
        sel0 = last_pg.locator(".el-select").first
        if sel0.count() > 0:
            sel0.click()
            page.wait_for_timeout(300)
            for opt in page.locator(".el-select-dropdown__item").all():
                try:
                    if opt.is_visible() and "80" in opt.inner_text():
                        opt.click()
                        page.wait_for_timeout(600)
                        break
                except Exception:
                    pass
            else:
                page.keyboard.press("Escape")
                page.wait_for_timeout(200)

    # 逐页搜索（最多翻 20 页）
    edit_clicked = False
    for _pn in range(20):
        # 在最后一个 tbody（Customized 表格）中搜索
        last_tbody = page.locator("tbody").last
        rows = last_tbody.locator("tr").all()
        for r in rows:
            try:
                if template_name in r.inner_text():
                    r.locator(".el-button--warning").first.click()
                    edit_clicked = True
                    break
            except Exception:
                pass
        if edit_clicked:
            break
        # 点击最后一个 pagination 的下一页
        if last_pg is None:
            break
        nb = last_pg.locator(".btn-next").first
        if nb.count() == 0:
            break
        cls = nb.get_attribute("class") or ""
        disa = nb.get_attribute("disabled")
        if disa is not None or "is-disabled" in cls:
            break
        nb.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

    assert edit_clicked, f"Template List 中未找到模板 '{template_name}'"
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Step 5: Block Table 分页切换验证
    for size in [10, 20, 40, 80]:
        switched = _switch_block_table_page_size(page, size)
        assert switched, f"未能切换 Block Table 到 {size} 条/页"

        # Block Table 是编辑页中的 tbody（最后一个 tbody）
        block_rows = page.locator("tbody").last.locator("tr").count()
        assert block_rows <= size, \
            f"Block Table 每页 {size} 条时，实际显示 {block_rows} 行，超出限制"
        assert block_rows > 0, \
            f"Block Table 每页 {size} 条时，应有数据显示"
