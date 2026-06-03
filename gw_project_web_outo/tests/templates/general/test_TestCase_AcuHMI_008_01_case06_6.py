import re
import time
import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case06_6
# 用例标题：用户自定义创建模板成功，10/20/40/80条/页切换查看
# 预置条件：
#   1. 服务启动正常，账号登录成功
#   2. Customized 区已有 85 个自定义模板（不足则自动补充）
# 测试步骤：
#   1. 进入 Template List，确认 Customized 区模板总数 >= 85
#   2. 依次切换每页显示数量：10 / 20 / 40 / 80
#   3. 每次切换后验证表格行数不超过所选数量
# 预期结果：
#   每页条数切换后，表格实际显示行数与所选条数一致（不超过）

TARGET_TOTAL = 85


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


def _click_visible_option(page, option_text: str):
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


def _create_template_batch(page, count: int):
    """Batch create `count` templates with minimum viable waits."""
    for i in range(count):
        # 每次都重新进入创建页，确保表单干净
        ntet = page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template")
        ntet.first.click()
        page.wait_for_load_state("networkidle")

        ts = str(int(time.time() * 1000))[-7:]
        name = f"Pg{i:03d}{ts}"

        # Device 区
        page.locator(".el-form-item").filter(has_text="Template Name").first.locator("input").fill(name)
        page.locator(".el-form-item").filter(has_text="Version").first.locator("input").fill("v1.00")

        page.keyboard.press("Escape")
        page.locator(".el-form-item").filter(has_text="Typical Model").first.locator(".el-select").click()
        page.wait_for_timeout(250)
        _click_visible_option(page, "")

        page.keyboard.press("Escape")
        page.locator(".el-form-item").filter(has_text="Wiring Configuration").first.locator(".el-select").click()
        page.wait_for_timeout(250)
        _click_visible_option(page, "")

        # Block 区（每次必须重新填写）
        page.keyboard.press("Escape")
        page.locator(".el-form-item").filter(has_text="Function").first.locator(".el-select").click()
        page.wait_for_timeout(250)
        _click_visible_option(page, "")

        page.keyboard.press("Escape")
        page.locator(".el-form-item").filter(has_text="Start").first.locator("input").fill("0001")
        page.locator(".el-form-item").filter(has_text="Count").first.locator("input").fill("1")

        page.get_by_role("button", name="Save Block").click()
        page.wait_for_timeout(800)

        page.get_by_role("button", name="Create Template").click()
        page.wait_for_timeout(1000)

        assert page.locator(".el-message--error").count() == 0, \
            f"第 {i+1} 个模板 '{name}' 创建失败"


def _get_customized_total(page) -> int:
    """Get total count of Customized templates from pagination info."""
    # El-Plus pagination shows total as ".el-pagination__total" text like "Total 85"
    totals = page.locator(".el-pagination__total").all()
    for el in reversed(totals):
        try:
            if el.is_visible():
                m = re.search(r'\d+', el.inner_text())
                if m:
                    return int(m.group())
        except Exception:
            pass

    # Fallback: switch Customized table to 80/page and count rows
    cust_pg = page.locator(".el-pagination").last
    sel = cust_pg.locator(".el-select").first
    if sel.count() > 0:
        sel.click()
        page.wait_for_timeout(300)
        all_items = page.locator(".el-select-dropdown__item").all()
        for item in all_items:
            try:
                if item.is_visible() and "80" in item.inner_text():
                    item.click()
                    page.wait_for_timeout(600)
                    break
            except Exception:
                pass

    tbodies = page.locator("tbody").all()
    if not tbodies:
        return 0
    rows = tbodies[-1].locator("tr").count()

    # Check if "next" button is still active (means there are more pages)
    next_btn = page.locator(".el-pagination .btn-next").last
    has_next = next_btn.count() > 0 and not next_btn.is_disabled()
    # If more pages exist, actual total > current rows; return conservative estimate
    return rows if not has_next else rows + 80  # rough upper bound


def _switch_customized_page_size(page, size: int) -> bool:
    """Switch Customized section page size. Returns True if successful."""
    cust_pg = page.locator(".el-pagination").last
    sel = cust_pg.locator(".el-select").first
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


def test_TestCase_AcuHMI_008_01_case06_6(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 进入 Template List，检查现有 Customized 模板数量
    _nav_to_templates(page)
    tl = page.locator(".el-menu-item").filter(has_text="Template List")
    tl.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    current = _get_customized_total(page)
    need = max(0, TARGET_TOTAL - current)

    # 如不足 85 个，批量补充创建
    if need > 0:
        _nav_to_templates(page)
        _create_template_batch(page, need)

        # 回到 Template List
        _nav_to_templates(page)
        tl = page.locator(".el-menu-item").filter(has_text="Template List")
        tl.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(800)

    # Step 2: 依次切换每页显示数量并验证
    for size in [10, 20, 40, 80]:
        switched = _switch_customized_page_size(page, size)
        assert switched, f"未能切换到 {size} 条/页"

        rows = page.locator("tbody").last.locator("tr").count()
        assert rows <= size, \
            f"每页 {size} 条时，Customized 表实际显示 {rows} 行，超出限制"
        assert rows > 0, f"每页 {size} 条时，Customized 表应有数据"
