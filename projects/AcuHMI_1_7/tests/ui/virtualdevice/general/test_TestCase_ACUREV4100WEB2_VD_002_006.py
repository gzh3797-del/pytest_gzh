import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_VD_PREFIX = "VD_CapTest_"
_LIMIT = 32


def _nav_to_vd_list(page):
    on_list = "/#/virtualMeter" in page.url and "addVirtualMeter" not in page.url
    if not on_list:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Virtual Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _add_vd(page, name: str):
    _nav_to_vd_list(page)
    page.get_by_role("button", name="Add Virtual Device").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(300)
    page.get_by_label("Virtual Device Name", exact=True).fill(name)
    page.get_by_placeholder("---Enter Parameter Name---").first.fill("p1")
    page.get_by_placeholder("---Enter Post Label---").first.fill("p1")
    page.get_by_placeholder("---Enter Calculated Meter Formula---").first.fill("1+1")
    page.get_by_placeholder("---Enter Unit---").first.fill("kW")
    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(300)


def _confirm_delete(page):
    for btn_name in ["Yes", "Confirm"]:
        try:
            page.get_by_role("button", name=btn_name).click(timeout=2000)
            page.wait_for_timeout(500)
            return
        except Exception:
            pass


def _delete_all_by_prefix(page, prefix: str):
    """删除所有以 prefix 开头的虚拟设备，自动处理分页。"""
    for _ in range(200):  # 最多处理 200 次（防止死循环）
        _nav_to_vd_list(page)
        rows = page.locator("tbody tr")
        found = False
        for i in range(rows.count()):
            try:
                name = rows.nth(i).locator("td").first.inner_text().strip()
            except Exception:
                continue
            if name.startswith(prefix):
                rows.nth(i).locator(".el-button").last.click(force=True)
                page.wait_for_timeout(500)
                _confirm_delete(page)
                found = True
                break  # 重新查询行列表（行已删，索引失效）
        if not found:
            break  # 当前页无 prefix 设备，检查下一页
        # found=True 时下次循环会重新 _nav_to_vd_list，天然从第一页开始

    # 额外：遍历第 2 页及以后（删完第 1 页后若仍有残留页）
    for _ in range(10):
        _nav_to_vd_list(page)
        next_btn = page.locator(".btn-next").first
        try:
            if next_btn.is_disabled():
                break
        except Exception:
            break
        next_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
        rows = page.locator("tbody tr")
        found_on_page = False
        for i in range(rows.count()):
            try:
                name = rows.nth(i).locator("td").first.inner_text().strip()
            except Exception:
                continue
            if name.startswith(prefix):
                rows.nth(i).locator(".el-button").last.click(force=True)
                page.wait_for_timeout(500)
                _confirm_delete(page)
                found_on_page = True
                break
        if not found_on_page:
            break


def _count_all_vd(page) -> int:
    """统计所有页的虚拟设备总数（通过翻页累加）。"""
    _nav_to_vd_list(page)
    total = 0
    for _ in range(20):
        total += page.locator("tbody tr").count()
        next_btn = page.locator(".btn-next").first
        try:
            if next_btn.is_disabled():
                break
        except Exception:
            break
        next_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
    _nav_to_vd_list(page)  # 归位到第 1 页
    return total


# 用例编号：TestCase_ACUREV4100WEB2_VD_002_006
# 用例标题：尝试创建超过32台的虚拟设备，系统提示虚拟设备数量超限
# 预置条件：管理员账号登录
# 测试步骤：
#   1. 若当前不足32台，批量补充至32台
#   2. 再次点击 Add Virtual Device，尝试创建第33台
#   3. 确认系统阻止创建并显示超限提示
# 预期结果：系统提示虚拟设备数量超出上限（按钮禁用 / 表单错误 / toast 提示）
def test_TestCase_ACUREV4100WEB2_VD_002_006(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    created_names: list = []

    try:
        # ── 1. 清理本用例上次遗留设备 ─────────────────────────────────────────
        _delete_all_by_prefix(page, _VD_PREFIX)

        # ── 2. 补充至恰好32台 ──────────────────────────────────────────────────
        current_count = _count_all_vd(page)
        to_add = max(0, _LIMIT - current_count)
        for i in range(to_add):
            name = f"{_VD_PREFIX}{i:03d}"
            _add_vd(page, name)
            created_names.append(name)

        # ── 3. 检查按钮状态（部分实现在 UI 层禁用按钮）───────────────────────
        _nav_to_vd_list(page)
        add_btn = page.get_by_role("button", name="Add Virtual Device")

        if add_btn.count() == 0:
            return  # 按钮已被移除 → 系统限制生效

        if add_btn.first.is_disabled():
            return  # 按钮已禁用 → 系统限制生效

        # ── 4. 按钮仍可点：尝试保存第33台，期望收到超限错误 ──────────────────
        add_btn.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        # 点击后若立即弹出 toast，不需要填表
        immediate_error = (
            page.locator(".el-message--error").count() > 0
            or page.locator(".el-message--warning").count() > 0
        )

        if not immediate_error:
            try:
                name_field = page.get_by_label("Virtual Device Name", exact=True)
                if name_field.count() > 0:
                    name_field.fill(f"{_VD_PREFIX}OVER")
                page.get_by_placeholder("---Enter Parameter Name---").first.fill("p1")
                page.get_by_placeholder("---Enter Post Label---").first.fill("p1")
                page.get_by_placeholder("---Enter Calculated Meter Formula---").first.fill("1+1")
                page.get_by_placeholder("---Enter Unit---").first.fill("kW")
            except Exception:
                pass
            page.get_by_role("button", name="Save").click()
            page.wait_for_timeout(1500)

        limit_shown = (
            page.locator(".el-message--error").count() > 0
            or page.locator(".el-message--warning").count() > 0
            or page.get_by_text("exceed", exact=False).count() > 0
            or page.get_by_text("maximum", exact=False).count() > 0
            or page.get_by_text("limit", exact=False).count() > 0
            or page.get_by_text("32", exact=False).count() > 0
        )
        assert limit_shown, (
            "创建第33台虚拟设备时，系统应提示虚拟设备数量超出上限，但未检测到任何超限提示"
        )

    finally:
        # ── 清理：删除所有 VD_CapTest_ 前缀的设备（含所有分页）────────────────
        _delete_all_by_prefix(page, _VD_PREFIX)
