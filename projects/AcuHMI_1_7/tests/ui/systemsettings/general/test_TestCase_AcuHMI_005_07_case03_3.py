# 用例编号: TestCase_AcuHMI_005_07_case03_3
# 用例标题: 可添加白名单数量验证，最多20条，第21条添加失败
# 预置条件: 1、管理权限登录AcuHMI网页
# 测试步骤:
#   1. 添加20个白名单 → 成功
#   2. 添加第21个白名单 → 失败，提示白名单添加数量到上限
# 预期结果: 最多20条白名单，第21条添加失败或按钮禁用

import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

MAX_WHITELIST = 20
TEST_IP_PREFIX = "192.168.100."


def _nav_to_settings_tab(page, tab: str):
    if "/systemSettings" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=tab).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _add_allow_list(page, ip_addr: str, description: str):
    page.get_by_role("button", name="Add Allow List").click()
    page.wait_for_timeout(500)
    # Switch to No (single IP mode) — placeholder becomes "Enter IP Address"
    page.locator(".el-dialog").locator(".el-radio").filter(has_text="No").click()
    page.wait_for_timeout(300)
    page.get_by_placeholder("Enter IP Address").fill(ip_addr)
    page.get_by_placeholder("Enter Description").fill(description)
    page.get_by_role("button", name="Confirm").click()
    page.wait_for_timeout(500)
    # 校验这条确实加进去了（无分页，匹配该 IP 的行应恰好一条）；
    # 必须按 IP 单元格"精确文本"匹配，不能用整行 has_text 子串：单 IP 模式下起止 IP
    # 同值连排（如 192.168.100.1|192.168.100.1），拼接处会把 192.168.100.11 当子串命中
    # 192.168.100.1 那行，导致 .11 被误判为 2 条。
    expect(
        page.locator("tbody tr").filter(
            has=page.get_by_role("cell", name=ip_addr, exact=True)
        )
    ).to_have_count(1)


def _delete_test_entries(page):
    """Delete all test entries (192.168.100.x) using el-popconfirm Yes button."""
    for _ in range(MAX_WHITELIST + 2):
        rows = page.locator("tbody").get_by_role("row")
        found = False
        for j in range(rows.count()):
            try:
                row_text = rows.nth(j).inner_text()
                if TEST_IP_PREFIX in row_text:
                    rows.nth(j).locator(".el-button").last.click(force=True)
                    page.wait_for_timeout(500)
                    page.locator(".el-popconfirm__action .el-button--primary").click(timeout=2000)
                    page.wait_for_timeout(500)
                    found = True
                    break
            except Exception:
                pass
        if not found:
            break


@pytest.mark.slow
def test_TestCase_AcuHMI_005_07_case03_3(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Access Control")

    # 确保 IP Allow List 处于 Enable 状态（非默认则切换），Add Allow List 按钮才出现
    allow_enable = page.locator(".el-form-item").filter(
        has_text="IP Allow List Enable"
    ).locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (allow_enable.get_attribute("class") or ""):
        allow_enable.click()
        page.wait_for_timeout(500)

    # 先清理上次运行可能残留的测试 IP（192.168.100.x），保证用例幂等、不与残留撞重；
    # 只删本测试段，不动真实生产数据。
    _delete_test_entries(page)

    # Count pre-existing entries (can't delete unknown production data)
    existing_count = page.locator("tbody").get_by_role("row").count()
    entries_to_add = MAX_WHITELIST - existing_count

    if entries_to_add <= 0:
        # Already at max — verify button is disabled and skip
        assert page.get_by_role("button", name="Add Allow List").is_disabled(), \
            f"已有{existing_count}条白名单，添加按钮应禁用"
        return

    try:
        # 步骤1: 填满剩余配额
        for i in range(entries_to_add):
            _add_allow_list(page, f"{TEST_IP_PREFIX}{i + 1}", f"test_{i + 1}")

        # 步骤2: 尝试添加第(20+1)条，应失败或按钮禁用
        add_btn = page.get_by_role("button", name="Add Allow List")
        is_disabled = add_btn.is_disabled()
        if not is_disabled:
            add_btn.click()
            page.wait_for_timeout(500)
            limit_msg = (
                page.get_by_text("maximum", exact=False)
                .or_(page.get_by_text("limit", exact=False))
                .or_(page.get_by_text("上限", exact=False))
            )
            assert limit_msg.count() > 0 or is_disabled, \
                "白名单数量已达上限，应添加失败或按钮禁用"
        else:
            assert is_disabled, "白名单数量达到上限后，添加按钮应禁用"

    finally:
        _delete_test_entries(page)
        # Disable IP Allow List to restore state
        try:
            page.locator(".el-form-item").filter(has_text="IP Allow List Enable").locator(
                ".el-radio"
            ).filter(has_text="Disable").click()
            page.wait_for_timeout(300)
            page.get_by_role("button", name="Save").click()
            page.wait_for_timeout(500)
        except Exception:
            pass
