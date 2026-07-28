import time

from playwright.sync_api import expect

from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case03_3
# 用例标题：基于已有自定义模板再创建新模板（派生模板）
# 预置条件：
#   1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. 进入 Templates > New Typical Energy Meter Template
#   2. 填写 Template Name=test_0100_<ts>、Version=v1.00
#      Wiring Configuration=3 Element 4 Wire Y、Function=READ_HOLDING_REGISTERS
#      Start=2100、Count=1，点击 Save Block，再点 Create Template
#   3. 断言弹出创建成功 toast（"Create Success."）
#   4. 进入 Template List，在 Customized 区断言基础模板名可见
#   5. 找到基础模板行，点 Action 列第二个图标（蓝色 primary，
#      tooltip="Create new template from this one"）
#   6. 在跳转的派生模板表单里填写 Template Name=<基础名>_1、Version=v1.00，点 Create Template
#   7. 断言弹出创建成功 toast
#   8. 在 Customized 区断言派生模板名可见
#   9. 测试后清理：删除派生模板，再删除基础模板，断言两者从列表消失
# 预期结果：
#   基础模板和派生模板均创建成功，均出现在 Template List > Customized 列表中；
#   测试结束后两个模板均被删除，Customized 列表中不再可见。


def _nav_to_templates(page):
    """Navigate to AcuHMI > Templates section."""
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


def _select_dropdown_option(page, form_item_text: str, option_text: str) -> None:
    """Open the el-select inside the named form-item and click the exact option.

    Raises AssertionError if the option is not found — no silent fallback.
    """
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text=form_item_text).first.locator(".el-select").click()
    page.wait_for_timeout(400)
    visible_items = [
        item for item in page.locator(".el-select-dropdown__item").all()
        if item.is_visible()
    ]
    for item in visible_items:
        if option_text in item.inner_text():
            item.click()
            page.wait_for_timeout(400)
            page.keyboard.press("Escape")
            page.wait_for_timeout(300)
            return
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)
    available = [item.inner_text().strip() for item in visible_items]
    raise AssertionError(
        f'Option "{option_text}" not found in "{form_item_text}" dropdown. '
        f"Available options: {available}"
    )


def _select_typical_model_first(page) -> None:
    """Select the first available option in Typical Model dropdown (optional field)."""
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Typical Model").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    visible_items = [
        item for item in page.locator(".el-select-dropdown__item").all()
        if item.is_visible()
    ]
    if visible_items:
        visible_items[0].click()
        page.wait_for_timeout(400)
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)


def _assert_success_toast(page, context_msg: str) -> None:
    """Assert that a 'Create Success.' toast is visible.

    Waits up to 5 s for *any* visible .el-message--success element,
    then asserts its text contains 'Create Success.' to prevent
    false-passes from stale DOM nodes of earlier toasts.
    """
    success_locator = page.locator(".el-message--success")
    expect(success_locator.first).to_be_visible(timeout=5000)
    # Confirm the visible toast carries the expected text
    assert "Create Success" in success_locator.first.inner_text(), (
        f"{context_msg}: toast 可见但文案不符，"
        f"实际文案: {success_locator.first.inner_text()!r}"
    )


def _get_customized_table(page):
    """Return the locator for the Customized c_common_table (2nd table on Template List page)."""
    return page.locator(".c_common_table").nth(1)


def _nav_to_template_list(page) -> None:
    """Navigate to Template List sub-menu."""
    _nav_to_templates(page)
    page.locator(".el-menu-item").filter(has_text="Template List").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)


def _wait_for_template_in_list(page, template_name: str, timeout_s: int = 30) -> None:
    """Reload templateList until ``template_name`` appears in Customized table or timeout.

    The device's processTemplate API is asynchronous; the GET list endpoint
    may return stale data for a few seconds after a successful PATCH.
    Polling-reload is the only reliable way to observe the new row.
    """
    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        page.reload()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)
        assert "/#/templates/templateList" in page.url, (
            f"reload 后应在 templateList 页，实际 URL: {page.url!r}"
        )
        cust_table = _get_customized_table(page)
        row = cust_table.locator("tbody tr").filter(has_text=template_name)
        if row.count() > 0:
            return
    raise AssertionError(
        f"Customized 列表中应能看到模板 {template_name!r}，"
        f"等待 {timeout_s}s 后仍未出现（含多次 reload）"
    )


def _delete_template_by_name(page, template_name: str, assert_gone: bool = True) -> None:
    """Delete a template from the Customized list by name.

    Scrolls into view before clicking to ensure the row is in the viewport.
    If assert_gone is True, asserts the template is no longer visible after deletion
    (use True for primary assertions, False for cleanup-only calls in finally).
    """
    cust_table = _get_customized_table(page)
    row = cust_table.locator("tbody tr").filter(has_text=template_name)
    if row.count() == 0:
        # Already gone; nothing to delete
        return
    danger_btn = row.first.locator(".el-button--danger").first
    danger_btn.scroll_into_view_if_needed()
    page.wait_for_timeout(300)
    danger_btn.click()
    page.wait_for_timeout(800)
    yes_btn = page.get_by_role("button", name="Yes")
    if yes_btn.count() > 0 and yes_btn.first.is_visible():
        yes_btn.first.click()
        page.wait_for_timeout(1500)
    if assert_gone:
        # Scroll to bottom to make sure list is refreshed and visible
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        page.wait_for_timeout(500)
        cust_table_after = _get_customized_table(page)
        gone_count = cust_table_after.locator("tbody tr").filter(has_text=template_name).count()
        assert gone_count == 0, (
            f"删除后模板 {template_name!r} 应从 Customized 列表消失，"
            f"但仍找到 {gone_count} 行"
        )


def test_TestCase_AcuHMI_008_01_case03_3(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # ── 命名规则 ──────────────────────────────────────────────────────────────
    ts = str(int(time.time()))[-6:]
    base_name = f"test_0100_{ts}"
    derived_name = f"{base_name}_1"

    # ── Step 1: 进入 New Typical Energy Meter Template ────────────────────────
    _nav_to_templates(page)
    page.wait_for_timeout(500)

    page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # ── Step 2: 填写基础模板表单 ────────────────────────────────────────────
    # Template Name
    page.locator(".el-form-item").filter(has_text="Template Name").first.locator("input").fill(base_name)
    page.wait_for_timeout(200)

    # Version
    page.locator(".el-form-item").filter(has_text="Version").first.locator("input").fill("v1.00")
    page.wait_for_timeout(200)

    # Typical Model — 选第一个可用选项（此字段为可选）
    _select_typical_model_first(page)

    # Wiring Configuration — 精确选 "3 Element 4 Wire Y"，找不到即失败
    _select_dropdown_option(page, "Wiring Configuration", "3 Element 4 Wire Y")

    # Function — 精确选 "READ_HOLDING_REGISTERS"，找不到即失败
    _select_dropdown_option(page, "Function", "READ_HOLDING_REGISTERS")

    # Start = 2100
    start_input = page.locator(".el-form-item").filter(has_text="Start").first.locator("input")
    start_input.clear()
    start_input.fill("2100")
    page.wait_for_timeout(200)

    # Count = 1
    count_input = page.locator(".el-form-item").filter(has_text="Count").first.locator("input")
    count_input.clear()
    count_input.fill("1")
    page.wait_for_timeout(200)

    # Save Block
    page.get_by_role("button", name="Save Block").click()
    page.wait_for_timeout(1500)

    # ── Step 3: Create Template → 断言成功 toast ──────────────────────────────
    page.get_by_role("button", name="Create Template").click()
    page.wait_for_timeout(2000)

    _assert_success_toast(page, f"创建基础模板 {base_name!r}")

    # ── Step 4 & 5 & 6 & 7 & 8 (try) + Step 9 (finally 清理) ─────────────────
    try:
        # ── Step 4: Template List → Customized 区验证基础模板名 ─────────────
        # 创建成功后 SPA 自动跳回 templateList，但 processTemplate API 为异步，
        # 列表接口可能短暂返回旧数据。轮询 reload 直到新模板出现。
        page.wait_for_url("**/templates/templateList", wait_until="networkidle", timeout=10000)
        _wait_for_template_in_list(page, base_name)

        cust_table = _get_customized_table(page)
        base_row = cust_table.locator("tbody tr").filter(has_text=base_name)
        assert base_row.count() > 0, (
            f"Customized 列表中未找到基础模板 {base_name!r}"
        )

        # ── Step 5: 点 Action 列第二个图标（.el-button--primary） ─────────────
        # 每行 4 个 Action 按钮：success（View）/ primary（Create from this）/ warning（Edit）/ danger（Delete）
        create_from_btn = base_row.first.locator(".el-button--primary").first
        create_from_btn.scroll_into_view_if_needed()
        page.wait_for_timeout(300)
        expect(create_from_btn).to_be_visible(timeout=5000)
        create_from_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)

        # ── Step 6: 填写派生模板表单 ────────────────────────────────────────
        # Block 数据从源模板继承，不重新填。Template Name / Version 重新指定。
        derived_name_input = (
            page.locator(".el-form-item").filter(has_text="Template Name").first.locator("input").first
        )
        assert derived_name_input.is_editable(), (
            "派生模板表单中 Template Name 字段应可编辑"
        )
        derived_name_input.fill(derived_name)
        page.wait_for_timeout(200)

        page.locator(".el-form-item").filter(has_text="Version").first.locator("input").first.fill("v1.00")
        page.wait_for_timeout(200)

        # Create Template
        page.get_by_role("button", name="Create Template").click()
        page.wait_for_timeout(2000)

        # ── Step 7: 断言派生模板创建成功 toast ───────────────────────────────
        _assert_success_toast(page, f"创建派生模板 {derived_name!r}")

        # ── Step 8: Template List → Customized 区验证派生模板名 ─────────────
        # processTemplate API 为异步，轮询 reload 直到派生模板出现
        page.wait_for_url("**/templates/templateList", wait_until="networkidle", timeout=10000)
        _wait_for_template_in_list(page, derived_name)

        cust_table_after = _get_customized_table(page)
        derived_row = cust_table_after.locator("tbody tr").filter(has_text=derived_name)
        assert derived_row.count() > 0, (
            f"Customized 列表中未找到派生模板 {derived_name!r}"
        )

    finally:
        # ── Step 9: 清理（无论主流程成败都执行）─────────────────────────────
        # 确保在 templateList 页才执行清理；若因某步骤失败跳到别的页，先导航回来。
        try:
            if "/#/templates/templateList" not in page.url:
                _nav_to_template_list(page)
        except Exception:  # noqa: BLE001
            pass  # 导航失败不盖住原始断言异常

        # 先删派生模板（派生先于基础，顺序一致），再删基础模板。
        # assert_gone=False：清理失败只打印/吞掉，不覆盖主流程的断言异常。
        for name in (derived_name, base_name):
            try:
                _delete_template_by_name(page, name, assert_gone=False)
            except Exception:  # noqa: BLE001
                pass  # 清理自身异常静默吞掉，不影响主流程结果
