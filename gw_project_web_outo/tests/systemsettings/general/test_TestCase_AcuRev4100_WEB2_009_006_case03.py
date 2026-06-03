import pytest
from pages.login_page import LoginPage

# 用例编号：TestCase_AcuRev4100_WEB2_009_006_case03
# 用例标题：Remote Access功能：点击Deregister取消注册，Registration Status显示Not Registration
# 预置条件：管理权限登录AcuHMI，设备已注册 Remote Access
# 测试步骤：
#   1. 进入 System Settings → Remote Access
#   2. 启用 Remote Access Enable（若未启用），点击 Save
#   3. 点击 Manual Register 进行注册，等待页面完全显示
#   4. 检查 Status 是否为 online；若不是，点击 Refresh Status 直到 online
#   5. 复制 Remote Access URL，新页面打开验证可进入项目登录页
#   6. 点击 Deregister 取消注册
#   7. 验证 Registration Status 显示 Not Registration / Not Registered


def _nav_to_remote_access(page):
    """Navigate to System Settings → Remote Access."""
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(600)
    ra = page.locator(".el-menu-item").filter(has_text="Remote Access").first
    ra.click(force=True)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)


def _ensure_remote_access_enabled(page):
    """Enable Remote Access if currently disabled, then save."""
    enable_item = page.locator(".el-form-item").filter(has_text="Remote Access Enable").first
    assert enable_item.count() > 0, "未找到 Remote Access Enable 字段"

    enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (enable_radio.get_attribute("class") or ""):
        enable_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(500)
        page.get_by_role("button", name="Save").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
        _nav_to_remote_access(page)


def _manual_register(page):
    """Click Manual Register and handle the registration dialog/process."""
    reg_btn = page.get_by_role("button", name="Manual Register")
    if reg_btn.count() == 0:
        reg_btn = page.locator("button").filter(has_text="Manual Register")
    if reg_btn.count() == 0:
        reg_btn = page.locator("button").filter(has_text="Register")

    assert reg_btn.count() > 0, "未找到 Manual Register / Register 按钮"
    reg_btn.first.click()
    page.wait_for_timeout(1000)

    # 处理可能弹出的确认对话框
    for btn_name in ["Yes, continue", "Yes,continue", "Yes", "Confirm", "确认", "OK"]:
        btn = page.get_by_role("button", name=btn_name)
        if btn.count() > 0:
            try:
                if btn.first.is_visible():
                    btn.first.click()
                    page.wait_for_timeout(800)
                    break
            except Exception:
                pass

    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)


def _wait_for_online_status(page, max_retries=10, interval_ms=3000):
    """Poll Refresh Status until Status shows 'online', up to max_retries times."""
    for attempt in range(max_retries):
        # 读取 Status 字段当前值
        status_fi = page.locator(".el-form-item").filter(has_text="Status").first
        if status_fi.count() == 0:
            # 尝试其他文本匹配
            status_fi = page.locator(".el-form-item").filter(has_text="status").first

        status_text = ""
        if status_fi.count() > 0:
            status_text = status_fi.first.inner_text().lower()

        if "online" in status_text:
            return True

        # 点击 Refresh Status 按钮
        refresh_btn = page.get_by_role("button", name="Refresh Status")
        if refresh_btn.count() == 0:
            refresh_btn = page.locator("button").filter(has_text="Refresh Status")
        if refresh_btn.count() == 0:
            refresh_btn = page.locator("button").filter(has_text="Refresh")

        if refresh_btn.count() > 0:
            refresh_btn.first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(interval_ms)
        else:
            page.wait_for_timeout(interval_ms)

    # 最后一次检查
    status_fi = page.locator(".el-form-item").filter(has_text="Status").first
    if status_fi.count() > 0:
        return "online" in status_fi.first.inner_text().lower()
    return False


def test_TestCase_AcuRev4100_WEB2_009_006_case03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 进入 Remote Access 页面
    _nav_to_remote_access(page)

    # Step 2: 确保 Remote Access Enable 已启用
    _ensure_remote_access_enabled(page)

    # Step 3: 点击 Manual Register 注册设备
    _manual_register(page)

    # 重新导航刷新页面状态
    _nav_to_remote_access(page)

    # Step 4: 等待 Status 变为 online（最多轮询 10 次，每次间隔 3s）
    is_online = _wait_for_online_status(page, max_retries=10, interval_ms=3000)
    assert is_online, "Manual Register 后 Status 未能变为 online（已轮询多次 Refresh Status）"

    # Step 5: 获取 Remote Access URL 并在新页面中验证可打开登录页
    url_fi = page.locator(".el-form-item").filter(has_text="Remote Access URL").first
    if url_fi.count() == 0:
        url_fi = page.locator(".el-form-item").filter(has_text="URL").first
    assert url_fi.count() > 0, "未找到 Remote Access URL 字段"

    # 读取 URL 值（input 或纯文本）
    url_input = url_fi.locator("input").first
    if url_input.count() > 0:
        remote_url = url_input.input_value().strip()
    else:
        # 过滤掉标签文本，只取值部分
        full_text = url_fi.first.inner_text().strip()
        lines = [l.strip() for l in full_text.splitlines() if l.strip()]
        remote_url = lines[-1] if lines else ""

    assert remote_url.startswith("http"), \
        f"Remote Access URL 格式不正确：'{remote_url}'"

    # 新标签页打开 Remote Access URL，验证可进入登录页
    new_page = page.context.new_page()
    new_page.goto(remote_url, timeout=30000)
    new_page.wait_for_load_state("networkidle")
    new_page.wait_for_timeout(2000)

    page_text = new_page.locator("body").inner_text().lower()
    has_login_elements = (
        new_page.locator("input[type='password']").count() > 0
        or new_page.locator("input[type='text']").count() > 0
        or "login" in page_text
        or "username" in page_text
        or "password" in page_text
        or "sign in" in page_text
        or "用户名" in page_text
        or "密码" in page_text
        or "登录" in page_text
    )
    assert has_login_elements, \
        f"Remote Access URL 打开后未显示登录页面，页面内容：{page_text[:300]}"
    new_page.close()

    # Step 6: 回到原页面，点击 Deregister 取消注册
    _nav_to_remote_access(page)

    deregister_btn = page.get_by_role("button", name="Deregister")
    if deregister_btn.count() == 0:
        deregister_btn = page.locator("button").filter(has_text="Deregister")

    assert deregister_btn.count() > 0, \
        "未找到 Deregister 按钮，注册可能未成功"

    deregister_btn.first.click()
    page.wait_for_timeout(1000)

    # 处理 Deregister 确认对话框
    for btn_name in ["Yes, continue", "Yes,continue", "Yes", "Confirm", "确认"]:
        btn = page.get_by_role("button", name=btn_name)
        if btn.count() > 0:
            try:
                if btn.first.is_visible():
                    btn.first.click()
                    page.wait_for_timeout(800)
                    break
            except Exception:
                pass

    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    # Step 7: 验证 Registration Status 显示 Not Registration / Not Registered
    status_fi = page.locator(".el-form-item").filter(has_text="Registration Status")
    if status_fi.count() > 0:
        status_text = status_fi.first.inner_text()
        assert "Not Registration" in status_text or "Not Registered" in status_text, \
            f"Deregister 后 Registration Status 应显示 Not Registration，实际：{status_text}"
    else:
        not_registered = (
            page.get_by_text("Not Registration", exact=False).count() > 0
            or page.get_by_text("Not Registered", exact=False).count() > 0
        )
        assert not_registered, \
            "Deregister 后页面上应显示 Not Registration 状态，但未找到"
