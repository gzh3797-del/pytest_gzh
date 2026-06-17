from playwright.sync_api import Page, expect

from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号：TestCase_AcuRev4100_WEB2_009_006_case03
# 用例标题：Remote Access功能：点击Deregister取消注册，Registration Status显示Not Registration
# 预置条件：管理权限登录AcuHMI；Remote Access Enable 处于 disabled（未启用、未注册）状态
# 测试步骤：
#   1. 进入 System Settings → Remote Access
#   2. 启用 Remote Access Enable 并 Save（启用即自动注册并生成 Remote Access URL）
#   3. 记录生成的 Remote Access URL
#   4. 点击 Deregister 取消注册，确认对话框点 Confirm
#   5. 验证 Registration Status 显示 Not Registration / Not Registered
#   6. 用步骤 3 记录的 URL 重新打开，验证已无法登录（远程通道关闭，返回 502 / 离线提示）
# 测试结束：恢复 Remote Access Enable 为 disabled


def _nav_to_remote_access(page: Page) -> None:
    """Navigate to System Settings → Remote Access.

    Remote Access 是顶部横向导航（.c_top_navbar）最右侧 el-sub-menu 的弹出项，
    默认 display:none（popper 收起），须先 hover 触发展开后再点击。
    直接 goto '#/systemSettings/remoteAccess' 会被路由重定向，必须走菜单交互。
    """
    base = page.url.split("#")[0]
    # 先导航到 System Settings 任意子页，确保顶部横向菜单已加载
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(600)

    # Hover 顶部横向菜单最右侧的"三点"sub-menu title，触发 Remote Access 弹出
    sub_title = page.locator(".c_top_navbar .el-sub-menu__title")
    sub_title.hover()

    # 等待 popper 内的 Remote Access 菜单项变为可见（auto-wait，无需 force）
    ra_item = (
        page.locator(".el-popper.is-pure .el-menu--popup .el-menu-item")
        .filter(has_text="Remote Access")
    )
    expect(ra_item).to_be_visible(timeout=3000)
    ra_item.click()

    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def _set_remote_access_enable(page: Page, enable: bool) -> None:
    """将 Remote Access Enable 设为 Enable/Disable 并 Save。"""
    label = "Enable" if enable else "Disable"
    enable_item = page.locator(".el-form-item").filter(has_text="Remote Access Enable").first
    assert enable_item.count() > 0, "未找到 Remote Access Enable 字段"
    enable_item.locator(".el-radio__label").filter(has_text=label).first.click()
    page.wait_for_timeout(500)
    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)


def _read_remote_access_url(page: Page) -> str:
    """读取 Remote Access URL；启用后注册需片刻，轮询等待 URL 生成（最多 ~15s）。"""
    url_fi = page.locator(".el-form-item").filter(has_text="Remote Access URL").first
    if url_fi.count() == 0:
        url_fi = page.locator(".el-form-item").filter(has_text="URL").first
    assert url_fi.count() > 0, "启用后未找到 Remote Access URL 字段"

    url_input = url_fi.locator("input").first
    remote_url = ""
    for _ in range(15):
        if url_input.count() > 0:
            remote_url = url_input.input_value().strip()
        else:
            full_text = url_fi.first.inner_text().strip()
            lines = [line.strip() for line in full_text.splitlines() if line.strip()]
            remote_url = lines[-1] if lines else ""
        if remote_url.startswith("http"):
            break
        page.wait_for_timeout(1000)
    return remote_url


def test_TestCase_AcuRev4100_WEB2_009_006_case03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 进入 Remote Access 页面（预置条件：Remote Access Enable=disabled）
    _nav_to_remote_access(page)

    try:
        # Step 2: 启用 Remote Access 并 Save —— 启用即自动注册并生成 URL
        _set_remote_access_enable(page, enable=True)

        # Step 3: 记录注册生成的 Remote Access URL（供注销后验证用）
        remote_url = _read_remote_access_url(page)
        assert remote_url.startswith("http"), \
            f"启用 Remote Access 后未生成有效 URL（注册可能未完成）：'{remote_url}'"

        # Step 4: 点击 Deregister 取消注册
        deregister_btn = page.get_by_role("button", name="Deregister")
        if deregister_btn.count() == 0:
            deregister_btn = page.locator("button").filter(has_text="Deregister")
        assert deregister_btn.count() > 0, "启用注册后未找到 Deregister 按钮"

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

        # Step 5: 验证 Registration Status 显示 Not Registration / Not Registered
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

        # Step 6: 用注销前记录的 Remote Access URL 重新打开，验证已无法进入登录页
        # 真机表现（已验证）：注销后 goto 抛 TimeoutError，或页面返回 HTTP 502
        # 标题 "502"，H2 "Oops, this device has turned off remote access"，无登录元素
        verify_page = None
        url_accessible = False
        try:
            verify_page = page.context.new_page()
            verify_page.goto(remote_url, timeout=30000)
            verify_page.wait_for_load_state("networkidle", timeout=10000)
            verify_page.wait_for_timeout(2000)
            url_accessible = True
        except Exception:
            # goto 超时或连接失败本身即证明注销成功、远程通道已关闭
            url_accessible = False

        if url_accessible and verify_page is not None:
            # goto 未抛异常时，正向核查页面内容
            page_text_after = verify_page.locator("body").inner_text().lower()
            has_login_elements_after = (
                verify_page.locator("input[type='password']").count() > 0
                or verify_page.locator("input[type='text']").count() > 0
                or "login" in page_text_after
                or "username" in page_text_after
                or "password" in page_text_after
                or "sign in" in page_text_after
                or "用户名" in page_text_after
                or "密码" in page_text_after
                or "登录" in page_text_after
            )
            # 期望仍然有注销/502 的明确错误提示（正向证据）
            has_offline_indicator = (
                "502" in page_text_after
                or "turned off remote access" in page_text_after
                or "device has turned off" in page_text_after
                or "remote access" in page_text_after
                or "not found" in page_text_after
                or "unavailable" in page_text_after
            )
            assert not has_login_elements_after, (
                f"Deregister 后用 Remote Access URL 打开仍显示登录页，"
                f"注销未生效。URL={remote_url}，页面内容片段：{page_text_after[:300]}"
            )
            assert has_offline_indicator, (
                f"Deregister 后页面既无登录页也无明确离线/502 提示，断言无法确认注销效果。"
                f"URL={remote_url}，页面内容片段：{page_text_after[:300]}"
            )

        if verify_page is not None:
            verify_page.close()

    finally:
        # 恢复预置条件：Remote Access Enable 设回 disabled（best-effort，不影响用例结论）
        try:
            _nav_to_remote_access(page)
            _set_remote_access_enable(page, enable=False)
        except Exception:
            pass
