import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case07_7
# 用例标题：Create from Typical Energy Meter创建模板成功，并可在Physical Devices看到设备断线，统计数据
# 预置条件：
#   1. 服务启动正常，账号登录成功
#   2. 需要真实物理设备连接（断线状态可见）
# 测试步骤：
#   1. 进入 Template List，在 Official 区找到 Typical Energy Meter 模板
#   2. 点击 "Create from Typical Energy Meter" 相关操作
#   3. 完成创建流程
#   4. 进入 Physical Devices 验证设备断线状态及统计数据
# 预期结果：
#   创建成功，Physical Devices 中可见设备处于断线状态，且有统计数据


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


@pytest.mark.xfail(strict=False, reason="依赖真实物理设备及 Physical Devices 功能，当前环境无法验证")
def test_TestCase_AcuHMI_008_01_case07_7(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 进入 Template List
    _nav_to_templates(page)
    tl = page.locator(".el-menu-item").filter(has_text="Template List")
    tl.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Step 2: 在 Official 区找到 Typical Energy Meter 类型模板行，点击绿色按钮（Create from...）
    # Official 表格第一个 tbody，Customized 表格第二个 tbody
    official_tbody = page.locator("tbody").first
    first_row = official_tbody.locator("tr").first
    assert first_row.count() > 0, "Official 模板列表为空，无法执行 Create from Typical Energy Meter"

    # 尝试找到 "Create from" 按钮（如绿色图标按钮或带文字的按钮）
    create_btn = page.get_by_role("button", name="Create from Typical Energy Meter")
    if create_btn.count() == 0:
        # fallback: 点击 Official 行中绿色按钮
        create_btn = first_row.locator(".el-button--success").first

    assert create_btn.count() > 0, "未找到 Create from Typical Energy Meter 入口，功能可能未在此版本实现"
    create_btn.first.click()
    page.wait_for_timeout(1000)

    # Step 3: 验证创建向导或跳转页面出现
    assert (
        page.locator(".el-dialog").count() > 0
        or page.get_by_text("Typical Energy Meter", exact=False).count() > 0
        or page.get_by_role("button", name="Next").count() > 0
    ), "点击后应显示创建向导或跳转相关页面"

    # Step 4: Physical Devices 断线验证（需真实设备，此处仅做框架占位）
    # 实际验证需跳转至 Physical Devices 页面，检查设备 Offline 状态及统计数据
    # 当前无真实设备，本步骤将触发 xfail
    raise NotImplementedError("Physical Devices 验证需真实设备连接，当前环境不支持")
