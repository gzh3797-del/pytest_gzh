from playwright.sync_api import Page, expect

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_001_003
# 用例标题：Alarm Acknowledgement Enable 开关默认状态展示
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 进入 Alarm 配置相关页面（System Settings → Alarm Notification）
#   2. 查看 Alarm Acknowledgement Enable 开关
# 预期结果：
#   2. 开关存在且默认状态符合需求设计（默认 Enable），位置展示正确
# 说明：出厂默认值无法在不恢复出厂的前提下严格验证，此处断言开关存在、
#      选项完整（Enable/Disable）且恰有一项选中；当前值应为需求默认 Enable。


def test_TestCase_ACUHMI17_BZ_001_003(app_page: Page):
    page = app_page

    # ── Step 1: 进入 System Settings → Alarm Notification ──
    ha.goto_alarm_notification(page)

    # ── Step 2: 开关存在、位置正确、默认状态 Enable ──
    form_item = page.locator(".el-form-item").filter(
        has_text="Alarm Acknowledgement Enable").first
    expect(form_item).to_be_visible()

    radios = form_item.locator(".el-radio")
    labels = [radios.nth(i).inner_text().strip() for i in range(radios.count())]
    assert labels == ["Enable", "Disable"], \
        f"开关选项应为 Enable/Disable，实际: {labels}"

    checked = [i for i in range(radios.count())
               if "is-checked" in (radios.nth(i).get_attribute("class") or "")]
    assert len(checked) == 1, f"开关应恰有一项选中，实际选中索引: {checked}"

    # 需求默认值为 Enable（若此前被人为改为 Disable，此断言会暴露环境偏离）
    assert checked == [0], (
        "Alarm Acknowledgement Enable 默认应为 Enable，"
        f"当前选中的是 {labels[checked[0]]!r}"
    )
