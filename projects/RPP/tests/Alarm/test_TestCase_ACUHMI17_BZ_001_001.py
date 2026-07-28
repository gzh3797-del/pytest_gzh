import re

from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_001_001
# 用例标题：Alarm 导航入口改名及告警数量显示
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 进入 Devices 页面
#   2. 查看原 Alarm Logs 导航入口的名称
#   3. 触发一条告警，重新查看入口
# 预期结果：
#   2. 导航入口名称已改为 Alarm
#   3. Alarm 后括号内显示当前未确认告警数量，数字与实际告警数一致
# 实装差异说明（2026-07-17 实测）：未确认数量角标实装在 Alarm 页内
#   Unacknowledged Alarms 二级 tab 上，而非左导航 "Alarm" 的括号内；
#   本脚本按实装位置校验数量一致性，位置差异是否接受由需求方判定。

_LABEL = "at_bz001001"


def test_TestCase_ACUHMI17_BZ_001_001(app_page: Page):
    page = app_page
    try:
        # ── Step 1-2: Devices 左导航入口名称应为 Alarm（不再是 Alarm Logs）──
        ha.ensure_devices_module(page)
        navs = page.locator(".left-nav-item")
        texts = [navs.nth(i).inner_text().strip() for i in range(navs.count())]
        names = [re.sub(r"\s*\(\d+\)\s*$", "", t) for t in texts]
        assert "Alarm" in names, f"Devices 导航中应有 'Alarm' 入口，实际: {texts}"
        assert "Alarm Logs" not in names, \
            f"导航入口应已由 'Alarm Logs' 改名为 'Alarm'，实际: {texts}"

        # ── Step 3: 触发一条告警后，Alarm 区域显示未确认告警数量 ──
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_unack(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，未在轮询周期内出现未确认告警"
        actual_unack = ha.unack_total(page)

        # 实装位置：Unacknowledged Alarms 二级 tab 的数量角标
        tab_text = page.get_by_role(
            "menuitem", name="Unacknowledged Alarms").first.inner_text().strip()
        m = re.search(r"(\d+)\s*$", tab_text)
        assert m, (
            "Alarm 区域应显示未确认告警数量角标"
            f"（Unacknowledged Alarms tab 实际文本: {tab_text!r}）"
        )
        assert int(m.group(1)) == actual_unack, (
            f"角标数量 {m.group(1)} 与 Unacknowledged Alarms 实际条数 "
            f"{actual_unack} 不一致"
        )
    finally:
        # 清理：确认并删除本用例触发的告警规则，恢复环境
        try:
            ha.ack_alarm(page, _LABEL)
        except Exception:
            pass
        ha.cleanup_test_rules(page)
