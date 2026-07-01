"""
FTS编号: FTS_WEB2_AWS_001_001
用例标题: AWS IoT 配置开启与关闭
用例级别: LV0

预置条件: 网关已上电，浏览器访问网关 Web UI 并以 admin 登录

测试步骤:
  1. 导航到 Protocols > AWS IoT 页面
  2. 验证默认状态为 Disable，配置字段隐藏
  3. 点击 Enable，验证配置字段展开可编辑
  4. 点击 Disable，验证配置字段再次隐藏

预期结果:
  - 默认 Disable，所有配置输入框不可见
  - Enable 后，Client ID / URL / Topic / Interval / 证书等字段全部可输入
  - Disable 后，字段再次隐藏
"""

import pytest


class TestCase_AcuHMI_1_7_AWS_001_001:

    @pytest.mark.aws_iot
    def test_enable_disable_toggle(self, aws_page):
        # 1. 默认 Disable，配置字段不可见
        assert not aws_page.is_enabled(), "默认状态应为 Disable"
        inputs = aws_page.page.locator(aws_page.CLIENT_ID_INPUT).all()
        assert not inputs or not inputs[0].is_visible(), "Disable 状态下 Client ID 不应显示"

        # 2. Enable → 字段展开可编辑
        aws_page.ensure_enabled()
        assert aws_page.is_enabled(), "点击 Enable 后应处于 Enable 状态"
        aws_page.page.wait_for_selector(aws_page.CLIENT_ID_INPUT, state="visible", timeout=10000)
        assert aws_page.page.locator(aws_page.CLIENT_ID_INPUT).is_visible()
        assert aws_page.page.locator(aws_page.URL_INPUT).is_visible()
        assert aws_page.page.locator(aws_page.TOPIC_INPUT).is_visible()

        # 3. Disable → 字段再次隐藏
        aws_page.disable()
        assert not aws_page.is_enabled(), "Disable 后应处于 Disable 状态"
        inputs = aws_page.page.locator(aws_page.CLIENT_ID_INPUT).all()
        assert not inputs or not inputs[0].is_visible(), "Disable 后 Client ID 不应显示"
