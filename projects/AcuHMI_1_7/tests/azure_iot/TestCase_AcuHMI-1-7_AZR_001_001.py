"""
FTS编号: FTS_AcuHMI_AZR_001_001
用例标题: Azure IoT 配置开启与关闭
用例级别: LV0

预置条件: 网关已上电，浏览器访问网关 Web UI 并以 admin 登录

测试步骤:
  1. 导航到 Protocols > Azure IoT 页面
  2. 验证默认状态为 Disable，配置字段隐藏
  3. 点击 Enable，验证配置字段展开可编辑
  4. 点击 Disable，验证配置字段再次隐藏
  5. 保存 Disable 状态

预期结果:
  - 默认 Disable，所有配置输入框不可见
  - Enable 后，Primary Connection String 等字段全部可输入
  - Disable 后，字段再次隐藏
"""

import pytest


class TestCase_AcuHMI_1_7_AZR_001_001:

    @pytest.mark.azure_iot
    def test_default_disable_state(self, azure_page):
        assert not azure_page.is_enabled(), "默认状态应为 Disable"
        inputs = azure_page.page.locator(azure_page.PRIMARY_CONN_STR_INPUT).all()
        assert not inputs or not inputs[0].is_visible(), "Disable 状态下 Primary Connection String 不应显示"

    @pytest.mark.azure_iot
    def test_enable_shows_fields(self, azure_page):
        azure_page.ensure_enabled()
        assert azure_page.is_enabled(), "点击 Enable 后应处于 Enable 状态"
        azure_page.page.wait_for_selector(azure_page.PRIMARY_CONN_STR_INPUT, state="visible", timeout=10000)
        assert azure_page.page.locator(azure_page.PRIMARY_CONN_STR_INPUT).is_visible()

    @pytest.mark.azure_iot
    def test_disable_hides_fields(self, azure_page):
        azure_page.ensure_enabled()
        azure_page.disable()
        assert not azure_page.is_enabled(), "Disable 后应处于 Disable 状态"
        inputs = azure_page.page.locator(azure_page.PRIMARY_CONN_STR_INPUT).all()
        assert not inputs or not inputs[0].is_visible(), "Disable 后 Primary Connection String 不应显示"
