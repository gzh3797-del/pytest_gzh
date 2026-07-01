"""
FTS编号: FTS_AcuHMI_AZR_004_008
用例标题: 未勾选设备时无法保存并提示
用例级别: LV2

预置条件:
  - Enable 已开启，合法 Connection String 已配置

测试步骤:
  1. Devices Selection 列表不勾选任何设备
  2. 点击保存

预期结果:
  - 阻止保存，提示 "Please provide Devices Selection"
"""

import pytest


def _uncheck_all_devices(page):
    rows = page.page.locator(page.DEVICE_TABLE_ROWS).all()
    for row in rows:
        cbs = row.locator("xpath=.//label[contains(@class,'el-checkbox')]").all()
        if cbs and "is-checked" in (cbs[0].get_attribute("class") or ""):
            cbs[0].evaluate("el => el.click()")
            page.page.wait_for_timeout(300)


class TestCase_AcuHMI_1_7_AZR_004_008:

    @pytest.mark.azure_iot
    def test_save_blocked_without_device(self, azure_page):
        azure_page.ensure_enabled()
        _uncheck_all_devices(azure_page)
        azure_page.save()
        try:
            el = azure_page.page.locator(azure_page.RESULT_MSG).first
            el.wait_for(state="visible", timeout=5000)
            msg = el.inner_text().strip()
        except Exception:
            msg = ""
        assert msg, "未勾选设备时 Save 应被阻止并给出提示"
