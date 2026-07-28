import pytest

from playwright.sync_api import Page


# 用例编号：TestCase_ACUHMI17_BZ_002_004
# 用例标题：物理 Alarm Reset 按钮已从设备 UI 删除
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 进入物理设备相关操作界面
#   2. 查找 Alarm Reset 按钮
# 预期结果：
#   2. Alarm Reset 按钮不可见，UI 中已删除，无法触发该操作


@pytest.mark.skip(reason="验证对象是设备本体物理面板 UI（非 Web 页面），"
                         "需人工现场目视确认 Alarm Reset 按钮已删除")
def test_TestCase_ACUHMI17_BZ_002_004(app_page: Page):
    pass
