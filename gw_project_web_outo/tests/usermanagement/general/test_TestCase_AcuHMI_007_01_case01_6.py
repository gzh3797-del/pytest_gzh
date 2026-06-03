import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_007_01_case01_6
# 用例标题：会话超时为60，保存配置并验证第60min用户自动退出
# 说明：需等待 60 分钟验证超时行为，属于时间依赖用例，不纳入自动化。
@pytest.mark.skip(reason="时间依赖用例：需等待60分钟验证超时，不纳入自动化（见 usermanagement_struct.md 澄清第16条）")
def test_TestCase_AcuHMI_007_01_case01_6(login_page: LoginPage):
    pass
