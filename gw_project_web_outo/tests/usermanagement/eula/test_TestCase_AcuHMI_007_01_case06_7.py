import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_007_01_case06_7
# 用例标题：已接受EULA用户，恢复出厂后登录系统
# 跳过原因：需要执行恢复出厂（Configuration Management → Reset），会破坏测试环境，不纳入自动化
@pytest.mark.skip(reason="需要恢复出厂设置（Factory Reset），会破坏测试环境，不纳入自动化")
def test_TestCase_AcuHMI_007_01_case06_7(login_page: LoginPage):
    pass
