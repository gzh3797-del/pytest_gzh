import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_007_01_case02_02
# 用例标题：添加用户，覆盖密码过期验证
# 跳过原因：需要修改系统时间（+1天/+2天），不适合自动化 CI 执行，时间依赖型用例
@pytest.mark.skip(reason="需修改系统时间进行验证（密码过期1天、2天后行为），属于时间依赖型用例，不纳入自动化")
def test_TestCase_AcuHMI_007_01_case02_02(login_page: LoginPage):
    pass
