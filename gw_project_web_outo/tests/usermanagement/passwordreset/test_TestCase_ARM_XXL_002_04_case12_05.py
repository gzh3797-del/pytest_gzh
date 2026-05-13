import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_ARM_XXL_002_04_case12_05
# 用例标题：用户锁定，使用临时密码登录系统，查看是否可以登录成功
# 跳过原因：需要外部工具生成临时密码，且依赖账户锁定状态，不纳入自动化
# 澄清汇总 #14：不实现自动化
@pytest.mark.skip(reason="依赖外部工具生成临时密码和账户锁定状态，不纳入自动化；澄清汇总 #14")
def test_TestCase_ARM_XXL_002_04_case12_05(login_page: LoginPage):
    pass
