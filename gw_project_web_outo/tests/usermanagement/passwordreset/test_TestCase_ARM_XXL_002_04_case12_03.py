import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_ARM_XXL_002_04_case12_03
# 用例标题：调整系统时间+1，使用生成的临时密码登录，登录失败
# 跳过原因：需要修改系统时间+外部工具生成临时密码，属于时间依赖型且需外部工具
# 澄清汇总 #14：不实现自动化
@pytest.mark.skip(reason="需要修改系统时间并依赖外部工具生成临时密码，不纳入自动化；澄清汇总 #14")
def test_TestCase_ARM_XXL_002_04_case12_03(login_page: LoginPage):
    pass
