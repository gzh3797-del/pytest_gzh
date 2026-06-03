import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_ARM_XXL_002_04_case12_06
# 用例标题：密码配置一天内不可修改，使用临时密码登录系统，查看是否可以修改密码
# 跳过原因：需要外部工具生成临时密码，且依赖密码最短期限配置，不纳入自动化
# 澄清汇总 #14：不实现自动化
@pytest.mark.skip(reason="依赖外部工具生成临时密码和密码最短期限策略，不纳入自动化；澄清汇总 #14")
def test_TestCase_ARM_XXL_002_04_case12_06(login_page: LoginPage):
    pass
