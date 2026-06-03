import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_ARM_XXL_002_04_case11_02
# 用例标题：恢复出厂后验证默认密码功能
# 跳过原因：需要执行恢复出厂（Factory Reset），会破坏测试环境，不纳入自动化
@pytest.mark.skip(reason="需要恢复出厂设置（Factory Reset），会破坏测试环境，不纳入自动化")
def test_TestCase_ARM_XXL_002_04_case11_02(login_page: LoginPage):
    pass
